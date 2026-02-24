"""
LINE Monitor — 主監控腳本
- 背景靜默執行（不顯示視窗）
- 訊息只寫入內網共用資料夾，零外部連線
- 每人每天一份 txt，依對話間隔自動分段
"""

import os, sys, time, sqlite3, re, json, socket
from datetime import datetime
from pathlib import Path

# ══════════════════════════════════════════════
#  設定檔路徑（從共用資料夾讀取 config.json）
# ══════════════════════════════════════════════
SCRIPT_DIR  = Path(sys.executable).parent if getattr(sys, 'frozen', False) else Path(__file__).parent
CONFIG_FILE = SCRIPT_DIR / "config.json"

def load_config():
    default = {
        "shared_folder": "",
        "watch_list":    [],       # 空 = 全部
        "interval_sec":  60,       # 掃描間隔（秒）
        "gap_minutes":   5,        # 幾分鐘沒新訊息視為對話段落結束
        "auto_dl_files": True,
        "sub_folder":    True,     # 依聯絡人建子資料夾
    }
    if CONFIG_FILE.exists():
        try:
            data = json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
            default.update(data)
        except:
            pass
    return default

# ══════════════════════════════════════════════
#  資安：封鎖外部網路（僅允許內網）
# ══════════════════════════════════════════════
_orig_getaddrinfo = socket.getaddrinfo

def _blocked_getaddrinfo(host, *args, **kwargs):
    """只允許私有 IP 或 localhost，其他全擋"""
    private_prefixes = ('192.168.', '10.', '172.', '127.', 'localhost', '::1')
    if not any(str(host).startswith(p) for p in private_prefixes):
        raise ConnectionRefusedError(f"[資安] 已封鎖外部連線嘗試：{host}")
    return _orig_getaddrinfo(host, *args, **kwargs)

socket.getaddrinfo = _blocked_getaddrinfo

# ══════════════════════════════════════════════
#  工具函式
# ══════════════════════════════════════════════
def now_str():  return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
def date_str(): return datetime.now().strftime("%Y-%m-%d")
def time_str(): return datetime.now().strftime("%H:%M:%S")

def log(msg):
    """內部 log，只寫到 script 同目錄的 run.log，不對外"""
    entry = f"[{now_str()}] {msg}\n"
    try:
        (SCRIPT_DIR / "run.log").open("a", encoding="utf-8").write(entry)
    except:
        pass

def contact_dir(cfg, sender):
    root = Path(cfg["shared_folder"])
    if cfg["sub_folder"]:
        d = root / sanitize(sender)
    else:
        d = root
    d.mkdir(parents=True, exist_ok=True)
    return d

def sanitize(name):
    """移除檔名不合法字元"""
    return re.sub(r'[\\/:*?"<>|]', '_', name).strip()

def should_watch(cfg, sender):
    wl = cfg.get("watch_list", [])
    if not wl:
        return True
    return any(n in sender for n in wl)

# ══════════════════════════════════════════════
#  訊息寫入（含對話分段邏輯）
# ══════════════════════════════════════════════
# last_msg_time[sender] = datetime of last message
last_msg_time: dict = {}

def write_message(cfg, sender, text):
    d     = contact_dir(cfg, sender)
    fpath = d / f"{sanitize(sender)}_{date_str()}.txt"
    now   = datetime.now()
    gap   = cfg.get("gap_minutes", 5)

    lines = []

    # 判斷是否需要分段
    if sender in last_msg_time:
        elapsed = (now - last_msg_time[sender]).total_seconds() / 60
        if elapsed >= gap:
            lines.append(f"\n── {now_str()} 新段落 ──────────────────────\n")
    else:
        # 第一次寫入，加上日期標頭
        if not fpath.exists():
            lines.append(f"═══ {sender}  {date_str()} ═══════════════════════\n\n")

    lines.append(f"[{time_str()}] {text}\n")
    last_msg_time[sender] = now

    with fpath.open("a", encoding="utf-8") as f:
        f.writelines(lines)

    log(f"MSG [{sender}] {text[:60]}")

# ══════════════════════════════════════════════
#  檔案下載
# ══════════════════════════════════════════════
seen_files: set = set()

def handle_file(cfg, src_path, sender):
    if not cfg.get("auto_dl_files"):
        return
    src = Path(src_path)
    if not src.exists() or str(src) in seen_files:
        return
    seen_files.add(str(src))

    d = contact_dir(cfg, sender) / "收到的檔案"
    d.mkdir(exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dst   = d / f"[{stamp}]_{src.name}"
    try:
        import shutil
        shutil.copy2(src, dst)
        write_message(cfg, sender, f"【收到檔案】{src.name}")
        log(f"FILE [{sender}] {src.name} → {dst.name}")
    except Exception as e:
        log(f"FILE_ERR {e}")

# ══════════════════════════════════════════════
#  讀取 Windows 通知資料庫（LINE 訊息來源）
# ══════════════════════════════════════════════
seen_notif: set = set()

def read_win_notifications(cfg):
    db_path = Path(os.environ.get("LOCALAPPDATA","")) / \
              "Microsoft/Windows/Notifications/wpndatabase.db"
    if not db_path.exists():
        return

    try:
        con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True, timeout=3)
        cur = con.cursor()
        cur.execute("""
            SELECT Payload, ArrivalTime
            FROM   Notification
            WHERE  AppId LIKE '%LINE%'
            ORDER  BY ArrivalTime DESC
            LIMIT  100
        """)
        rows = cur.fetchall()
        con.close()
    except Exception as e:
        log(f"NOTIF_DB_ERR {e}")
        return

    for payload, arrival in rows:
        key = f"{arrival}_{hash(str(payload))}"
        if key in seen_notif:
            continue
        seen_notif.add(key)
        parse_payload(cfg, payload)

def parse_payload(cfg, payload):
    try:
        s = payload.decode("utf-8", "ignore") if isinstance(payload, bytes) else str(payload)
        texts = re.findall(r'<text[^>]*>([^<]+)</text>', s)
        if len(texts) >= 2:
            sender, msg = texts[0].strip(), texts[1].strip()
            if sender and msg and should_watch(cfg, sender):
                write_message(cfg, sender, msg)
    except Exception as e:
        log(f"PARSE_ERR {e}")

# ══════════════════════════════════════════════
#  掃描下載資料夾（接收到的檔案）
# ══════════════════════════════════════════════
def scan_downloads(cfg):
    dl = Path.home() / "Downloads"
    if not dl.exists():
        return
    cutoff = time.time() - cfg["interval_sec"] * 2
    try:
        for f in dl.iterdir():
            if f.is_file() and f.stat().st_mtime > cutoff:
                handle_file(cfg, f, "LINE用戶")
    except Exception as e:
        log(f"DL_SCAN_ERR {e}")

# ══════════════════════════════════════════════
#  設定初始化（建立 config.json 若不存在）
# ══════════════════════════════════════════════
def ensure_config():
    if not CONFIG_FILE.exists():
        default = {
            "shared_folder": "",
            "watch_list":    [],
            "interval_sec":  60,
            "gap_minutes":   5,
            "auto_dl_files": True,
            "sub_folder":    True,
        }
        CONFIG_FILE.write_text(
            json.dumps(default, ensure_ascii=False, indent=2),
            encoding="utf-8"
        )
        log("建立預設 config.json，請先用控制台設定後再啟動監控")
        return False
    return True

# ══════════════════════════════════════════════
#  主迴圈
# ══════════════════════════════════════════════
def main():
    if not ensure_config():
        return

    cfg = load_config()

    if not cfg["shared_folder"]:
        log("ERROR: shared_folder 未設定，請先執行控制台")
        return

    shared = Path(cfg["shared_folder"])
    try:
        shared.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        log(f"ERROR: 無法存取共用資料夾 {shared} — {e}")
        return

    log(f"=== 監控啟動 | 路徑：{shared} | 對象：{cfg['watch_list'] or '全部'} ===")

    # 預熱：把現有通知標記為已讀，避免啟動時重複寫入
    read_win_notifications.__globals__["seen_notif"] = set()
    _warmup_cfg = dict(cfg); _warmup_cfg["watch_list"] = []
    # 先靜默跑一次，只記錄 key 不寫入
    db_path = Path(os.environ.get("LOCALAPPDATA","")) / \
              "Microsoft/Windows/Notifications/wpndatabase.db"
    if db_path.exists():
        try:
            con = sqlite3.connect(f"file:{db_path}?mode=ro", uri=True, timeout=3)
            cur = con.cursor()
            cur.execute("SELECT Payload, ArrivalTime FROM Notification WHERE AppId LIKE '%LINE%' ORDER BY ArrivalTime DESC LIMIT 200")
            for _, arr in cur.fetchall():
                seen_notif.add(f"{arr}_{hash('')}")  # 粗略標記
            # 更精確
            cur.execute("SELECT Payload, ArrivalTime FROM Notification WHERE AppId LIKE '%LINE%' ORDER BY ArrivalTime DESC LIMIT 200")
            for p, arr in cur.fetchall():
                seen_notif.add(f"{arr}_{hash(str(p))}")
            con.close()
        except:
            pass
    log("預熱完成，開始監控新訊息")

    try:
        while True:
            cfg = load_config()  # 每次循環重新讀設定，支援熱更新
            read_win_notifications(cfg)
            scan_downloads(cfg)
            time.sleep(cfg["interval_sec"])
    except KeyboardInterrupt:
        log("=== 監控停止 ===")

if __name__ == "__main__":
    main()
