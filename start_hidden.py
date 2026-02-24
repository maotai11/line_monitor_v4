"""
啟動器：以完全隱藏視窗方式執行 line_monitor.py
本程式自身也不顯示視窗
"""
import sys, subprocess, os
from pathlib import Path

# 這個啟動器本身要隱藏（用 pythonw.exe 或 VBScript 呼叫時就沒視窗）
SCRIPT = Path(__file__).parent / "line_monitor.py"
PYTHON = Path(sys.executable)

# 優先用 pythonw.exe（無視窗版）
pythonw = PYTHON.parent / "pythonw.exe"
if not pythonw.exists():
    pythonw = PYTHON

subprocess.Popen(
    [str(pythonw), str(SCRIPT)],
    creationflags=0x08000000,   # CREATE_NO_WINDOW
    close_fds=True,
    stdin=subprocess.DEVNULL,
    stdout=subprocess.DEVNULL,
    stderr=subprocess.DEVNULL,
)
