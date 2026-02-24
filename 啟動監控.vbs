' 【停止監控】關閉背景執行的監控程式

Dim objShell, objWMI, colProcess
Set objShell = CreateObject("WScript.Shell")
Set objWMI   = GetObject("winmgmts:\\.\root\cimv2")

Set colProcess = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name='pythonw.exe' OR Name='python.exe'")

Dim nKilled : nKilled = 0
For Each oProc In colProcess
    If InStr(LCase(oProc.CommandLine), "line_monitor") > 0 Then
        oProc.Terminate()
        nKilled = nKilled + 1
    End If
Next

If nKilled > 0 Then
    MsgBox "已停止 " & nKilled & " 個監控程序。", 64, "LINE Monitor"
Else
    MsgBox "找不到正在執行的監控程序。", 48, "LINE Monitor"
End If
