' 【啟動監控】雙擊此檔案即可，完全靜默，工作列不出現任何視窗
' 需要 Python Embeddable 放在同目錄 python\ 資料夾

Dim objShell
Set objShell = CreateObject("WScript.Shell")

' 取得當前目錄
Dim sDir
sDir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

' Python 路徑（優先找 pythonw，其次找 python）
Dim sPythonW, sPython
sPythonW = sDir & "\python\pythonw.exe"
sPython  = sDir & "\python\python.exe"

Dim sPyExe
If CreateObject("Scripting.FileSystemObject").FileExists(sPythonW) Then
    sPyExe = sPythonW
ElseIf CreateObject("Scripting.FileSystemObject").FileExists(sPython) Then
    sPyExe = sPython
Else
    MsgBox "找不到 Python，請確認 python\ 資料夾已放入 Python Embeddable", 16, "LINE Monitor"
    WScript.Quit
End If

Dim sScript
sScript = sDir & "\start_hidden.py"

' 0 = 完全隱藏視窗
objShell.Run """" & sPyExe & """ """ & sScript & """", 0, False

Set objShell = Nothing
