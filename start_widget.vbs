Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
localApp = WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")

pyPath = ""

If fso.FileExists("C:\Python313\pythonw.exe") Then
    pyPath = "C:\Python313\pythonw.exe"
ElseIf fso.FileExists("C:\Python312\pythonw.exe") Then
    pyPath = "C:\Python312\pythonw.exe"
ElseIf fso.FileExists("C:\Python311\pythonw.exe") Then
    pyPath = "C:\Python311\pythonw.exe"
ElseIf fso.FileExists("C:\Python310\pythonw.exe") Then
    pyPath = "C:\Python310\pythonw.exe"
ElseIf fso.FileExists(localApp & "\Programs\Python\Python313\pythonw.exe") Then
    pyPath = localApp & "\Programs\Python\Python313\pythonw.exe"
ElseIf fso.FileExists(localApp & "\Programs\Python\Python312\pythonw.exe") Then
    pyPath = localApp & "\Programs\Python\Python312\pythonw.exe"
ElseIf fso.FileExists(localApp & "\Programs\Python\Python311\pythonw.exe") Then
    pyPath = localApp & "\Programs\Python\Python311\pythonw.exe"
ElseIf fso.FileExists(localApp & "\Programs\Python\Python310\pythonw.exe") Then
    pyPath = localApp & "\Programs\Python\Python310\pythonw.exe"
End If

If pyPath = "" Then
    MsgBox "Python을 찾을 수 없습니다." & vbCrLf & vbCrLf & "설치_및_실행.bat 을 먼저 실행해 주세요.", vbExclamation, "Calendar and To do list"
    WScript.Quit 1
End If

WshShell.Run """" & pyPath & """ """ & scriptDir & "\main.py""", 0, False
