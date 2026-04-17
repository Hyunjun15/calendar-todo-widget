' Calendar and To do list — 자동 시작 스크립트
' 콘솔 창 없이 백그라운드 실행 (어느 PC에서든 작동)

Option Explicit

Dim WshShell, fso, scriptDir

Set WshShell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' 스크립트가 있는 폴더를 작업 디렉토리로 설정 (상대경로)
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)
WshShell.CurrentDirectory = scriptDir

' pythonw (콘솔 없음) 또는 python (폴백) 으로 실행
' 0 = 창 숨김, False = 비동기 (기다리지 않음)
On Error Resume Next
WshShell.Run "pythonw """ & scriptDir & "\main.py""", 0, False
If Err.Number <> 0 Then
    Err.Clear
    WshShell.Run "python """ & scriptDir & "\main.py""", 0, False
End If
On Error GoTo 0

Set fso = Nothing
Set WshShell = Nothing
