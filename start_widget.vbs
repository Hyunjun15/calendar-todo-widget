' Calendar and To do list — 자동 시작 스크립트
' 컴퓨터 로그인 시 콘솔 창 없이 백그라운드 실행

Option Explicit

Dim WshShell, scriptDir

Set WshShell = CreateObject("WScript.Shell")

' 스크립트 위치를 작업 디렉토리로 설정
scriptDir = "C:\Users\Admin\Desktop\Claude\To_do_list_and_calender"

' 이미 실행 중이면 중복 실행 방지
Dim objWMI, colProcesses
Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name='pythonw.exe'")

WshShell.CurrentDirectory = scriptDir

' 0 = 창 숨김, False = 비동기 (기다리지 않음)
WshShell.Run "C:\Python312\pythonw.exe main.py", 0, False

Set WshShell = Nothing
Set objWMI = Nothing
