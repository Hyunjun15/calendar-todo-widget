# -*- coding: utf-8 -*-
"""
Calendar and To do list — 배포 패키지 빌드 스크립트 (PyInstaller, 독립 실행형 .exe)
실행:  python build_dist.py
"""

import subprocess, shutil, sys, os, re
from pathlib import Path

# Windows cmd cp949 환경에서 유니코드 출력 오류 방지
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE      = Path(__file__).parent
DIST_ROOT = BASE / "배포패키지"

# PyInstaller --name 은 ASCII 만 안전 → 빌드 후 rename
_EXE_ASCII = "productivity_widget"
_EXE_FINAL = "Calendar_and_To_do_list"
APP_DIR    = DIST_ROOT / _EXE_FINAL   # 최종 폴더명


def _read_version() -> tuple[str, str]:
    """main.py 에서 APP_VERSION, APP_VERSION_DATE 를 직접 읽어옴"""
    src = (BASE / "main.py").read_text(encoding="utf-8")
    ver  = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"',  src, re.M)
    date = re.search(r'^APP_VERSION_DATE\s*=\s*"([^"]+)"', src, re.M)
    return (ver.group(1) if ver else "?"), (date.group(1) if date else "?")


def step(n: int, msg: str):
    print(f"\n{'─'*52}\n[{n}] {msg}")


def main():
    app_ver, app_date = _read_version()
    print("=" * 52)
    print(f"  Calendar and To do list  배포 패키지 빌더  {app_ver}  ({app_date})")
    print("=" * 52)

    # ── 1. PyInstaller 확인 ────────────────────────────
    step(1, "PyInstaller 확인")
    try:
        import PyInstaller  # noqa
        print("    설치됨 :", PyInstaller.__version__)
    except ImportError:
        print("    없음 → 설치 중...")
        subprocess.run(
            [sys.executable, "-m", "pip", "install", "pyinstaller", "-q"],
            check=True)

    # ── 2. 이전 빌드 정리 ──────────────────────────────
    step(2, "이전 빌드 정리")
    shutil.rmtree(DIST_ROOT, ignore_errors=True)
    shutil.rmtree(BASE / "build_tmp", ignore_errors=True)
    for f in (BASE / f"{_EXE_ASCII}.spec",):
        if f.exists(): f.unlink()
    print("    완료")

    # ── 3. PyInstaller 빌드 ────────────────────────────
    step(3, "PyInstaller 빌드 중 (단일 파일 방식, 1~3분 소요)...")
    result = subprocess.run([
        sys.executable, "-m", "PyInstaller",
        "--onefile",
        "--windowed",
        "--name", _EXE_ASCII,          # ASCII 이름으로 빌드 → 후처리로 rename
        "--distpath", str(DIST_ROOT),
        "--workpath", str(BASE / "build_tmp"),
        "--specpath", str(BASE),
        str(BASE / "main.py"),
    ], cwd=str(BASE))

    if result.returncode != 0:
        print("\n    [오류] 빌드 실패! 위 로그를 확인하세요.")
        return

    # dist 폴더 하위에 바로 EXE가 생김 (onedir와 다름)
    ascii_exe = DIST_ROOT / f"{_EXE_ASCII}.exe"
    if not ascii_exe.exists():
        print("\n    [오류] 빌드된 EXE 파일을 찾을 수 없습니다.")
        return

    # 최종 폴더 생성 및 EXE 이동/이름 변경
    APP_DIR.mkdir(parents=True, exist_ok=True)
    final_exe = APP_DIR / f"{_EXE_FINAL}.exe"
    if final_exe.exists(): final_exe.unlink()
    shutil.move(ascii_exe, final_exe)
    
    print("    빌드 및 파일 이동 성공")

    # ── 4. assets 복사 ─────────────────────────────────
    step(4, "assets 폴더 복사")
    shutil.copytree(BASE / "assets", APP_DIR / "assets", dirs_exist_ok=True)
    print("    완료")

    # ── 5. Update works 복사 ───────────────────────────
    step(5, "Update works 폴더 복사")
    uw_src = BASE / "Update works"
    uw_dst = APP_DIR / "Update works"
    if uw_src.exists():
        shutil.copytree(uw_src, uw_dst, dirs_exist_ok=True)
        print(f"    {len(list(uw_dst.iterdir()))}개 파일 복사")
    else:
        uw_dst.mkdir(exist_ok=True)
        print("    빈 폴더 생성")

    # ── 6. 바탕화면 바로가기 bat 생성 ─────────────────
    step(6, "바탕화면_바로가기.bat 생성")
    bat_path = APP_DIR / "바탕화면_바로가기.bat"
    # ★ 핵심 변경:
    #   - PowerShell 멀티라인 호출 제거 (배치파일 내 한글 변수 전달 실패 원인)
    #   - .exe 이름 하드코딩 제거 → for 루프로 자동탐지
    #   - VBScript 임시파일 방식 사용 (한글 경로 가장 안정적)
    bat_lines = [
        "@echo off",
        "title Calendar and To do list - Desktop Shortcut Setup",
        "color 0A",
        "cls",
        "",
        "echo.",
        "echo  ================================================",
        "echo   Calendar and To do list  -  Desktop Shortcut Setup",
        "echo  ================================================",
        "echo.",
        "",
        'set "DIR=%~dp0"',
        'set "EXE_PATH="',
        "",
        'for %%F in ("%DIR%*.exe") do (',
        '    set "EXE_PATH=%%~fF"',
        '    set "EXE_NAME=%%~nF"',
        ")",
        "",
        "if not defined EXE_PATH (",
        "    echo  [Error] No .exe found in this folder.",
        "    echo  Path: %DIR%",
        "    echo.",
        "    pause",
        "    exit /b 1",
        ")",
        "",
        "echo  Found: %EXE_NAME%.exe",
        "echo  Creating shortcut on Desktop...",
        "echo.",
        "",
        'set "VBS=%TEMP%\\make_shortcut_%RANDOM%.vbs"',
        'echo Set ws = CreateObject^("WScript.Shell"^) > "%VBS%"',
        'echo dst = ws.SpecialFolders^("Desktop"^) >> "%VBS%"',
        'echo lnk = dst ^& "\\" ^& "%EXE_NAME%" ^& ".lnk" >> "%VBS%"',
        'echo Set sc = ws.CreateShortcut^(lnk^) >> "%VBS%"',
        'echo sc.TargetPath = "%EXE_PATH%" >> "%VBS%"',
        'echo sc.WorkingDirectory = "%DIR%" >> "%VBS%"',
        'echo sc.Description = "Calendar and To do list" >> "%VBS%"',
        'echo sc.Save >> "%VBS%"',
        "",
        'cscript //nologo "%VBS%"',
        'del "%VBS%" >nul 2>&1',
        "",
        'if exist "%USERPROFILE%\\Desktop\\%EXE_NAME%.lnk" (',
        "    echo  [OK] Shortcut created on Desktop!",
        "    echo  Double-click the icon to launch the widget.",
        ") else (",
        "    echo  [FAIL] Shortcut creation failed.",
        "    echo  Try running this file as Administrator.",
        ")",
        "",
        "echo.",
        "pause",
        "",
    ]
    bat_content = "\r\n".join(bat_lines)
    bat_path.write_bytes(bat_content.encode("ascii"))
    print("    완료")

    # ── 7. 사용_안내.txt 복사 ─────────────────────────
    step(7, "사용_안내.txt 복사")
    _guide_src = BASE / "사용_안내.txt"
    if _guide_src.exists():
        shutil.copy2(str(_guide_src), str(APP_DIR / "사용_안내.txt"))
    else:
        print("    ⚠ 사용_안내.txt가 프로젝트 루트에 없습니다.")
    print("    완료")

    # ── 8. 임시 파일 정리 ─────────────────────────────
    step(8, "임시 파일 정리")
    shutil.rmtree(BASE / "build_tmp", ignore_errors=True)
    for f in (BASE / f"{_EXE_ASCII}.spec", BASE / f"{_EXE_FINAL}.spec"):
        if f.exists(): f.unlink()
    print("    완료")

    # ── 완료 ──────────────────────────────────────────
    print("\n" + "=" * 52)
    print("  [완료] 배포 패키지 생성 성공!")
    print(f"  경로: {APP_DIR}")
    print()
    print("  배포 방법:")
    print("  1. '배포패키지/Calendar_and_To_do_list' 폴더 전체를 zip 압축")
    print("  2. zip 파일 전달")
    print("  3. 받는 사람: 압축 해제 후 '바탕화면_바로가기.bat' 실행")
    print("     또는 'Calendar_and_To_do_list.exe' 직접 실행")
    print("=" * 52)

    # 탐색기로 결과 폴더 열기
    os.startfile(str(APP_DIR))


if __name__ == "__main__":
    main()
