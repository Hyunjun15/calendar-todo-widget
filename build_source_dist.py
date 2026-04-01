# -*- coding: utf-8 -*-
"""
Calendar and To do list — 소스코드 배포 패키지 빌드 스크립트
실행:  python build_source_dist.py
"""

import shutil, sys, os, re
from pathlib import Path
from datetime import date

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except AttributeError:
    pass

BASE      = Path(__file__).parent
DIST_ROOT = BASE / "배포패키지_소스"
APP_DIR   = DIST_ROOT / "Calendar_and_To_do_list"


def _read_version() -> tuple[str, str]:
    src   = (BASE / "main.py").read_text(encoding="utf-8")
    ver   = re.search(r'^APP_VERSION\s*=\s*"([^"]+)"',      src, re.M)
    date_ = re.search(r'^APP_VERSION_DATE\s*=\s*"([^"]+)"', src, re.M)
    return (ver.group(1) if ver else "?"), (date_.group(1) if date_ else "?")


def step(n: int, msg: str):
    print(f"\n{'─'*52}\n[{n}] {msg}")


def main():
    app_ver, app_date = _read_version()
    print("=" * 52)
    print(f"  Calendar and To do list  소스 배포 빌더  {app_ver}  ({app_date})")
    print("=" * 52)

    # ── 1. 이전 빌드 정리 ──────────────────────────────
    step(1, "이전 빌드 정리")
    shutil.rmtree(DIST_ROOT, ignore_errors=True)
    APP_DIR.mkdir(parents=True, exist_ok=True)
    print("    완료")

    # ── 2. main.py 복사 ────────────────────────────────
    step(2, "main.py 복사")
    shutil.copy2(BASE / "main.py", APP_DIR / "main.py")
    print("    완료")

    # ── 3. assets 복사 ─────────────────────────────────
    step(3, "assets 폴더 복사")
    shutil.copytree(BASE / "assets", APP_DIR / "assets", dirs_exist_ok=True)
    print("    완료")

    # ── 4. Update works 템플릿 ────────────────────────
    step(4, "Update works 폴더 및 메모장 템플릿 생성")
    uw_dst = APP_DIR / "Update works"
    uw_dst.mkdir(exist_ok=True)
    today_str = date.today().strftime("%Y.%m.%d")
    template_path = uw_dst / f"{today_str}.txt"
    template_content = (
        "[과제 및 To do list]\n"
        "1. 예시 과제\n"
        "\t내용: 과제 내용을 여기에 작성하세요\n"
        "\t목표: 달성 목표를 작성하세요\n"
        "\t마감기한: YYYY-MM-DD\n"
        "2. 두 번째 과제\n"
        "\t내용: \n"
        "\t목표: \n"
        "\t마감기한: \n"
        "\n"
        "[이번주/차주 긴급 업무]\n"
        "1. 긴급 업무 내용을 여기에 작성하세요\n"
        "\n"
        "[기타]\n"
        "1. 기타 메모 제목\n"
        "내용을 자유롭게 작성하세요\n"
    )
    template_path.write_text(template_content, encoding="utf-8")
    print(f"    템플릿 생성: {today_str}.txt")

    # ── 5. requirements.txt ───────────────────────────
    step(5, "requirements.txt 생성")
    (APP_DIR / "requirements.txt").write_text("PySide6>=6.5.0\n", encoding="utf-8")
    print("    완료")

    # ── 6. 설치_및_실행.bat 생성 (Python 자동 설치 포함) ──
    step(6, "설치_및_실행.bat 생성")

    # Python 3.12.10 공식 다운로드 URL
    PY_URL = "https://www.python.org/ftp/python/3.13.3/python-3.13.3-amd64.exe"

    install_lines = [
        "@echo off",
        "title Calendar and To do list - Setup",
        "color 0A",
        "cls",
        "echo.",
        "echo  ================================================",
        "echo   Calendar and To do list  -  Setup",
        "echo  ================================================",
        "echo.",
        "",
        ":: Log file for debugging",
        'set "LOG=%~dp0install_log.txt"',
        "echo Setup started > %LOG%",
        "",
        ":: ── Search for Python ──────────────────────────",
        'set "PYTHON_CMD="',
        "",
        "where python >nul 2>&1",
        "if %errorlevel%==0 (",
        '    set "PYTHON_CMD=python"',
        "    goto :check_version",
        ")",
        "where py >nul 2>&1",
        "if %errorlevel%==0 (",
        '    set "PYTHON_CMD=py"',
        "    goto :check_version",
        ")",
        "goto :install_python",
        "",
        ":: ── Version check (3.10+) ──────────────────────",
        ":check_version",
        'for /f "tokens=2 delims= " %%v in (\'%PYTHON_CMD% --version 2^>^&1\') do set "PY_VER=%%v"',
        'for /f "tokens=1 delims=." %%m in ("%PY_VER%") do set "PY_MAJOR=%%m"',
        'for /f "tokens=2 delims=." %%n in ("%PY_VER%") do set "PY_MINOR=%%n"',
        "echo  Python %PY_VER% found",
        "echo Python %PY_VER% found >> %LOG%",
        "if %PY_MAJOR% LSS 3 goto :install_python",
        "if %PY_MAJOR%==3 if %PY_MINOR% LSS 10 goto :install_python",
        "goto :install_packages",
        "",
        ":: ── Download and install Python 3.13 ───────────",
        ":install_python",
        "echo  Python not found or version too old. Installing Python 3.13...",
        "echo  (This may take a few minutes)",
        "echo.",
        f'set "PY_INSTALLER=%TEMP%\\python-3.13.3-setup.exe"',
        f'set "PY_URL={PY_URL}"',
        "",
        "echo  [1/3] Downloading Python 3.13.3 (~25MB)...",
        "echo Downloading Python... >> %LOG%",
        'powershell -NoProfile -ExecutionPolicy Bypass -Command "(New-Object Net.WebClient).DownloadFile(\'%PY_URL%\', \'%PY_INSTALLER%\')"',
        "if %errorlevel% neq 0 (",
        "    echo.",
        "    echo  [ERROR] Download failed.",
        "    echo  Please install Python 3.10+ manually: https://www.python.org/downloads/",
        "    echo  (Check 'Add Python to PATH' during install)",
        "    echo Download failed >> %LOG%",
        "    pause",
        "    exit /b 1",
        ")",
        "",
        "echo  [2/3] Installing Python (please wait)...",
        "echo Installing Python... >> %LOG%",
        '"%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_launcher=1',
        "if %errorlevel% neq 0 (",
        "    echo.",
        "    echo  [ERROR] Install failed. Try running as Administrator.",
        "    echo Install failed >> %LOG%",
        "    pause",
        "    exit /b 1",
        ")",
        'del "%PY_INSTALLER%" >nul 2>&1',
        "echo  [3/3] Python installed!",
        "echo Python installed >> %LOG%",
        "echo.",
        "",
        ":: Refresh PATH from registry",
        'for /f "tokens=2*" %%a in (\'reg query "HKCU\\Environment" /v PATH 2^>nul\') do set "USR_PATH=%%b"',
        'set "PATH=%PATH%;%USR_PATH%;%LOCALAPPDATA%\\Programs\\Python\\Python313;%LOCALAPPDATA%\\Programs\\Python\\Python313\\Scripts"',
        'set "PYTHON_CMD=python"',
        "",
        ":: ── Install PySide6 ─────────────────────────────",
        ":install_packages",
        "echo  Checking packages...",
        "%PYTHON_CMD% -c \"import PySide6\" >nul 2>&1",
        "if %errorlevel% neq 0 (",
        "    echo  Installing PySide6 (first time only, 2-5 min)...",
        "    echo Installing PySide6... >> %LOG%",
        '    %PYTHON_CMD% -m pip install PySide6 -q --disable-pip-version-check',
        "    if %errorlevel% neq 0 (",
        "        echo.",
        "        echo  [ERROR] PySide6 install failed. Try running as Administrator.",
        "        echo PySide6 failed >> %LOG%",
        "        pause",
        "        exit /b 1",
        "    )",
        "    echo  PySide6 installed!",
        ") else (",
        "    echo  PySide6 already installed.",
        ")",
        "echo PySide6 OK >> %LOG%",
        "",
        ":: ── Launch widget ───────────────────────────────",
        'set "SCRIPT_DIR=%~dp0"',
        "echo.",
        "echo  Launching Calendar and To do list...",
        "echo Launching... >> %LOG%",
        "",
        'for /f "tokens=*" %%p in (\'where pythonw 2^>nul\') do set "PYTHONW=%%p"',
        "if defined PYTHONW (",
        '    start "" "%PYTHONW%" "%SCRIPT_DIR%main.py"',
        ") else (",
        '    start "" %PYTHON_CMD% "%SCRIPT_DIR%main.py"',
        ")",
        "",
        "echo  Done! Check the right side of your screen.",
        "echo Done >> %LOG%",
        "echo.",
        "timeout /t 3 >nul",
        "",
    ]

    install_content = "\r\n".join(install_lines)
    # CP949(ANSI)로 저장 - 한국어 Windows cmd 기본 인코딩
    (APP_DIR / "설치_및_실행.bat").write_bytes(install_content.encode("cp949", errors="replace"))
    print("    완료")

    # ── 7. 사용_안내.txt (프로젝트 루트의 파일 복사) ─
    step(7, "사용_안내.txt 복사")
    guide_src = BASE / "사용_안내.txt"
    if guide_src.exists():
        shutil.copy2(guide_src, APP_DIR / "사용_안내.txt")
        print("    완료")
    else:
        print("    [경고] 사용_안내.txt 없음 — 건너뜀")

    # ── 완료 ──────────────────────────────────────────
    print("\n" + "=" * 52)
    print("  [완료] 소스 배포 패키지 생성 성공!")
    print(f"  경로: {APP_DIR}")
    print()
    print("  배포 방법:")
    print("  1. '배포패키지_소스/Calendar_and_To_do_list' 폴더를 zip 압축")
    print("  2. zip 전달")
    print("  3. 받는 사람: 압축 해제 → '설치_및_실행.bat' 더블클릭")
    print("=" * 52)

    os.startfile(str(APP_DIR))


if __name__ == "__main__":
    main()
