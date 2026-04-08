"""
ui_test_lib.py — Calendar and To do list 앱 UI 자동화 테스트 라이브러리
pywinauto + win32 기반. 에이전트가 직접 import해서 사용한다.
"""
import time, os, sys
from pathlib import Path
from PIL import Image
import win32gui, win32ui, win32con, win32api
import ctypes

# ─── 상수 ───────────────────────────────────────────────────────────────────
APP_TITLE   = "Calendar and To do list"
SHOT_DIR    = Path(__file__).parent / "troubleshooting_screenshots"
SHOT_DIR.mkdir(exist_ok=True)

# ─── 창 찾기 ─────────────────────────────────────────────────────────────────
def find_app_hwnd() -> int:
    """앱 메인 창 핸들 반환. 없으면 0."""
    result = []
    def cb(hwnd, _):
        if win32gui.GetWindowText(hwnd) == APP_TITLE \
           and win32gui.GetClassName(hwnd) == "Qt6110QWindowIcon":
            result.append(hwnd)
        return True
    win32gui.EnumWindows(cb, None)
    return result[0] if result else 0

def app_rect() -> tuple[int,int,int,int]:
    """(left, top, right, bottom) 절대 좌표."""
    hwnd = find_app_hwnd()
    if not hwnd:
        raise RuntimeError("앱이 실행 중이지 않습니다.")
    return win32gui.GetWindowRect(hwnd)

def win_size() -> tuple[int,int]:
    l,t,r,b = app_rect()
    return r-l, b-t

# ─── 스크린샷 ────────────────────────────────────────────────────────────────
def _capture_hwnd(hwnd: int) -> Image.Image:
    """특정 hwnd를 PrintWindow로 캡처."""
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r-l, b-t
    hwndDC = win32gui.GetWindowDC(hwnd)
    mfcDC  = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
    bmp    = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfcDC, w, h)
    saveDC.SelectObject(bmp)
    if not ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 2):
        ctypes.windll.user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), 0)
    info = bmp.GetInfo()
    data = bmp.GetBitmapBits(True)
    img  = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]),
                            data, "raw", "BGRX", 0, 1)
    win32gui.DeleteObject(bmp.GetHandle())
    saveDC.DeleteDC(); mfcDC.DeleteDC()
    win32gui.ReleaseDC(hwnd, hwndDC)
    return img

def screenshot(name: str = "snap") -> Image.Image:
    """앱 창 캡처. Qt6110QWindow 다이얼로그가 있으면 합성."""
    import numpy as np
    app_hwnd = find_app_hwnd()
    img = _capture_hwnd(app_hwnd)
    img_arr = np.array(img)

    # Qt6110QWindow 다이얼로그 합성
    dialogs = find_dialogs()
    app_l, app_t, _, _ = win32gui.GetWindowRect(app_hwnd)
    for dlg_hwnd in dialogs:
        try:
            dl, dt, dr, db = win32gui.GetWindowRect(dlg_hwnd)
            dlg_img = _capture_hwnd(dlg_hwnd)
            dlg_arr = np.array(dlg_img)
            if dlg_arr.max() == 0:
                continue
            # 앱 창 상대 좌표
            rx, ry = dl - app_l, dt - app_t
            dh, dw = dlg_arr.shape[:2]
            x1 = max(0, rx); y1 = max(0, ry)
            x2 = min(img_arr.shape[1], rx + dw)
            y2 = min(img_arr.shape[0], ry + dh)
            if x2 > x1 and y2 > y1:
                sx1 = x1 - rx; sy1 = y1 - ry
                img_arr[y1:y2, x1:x2] = dlg_arr[sy1:sy1+(y2-y1), sx1:sx1+(x2-x1)]
        except Exception:
            pass

    img = Image.fromarray(img_arr)
    path = SHOT_DIR / f"{name}.png"
    img.save(path)
    return img

# ─── 포커스 ──────────────────────────────────────────────────────────────────
def focus_app():
    """앱 창을 포그라운드로 가져오고 포커스를 준다."""
    hwnd = find_app_hwnd()
    if not hwnd:
        return
    try:
        win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        # 포그라운드 전환 실패 시 마우스 클릭으로 포커스 확보
        l, t, r, b = win32gui.GetWindowRect(hwnd)
        cx, cy = (l+r)//2, t + 20
        win32api.SetCursorPos((cx, cy))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, cx, cy, 0, 0)
        time.sleep(0.05)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, cx, cy, 0, 0)
    time.sleep(0.2)

# ─── 마우스 조작 ─────────────────────────────────────────────────────────────
def _abs(rx: int, ry: int) -> tuple[int,int]:
    """창 내 상대 좌표 → 화면 절대 좌표."""
    l, t, _, _ = app_rect()
    return l + rx, t + ry

def click(rx: int, ry: int, button: str = "left"):
    """창 내 (rx, ry) 좌표 클릭."""
    x, y = _abs(rx, ry)
    win32api.SetCursorPos((x, y))
    time.sleep(0.05)
    if button == "left":
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,  x, y, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,    x, y, 0, 0)
    else:
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN, x, y, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,   x, y, 0, 0)
    time.sleep(0.15)

def double_click(rx: int, ry: int):
    click(rx, ry); time.sleep(0.08); click(rx, ry)

def right_click(rx: int, ry: int):
    click(rx, ry, button="right")

def drag(rx1, ry1, rx2, ry2, steps: int = 20, delay: float = 0.02):
    """(rx1,ry1) → (rx2,ry2) 드래그."""
    x1, y1 = _abs(rx1, ry1)
    x2, y2 = _abs(rx2, ry2)
    win32api.SetCursorPos((x1, y1))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, x1, y1, 0, 0)
    time.sleep(0.05)
    for i in range(1, steps + 1):
        nx = x1 + (x2 - x1) * i // steps
        ny = y1 + (y2 - y1) * i // steps
        win32api.SetCursorPos((nx, ny))
        time.sleep(delay)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, x2, y2, 0, 0)
    time.sleep(0.2)

# ─── 키보드 조작 ─────────────────────────────────────────────────────────────
def key(vk: int):
    win32api.keybd_event(vk, 0, 0, 0)
    time.sleep(0.03)
    win32api.keybd_event(vk, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.05)

def hotkey(vk1: int, vk2: int):
    win32api.keybd_event(vk1, 0, 0, 0)
    time.sleep(0.02)
    win32api.keybd_event(vk2, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(vk2, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(vk1, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.1)

def _send(keys_str: str, pause: float = 0.05):
    """pywinauto send_keys — 포커스를 건드리지 않고 전송."""
    from pywinauto.keyboard import send_keys
    send_keys(keys_str, pause=pause)
    time.sleep(0.1)

def escape():
    """ESC 키 전송 (현재 포커스된 창에 전달)."""
    _send_input_key(0x1B)  # VK_ESCAPE

def enter():
    """Enter 키 전송."""
    _send_input_key(0x0D)  # VK_RETURN

def _send_input_key(vk: int):
    """SendInput으로 키 이벤트 전송 (포커스 변경 없음)."""
    import ctypes as _ct
    from ctypes import wintypes as _wt

    class _KI(_ct.Structure):
        _fields_ = [("wVk", _wt.WORD), ("wScan", _wt.WORD),
                    ("dwFlags", _wt.DWORD), ("time", _wt.DWORD),
                    ("dwExtraInfo", _ct.POINTER(_ct.c_ulong))]

    class _IN(_ct.Structure):
        class _U(_ct.Union):
            _fields_ = [("ki", _KI)]
        _anonymous_ = ("_u",)
        _fields_ = [("type", _wt.DWORD), ("_u", _U)]

    def _mk(vk_, up=False):
        inp = _IN(type=1)
        inp.ki.wVk = vk_
        inp.ki.dwFlags = 2 if up else 0
        return inp

    seq = [_mk(vk), _mk(vk, up=True)]
    arr = (_IN * len(seq))(*seq)
    _ct.windll.user32.SendInput(len(seq), arr, _ct.sizeof(_IN))
    time.sleep(0.08)

def type_text(text: str):
    """텍스트 타이핑 (영문/숫자/한글 혼용 지원)."""
    from pywinauto.keyboard import send_keys
    focus_app()
    send_keys(text, with_spaces=True, pause=0.02)

# ─── 색상/픽셀 분석 ──────────────────────────────────────────────────────────
def pixel_color(img: Image.Image, rx: int, ry: int) -> tuple[int,int,int]:
    """이미지 내 특정 픽셀의 RGB 값."""
    return img.getpixel((rx, ry))[:3]

def region_avg_brightness(img: Image.Image,
                           x0, y0, x1, y1) -> float:
    """지정 영역의 평균 밝기 (0~255)."""
    crop = img.crop((x0, y0, x1, y1)).convert("L")
    pixels = list(crop.getdata())
    return sum(pixels) / len(pixels)

def has_dark_background(img: Image.Image, x0, y0, x1, y1,
                         threshold: float = 128) -> bool:
    return region_avg_brightness(img, x0, y0, x1, y1) < threshold

def find_color_region(img: Image.Image, target_rgb: tuple,
                      tolerance: int = 20) -> list[tuple[int,int]]:
    """이미지에서 특정 색상에 가까운 픽셀 좌표 목록 반환."""
    r, g, b = target_rgb
    matches = []
    for y in range(0, img.height, 4):
        for x in range(0, img.width, 4):
            pr, pg, pb = img.getpixel((x, y))[:3]
            if abs(pr-r) < tolerance and abs(pg-g) < tolerance and abs(pb-b) < tolerance:
                matches.append((x, y))
    return matches

# ─── 앱 실행 / 종료 ──────────────────────────────────────────────────────────
def launch_app():
    import subprocess
    subprocess.Popen(
        [r"C:\Python312\pythonw.exe",
         r"C:\Users\Admin\Desktop\Claude\To_do_list_and_calender\main.py"],
        creationflags=subprocess.DETACHED_PROCESS
    )
    for _ in range(15):
        time.sleep(1)
        if find_app_hwnd():
            time.sleep(1.5)  # UI 렌더링 대기
            return True
    return False

def kill_app():
    import subprocess
    subprocess.run(["cmd", "/c", "taskkill /F /IM pythonw.exe"], capture_output=True)
    time.sleep(1)

def ensure_app_running() -> bool:
    hwnd = find_app_hwnd()
    if hwnd and win32gui.IsWindowVisible(hwnd):
        return True
    # 앱이 없거나 숨겨진 경우 재시작
    kill_app()
    time.sleep(1)
    return launch_app()

# ─── 창 크기 조절 ────────────────────────────────────────────────────────────
def resize_window(dw: int, dh: int):
    """현재 창 크기에서 (dw, dh) 만큼 변경."""
    hwnd = find_app_hwnd()
    l, t, r, b = win32gui.GetWindowRect(hwnd)
    w, h = r-l, b-t
    win32gui.SetWindowPos(hwnd, None, l, t, w+dw, h+dh,
                          win32con.SWP_NOZORDER | win32con.SWP_NOMOVE)
    time.sleep(0.3)

# ─── 다이얼로그 탐지 ─────────────────────────────────────────────────────────
def _get_app_pid() -> int:
    """앱 프로세스 PID 반환."""
    import win32process
    hwnd = find_app_hwnd()
    if not hwnd:
        return 0
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return pid

def find_dialogs() -> list[int]:
    """앱 프로세스의 다이얼로그 창 목록 (Qt6110QWindow 기반)."""
    import win32process
    app_hwnd = find_app_hwnd()
    if not app_hwnd:
        return []
    _, app_pid = win32process.GetWindowThreadProcessId(app_hwnd)
    result = []
    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        if hwnd == app_hwnd:
            return True
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        if pid == app_pid:
            cls = win32gui.GetClassName(hwnd)
            # Qt6110QWindow: 다이얼로그가 별도 OS 창으로 열릴 때
            if 'Qt6110QWindow' in cls:
                result.append(hwnd)
        return True
    win32gui.EnumWindows(cb, None)
    return result

def dialog_opened(before_handles: list[int]) -> bool:
    """액션 전/후 핸들 비교로 새 다이얼로그 열렸는지 확인."""
    after = find_dialogs()
    new = [h for h in after if h not in before_handles]
    return len(new) > 0

def visual_changed(img_before: "Image.Image", img_after: "Image.Image",
                   threshold: int = 5000) -> bool:
    """두 스크린샷 픽셀 비교로 UI 변화 감지 (numpy 사용)."""
    import numpy as np
    from PIL import ImageChops
    diff = ImageChops.difference(img_before, img_after)
    return int(np.array(diff).sum()) > threshold

# ─── 리포트 ──────────────────────────────────────────────────────────────────
_results: list[dict] = []

def _p(text: str):
    """인코딩 안전 출력."""
    sys.stdout.buffer.write((text + "\n").encode("utf-8", errors="replace"))
    sys.stdout.buffer.flush()

def check(name: str, passed: bool, detail: str = ""):
    status = "[PASS]" if passed else "[FAIL]"
    _results.append({"name": name, "passed": passed, "detail": detail})
    suffix = f" -- {detail}" if detail else ""
    _p(f"  {status}  {name}{suffix}")

def report() -> dict:
    total  = len(_results)
    passed = sum(1 for r in _results if r["passed"])
    failed = total - passed
    _p("\n" + "="*50)
    _p(f"결과: {passed}/{total} PASS  |  {failed} FAIL")
    if failed:
        _p("\n실패 항목:")
        for r in _results:
            if not r["passed"]:
                _p(f"  [FAIL] {r['name']}: {r['detail']}")
    return {"total": total, "passed": passed, "failed": failed, "results": _results}

def reset():
    _results.clear()
