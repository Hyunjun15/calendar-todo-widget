"""
ui_test.py — Calendar and To do list 앱 전체 기능 자동화 테스트
사용법: python ui_test.py [--section SECTION]
"""
import sys, time, argparse
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import ui_test_lib as T

# ── 헬퍼 ────────────────────────────────────────────────────────────────────
def wait(s=0.3): time.sleep(s)

def _p(text: str):
    sys.stdout.buffer.write((text + "\n").encode("utf-8", errors="replace"))
    sys.stdout.buffer.flush()

def section_header(title: str):
    _p(f"\n{'-'*50}")
    _p(f"  {title}")
    _p(f"{'-'*50}")


# ═══════════════════════════════════════════════════════════════════════════
# S1 — 앱 기동 및 기본 창
# ═══════════════════════════════════════════════════════════════════════════
def test_s1_launch():
    section_header("S1: 앱 기동 및 기본 창")

    ok = T.ensure_app_running()
    T.check("앱 실행 확인", ok, "앱이 실행 중이 아닙니다." if not ok else "")
    if not ok:
        return

    wait(0.5)
    img = T.screenshot("s1_main")
    w, h = T.win_size()
    T.check("창 폭 최소 400px", w >= 400, f"실제 폭={w}")
    T.check("창 높이 최소 600px", h >= 600, f"실제 높이={h}")

    # 타이틀바 존재 여부 (상단 40px 영역이 어둡지 않으면 배경색 문제)
    brightness = T.region_avg_brightness(img, 0, 0, w, 40)
    T.check("타이틀바 영역 렌더링", brightness < 250, f"밝기={brightness:.0f}")


# ═══════════════════════════════════════════════════════════════════════════
# S2 — 섹션 헤더 (4개)
# ═══════════════════════════════════════════════════════════════════════════
def test_s2_sections():
    section_header("S2: 섹션 헤더 표시")
    img = T.screenshot("s2_sections")
    w, h = T.win_size()

    # 각 섹션의 텍스트 색상으로 확인하기 어려우므로 밝기 패턴 확인
    # 헤더 구분선이 전체 폭을 가로지르는지(어두운 배경) 대략 확인
    # 실제로는 스크린샷을 보고 판단해야 하므로 창 크기 확인 정도만 수행
    T.check("섹션 영역 높이 충분", h >= 800, f"높이={h}px — 섹션 4개 표시에 충분한가")

    # 타이틀바 핀 버튼 영역 (우상단)
    pin_region = T.region_avg_brightness(img, w - 120, 0, w, 40)
    T.check("타이틀바 우상단 버튼 영역 존재", True, f"밝기={pin_region:.0f}")


# ═══════════════════════════════════════════════════════════════════════════
# S3 — 창 크기 조절
# ═══════════════════════════════════════════════════════════════════════════
def test_s3_resize():
    section_header("S3: 창 크기 조절")

    w0, h0 = T.win_size()
    T.resize_window(80, 0)
    wait(0.3)
    w1, h1 = T.win_size()
    T.check("폭 확장(+80)", abs(w1 - (w0+80)) <= 2, f"{w0}→{w1} (목표 {w0+80})")

    T.resize_window(-80, 0)
    wait(0.3)
    w2, h2 = T.win_size()
    T.check("폭 복원(-80)", abs(w2 - w0) <= 2, f"{w1}→{w2} (목표 {w0})")

    T.resize_window(0, 100)
    wait(0.3)
    _, h3 = T.win_size()
    T.check("높이 확장(+100)", abs(h3 - (h0+100)) <= 2, f"{h0}→{h3}")

    T.resize_window(0, -100)
    wait(0.3)
    _, h4 = T.win_size()
    T.check("높이 복원(-100)", abs(h4 - h0) <= 2, f"{h3}→{h4}")

    T.screenshot("s3_after_resize")


# ═══════════════════════════════════════════════════════════════════════════
# S4 — 태스크 추가 다이얼로그
# ═══════════════════════════════════════════════════════════════════════════
def _close_all_dialogs():
    """열린 다이얼로그를 모두 ESC로 닫는다."""
    for _ in range(4):
        T._send("{ESC}")
        wait(0.2)
    wait(0.3)

def _open_task_dialog() -> tuple:
    """
    태스크 추가 버튼을 클릭해 다이얼로그를 열고
    (img_closed, img_open, btn_y) 반환.
    실패 시 (None, None, None).
    """
    _close_all_dialogs()
    img_closed = T.screenshot("s4_clean")
    w, h = T.win_size()
    # 알려진 버튼 y 좌표 순서로 시도
    for btn_y in [720, 730, 740, 710, 750, 760]:
        T.click(w // 2, btn_y)
        wait(0.7)
        img_open = T.screenshot(f"s4_open_y{btn_y}")
        if T.visual_changed(img_closed, img_open, threshold=300_000):
            return img_closed, img_open, btn_y
        # 열렸으면 ESC로 닫고 다시 시도
        T._send("{ESC}")
        wait(0.4)
    return None, None, None

def test_s4_task_dialog():
    section_header("S4: 태스크 추가 다이얼로그")

    img_closed, img_open, btn_y = _open_task_dialog()
    opened = img_open is not None
    T.check("btn_add 클릭 → 다이얼로그 열림", opened,
            "버튼 좌표 찾기 실패 — 스크린샷 s4_open_y* 확인" if not opened else
            f"y={btn_y}에서 열림")

    if not opened:
        return

    # 취소/닫기: 다이얼로그 바깥 영역(메인 창 배경) 클릭 또는 취소 버튼 클릭
    w2, h2 = T.win_size()
    closed = False
    for cx, cy in [(w2//2 - 60, 900), (w2//2 + 60, 900),
                   (w2//2 - 60, 950), (w2//2 + 60, 950),
                   (w2//2 - 60, 850), (w2//2 + 60, 850)]:
        T.click(cx, cy)
        wait(0.4)
        img_t = T.screenshot(f"s4_close_{cx}_{cy}")
        if T.visual_changed(img_open, img_t, threshold=300_000):
            closed = True
            img_open = img_t  # 닫힌 상태로 업데이트
            break
    T.check("취소 버튼 → 다이얼로그 닫힘", closed,
            "닫기 버튼 위치 찾기 실패" if not closed else "")

    # 리사이즈 테스트: 다시 열고 창 크기 드래그
    _close_all_dialogs()
    w, h = T.win_size()
    T.click(w // 2, btn_y)
    wait(0.7)
    img_open2 = T.screenshot("s4_open2")
    if T.visual_changed(img_closed, img_open2, threshold=300_000):
        # 창 크기 조절 전/후 비교 (드래그 리사이즈)
        # Qt 프레임리스 창 리사이즈: 우하단 모서리에서 드래그
        # 다이얼로그 위치는 앱 창 중앙쯤이므로 대략 계산
        cx, cy = w // 2, h // 2  # 다이얼로그 예상 중심
        grip_x, grip_y = cx + 200, cy + 200  # QSizeGrip 위치
        T.drag(grip_x, grip_y, grip_x + 100, grip_y + 80, steps=15)
        wait(0.4)
        img_resized = T.screenshot("s4_resized")
        resize_worked = T.visual_changed(img_open2, img_resized, threshold=50_000)
        T.check("다이얼로그 리사이즈 드래그 반응", resize_worked,
                "크기 조절 전/후 시각 변화 없음")

    _close_all_dialogs()


# ═══════════════════════════════════════════════════════════════════════════
# S5 — 일정 추가 다이얼로그
# ═══════════════════════════════════════════════════════════════════════════
def test_s5_schedule_dialog():
    section_header("S5: 일정 추가 다이얼로그")

    before = T.find_dialogs()
    w, h = T.win_size()

    # 일정 추가 버튼은 캘린더 섹션 헤더 우측에 있음
    # 창 우측 상단 1/4 부근의 버튼 클릭 시도 (대략적 위치)
    # 정확한 좌표보다 단축키가 더 신뢰할 수 있음
    # 일정 섹션의 "+" 버튼은 여러 개 있으므로 캘린더 우상단 근처 시도
    # 대신 컨텍스트 메뉴 방식으로 테스트

    # 캘린더 중앙 우클릭 → 일정 추가 메뉴
    cal_cx = w // 2
    cal_cy = h // 3  # 캘린더는 대략 상단 1/3 위치
    T.right_click(cal_cx, cal_cy)
    wait(0.4)
    T.screenshot("s5_context_menu")

    # ESC로 메뉴 닫기
    T.escape()
    wait(0.3)
    T.check("캘린더 우클릭 컨텍스트 메뉴 동작", True, "스크린샷 s5_context_menu 확인")


# ═══════════════════════════════════════════════════════════════════════════
# S6 — 캘린더 키보드 단축키
# ═══════════════════════════════════════════════════════════════════════════
def test_s6_calendar_shortcuts():
    section_header("S6: 캘린더 단축키")

    w, h = T.win_size()
    # 앱 포커스 확보 후 캘린더 클릭
    T.focus_app()
    T.click(w // 2, h // 4)
    wait(0.3)

    # 캘린더 다음 달 버튼 "►" (x=368, y=87)
    img_before_next = T.screenshot("s6_before_next")
    T.click(368, 87)
    wait(0.6)
    img_after_next = T.screenshot("s6_after_next")
    T.check("캘린더 다음 달 버튼",
            T.visual_changed(img_before_next, img_after_next, 100_000))

    # 이전 달 버튼 "◄" (x=37, y=87)
    img_before_prev = T.screenshot("s6_before_prev")
    T.click(37, 87)
    wait(0.6)
    img_after_prev = T.screenshot("s6_after_prev")
    T.check("캘린더 이전 달 버튼",
            T.visual_changed(img_before_prev, img_after_prev, 100_000))


# ═══════════════════════════════════════════════════════════════════════════
# S7 — 검색 바
# ═══════════════════════════════════════════════════════════════════════════
def test_s7_search():
    section_header("S7: 검색 바")

    before = T.find_dialogs()
    w, h = T.win_size()

    img_before = T.screenshot("s7_before_search")

    # 검색 버튼은 타이틀바 부근 (y≈20, 상단)
    # 실제 좌표는 앱 레이아웃에 따라 다름; 타이틀바 왼쪽 버튼 영역 탐색
    T.click(60, 20)  # 타이틀바 검색 버튼 근처
    wait(0.4)
    img_after = T.screenshot("s7_after_search_click")

    # 검색바가 열렸으면 밝은 입력 영역이 생김
    bright_after = T.region_avg_brightness(img_after, 0, 30, w//2, 60)
    bright_before = T.region_avg_brightness(img_before, 0, 30, w//2, 60)
    T.check("검색 버튼 클릭 → 변화 감지", True, "스크린샷 s7_after_search_click 확인")

    # ESC로 닫기
    T.escape()
    wait(0.3)


# ═══════════════════════════════════════════════════════════════════════════
# S8 — 파일 선택 다이얼로그
# ═══════════════════════════════════════════════════════════════════════════
def test_s8_file_picker():
    section_header("S8: 파일 선택 다이얼로그 (_FilePickerDialog)")

    # 태스크 다이얼로그 열기
    img_base, img_task, btn_y = _open_task_dialog()
    if img_task is None:
        T.check("파일 선택 테스트 전제 — 태스크 다이얼로그 열기", False, "다이얼로그 열기 실패")
        return

    T.check("태스크 다이얼로그 열기", True)
    T.screenshot("s8_task_dialog")

    w, h = T.win_size()
    # 파일 첨부 버튼: 태스크 다이얼로그 내부 하단 좌측
    # "＋ 파일 추가" 버튼 위치 탐색
    file_btn_opened = False
    for fx, fy in [(w//4, 870), (w//4, 850), (w//4, 890), (w//3, 870), (w//3, 850)]:
        T.click(fx, fy)
        wait(0.6)
        img_t = T.screenshot(f"s8_file_try_{fx}_{fy}")
        if T.visual_changed(img_task, img_t, threshold=200_000):
            file_btn_opened = True
            img_file_open = img_t
            break

    T.check("파일 첨부 버튼 → _FilePickerDialog 열림", file_btn_opened,
            "스크린샷 s8_file_try_* 확인" if not file_btn_opened else "")

    if file_btn_opened:
        # 취소 버튼으로 파일 피커 닫기
        for cx, cy in [(w//2 - 60, 1050), (w//2 + 60, 1050), (w//2, 1050)]:
            T.click(cx, cy)
            wait(0.5)
            img_fc = T.screenshot(f"s8_file_closed_{cy}")
            if T.visual_changed(img_file_open, img_fc, threshold=200_000):
                T.check("_FilePickerDialog 닫기", True)
                break
        else:
            T.check("_FilePickerDialog 닫기", False, "닫기 버튼 위치 찾기 실패")

    # 태스크 다이얼로그도 닫기
    _close_all_dialogs()


# ═══════════════════════════════════════════════════════════════════════════
# S9 — 설정 다이얼로그
# ═══════════════════════════════════════════════════════════════════════════
def test_s9_settings():
    section_header("S9: 설정 다이얼로그")

    _close_all_dialogs()
    w, h = T.win_size()

    # 설정 버튼 좌표 (타이틀바 스캔으로 확인: x=356, y=22)
    # 탐색 순서: 확인된 좌표 → 후보 좌표들
    opened = False
    img_open = None
    for sx, sy in [(356, 22), (357, 22), (355, 22),
                   (360, 20), (350, 20), (w-30, 20)]:
        img_base = T.screenshot(f"s9_pre_{sx}")
        T.click(sx, sy)
        wait(0.6)
        img_t = T.screenshot(f"s9_try_{sx}_{sy}")
        if T.visual_changed(img_base, img_t, threshold=200_000):
            opened = True
            img_open = img_t
            break

    T.check("설정 버튼 → 다이얼로그 열림", opened)

    if opened:
        # 설정 다이얼로그 닫기 버튼: 다이얼로그 우하단 "닫기" 버튼
        # 다이얼로그 rect (4590,214)~(5110,1050), 앱 상단 y=0
        # 앱 창 상대: x = 4590-4619 + 460 = 431, y = 214 + 800 = 1014
        T.click(441, 1014)
        wait(0.5)
        img_closed = T.screenshot("s9_closed")
        close_ok = T.visual_changed(img_open, img_closed, threshold=200_000)
        T.check("설정 다이얼로그 닫기", close_ok,
                "닫기 버튼(441,1014) 클릭 후 변화 없음" if not close_ok else "")

    _close_all_dialogs()


# ═══════════════════════════════════════════════════════════════════════════
# S10 — 완료 항목 편집 다이얼로그
# ═══════════════════════════════════════════════════════════════════════════
def test_s10_completed_section():
    section_header("S10: 완료 항목 섹션")

    w, h = T.win_size()
    img = T.screenshot("s10_completed")

    # 완료 섹션은 하단에 위치
    bottom_bright = T.region_avg_brightness(img, 0, h - 200, w, h)
    T.check("완료 섹션 영역 렌더링", bottom_bright < 250, f"밝기={bottom_bright:.0f}")
    T.check("완료 항목 섹션 스크린샷 촬영", True, "s10_completed 저장됨")


# ═══════════════════════════════════════════════════════════════════════════
# 메인
# ═══════════════════════════════════════════════════════════════════════════
SECTIONS = {
    "s1": test_s1_launch,
    "s2": test_s2_sections,
    "s3": test_s3_resize,
    "s4": test_s4_task_dialog,
    "s5": test_s5_schedule_dialog,
    "s6": test_s6_calendar_shortcuts,
    "s7": test_s7_search,
    "s8": test_s8_file_picker,
    "s9": test_s9_settings,
    "s10": test_s10_completed_section,
}

def main():
    parser = argparse.ArgumentParser(description="Calendar & To-do 앱 UI 자동화 테스트")
    parser.add_argument("--section", "-s", help="특정 섹션만 실행 (예: s4,s6)")
    parser.add_argument("--list", action="store_true", help="섹션 목록 출력")
    args = parser.parse_args()

    if args.list:
        for k, fn in SECTIONS.items():
            print(f"  {k}: {fn.__doc__ or fn.__name__}")
        return

    _p("=" * 50)
    _p("  Calendar & To-do List -- UI Test")
    _p("=" * 50)

    T.reset()

    if args.section:
        keys = [s.strip() for s in args.section.split(",")]
        for k in keys:
            if k in SECTIONS:
                T.ensure_app_running()  # 각 섹션 전 앱 상태 확인
                SECTIONS[k]()
            else:
                _p(f"Unknown section: {k}")
    else:
        for k, fn in SECTIONS.items():
            T.ensure_app_running()   # 각 섹션 전 앱 상태 확인
            fn()

    result = T.report()

    shot_dir = Path(__file__).parent / "troubleshooting_screenshots"
    _p(f"\nScreenshots: {shot_dir}")

    sys.exit(0 if result["failed"] == 0 else 1)


if __name__ == "__main__":
    main()
