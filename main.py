"""
Calendar and To do list v2 — 단일 파일 (all-in-one)

Update works 폴더의 날짜별 .txt 파일을 읽어
할 일 목록, 긴급 업무, 기타, 개인업무 4개 섹션을 자동/수동으로 갱신.

파일 형식 (Update works/YYYY.MM.DD.txt):
  [과제 및 To do list]
  N. 제목
  \t내용: ...
  \t목표: ...
  \t마감기한: ...

  [이번주/차주 긴급 업무]
  N. 내용

  [기타]
  N. 제목
  자유 텍스트 (들여쓰기 없음)

  [개인업무] ← 선택 항목, 위젯에서 직접 입력 가능
"""

# ═══════════════════════════════════════════════════════════════════════════
# 1. IMPORTS
# ═══════════════════════════════════════════════════════════════════════════
import sys, os, re, calendar, sqlite3, shutil, logging
import urllib.request, urllib.error
from datetime import date, datetime, timedelta
from pathlib import Path
from logging.handlers import RotatingFileHandler

from PySide6.QtWidgets import (
    QApplication, QWidget, QDialog,
    QVBoxLayout, QHBoxLayout, QGridLayout,
    QScrollArea, QFrame, QSizePolicy, QSplitter,
    QLabel, QLineEdit, QTextEdit, QPlainTextEdit,
    QCheckBox, QPushButton, QComboBox, QProgressBar,
    QDateEdit, QMessageBox, QMenu, QSystemTrayIcon,
    QTabWidget, QSlider, QSpinBox, QFileDialog,
)
from PySide6.QtCore import (
    Qt, Signal, QPoint, QSettings,
    QFileSystemWatcher, QTimer, QDate, QLockFile, QEvent,
)
from PySide6.QtGui import (
    QFont, QKeySequence, QShortcut,
    QIcon, QPixmap, QColor, QPainter, QAction,
    QDrag,
)
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import QUrl
from PySide6.QtCore import QMimeData


# ═══════════════════════════════════════════════════════════════════════════
# 2. CONSTANTS & PATHS
# ═══════════════════════════════════════════════════════════════════════════

APP_VERSION      = "v3.28"
APP_VERSION_DATE = "2026-04-16"

def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller
    (internal resources inside EXE in onefile mode)
    """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# 1) 외부 폴더 (사용자가 직접 수정하는 폴더 - EXE와 같은 상위 폴더에 위치)
if getattr(sys, 'frozen', False):
    # PyInstaller 배포 시: EXE 파일이 있는 폴더 기준
    BASE_DIR = Path(sys.executable).parent
else:
    # 개발 환경: 소스 파일이 있는 폴더 기준
    BASE_DIR = Path(__file__).parent

ASSETS_DIR      = BASE_DIR / "assets"
EXPORT_DIR      = BASE_DIR / "Export"
ATTACHMENTS_DIR = Path.home() / ".productivity_widget" / "attachments"
BACKUPS_DIR     = Path.home() / ".productivity_widget" / "backups"

# 2) 내부 리소스 (아이콘 등 EXE 안에 박혀있는 파일들 - 나중에 필요 시 사용)
# INTERNAL_ASSETS = Path(resource_path("assets"))

WINDOW_WIDTH    = 720
WINDOW_HEIGHT   = 1440

PRIORITY_HIGH   = 1
PRIORITY_MEDIUM = 2
PRIORITY_LOW    = 3

PRIORITY_LABELS = {1: "높음", 2: "보통", 3: "낮음"}
PRIORITY_COLORS = {1: "#f38ba8", 2: "#fab387", 3: "#a6e3a1"}

TASK_TODO     = "todo"
TASK_URGENT   = "urgent"
TASK_MISC     = "misc"
TASK_PERSONAL = "personal"

SOURCE_FILE   = "file"    # txt 파일에서 로드된 항목
SOURCE_MANUAL = "manual"  # 위젯에서 직접 입력된 항목

C = {
    "base":     "#1e1e2e", "mantle":   "#181825", "crust":    "#11111b",
    "surface0": "#313244", "surface1": "#45475a", "surface2": "#585b70",
    "overlay0": "#6c7086", "overlay1": "#7f849c",
    "text":     "#cdd6f4", "subtext":  "#a6adc8",
    "blue":     "#89b4fa", "lavender": "#b4befe",
    "green":    "#a6e3a1", "yellow":   "#f9e2af",
    "red":      "#f38ba8", "peach":    "#fab387",
    "teal":     "#94e2d5", "sky":      "#89dceb",
    "mauve":    "#cba6f7",
}

# ── 태스크 색상 프리셋 (todo / personal 태스크용) ─────────────────────────
TASK_COLORS = [
    None,       # 기본 (우선순위 색상)
    "#f38ba8",  # 빨강
    "#fab387",  # 주황
    "#f9e2af",  # 노랑
    "#a6e3a1",  # 초록
    "#89dceb",  # 하늘
    "#89b4fa",  # 파랑
    "#cba6f7",  # 보라
    "#f5c2e7",  # 분홍
    "#94e2d5",  # 민트
]

# ── UI 테마 정의 ──────────────────────────────────────────────────────────
THEMES: dict[str, dict] = {
    "dark": {
        "name": "다크 (기본)",
        "base": "#1e1e2e", "mantle": "#181825",
        "surface0": "#313244", "surface1": "#45475a",
        "overlay0": "#6c7086", "overlay1": "#7f849c",   # 추가: 보조 텍스트 계층
        "text": "#cdd6f4", "subtext": "#a6adc8",
        "blue": "#89b4fa", "task_bg": "#27273a",
        "task_hover": "#38386a",   # hover delta 강화: 1.1:1 → ~1.4:1
        "input_bg": "#313244", "border": "#45475a",
    },
    "black": {
        "name": "블랙",
        "base": "#0a0a0a", "mantle": "#111111",
        "surface0": "#1c1c1c", "surface1": "#252525",
        "overlay0": "#777777", "overlay1": "#8a8a8a",
        "text": "#e8e8e8", "subtext": "#999999",
        "blue": "#5f9cf5", "task_bg": "#181818",
        "task_hover": "#2e2e2e",   # hover delta 강화
        "input_bg": "#1c1c1c", "border": "#333333",
    },
    "latte": {
        "name": "라떼 (라이트)",
        "base": "#eff1f5", "mantle": "#e6e9ef",
        "surface0": "#ccd0da", "surface1": "#bcc0cc",
        "overlay0": "#9ca0b0", "overlay1": "#8c8fa1",
        "text": "#4c4f69", "subtext": "#6c6f85",
        "blue": "#1e66f5", "task_bg": "#e6e9ef",
        "task_hover": "#dce0e8",
        "input_bg": "#eff1f5", "border": "#bcc0cc",
    },
    "navy": {
        "name": "네이비",
        "base": "#0f172a", "mantle": "#0a1020",
        "surface0": "#1e293b", "surface1": "#273448",
        "overlay0": "#6b7fa0", "overlay1": "#7a8fab",
        "text": "#e2e8f0", "subtext": "#94a3b8",
        "blue": "#60a5fa", "task_bg": "#1a2535",
        "task_hover": "#2a3f5e",   # hover delta 강화
        "input_bg": "#1e293b", "border": "#334155",
    },
    "gruvbox": {
        "name": "그루박스",
        "base": "#282828", "mantle": "#1d2021",
        "surface0": "#3c3836", "surface1": "#504945",
        "overlay0": "#928374", "overlay1": "#a89984",
        "text": "#ebdbb2", "subtext": "#d5c4a1",
        "blue": "#83a598", "task_bg": "#32302f",
        "task_hover": "#45403d",
        "input_bg": "#3c3836", "border": "#504945",
    },
    "tokyo": {
        "name": "도쿄 나이트",
        "base": "#1a1b26", "mantle": "#13141f",
        "surface0": "#24283b", "surface1": "#414868",
        "overlay0": "#565f89", "overlay1": "#6b7db3",
        "text": "#c0caf5", "subtext": "#a9b1d6",
        "blue": "#7aa2f7", "task_bg": "#1f2335",
        "task_hover": "#2d3149",
        "input_bg": "#24283b", "border": "#414868",
    },
}


# ── 일정 유형 ──────────────────────────────────────────────────────────────
SCHED_SINGLE   = "schedule"   # 단기 일정
SCHED_VACATION = "vacation"   # 연차/휴가
SCHED_TRAINING = "training"   # 교육
SCHED_TRIP     = "trip"       # 출장

SCHED_LABELS = {
    SCHED_SINGLE:   "단기 일정",
    SCHED_VACATION: "연차/휴가",
    SCHED_TRAINING: "교육",
    SCHED_TRIP:     "출장",
}
SCHED_ICONS  = {
    SCHED_SINGLE:   "📅",
    SCHED_VACATION: "🏖",
    SCHED_TRAINING: "📚",
    SCHED_TRIP:     "✈️",
}
SCHED_COLORS = {
    SCHED_SINGLE:   "#89b4fa",   # 파랑
    SCHED_VACATION: "#f9e2af",   # 노랑
    SCHED_TRAINING: "#a6e3a1",   # 초록
    SCHED_TRIP:     "#fab387",   # 주황
}
# 달력 셀 하단 도트 색상 (배경 틴트 대신 도트 방식 사용)
PERSONAL_CAL_COLOR  = "#cba6f7"  # 보라 (개인업무)
KAKAOWORK_CAL_COLOR = "#89dceb"  # 하늘 (카카오워크 팀 일정)

# ── iCal 이벤트 표시 색상 ──────────────────────────────────────────────────
ICAL_COLOR_BOSS     = "#89b4fa"  # 소장님일정 — 밝은 파랑 (최상단 강조, 배경에 묻히지 않음)
ICAL_COLOR_VACATION = "#a6e3a1"  # 연차/반차/예비군/휴가 — 초록
ICAL_COLOR_ALLDAY   = "#fab387"  # 종일 기타 (출장/교육/파견 등) — 주황
ICAL_COLOR_TIMED    = "#ffe5b4"  # 시간 지정 — 파스텔 노랑

ICAL_BOSS_KEYWORD = "소장님일정"
ICAL_VACATION_KEYWORDS = ["연차", "예비군", "오후반차", "오전반차", "출산휴가",
                           "경조휴가", "병가", "휴가"]


# ── iCal URL 보호 (base64 obfuscation) ────────────────────────────────────
import base64 as _b64

def _ical_url_encode(url: str) -> str:
    """저장 시 URL을 base64로 인코딩"""
    return _b64.b64encode(url.encode()).decode()

def _ical_url_decode(encoded: str) -> str:
    """읽을 때 base64 디코딩, 실패 시 원본 반환 (구버전 호환)"""
    try:
        return _b64.b64decode(encoded.encode()).decode()
    except Exception:
        return encoded  # 구버전 평문 URL 호환


def _ical_time_label(ev) -> str:
    """시간 지정 이벤트: '시작 - 종료, 장소', 종일: '' (빈 문자열)
    sqlite3.Row 및 dict 모두 지원.
    """
    def _get(k):
        try: return ev[k] or ""
        except (KeyError, IndexError): return ""
    s = _get("start_time_str")
    if not s:
        return ""
    e = _get("end_time_str")
    loc = _get("location")
    time_part = f"{s} - {e}" if e else s
    return f"{time_part}, {loc}" if loc else time_part


def _ical_classify(summary: str, start_time_str) -> tuple[int, str]:
    """iCal 이벤트 분류 → (정렬우선순위, 색상코드)
    0: 소장님일정(최상단, 파랑)  1: 휴가류(초록)
    2: 종일 기타(주황)           3: 시간지정(파스텔노랑)
    """
    if ICAL_BOSS_KEYWORD in summary:
        return (0, ICAL_COLOR_BOSS)
    if start_time_str:
        return (3, ICAL_COLOR_TIMED)
    if any(k in summary for k in ICAL_VACATION_KEYWORDS):
        return (1, ICAL_COLOR_VACATION)
    return (2, ICAL_COLOR_ALLDAY)

# 태스크 아이템 파스텔 색상 (ID 기반 순환 — 모든 섹션 공통)
ITEM_PASTEL_COLORS = [
    "#ff9eb5",  # 코랄 핑크
    "#ffb347",  # 피치 오렌지
    "#ffe066",  # 레몬 옐로
    "#b5ead7",  # 민트 그린
    "#aecbfa",  # 스카이 블루
    "#d4a8e0",  # 라벤더
    "#ffd6a5",  # 살구
    "#9deac1",  # 에메랄드
    "#fdcfe8",  # 베이비 핑크
    "#a0d8f1",  # 파우더 블루
    "#c8f7c5",  # 연초록
    "#e0c3fc",  # 연보라
    "#b3f0e0",  # 민트블루
    "#ffe5b4",  # 버터
]
# 하위 호환
URGENT_PASTEL_COLORS = ITEM_PASTEL_COLORS


def _theme_is_light(base_hex: str) -> bool:
    """배경색 밝기 판별 (luminance > 128 → 라이트)"""
    h = base_hex.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return (0.299 * r + 0.587 * g + 0.114 * b) > 128


def build_theme_qss(theme_key: str, font_size: int = 10, font_family: str = "맑은 고딕") -> str:
    """
    선택된 테마로 QSS 오버라이드 생성.
    style.qss의 하드코딩 색상(다크 전용)을 테마별로 덮어씌움.
    overlay0/overlay1이 없는 테마는 subtext로 폴백.
    """
    t    = THEMES.get(theme_key, THEMES["dark"])
    ov0  = t.get("overlay0", t["subtext"])
    ov1  = t.get("overlay1", t["subtext"])
    is_light = _theme_is_light(t["base"])
    fs   = max(8, min(18, font_size))       # 기본 글자 크기
    fs_s = max(7, fs - 1)                   # 보조 레이블 (작은 글자)
    fs_l = fs + 2                           # 제목/헤더 (큰 글자)
    ff   = font_family or "맑은 고딕"        # 글씨체

    # 라이트 테마: 달력 마감일 버튼 — 배경이 밝으므로 어두운 텍스트 사용
    if is_light:
        cal_deadline_css = f"""
QPushButton#CalDayHasTasks {{
    background: rgba(180,20,60,0.13);
    border: 1px solid rgba(180,20,60,0.40);
    border-radius: 17px;
    font-size: 12px;
    color: #b0142a;
    font-weight: bold;
}}
QPushButton#CalDayHasTasks:hover {{
    background: rgba(180,20,60,0.24);
    color: #8a0e1e;
    border-color: rgba(180,20,60,0.60);
}}
"""
    else:
        # 다크 테마: build_theme_qss 마지막에 명시적 오버라이드 (style.qss rgba 블렌딩 불확실성 제거)
        cal_deadline_css = f"""
QPushButton#CalDayHasTasks {{
    background: rgba(243,139,168,0.32);
    border: 1px solid rgba(243,139,168,0.55);
    border-radius: 17px; font-size: 12px;
    color: #f06080; font-weight: bold;
}}
QPushButton#CalDayHasTasks:hover {{
    background: #6b1529;
    color: #ffffff;
    border-color: #a02040;
}}
"""

    return f"""
* {{ color: {t['text']}; font-family: "{ff}"; }}
QWidget {{ background-color: {t['base']}; }}
QWidget#MainWindow  {{ background-color: {t['base']}; border: 1px solid {t['surface0']}; }}
QWidget#TitleBar    {{ background-color: {t['mantle']}; border-bottom: 1px solid {t['surface0']}; }}
QWidget#CalendarWidget {{ background-color: {t['mantle']}; }}
QWidget#SectionWidget  {{ background-color: {t['mantle']}; border: 1px solid {t['surface0']}; }}
QWidget#SectionHeader  {{ background-color: {t['mantle']}; }}

/* ── 카드 배경 (수정 4: 라이트 테마 충돌 해결) ── */
QFrame#TaskItem          {{ background-color: {t['task_bg']}; border: 1px solid {t['border']}; }}
QFrame#TaskItem:hover    {{ background-color: {t['task_hover']}; border-color: {t['surface1']}; }}
QFrame#TaskItemCompleted {{ background-color: {t['base']}; border: 1px solid {t['surface0']}; }}
QFrame#TaskItemCompleted:hover {{ background-color: {t['task_hover']}; }}
QFrame#MiscItem          {{ background-color: {t['task_bg']}; border: 1px solid {t['border']}; border-left: 3px solid {t['surface1']}; }}
QFrame#MiscItem:hover    {{ background-color: {t['task_hover']}; border-left-color: {t['blue']}; }}
QFrame#LogItem           {{ background: {t['task_bg']}; border: 1px solid {t['border']}; }}
QFrame#TaskInfoBox       {{ background: {t['task_bg']}; border-left: 3px solid {t['blue']}; border-top: 1px solid {t['border']}; border-right: 1px solid {t['border']}; border-bottom: 1px solid {t['border']}; }}
QFrame#ScheduleItem      {{ background-color: {t['task_bg']}; border: 1px solid {t['border']}; }}
QFrame#ScheduleItem:hover {{ background-color: {t['task_hover']}; border-color: {t['surface1']}; }}

QDialog {{ background-color: {t['base']}; border: 1px solid {t['surface1']}; }}
QLineEdit, QTextEdit, QPlainTextEdit {{
    background: {t['input_bg']}; border: 1px solid {t['border']}; color: {t['text']};
}}
QComboBox {{ background: {t['input_bg']}; border: 1px solid {t['border']}; color: {t['text']}; }}
QComboBox QAbstractItemView {{ background: {t['input_bg']}; color: {t['text']}; border: 1px solid {t['border']}; }}
QDateEdit {{ background: {t['input_bg']}; border: 1px solid {t['border']}; color: {t['text']}; }}
QPushButton#PrimaryBtn   {{ background: {t['blue']}; }}
QPushButton#CalTodayBtn  {{ background: {t['surface0']}; color: {t['blue']}; }}
QPushButton#CalNavBtn    {{ background: {t['surface0']}; }}
QPushButton#RefreshBtn   {{ background: {t['surface0']}; color: {t['blue']}; }}

/* ── 텍스트 색상 (수정 2: 테마 연동으로 교체) ── */
QLabel#TitleLabel, QLabel#SectionTitle, QLabel#DialogTitle,
QLabel#TaskTitle, QLabel#LogContent, QLabel#TaskInfoTitle,
QLabel#CalHeaderLabel, QLabel#MiscTitle {{ color: {t['text']}; }}
QLabel#SectionStats      {{ color: {ov0}; }}
QLabel#FormLabel         {{ color: {t['subtext']}; }}
QLabel#CalDow            {{ color: {t['subtext']}; }}
QLabel#TaskInfoDesc, QLabel#MiscContent {{ color: {ov0}; }}
QLabel#SourceBadge       {{ color: {t['subtext']}; background: {t['surface0']}; }}
QLabel#TaskGoal          {{ color: {t['text']}; font-style: italic; }}
QLabel#DueBadgeFuture    {{ color: {t['subtext']}; }}
QLabel#UpdateTime        {{ color: {ov0}; }}
QLabel#VersionLabel      {{ color: {t['subtext']}; background: {t['mantle']}; border: 1px solid {t['surface0']}; }}
QPushButton#AddTaskBtn   {{ color: {t['surface1']}; border: 1px dashed {t['surface0']}; }}

/* ── 수정 3: 완료 태스크 가독성 (~4.5:1) ── */
QLabel#TaskTitleDone     {{ color: {ov1}; }}

/* ── 수정 7: 일정 아이템 보조 텍스트 ── */
QLabel#ScheduleItemName    {{ color: {t['text']}; }}
QLabel#ScheduleItemMeta    {{ color: {t['subtext']}; }}
QLabel#ScheduleItemContent {{ color: {ov1}; }}

QLabel#LogTimestamp {{ color: {t['blue']}; }}
QScrollBar::handle:vertical {{ background: {t['surface1']}; }}
QProgressBar {{ background: {t['surface0']}; }}
QProgressBar::chunk {{ background: {t['blue']}; }}
{cal_deadline_css}
/* ── 동적 글자 크기 오버라이드 ── */
QLabel#TitleLabel, QLabel#SectionTitle, QLabel#DialogTitle,
QLabel#CalHeaderLabel, QLabel#MiscTitle,
QLabel#TaskTitle, QLabel#LogContent, QLabel#TaskInfoTitle,
QLabel#ScheduleItemName, QLabel#EventPopupDate {{ font-size: {fs}pt; }}
QLabel#SectionStats, QLabel#FormLabel, QLabel#CalDow,
QLabel#TaskTitleDone, QLabel#UpdateTime,
QLabel#VersionLabel, QLabel#SourceBadge, QLabel#TaskGoal,
QLabel#ScheduleItemMeta, QLabel#ScheduleItemContent,
QLabel#DueBadgeFuture, QLabel#TaskInfoDesc,
QLabel#MiscContent, QLabel#LogTimestamp {{ font-size: {fs_s}pt; }}
QPushButton#RefreshBtn {{ font-size: {fs_s}pt; }}
QLineEdit, QTextEdit, QPlainTextEdit {{ font-size: {fs}pt; }}
QComboBox, QSpinBox, QDateEdit {{ font-size: {fs}pt; }}
QComboBox#SortCombo {{ font-size: 12px; padding: 2px 8px; }}
"""


# ═══════════════════════════════════════════════════════════════════════════
# 3. LOGGING SETUP
# ═══════════════════════════════════════════════════════════════════════════

def _setup_logging() -> logging.Logger:
    """앱 로거 설정 — ~/.productivity_widget/app.log (최대 1MB × 3개)"""
    log_dir = Path.home() / ".productivity_widget"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "app.log"

    logger = logging.getLogger("CalendarTodo")
    logger.setLevel(logging.DEBUG)

    if not logger.handlers:
        fh = RotatingFileHandler(
            str(log_path),
            maxBytes=1_048_576,   # 1 MB
            backupCount=3,
            encoding="utf-8",
        )
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(logging.Formatter(
            "%(asctime)s [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        ))
        logger.addHandler(fh)

    return logger


# 전역 로거 (모듈 임포트 시 초기화)
_log = _setup_logging()


# ═══════════════════════════════════════════════════════════════════════════
# 4. DATABASE
# ═══════════════════════════════════════════════════════════════════════════

class Database:
    """SQLite 데이터베이스 래퍼 (태스크 + 진행 로그)"""

    def __init__(self):
        app_dir = Path.home() / ".productivity_widget"
        app_dir.mkdir(exist_ok=True)
        self.db_path = app_dir / "tasks.db"
        self.conn = sqlite3.connect(str(self.db_path), check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self._create_tables()
        self._migrate()
        _log.info("Database opened: %s", self.db_path)

    def _create_tables(self):
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS tasks (
                id           INTEGER PRIMARY KEY AUTOINCREMENT,
                title        TEXT    NOT NULL,
                description  TEXT    DEFAULT '',
                goal         TEXT    DEFAULT '',
                task_type    TEXT    NOT NULL,
                priority     INTEGER DEFAULT 2,
                due_date     TEXT    DEFAULT NULL,
                is_completed INTEGER DEFAULT 0,
                created_at   TEXT    NOT NULL,
                completed_at TEXT    DEFAULT NULL,
                sort_order   INTEGER DEFAULT 0,
                source          TEXT    DEFAULT 'manual',
                color           TEXT    DEFAULT NULL,
                file_path       TEXT    DEFAULT NULL,
                is_user_deleted INTEGER DEFAULT 0,
                linked_todo_id  INTEGER DEFAULT NULL
            );
            CREATE TABLE IF NOT EXISTS task_files (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id       INTEGER NOT NULL,
                original_path TEXT    NOT NULL,
                copy_path     TEXT    DEFAULT NULL,
                filename      TEXT    NOT NULL,
                added_at      TEXT    NOT NULL,
                FOREIGN KEY (task_id) REFERENCES tasks(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS task_logs (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id    INTEGER NOT NULL,
                content    TEXT    NOT NULL,
                created_at         TEXT    NOT NULL,
                file_path          TEXT    DEFAULT NULL,
                log_type           TEXT    DEFAULT 'general',
                progress_group_id  INTEGER DEFAULT NULL,
                FOREIGN KEY (task_id) REFERENCES tasks(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS db_meta (
                key   TEXT PRIMARY KEY,
                value TEXT
            );
            CREATE TABLE IF NOT EXISTS progress_groups (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id          INTEGER NOT NULL,
                title            TEXT    NOT NULL,
                created_at       TEXT    NOT NULL,
                source_urgent_id INTEGER DEFAULT NULL,
                FOREIGN KEY (task_id) REFERENCES tasks(id) ON DELETE CASCADE
            );
            CREATE TABLE IF NOT EXISTS schedules (
                id         INTEGER PRIMARY KEY AUTOINCREMENT,
                name       TEXT    NOT NULL,
                event_date TEXT    NOT NULL,
                end_date   TEXT    DEFAULT NULL,
                start_time TEXT    DEFAULT NULL,
                location   TEXT    DEFAULT '',
                content    TEXT    DEFAULT '',
                event_type TEXT    DEFAULT 'schedule'
            );
            CREATE TABLE IF NOT EXISTS ical_events (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                uid            TEXT    UNIQUE NOT NULL,
                summary        TEXT    NOT NULL DEFAULT '',
                dtstart        TEXT    NOT NULL,
                dtend          TEXT    DEFAULT NULL,
                start_time_str TEXT    DEFAULT NULL,
                end_time_str   TEXT    DEFAULT NULL,
                location       TEXT    DEFAULT '',
                description    TEXT    DEFAULT '',
                organizer      TEXT    DEFAULT ''
            );
        """)
        self.conn.commit()

    def _migrate(self):
        """기존 DB 스키마 마이그레이션 (컬럼 누락 시 추가)"""
        for col, defn in [("goal",      "TEXT DEFAULT ''"),
                          ("source",    "TEXT DEFAULT 'manual'"),
                          ("color",     "TEXT DEFAULT NULL"),
                          ("file_path", "TEXT DEFAULT NULL")]:
            try:
                self.conn.execute(f"ALTER TABLE tasks ADD COLUMN {col} {defn}")
                self.conn.commit()
            except sqlite3.OperationalError:
                pass  # 이미 존재

        # task_logs file_path 마이그레이션
        try:
            self.conn.execute("ALTER TABLE task_logs ADD COLUMN file_path TEXT DEFAULT NULL")
            self.conn.commit()
        except sqlite3.OperationalError:
            pass

        # ical_events 컬럼 마이그레이션
        for col, defn in [("start_time_str", "TEXT DEFAULT NULL"),
                          ("end_time_str",   "TEXT DEFAULT NULL")]:
            try:
                self.conn.execute(f"ALTER TABLE ical_events ADD COLUMN {col} {defn}")
                self.conn.commit()
            except sqlite3.OperationalError:
                pass

        # is_user_deleted 컬럼 마이그레이션
        try:
            self.conn.execute("ALTER TABLE tasks ADD COLUMN is_user_deleted INTEGER DEFAULT 0")
            self.conn.commit()
        except sqlite3.OperationalError:
            pass

        # linked_todo_id 컬럼 (긴급업무→과제 연결)
        try:
            self.conn.execute("ALTER TABLE tasks ADD COLUMN linked_todo_id INTEGER DEFAULT NULL")
            self.conn.commit()
        except sqlite3.OperationalError:
            pass

        # task_files 테이블 마이그레이션
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS task_files (
                id            INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id       INTEGER NOT NULL,
                original_path TEXT    NOT NULL,
                copy_path     TEXT    DEFAULT NULL,
                filename      TEXT    NOT NULL,
                added_at      TEXT    NOT NULL,
                FOREIGN KEY (task_id) REFERENCES tasks(id) ON DELETE CASCADE
            );
        """)
        self.conn.commit()

        # progress_groups 테이블 마이그레이션
        self.conn.executescript("""
            CREATE TABLE IF NOT EXISTS progress_groups (
                id               INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id          INTEGER NOT NULL,
                title            TEXT    NOT NULL,
                created_at       TEXT    NOT NULL,
                source_urgent_id INTEGER DEFAULT NULL,
                FOREIGN KEY (task_id) REFERENCES tasks(id) ON DELETE CASCADE
            );
        """)
        self.conn.commit()

        # task_logs: log_type / progress_group_id 컬럼 마이그레이션
        for col, defn in [("log_type",          "TEXT DEFAULT 'general'"),
                          ("progress_group_id", "INTEGER DEFAULT NULL")]:
            try:
                self.conn.execute(f"ALTER TABLE task_logs ADD COLUMN {col} {defn}")
                self.conn.commit()
            except sqlite3.OperationalError:
                pass

        # task_type CHECK 제약이 있는 구형 DB → 재생성
        try:
            self.conn.execute(
                "INSERT INTO tasks (title,task_type,created_at) "
                "VALUES ('__probe__','misc','x')"
            )
            self.conn.execute("DELETE FROM tasks WHERE title='__probe__'")
            self.conn.commit()
        except sqlite3.IntegrityError:
            self._recreate_tasks_table()

    def _recreate_tasks_table(self):
        """CHECK 제약 제거를 위해 tasks 테이블 재생성 (데이터 보존)"""
        self.conn.executescript("""
            CREATE TABLE tasks_v2 (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL, description TEXT DEFAULT '',
                goal TEXT DEFAULT '', task_type TEXT NOT NULL,
                priority INTEGER DEFAULT 2, due_date TEXT DEFAULT NULL,
                is_completed INTEGER DEFAULT 0, created_at TEXT NOT NULL,
                completed_at TEXT DEFAULT NULL, sort_order INTEGER DEFAULT 0,
                source TEXT DEFAULT 'manual', color TEXT DEFAULT NULL,
                file_path TEXT DEFAULT NULL,
                is_user_deleted INTEGER DEFAULT 0,
                linked_todo_id INTEGER DEFAULT NULL
            );
            INSERT INTO tasks_v2
                SELECT id,title,description,
                       COALESCE(goal,''), task_type, priority, due_date,
                       is_completed, created_at, completed_at, sort_order,
                       COALESCE(source,'manual'), NULL,
                       COALESCE(is_user_deleted, 0), NULL
                FROM tasks;
            DROP TABLE tasks;
            ALTER TABLE tasks_v2 RENAME TO tasks;
        """)
        self.conn.commit()

    # ── 태스크 CRUD ──────────────────────────────────────────────────────
    def add_task(self, title, description="", goal="", task_type=TASK_TODO,
                 priority=PRIORITY_MEDIUM, due_date=None, source=SOURCE_MANUAL,
                 color=None, file_path=None, linked_todo_id=None) -> int:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = self.conn.execute(
            "INSERT INTO tasks (title,description,goal,task_type,priority,"
            "due_date,created_at,source,color,file_path,linked_todo_id) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
            (title, description, goal, task_type, priority, due_date, now, source, color, file_path, linked_todo_id)
        )
        self.conn.commit()
        return cur.lastrowid

    def get_tasks(self, task_type=None, completed=None) -> list:
        where_parts: list[str] = ["is_user_deleted=0"]
        params: list = []
        if task_type:
            where_parts.append("task_type=?"); params.append(task_type)
        if completed is not None:
            where_parts.append("is_completed=?"); params.append(1 if completed else 0)
        where = f"WHERE {' AND '.join(where_parts)}"
        cur = self.conn.execute(
            f"SELECT * FROM tasks {where} "
            "ORDER BY priority ASC, sort_order ASC, created_at ASC",
            params
        )
        return cur.fetchall()

    def get_task(self, task_id) -> sqlite3.Row:
        cur = self.conn.execute("SELECT * FROM tasks WHERE id=?", (task_id,))
        return cur.fetchone()

    def update_task(self, task_id, **kwargs):
        if not kwargs:
            return
        clause = ", ".join(f"{k}=?" for k in kwargs)
        self.conn.execute(f"UPDATE tasks SET {clause} WHERE id=?",
                          list(kwargs.values()) + [task_id])
        self.conn.commit()

    def toggle_complete(self, task_id, completed: bool):
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S") if completed else None
        self.conn.execute(
            "UPDATE tasks SET is_completed=?, completed_at=? WHERE id=?",
            (1 if completed else 0, now, task_id)
        )
        self.conn.commit()

    def update_sort_order(self, ordered_ids: list):
        """태스크 목록을 순서대로 sort_order 업데이트"""
        for idx, tid in enumerate(ordered_ids):
            self.conn.execute("UPDATE tasks SET sort_order=? WHERE id=?", (idx, tid))
        self.conn.commit()

    def delete_task(self, task_id):
        task = self.get_task(task_id)
        if task and task["source"] == SOURCE_FILE:
            # 파일에서 가져온 항목: soft delete (재삽입 차단)
            self.conn.execute(
                "UPDATE tasks SET is_user_deleted=1 WHERE id=?", (task_id,)
            )
        else:
            # 직접 입력 항목: hard delete (tasks 먼저 → CASCADE로 로그도 삭제)
            self.conn.execute("DELETE FROM tasks WHERE id=?", (task_id,))
        self.conn.commit()

    def get_linked_todo_title(self, urgent_task_id: int) -> str | None:
        """긴급업무의 연결된 과제 제목 반환 (없으면 None)"""
        row = self.conn.execute(
            "SELECT t2.title FROM tasks t1 "
            "JOIN tasks t2 ON t1.linked_todo_id=t2.id "
            "WHERE t1.id=? AND t2.is_user_deleted=0",
            (urgent_task_id,)
        ).fetchone()
        return row[0] if row else None

    def get_task_stats(self, task_type) -> tuple[int, int]:
        total = self.conn.execute(
            "SELECT COUNT(*) FROM tasks WHERE task_type=? AND is_user_deleted=0",
            (task_type,)
        ).fetchone()[0]
        done = self.conn.execute(
            "SELECT COUNT(*) FROM tasks WHERE task_type=? AND is_completed=1 AND is_user_deleted=0",
            (task_type,)
        ).fetchone()[0]
        return total, done

    # ── 파일 가져오기 ─────────────────────────────────────────────────────
    def sync_from_file(self, task_type: str, parsed_tasks: list[dict]) -> int:
        """
        파일에서 DB에 없는 신규 항목만 추가 (단방향 가져오기).
        - 기존 DB 항목은 source 구분 없이 절대 삭제·수정하지 않음
        - 중복 제목이 없는 항목만 INSERT (제목 기준 중복 체크)
        - 한 번 가져온 항목은 이후 사용자가 직접 관리
        """
        # 현재 DB의 모든 제목 — soft-deleted 포함, 완료 여부 무관하게 체크
        all_titles = {
            r["title"] for r in self.conn.execute(
                "SELECT title FROM tasks WHERE task_type=?", (task_type,)
            ).fetchall()
        }
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        inserted = 0
        for t in parsed_tasks:
            if t["title"] not in all_titles:
                self.conn.execute(
                    "INSERT INTO tasks (title,description,goal,task_type,"
                    "priority,due_date,created_at,source) VALUES (?,?,?,?,?,?,?,?)",
                    (t["title"], t["description"], t["goal"], task_type,
                     t["priority"], t["due_date"], now, SOURCE_FILE)
                )
                inserted += 1
        self.conn.commit()
        return inserted

    # ── 로그 CRUD ────────────────────────────────────────────────────────
    def add_log(self, task_id, content, file_path=None) -> int:
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        cur = self.conn.execute(
            "INSERT INTO task_logs (task_id,content,created_at,file_path) VALUES (?,?,?,?)",
            (task_id, content, now, file_path or None)
        )
        self.conn.commit()
        return cur.lastrowid

    def get_logs(self, task_id) -> list:
        cur = self.conn.execute(
            "SELECT * FROM task_logs WHERE task_id=? ORDER BY created_at ASC",
            (task_id,)
        )
        return cur.fetchall()

    def delete_log(self, log_id):
        self.conn.execute("DELETE FROM task_logs WHERE id=?", (log_id,))
        self.conn.commit()

    def update_log(self, log_id, content, file_path=None):
        self.conn.execute(
            "UPDATE task_logs SET content=?, file_path=? WHERE id=?",
            (content, file_path or None, log_id)
        )
        self.conn.commit()

    # ── 첨부 파일 CRUD ───────────────────────────────────────────────────────
    def _ensure_attachments_dir(self):
        ATTACHMENTS_DIR.mkdir(parents=True, exist_ok=True)

    def add_task_file(self, task_id: int, original_path: str) -> int:
        """파일을 attachments 폴더에 복사하고 DB에 기록"""
        self._ensure_attachments_dir()
        src = Path(original_path)
        filename = src.name
        dst = ATTACHMENTS_DIR / f"task_{task_id}_{filename}"
        copy_path = None
        if src.exists():
            try:
                shutil.copy2(str(src), str(dst))
                copy_path = str(dst)
            except Exception:
                pass
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = self.conn.execute(
            "INSERT INTO task_files (task_id, original_path, copy_path, filename, added_at) "
            "VALUES (?,?,?,?,?)",
            (task_id, original_path, copy_path, filename, now)
        )
        self.conn.commit()
        return cur.lastrowid

    def get_task_files(self, task_id: int) -> list:
        return self.conn.execute(
            "SELECT * FROM task_files WHERE task_id=? ORDER BY added_at ASC",
            (task_id,)
        ).fetchall()

    def delete_task_file(self, file_id: int):
        """DB에서 삭제 + copy_path 파일도 삭제"""
        row = self.conn.execute(
            "SELECT copy_path FROM task_files WHERE id=?", (file_id,)
        ).fetchone()
        if row and row["copy_path"]:
            try:
                Path(row["copy_path"]).unlink(missing_ok=True)
            except Exception:
                pass
        self.conn.execute("DELETE FROM task_files WHERE id=?", (file_id,))
        self.conn.commit()

    def get_missing_file_tasks(self) -> list[dict]:
        """original_path가 존재하지 않는 파일 목록 반환"""
        rows = self.conn.execute(
            "SELECT tf.id, tf.task_id, tf.filename, tf.original_path, tf.copy_path, "
            "t.title as task_title "
            "FROM task_files tf JOIN tasks t ON tf.task_id=t.id "
            "WHERE t.is_user_deleted=0 AND t.is_completed=0"
        ).fetchall()
        missing = []
        for r in rows:
            if not Path(r["original_path"]).exists():
                missing.append(dict(r))
        return missing

    # ── 진행 그룹 CRUD ───────────────────────────────────────────────────────
    def add_progress_group(self, task_id: int, title: str,
                           source_urgent_id: int | None = None) -> int:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = self.conn.execute(
            "INSERT INTO progress_groups (task_id,title,created_at,source_urgent_id) "
            "VALUES (?,?,?,?)",
            (task_id, title, now, source_urgent_id)
        )
        self.conn.commit()
        return cur.lastrowid

    def get_progress_groups(self, task_id: int) -> list:
        return self.conn.execute(
            "SELECT * FROM progress_groups WHERE task_id=? ORDER BY created_at ASC",
            (task_id,)
        ).fetchall()

    def delete_progress_group(self, group_id: int):
        """그룹 삭제 + 소속 로그 항목 삭제"""
        self.conn.execute(
            "DELETE FROM task_logs WHERE progress_group_id=?", (group_id,)
        )
        self.conn.execute(
            "DELETE FROM progress_groups WHERE id=?", (group_id,)
        )
        self.conn.commit()

    def get_progress_logs(self, group_id: int) -> list:
        return self.conn.execute(
            "SELECT * FROM task_logs WHERE progress_group_id=? ORDER BY created_at ASC",
            (group_id,)
        ).fetchall()

    def add_progress_log(self, task_id: int, group_id: int, content: str) -> int:
        now = datetime.now().strftime("%Y-%m-%d %H:%M")
        cur = self.conn.execute(
            "INSERT INTO task_logs (task_id,content,created_at,log_type,progress_group_id) "
            "VALUES (?,?,?,'progress',?)",
            (task_id, content, now, group_id)
        )
        self.conn.commit()
        return cur.lastrowid

    def update_progress_group_title(self, group_id: int, title: str):
        self.conn.execute(
            "UPDATE progress_groups SET title=? WHERE id=?", (title, group_id)
        )
        self.conn.commit()

    def update_progress_log(self, log_id: int, content: str):
        self.conn.execute(
            "UPDATE task_logs SET content=? WHERE id=?", (content, log_id)
        )
        self.conn.commit()

    def delete_progress_log(self, log_id: int):
        self.conn.execute("DELETE FROM task_logs WHERE id=?", (log_id,))
        self.conn.commit()

    def get_general_logs(self, task_id: int) -> list:
        """일반 로그만 반환 (log_type='general' 또는 NULL)"""
        return self.conn.execute(
            "SELECT * FROM task_logs WHERE task_id=? "
            "AND (log_type='general' OR log_type IS NULL) "
            "AND progress_group_id IS NULL "
            "ORDER BY created_at ASC",
            (task_id,)
        ).fetchall()

    def get_urgent_progress_groups(self, urgent_task_id: int) -> list:
        """긴급업무 완료 시 생성된 진행 그룹 목록"""
        return self.conn.execute(
            "SELECT * FROM progress_groups WHERE source_urgent_id=?",
            (urgent_task_id,)
        ).fetchall()

    # ── 일정 CRUD ────────────────────────────────────────────────────────────
    def add_schedule(self, name, event_date, end_date=None, start_time=None,
                     location='', content='', event_type=SCHED_SINGLE) -> int:
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur = self.conn.execute(
            "INSERT INTO schedules (name,event_date,end_date,start_time,location,content,event_type)"
            " VALUES (?,?,?,?,?,?,?)",
            (name, event_date, end_date, start_time, location, content, event_type)
        )
        self.conn.commit()
        return cur.lastrowid

    def get_schedules(self) -> list:
        return self.conn.execute(
            "SELECT * FROM schedules ORDER BY event_date ASC"
        ).fetchall()

    def get_schedule_by_id(self, sched_id) -> sqlite3.Row:
        cur = self.conn.execute("SELECT * FROM schedules WHERE id=?", (sched_id,))
        return cur.fetchone()

    def get_schedules_for_date(self, date_str: str) -> list:
        """단일 날짜에 해당하는 일정 반환 (범위 포함)"""
        return self.conn.execute(
            "SELECT * FROM schedules WHERE event_date=? "
            "OR (end_date IS NOT NULL AND event_date<=? AND end_date>=?)"
            " ORDER BY start_time ASC",
            (date_str, date_str, date_str)
        ).fetchall()

    def get_schedule_date_map(self) -> dict:
        """date_str → list[row] 매핑 (달력 표시용)"""
        rows = self.get_schedules()
        result: dict[str, list] = {}
        today = date.today()
        for r in rows:
            try:
                start = date.fromisoformat(r["event_date"])
                end   = date.fromisoformat(r["end_date"]) if r["end_date"] else start
            except (ValueError, TypeError):
                continue
            d = start
            while d <= end:
                ds = d.isoformat()
                result.setdefault(ds, []).append(r)
                d += timedelta(days=1)
        return result

    def update_schedule(self, sched_id, **kwargs):
        if not kwargs: return
        clause = ", ".join(f"{k}=?" for k in kwargs)
        self.conn.execute(f"UPDATE schedules SET {clause} WHERE id=?",
                          list(kwargs.values()) + [sched_id])
        self.conn.commit()

    def delete_schedule(self, sched_id):
        self.conn.execute("DELETE FROM schedules WHERE id=?", (sched_id,))
        self.conn.commit()

    # ── iCal 이벤트 CRUD ─────────────────────────────────────────────────────
    def sync_ical_events(self, events: list[dict]) -> int:
        """iCal 이벤트 upsert (uid 기준) — 기존 이벤트 업데이트, 신규 삽입"""
        inserted = 0
        for e in events:
            uid = e.get("uid", "").strip()
            if not uid or not e.get("dtstart"):
                continue
            existing = self.conn.execute(
                "SELECT id FROM ical_events WHERE uid=?", (uid,)
            ).fetchone()
            vals = (e.get("summary",""), e.get("dtstart",""), e.get("dtend"),
                    e.get("start_time_str"), e.get("end_time_str"),
                    e.get("location",""), e.get("description",""), e.get("organizer",""))
            if existing:
                self.conn.execute(
                    "UPDATE ical_events SET summary=?,dtstart=?,dtend=?,"
                    "start_time_str=?,end_time_str=?,location=?,description=?,organizer=? WHERE uid=?",
                    vals + (uid,)
                )
            else:
                self.conn.execute(
                    "INSERT INTO ical_events "
                    "(uid,summary,dtstart,dtend,start_time_str,end_time_str,"
                    "location,description,organizer) "
                    "VALUES (?,?,?,?,?,?,?,?,?)",
                    (uid,) + vals
                )
                inserted += 1
        self.conn.commit()
        return inserted

    def clear_ical_events(self):
        self.conn.execute("DELETE FROM ical_events")
        self.conn.commit()

    def get_ical_events(self) -> list:
        return self.conn.execute(
            "SELECT * FROM ical_events ORDER BY dtstart ASC"
        ).fetchall()

    def get_ical_date_map(self) -> dict:
        """date_str → list[row] 매핑 (달력 표시용, iCal DTEND exclusive 처리)"""
        result: dict[str, list] = {}
        for r in self.get_ical_events():
            try:
                start = date.fromisoformat(r["dtstart"])
                if r["dtend"]:
                    end = date.fromisoformat(r["dtend"])
                    # iCal all-day events: DTEND is exclusive → subtract 1 day
                    if end > start:
                        end = end - timedelta(days=1)
                else:
                    end = start
            except (ValueError, TypeError):
                continue
            d = start
            while d <= end:
                result.setdefault(d.isoformat(), []).append(r)
                d += timedelta(days=1)
        return result

    def close(self):
        if self.conn:
            self.conn.close()
            _log.info("Database closed.")


# ═══════════════════════════════════════════════════════════════════════════
# 4. FILE PARSER
# ═══════════════════════════════════════════════════════════════════════════

def _parse_korean_date(text: str) -> str | None:
    """
    한국어 날짜/기한 → YYYY-MM-DD 변환.
    지원: "7월 31일", "4월말", "4월 중순", "4월 초", "5-6월 내"
    불가: "미정", "확인 필요", "상시대기" → None 반환
    """
    if not text:
        return None
    t = text.strip()

    # 불확실 키워드 → None
    if any(k in t for k in ["미정", "확인", "상시", "TBD", "tbd", "결정"]):
        return None

    today = date.today()
    yr = today.year

    # "N월말" / "N월 말"
    m = re.search(r'(\d+)월\s*말', t)
    if m:
        mo = int(m.group(1))
        if mo < today.month:
            yr += 1
        last = calendar.monthrange(yr, mo)[1]
        return f"{yr}-{mo:02d}-{last:02d}"

    # "N월 중순"
    m = re.search(r'(\d+)월\s*중순', t)
    if m:
        mo = int(m.group(1))
        if mo < today.month:
            yr += 1
        return f"{yr}-{mo:02d}-20"

    # "N월 초"
    m = re.search(r'(\d+)월\s*초', t)
    if m:
        mo = int(m.group(1))
        if mo < today.month:
            yr += 1
        return f"{yr}-{mo:02d}-10"

    # "N-M월" 범위 → 뒤 달 말
    m = re.search(r'(\d+)-(\d+)월', t)
    if m:
        mo = int(m.group(2))
        if mo < today.month:
            yr += 1
        last = calendar.monthrange(yr, mo)[1]
        return f"{yr}-{mo:02d}-{last:02d}"

    # "N월 N일"
    m = re.search(r'(\d+)월\s*(\d+)일', t)
    if m:
        mo, dy = int(m.group(1)), int(m.group(2))
        try:
            d = date(yr, mo, dy)
            if d < today:
                d = date(yr + 1, mo, dy)
            return d.isoformat()
        except ValueError:
            return None

    return None


# 4-B. ICAL PARSER
# ═══════════════════════════════════════════════════════════════════════════

class ICalParser:
    """
    RFC 5545 iCal (.ics) 파서 — 외부 라이브러리 없이 표준 라이브러리만 사용.
    VEVENT 블록에서 필요한 필드(UID, SUMMARY, DTSTART, DTEND, LOCATION,
    DESCRIPTION, ORGANIZER)만 추출.
    """

    def parse(self, text: str) -> list[dict]:
        """iCal 텍스트 → event dict 목록 반환"""
        # CRLF 정규화 + 줄 접기(line folding) 해제
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        unfolded: list[str] = []
        for line in text.split("\n"):
            if line and line[0] in (" ", "\t") and unfolded:
                unfolded[-1] += line[1:]
            else:
                unfolded.append(line)

        events: list[dict] = []
        current: dict | None = None

        for line in unfolded:
            line = line.strip()
            if line == "BEGIN:VEVENT":
                current = {"uid": "", "summary": "", "dtstart": "",
                           "dtend": None, "start_time_str": None, "end_time_str": None,
                           "location": "", "description": "", "organizer": ""}
            elif line == "END:VEVENT":
                if current and current["dtstart"]:
                    events.append(current)
                current = None
            elif current is not None and ":" in line:
                colon_idx = line.index(":")
                key_full  = line[:colon_idx]
                value     = line[colon_idx + 1:]
                key       = key_full.split(";")[0].upper()

                if key == "UID":
                    current["uid"] = value.strip()
                elif key == "SUMMARY":
                    current["summary"] = self._unescape(value)
                elif key == "DTSTART":
                    d_str, t_str = self._parse_datetime(value)
                    current["dtstart"] = d_str or ""
                    current["start_time_str"] = t_str  # None for all-day
                elif key == "DTEND":
                    d_str, t_str = self._parse_datetime(value)
                    current["dtend"] = d_str
                    current["end_time_str"] = t_str  # None for all-day
                elif key == "LOCATION":
                    current["location"] = self._unescape(value)
                elif key == "DESCRIPTION":
                    current["description"] = self._unescape(value)
                elif key == "ORGANIZER":
                    cn_m = re.search(r'CN=([^;:]+)', key_full)
                    if cn_m:
                        current["organizer"] = cn_m.group(1).strip()

        return events

    @staticmethod
    def _parse_datetime(value: str) -> tuple[str | None, str | None]:
        """iCal date/datetime → (YYYY-MM-DD, HH:MM) 또는 (YYYY-MM-DD, None) for all-day.
        UTC(Z suffix) 또는 오프셋(+HHMM/-HHMM) 포함 시 로컬 타임존으로 변환."""
        import datetime as _dt_mod
        value = value.strip()
        if "T" in value:
            t_idx = value.index("T")
            date_part = value[:t_idx]
            time_part = value[t_idx + 1:]
            if len(date_part) == 8 and date_part.isdigit():
                digits = time_part[:6]
                if len(digits) == 6 and digits.isdigit():
                    hour, minute, second = int(digits[0:2]), int(digits[2:4]), int(digits[4:6])
                    suffix = time_part[6:]
                    try:
                        if suffix == "Z":
                            # UTC → 로컬 변환
                            aware = _dt_mod.datetime(
                                int(date_part[:4]), int(date_part[4:6]), int(date_part[6:8]),
                                hour, minute, second, tzinfo=_dt_mod.timezone.utc)
                            local = aware.astimezone()
                            return local.strftime("%Y-%m-%d"), local.strftime("%H:%M")
                        elif suffix and suffix[0] in ('+', '-'):
                            # 오프셋 (+0900, -0500 등)
                            sign   = 1 if suffix[0] == '+' else -1
                            off    = suffix[1:].replace(":", "")
                            if len(off) >= 4 and off[:4].isdigit():
                                off_td = _dt_mod.timedelta(
                                    hours=sign * int(off[:2]),
                                    minutes=sign * int(off[2:4]))
                                tz     = _dt_mod.timezone(off_td)
                                aware  = _dt_mod.datetime(
                                    int(date_part[:4]), int(date_part[4:6]), int(date_part[6:8]),
                                    hour, minute, second, tzinfo=tz)
                                local  = aware.astimezone()
                                return local.strftime("%Y-%m-%d"), local.strftime("%H:%M")
                    except Exception:
                        pass
                    # 타임존 없는 로컬 시간
                    date_str = f"{date_part[0:4]}-{date_part[4:6]}-{date_part[6:8]}"
                    return date_str, f"{hour:02d}:{minute:02d}"
                else:
                    date_str = f"{date_part[0:4]}-{date_part[4:6]}-{date_part[6:8]}"
                    return date_str, None
        elif len(value) == 8 and value.isdigit():
            return f"{value[0:4]}-{value[4:6]}-{value[6:8]}", None
        return None, None

    @staticmethod
    def _parse_date(value: str) -> str | None:
        """하위 호환용 — date만 반환"""
        v = value.strip()
        if len(v) >= 8 and v[:8].isdigit():
            return f"{v[0:4]}-{v[4:6]}-{v[6:8]}"
        return None

    @staticmethod
    def _unescape(value: str) -> str:
        return (value.replace("\\n", "\n").replace("\\,", ",")
                     .replace("\\;", ";").replace("\\\\", "\\"))


# ═══════════════════════════════════════════════════════════════════════════
# 5. HELPER FUNCTIONS
# ═══════════════════════════════════════════════════════════════════════════

def due_badge(due_date_str: str | None) -> tuple[str, str] | None:
    """마감일 뱃지 (텍스트, QSS objectName) 반환. 없으면 None."""
    if not due_date_str:
        return None
    try:
        due  = date.fromisoformat(due_date_str)
    except ValueError:
        return None
    diff = (due - date.today()).days
    if diff < 0:
        return f"D+{abs(diff)}일 초과", "DueBadgeOverdue"
    elif diff == 0:
        return "오늘 마감", "DueBadgeToday"
    elif diff == 1:
        return "내일 마감", "DueBadgeSoon"
    elif diff <= 7:
        return f"D-{diff}일", "DueBadgeNormal"
    else:
        return due.strftime("%m/%d"), "DueBadgeFuture"


def open_file_path(path: str, parent=None):
    """파일 경로 열기 — 파일이면 탐색기에서 선택, 폴더면 폴더 열기"""
    import subprocess
    if not path:
        return
    if not os.path.exists(path):
        QMessageBox.warning(parent, "경로 오류",
                            f"경로를 찾을 수 없습니다:\n{path}")
        return
    if os.path.isdir(path):
        os.startfile(path)
    else:
        subprocess.Popen(f'explorer /select,"{os.path.normpath(path)}"')


# ═══════════════════════════════════════════════════════════════════════════
# 6. CALENDAR WIDGET
# ═══════════════════════════════════════════════════════════════════════════

_DEADLINE_BTN_HOVER_QSS = (
    "QPushButton {"
    " background: #6b1529;"
    " color: #ffffff;"
    " border: 1px solid #a02040;"
    " border-radius: 17px;"
    " font-size: 12px;"
    " font-weight: bold;"
    "}"
)

class CalDayButton(QPushButton):
    """달력 날짜 버튼 — 호버/더블클릭/이벤트 도트 + 기간 바 표시"""
    hovered        = Signal(object, object)  # date, QPoint (global topleft)
    unhovered      = Signal()
    double_clicked = Signal(object)          # date

    def __init__(self, day: int, d: date,
                 event_types: list | None = None,
                 has_personal: bool = False,
                 has_deadline: bool = False,
                 has_cowork: bool = False,
                 period_bars: list | None = None,
                 parent=None):
        super().__init__(str(day), parent)
        self._date         = d
        self._event_types  = event_types or []
        self._has_personal = has_personal
        self._has_deadline = has_deadline
        self._has_cowork   = has_cowork
        # period_bars: list of (color, pos) where pos in ("start","middle","end","single")
        self._period_bars  = period_bars or []

    def enterEvent(self, e):
        super().enterEvent(e)
        # QSS :hover cascade 의존 대신 직접 stylesheet 적용 (플랫폼 호환성)
        if self._has_deadline and self.objectName() == "CalDayHasTasks":
            self.setStyleSheet(_DEADLINE_BTN_HOVER_QSS)
        self.hovered.emit(self._date, self.mapToGlobal(self.rect().topLeft()))

    def leaveEvent(self, e):
        super().leaveEvent(e)
        if self._has_deadline and self.objectName() == "CalDayHasTasks":
            self.setStyleSheet("")  # 개별 스타일 제거 → 앱 전역 QSS 복원
        self.unhovered.emit()

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.double_clicked.emit(self._date)
        super().mouseDoubleClickEvent(e)

    def paintEvent(self, e):
        super().paintEvent(e)
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # ── 기간 바 (개인 일정 기간 표시) ───────────────────────────
        bar_h   = 5
        bar_gap = 2
        bar_y_start = 22
        w = self.width()
        # 투명도: start/end=220(강조), middle=140(이어짐), single=230
        _bar_alpha = {"start": 220, "end": 220, "middle": 140, "single": 230}
        for i, (color, pos) in enumerate(self._period_bars[:4]):
            y = bar_y_start + i * (bar_h + bar_gap)
            c = QColor(color)
            c.setAlpha(_bar_alpha.get(pos, 180))
            painter.setBrush(c)
            painter.setPen(Qt.PenStyle.NoPen)
            if pos == "start":
                # 오른쪽 절반부터 시작 (시작점 강조)
                painter.drawRoundedRect(w // 2, y, w // 2 - 1, bar_h, bar_h // 2, bar_h // 2)
                painter.drawRect(w // 2, y, w // 4, bar_h)      # 오른쪽 평탄화
            elif pos == "end":
                # 왼쪽 절반까지 (끝점 강조)
                painter.drawRoundedRect(1, y, w // 2, bar_h, bar_h // 2, bar_h // 2)
                painter.drawRect(w // 4, y, w // 4, bar_h)      # 왼쪽 평탄화
            elif pos == "single":
                painter.drawRoundedRect(4, y, w - 8, bar_h, bar_h // 2, bar_h // 2)
            else:  # middle
                painter.drawRect(0, y, w, bar_h)                # 양끝 이음새 없이 전체 채움

        # ── 원형 이벤트 도트 (기간 바로 표시된 타입 제외) ────────
        dots = []
        for et in [SCHED_TRIP, SCHED_VACATION, SCHED_TRAINING, SCHED_SINGLE]:
            if et in self._event_types:
                dots.append(SCHED_COLORS[et])
        if self._has_personal and not self._period_bars:
            # 기간 바가 없을 때만 개인 일정 도트 표시 (바와 중복 방지)
            dots.append(PERSONAL_CAL_COLOR)
        if self._has_cowork:
            dots.append(KAKAOWORK_CAL_COLOR)

        if dots:
            d = 6          # 원형 지름
            gap = 3
            total = len(dots) * d + (len(dots) - 1) * gap
            x = (self.width() - total) // 2
            y = self.height() - 8
            for color in dots:
                painter.setBrush(QColor(color))
                painter.setPen(Qt.PenStyle.NoPen)
                painter.drawEllipse(x, y, d, d)
                x += d + gap

        painter.end()


class EventPopup(QWidget):
    """달력 날짜 호버 시 캘린더 왼편에 떠오르는 일정 상세 팝업 (테마 동적 적용)"""

    def __init__(self, parent=None):
        super().__init__(parent, Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setObjectName("EventPopup")
        self.setMinimumWidth(340)
        self.setMaximumWidth(520)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 12, 14, 12)
        lay.setSpacing(8)
        self._lay = lay
        self._hide_timer = QTimer(self)
        self._hide_timer.setSingleShot(True)
        self._hide_timer.setInterval(150)
        self._hide_timer.timeout.connect(self.hide)
        # 테마 색상 캐시 (기본 다크)
        self._tc: dict = {}
        self.apply_theme("dark")

    def enterEvent(self, e):
        super().enterEvent(e)
        self._hide_timer.stop()  # 팝업 위에 마우스 올리면 사라지지 않음

    def leaveEvent(self, e):
        super().leaveEvent(e)
        self._hide_timer.start()  # 팝업 벗어나면 150ms 후 숨김

    @staticmethod
    def _is_light(hex_color: str) -> bool:
        """배경색이 밝은 계열인지 판단 (밝기 0~255, 128 기준)"""
        h = hex_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return (0.299 * r + 0.587 * g + 0.114 * b) > 128

    def apply_theme(self, theme_key: str):
        """테마 변경 시 호출 — 팝업 배경·텍스트 색 업데이트"""
        t = THEMES.get(theme_key, THEMES["dark"])
        light = self._is_light(t["base"])
        ov0   = t.get("overlay0", t["subtext"])

        if light:
            # 라이트 계열: 팝업도 밝은 배경
            popup_bg  = t["mantle"]        # base보다 약간 어두운 배경
            border_c  = t["surface1"]
            card_bg   = t["base"]
            card_bord = t["surface0"]
            hdr_color = t["blue"]
            sep_color = t["surface1"]
            meta_c    = t["subtext"]
            cont_c    = ov0
        else:
            # 다크 계열: 기존 스타일 유지
            popup_bg  = t["mantle"]
            border_c  = t.get("surface1", "#3d3d58")
            card_bg   = t.get("surface0", "#252535")
            card_bord = t.get("surface0", "#2e2e48")
            hdr_color = t["blue"]
            sep_color = t.get("surface0", "#2e2e48")
            meta_c    = ov0
            cont_c    = t["subtext"]

        self._tc = {
            "popup_bg": popup_bg, "border": border_c,
            "card_bg":  card_bg,  "card_bord": card_bord,
            "hdr":      hdr_color, "sep": sep_color,
            "meta":     meta_c,   "cont": cont_c,
            "text":     t["text"],
        }
        self.setStyleSheet(
            f"QWidget#EventPopup{{background:{popup_bg};border:1px solid {border_c};border-radius:12px;}}"
            f"QFrame#EventPopupCard{{border-radius:8px;}}"
            f"QLabel{{background:transparent;color:{t['text']};}}"
        )

    def show_for(self, d: date, events: list, personal_tasks: list, btn_global_pos,
                 deadline_tasks=None, ical_events=None):
        """일정+개인업무 목록으로 팝업 갱신 후 달력 왼쪽에 표시"""
        tc = self._tc
        while self._lay.count():
            item = self._lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        # 날짜 헤더
        hdr = QLabel(d.strftime("%Y년 %m월 %d일 (%a)"))
        hdr.setObjectName("EventPopupDate")
        hdr.setStyleSheet(f"color:{tc['hdr']};font-weight:bold;font-size:12px;background:transparent;")
        self._lay.addWidget(hdr)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet(f"background:{tc['sep']};max-height:1px;border:none;")
        self._lay.addWidget(sep)

        for ev in events:
            card = QFrame(); card.setObjectName("EventPopupCard")
            cl = QVBoxLayout(card)
            cl.setContentsMargins(10, 8, 10, 8); cl.setSpacing(3)

            etype = ev["event_type"]
            icon  = SCHED_ICONS.get(etype, "📅")
            color = SCHED_COLORS.get(etype, tc['hdr'])

            name_lbl = QLabel(f"{icon} {ev['name']}")
            name_lbl.setStyleSheet(f"color:{color};font-weight:bold;font-size:12px;background:transparent;")
            name_lbl.setWordWrap(True)
            cl.addWidget(name_lbl)

            if ev["start_time"]:
                tl = QLabel(f"🕐 {ev['start_time']}")
                tl.setStyleSheet(f"color:{tc['meta']};font-size:11px;background:transparent;")
                cl.addWidget(tl)
            if ev["location"]:
                ll = QLabel(f"📍 {ev['location']}")
                ll.setStyleSheet(f"color:{tc['meta']};font-size:11px;background:transparent;")
                cl.addWidget(ll)
            if ev["end_date"] and ev["end_date"] != ev["event_date"]:
                dl = QLabel(f"📆 {ev['event_date']} ~ {ev['end_date']}")
                dl.setStyleSheet(f"color:{tc['meta']};font-size:11px;background:transparent;")
                cl.addWidget(dl)
            if ev["content"]:
                cont = QLabel(ev["content"])
                cont.setStyleSheet(f"color:{tc['cont']};font-size:11px;background:transparent;")
                cont.setWordWrap(True); cl.addWidget(cont)

            card.setStyleSheet(
                f"QFrame#EventPopupCard{{background:{tc['card_bg']};border-radius:8px;"
                f"border-left:3px solid {color};border-top:1px solid {tc['card_bord']};"
                f"border-right:1px solid {tc['card_bord']};border-bottom:1px solid {tc['card_bord']};}}"
            )
            self._lay.addWidget(card)

        # 개인업무 카드
        for task in personal_tasks:
            card = QFrame(); card.setObjectName("EventPopupCard")
            cl = QVBoxLayout(card)
            cl.setContentsMargins(10, 8, 10, 8); cl.setSpacing(2)

            name_lbl = QLabel(f"👤 {task['title']}")
            name_lbl.setStyleSheet(f"color:#cba6f7;font-weight:bold;font-size:12px;background:transparent;")
            name_lbl.setWordWrap(True)
            cl.addWidget(name_lbl)
            if task["description"]:
                d_lbl = QLabel(task["description"])
                d_lbl.setStyleSheet(f"color:{tc['meta']};font-size:11px;background:transparent;")
                d_lbl.setWordWrap(True); cl.addWidget(d_lbl)

            card.setStyleSheet(
                f"QFrame#EventPopupCard{{background:{tc['card_bg']};border-radius:8px;"
                f"border-left:3px solid #cba6f7;border-top:1px solid {tc['card_bord']};"
                f"border-right:1px solid {tc['card_bord']};border-bottom:1px solid {tc['card_bord']};}}"
            )
            self._lay.addWidget(card)

        # 마감 태스크 카드: deadline_map(dict) 교체로 태스크 상세 정보를 받아 렌더링
        for task in (deadline_tasks or []):
            card = QFrame(); card.setObjectName("EventPopupCard")
            cl = QVBoxLayout(card)
            cl.setContentsMargins(10, 8, 10, 8); cl.setSpacing(2)

            # 우선순위별 색상 (높음=빨강, 보통=주황, 낮음=초록)
            prio_colors = {1: "#f38ba8", 2: "#fab387", 3: "#a6e3a1"}
            pcolor = prio_colors.get(task["priority"], "#f38ba8")

            # 제목
            name_lbl = QLabel(f"📌 {task['title']}")
            name_lbl.setStyleSheet(
                f"color:{pcolor};font-weight:bold;font-size:12px;background:transparent;")
            name_lbl.setWordWrap(True)
            cl.addWidget(name_lbl)

            # 섹션 타입 + 마감일
            type_labels = {TASK_TODO: "할 일", TASK_URGENT: "긴급 업무"}
            type_str = type_labels.get(task["task_type"], task["task_type"])
            tl = QLabel(f"[{type_str}]  마감: {task['due_date']}")
            tl.setStyleSheet(
                f"color:{tc['meta']};font-size:11px;background:transparent;")
            cl.addWidget(tl)

            # 목표 (있을 때만) — sqlite3.Row는 .get() 없으므로 직접 접근
            if task["goal"]:
                gl = QLabel(f"▸ {task['goal']}")
                gl.setStyleSheet(
                    f"color:{tc['cont']};font-size:11px;background:transparent;")
                gl.setWordWrap(True)
                cl.addWidget(gl)

            card.setStyleSheet(
                f"QFrame#EventPopupCard{{background:{tc['card_bg']};border-radius:8px;"
                f"border-left:3px solid {pcolor};border-top:1px solid {tc['card_bord']};"
                f"border-right:1px solid {tc['card_bord']};border-bottom:1px solid {tc['card_bord']};}}"
            )
            self._lay.addWidget(card)

        # 카카오워크 팀 일정 카드 (소장님→휴가→종일→시간 순)
        def _ical_sort_key(ev):
            pri, _ = _ical_classify(ev["summary"], ev["start_time_str"])
            return (pri, ev["start_time_str"] or "")
        for ev in sorted(ical_events or [], key=_ical_sort_key):
            summary  = ev["summary"]
            pri, color = _ical_classify(summary, ev["start_time_str"])

            card = QFrame(); card.setObjectName("EventPopupCard")
            cl = QVBoxLayout(card)
            cl.setContentsMargins(10, 7, 10, 7); cl.setSpacing(2)

            name_lbl = QLabel(f"🏢 {summary}")
            name_lbl.setStyleSheet(
                f"color:{color};font-weight:bold;font-size:12px;background:transparent;")
            name_lbl.setWordWrap(True)
            cl.addWidget(name_lbl)

            # 시간 지정 → 시작-종료 (장소), 종일 → 장소만 (있으면)
            time_label = _ical_time_label(ev)
            if time_label:
                tl = QLabel(f"🕐 {time_label}")
                tl.setStyleSheet(f"color:{tc['meta']};font-size:11px;background:transparent;")
                cl.addWidget(tl)
            elif ev.get("location"):
                ll = QLabel(f"📍 {ev['location']}")
                ll.setStyleSheet(f"color:{tc['meta']};font-size:11px;background:transparent;")
                cl.addWidget(ll)

            card.setStyleSheet(
                f"QFrame#EventPopupCard{{background:{tc['card_bg']};border-radius:8px;"
                f"border-left:3px solid {color};"
                f"border-top:1px solid {tc['card_bord']};"
                f"border-right:1px solid {tc['card_bord']};"
                f"border-bottom:1px solid {tc['card_bord']};}}"
            )
            self._lay.addWidget(card)

        self.adjustSize()

        from PySide6.QtGui import QGuiApplication
        from PySide6.QtCore import QPoint as _QPoint
        tgt_screen = QGuiApplication.screenAt(btn_global_pos)
        screen_geo = (tgt_screen if tgt_screen else QGuiApplication.primaryScreen()).availableGeometry()

        # 팝업은 항상 메인 창 왼쪽 밖에 표시 — 달력 셀을 가리지 않음
        parent_win = self.parent()
        if parent_win is not None:
            win_left  = parent_win.mapToGlobal(_QPoint(0, 0)).x()
            win_right = parent_win.mapToGlobal(_QPoint(parent_win.width(), 0)).x()
            x = win_left - self.width() - 8
            if x < screen_geo.left():
                # 왼쪽 공간 부족 → 메인 창 오른쪽 밖
                x = win_right + 8
        else:
            x = btn_global_pos.x() - self.width() - 14
            if x < screen_geo.left():
                x = btn_global_pos.x() + 40

        # Y: 호버 버튼 위치 기준, 화면 아래 넘치면 위로 밀어올림
        y = btn_global_pos.y()
        if y + self.height() > screen_geo.bottom():
            y = screen_geo.bottom() - self.height() - 10
        if x + self.width() > screen_geo.right():
            x = screen_geo.right() - self.width() - 4

        self._hide_timer.stop()
        self.move(x, y)
        self.show()
        self.raise_()  # 메인 윈도우 위로 올림 (Windows z-order 보정)

    def schedule_hide(self):
        self._hide_timer.start()

    def cancel_hide(self):
        self._hide_timer.stop()


class _CoworkResizeHandle(QFrame):
    """연구실 일정 패널 하단 드래그 핸들 — 위아래로 드래그해 패널 높이 조절"""
    def __init__(self, target: QWidget, parent=None):
        super().__init__(parent)
        self._target = target
        self._drag_y: int | None = None
        self._start_h: int = 0
        self.setFixedHeight(8)
        self.setCursor(Qt.CursorShape.SizeVerCursor)
        self.setToolTip("드래그하여 높이 조절")
        self.setStyleSheet(
            "QFrame{background:#2e2e48;border-radius:4px;margin:0 6px;}"
            "QFrame:hover{background:#45475a;}"
        )

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self._drag_y  = e.globalPosition().toPoint().y()
            self._start_h = self._target.height()

    def mouseMoveEvent(self, e):
        if self._drag_y is not None:
            dy    = e.globalPosition().toPoint().y() - self._drag_y
            new_h = max(40, min(500, self._start_h + dy))
            self._target.setFixedHeight(new_h)

    def mouseReleaseEvent(self, e):
        self._drag_y = None


class CalendarWidget(QWidget):
    date_selected              = Signal(date)
    add_schedule_requested     = Signal(date)   # 일정 추가 요청
    add_personal_task_requested = Signal(date)  # 개인업무 추가 요청

    def __init__(self, db: Database, parent=None):
        super().__init__(parent)
        self.db = db
        self.setObjectName("CalendarWidget")
        today = date.today()
        self._year, self._month = today.year, today.month
        self._today    = today
        self._selected = today
        # _deadline_dates(set) → _deadline_map(dict): hover 시 태스크 목록 전달을 위해 교체
        self._deadline_map: dict[str, list] = {}     # due_date → [task row, ...]
        self._personal_map: dict[str, list] = {}    # date_str → [personal task rows]
        self._sched_map: dict[str, list] = {}        # date_str → [schedule rows]
        self._ical_map: dict[str, list] = {}         # date_str → [ical_event rows]
        self._show_cowork: bool = True               # Co-work 모드 표시 여부
        self._day_buttons: list[CalDayButton] = []
        self._popup = EventPopup(self)  # parent 전달 → 메인 윈도우 위 z-order 보장
        self._setup_ui()
        self.refresh()
        # 키보드 단축키 (Ctrl+Left/Right: 월 이동, Home: 오늘)
        QShortcut(QKeySequence("Ctrl+Left"),  self, self._prev,     context=Qt.ShortcutContext.WidgetWithChildrenShortcut)
        QShortcut(QKeySequence("Ctrl+Right"), self, self._next,     context=Qt.ShortcutContext.WidgetWithChildrenShortcut)
        QShortcut(QKeySequence("Home"),       self, self._goto_today, context=Qt.ShortcutContext.WidgetWithChildrenShortcut)

    def _setup_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 10, 12, 12)
        lay.setSpacing(8)

        # 네비게이션 헤더
        hdr = QHBoxLayout()
        self.btn_prev = QPushButton("◀")
        self.btn_prev.setObjectName("CalNavBtn")
        self.btn_prev.setFixedSize(30, 30)
        self.btn_prev.clicked.connect(self._prev)
        hdr.addWidget(self.btn_prev)

        self.lbl_ym = QLabel()
        self.lbl_ym.setObjectName("CalHeaderLabel")
        self.lbl_ym.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_ym.setFont(QFont("맑은 고딕", 13, QFont.Weight.Bold))
        self.lbl_ym.setToolTip("날짜를 더블클릭하면 일정을 추가할 수 있습니다\n우클릭하면 일정·개인업무 추가 옵션이 나타납니다")
        hdr.addWidget(self.lbl_ym, 1)

        self.btn_next = QPushButton("▶")
        self.btn_next.setObjectName("CalNavBtn")
        self.btn_next.setFixedSize(30, 30)
        self.btn_next.clicked.connect(self._next)
        hdr.addWidget(self.btn_next)

        self.btn_today = QPushButton("오늘")
        self.btn_today.setObjectName("CalTodayBtn")
        self.btn_today.setFixedHeight(30)
        self.btn_today.clicked.connect(self._goto_today)
        hdr.addWidget(self.btn_today)

        self.btn_cowork = QPushButton("🏢")
        self.btn_cowork.setObjectName("CalTodayBtn")
        self.btn_cowork.setFixedSize(30, 30)
        self.btn_cowork.setCheckable(True)
        self.btn_cowork.setChecked(True)
        self.btn_cowork.setToolTip("Co-work 모드 (카카오워크 팀 일정 표시 토글)")
        self.btn_cowork.toggled.connect(self._toggle_cowork)
        hdr.addWidget(self.btn_cowork)
        lay.addLayout(hdr)

        # 요일 헤더
        dow_grid = QGridLayout()
        dow_grid.setSpacing(2)
        for col, (name, obj) in enumerate(zip(
            ["일","월","화","수","목","금","토"],
            ["CalDowSun"]+["CalDow"]*5+["CalDowSat"]
        )):
            l = QLabel(name)
            l.setObjectName(obj)
            l.setAlignment(Qt.AlignmentFlag.AlignCenter)
            l.setFixedHeight(22)
            dow_grid.addWidget(l, 0, col)
        lay.addLayout(dow_grid)

        self.day_grid = QGridLayout()
        self.day_grid.setSpacing(2)
        lay.addLayout(self.day_grid)

        # 캘린더 사용 힌트
        _cal_hint = QLabel("💡 날짜 우클릭 → 일정 추가  |  날짜 클릭 → 해당일 항목 강조")
        _cal_hint.setAlignment(Qt.AlignmentFlag.AlignRight)
        _cal_hint.setStyleSheet("color:#7f849c;font-size:11px;padding:1px 4px 3px 0;")
        lay.addWidget(_cal_hint)

        # ── 팀 일정 패널 (날짜 클릭 시 표시) ────────────────────────────
        self._cowork_sep = QFrame()
        self._cowork_sep.setFrameShape(QFrame.Shape.HLine)
        self._cowork_sep.setMaximumHeight(1)
        self._cowork_sep.setStyleSheet("background:#2e2e48;border:none;")
        self._cowork_sep.hide()
        lay.addWidget(self._cowork_sep)

        self._cowork_scroll = QScrollArea()
        self._cowork_scroll.setWidgetResizable(True)
        self._cowork_scroll.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self._cowork_scroll.setStyleSheet(
            "QScrollArea{border:none;background:transparent;}")
        self._cowork_cont = QWidget()
        self._cowork_cont.setStyleSheet("background:transparent;")
        self._cowork_inner = QVBoxLayout(self._cowork_cont)
        self._cowork_inner.setContentsMargins(0, 4, 0, 4)
        self._cowork_inner.setSpacing(3)
        self._cowork_scroll.setWidget(self._cowork_cont)
        self._cowork_scroll.hide()
        lay.addWidget(self._cowork_scroll)

        # 드래그 핸들 (숨김 상태로 시작)
        self._cowork_handle = _CoworkResizeHandle(self._cowork_scroll)
        self._cowork_handle.hide()
        lay.addWidget(self._cowork_handle)

    def _build(self):
        for b in self._day_buttons:
            self.day_grid.removeWidget(b)
            b.deleteLater()
        self._day_buttons.clear()
        self.lbl_ym.setText(f"{self._year}년  {self._month}월")
        first     = date(self._year, self._month, 1)
        start_col = (first.weekday() + 1) % 7
        days      = calendar.monthrange(self._year, self._month)[1]
        row, col  = 0, start_col
        for d_num in range(1, days + 1):
            d  = date(self._year, self._month, d_num)
            ds = d.isoformat()
            sched_types  = [r["event_type"] for r in self._sched_map.get(ds, [])]
            has_deadline = ds in self._deadline_map
            has_personal = ds in self._personal_map
            has_cowork   = self._show_cowork and bool(self._ical_map.get(ds))

            # 기간 일정 바 계산 (개인 직접 추가 일정만, iCal 제외)
            period_bars = []
            period_bar_types = set()
            for r in self._sched_map.get(ds, []):
                if not r["end_date"] or r["end_date"] == r["event_date"]:
                    continue  # 단일일 일정은 도트로만 표시
                color = SCHED_COLORS.get(r["event_type"], "#89b4fa")
                try:
                    s = date.fromisoformat(r["event_date"])
                    en = date.fromisoformat(r["end_date"])
                except (ValueError, TypeError):
                    continue
                if d == s and d == en:
                    pos = "single"
                elif d == s:
                    pos = "start"
                elif d == en:
                    pos = "end"
                else:
                    pos = "middle"
                period_bars.append((color, pos))
                period_bar_types.add(r["event_type"])

            # 기간 바로 이미 표시되는 이벤트 타입은 도트에서 제외 (중복 방지)
            dot_sched_types = [et for et in sched_types if et not in period_bar_types]

            btn = CalDayButton(d_num, d,
                               event_types=list(set(dot_sched_types)),
                               has_personal=has_personal,
                               has_deadline=has_deadline,
                               has_cowork=has_cowork,
                               period_bars=period_bars)
            btn.setFixedSize(42, 64)   # 셀 크기 고정 (텍스트+바+도트 여유 확보)
            btn.clicked.connect(lambda _, dt=d: self._click(dt))
            btn.double_clicked.connect(lambda _, dt=d: self.add_schedule_requested.emit(dt))
            btn.hovered.connect(self._on_hover)
            btn.unhovered.connect(self._popup.schedule_hide)

            # objectName 결정
            if d == self._today:
                obj = "CalDayToday"
            elif d == self._selected and d != self._today:
                obj = "CalDaySelected"
            elif has_deadline:
                obj = "CalDayHasTasks"
            elif col == 0:
                obj = "CalDaySun"
            elif col == 6:
                obj = "CalDaySat"
            else:
                obj = "CalDay"
            btn.setObjectName(obj)

            # 툴팁 제거: EventPopup이 마감 태스크도 직접 렌더링하므로 시스템 툴팁 불필요
            self.day_grid.addWidget(btn, row, col)
            self._day_buttons.append(btn)
            col += 1
            if col > 6:
                col = 0; row += 1

    def _on_hover(self, d: date, global_pos):
        """날짜 호버 — 일정/개인업무/마감 태스크/카카오워크 이벤트 있으면 팝업 표시"""
        ds       = d.isoformat()
        events   = self._sched_map.get(ds, [])
        personal = self._personal_map.get(ds, [])
        deadline = self._deadline_map.get(ds, [])
        ical     = self._ical_map.get(ds, []) if self._show_cowork else []
        if events or personal or deadline or ical:
            self._popup.show_for(d, events, personal, global_pos,
                                 deadline_tasks=deadline, ical_events=ical)
        else:
            self._popup.schedule_hide()

    def _auto_fit_cowork_height(self):
        """컨텐츠 실제 높이 기반으로 스크롤 영역 높이 자동 설정 (드래그 미사용 시)"""
        self._cowork_cont.adjustSize()
        content_h = self._cowork_cont.sizeHint().height()
        fit_h = max(40, min(160, content_h + 8))
        self._cowork_scroll.setFixedHeight(fit_h)

    def _toggle_cowork(self, on: bool):
        self._show_cowork = on
        self._build()
        self._update_cowork_panel(self._selected)

    def _update_cowork_panel(self, d: date | None):
        """선택된 날짜의 카카오워크 팀 일정을 달력 하단 패널에 표시"""
        # 기존 내용 초기화
        while self._cowork_inner.count():
            item = self._cowork_inner.takeAt(0)
            if item.widget(): item.widget().deleteLater()

        if not d or not self._show_cowork:
            self._cowork_scroll.hide(); self._cowork_sep.hide()
            self._cowork_handle.hide(); return

        events = self._ical_map.get(d.isoformat(), [])
        if not events:
            self._cowork_scroll.hide(); self._cowork_sep.hide()
            self._cowork_handle.hide(); return

        # 정렬: 종일(시간없음) 상단 → 시간순
        def sort_key(ev):
            t = ev["start_time_str"] or ""
            return (1 if t else 0, t)
        events = sorted(events, key=sort_key)

        # 헤더
        hdr = QLabel(f"🏢  {d.strftime('%m/%d')} 연구실 일정  ({len(events)}건)")
        hdr.setStyleSheet(
            f"font-weight:bold;font-size:11px;color:{KAKAOWORK_CAL_COLOR};"
            "padding:2px 4px;background:transparent;")
        self._cowork_inner.addWidget(hdr)

        def _sort_key(ev):
            pri, _ = _ical_classify(ev["summary"], ev["start_time_str"])
            return (pri, ev["start_time_str"] or "")

        for ev in sorted(events, key=_sort_key):
            summary    = ev["summary"]
            pri, color = _ical_classify(summary, ev["start_time_str"])
            time_label = _ical_time_label(ev)

            row_w = QWidget()
            row_w.setStyleSheet("QWidget{background:transparent;margin:1px 0;}")
            row_l = QHBoxLayout(row_w)
            row_l.setContentsMargins(2, 1, 4, 1); row_l.setSpacing(6)

            bar = QFrame()
            bar.setFrameShape(QFrame.Shape.VLine)
            bar.setFixedWidth(3)
            bar.setStyleSheet(f"background:{color};border:none;")
            row_l.addWidget(bar)

            txt = summary if not time_label else f"{time_label}  {summary}"
            lbl = QLabel(txt)
            lbl.setStyleSheet(f"color:{color};font-size:11px;font-weight:bold;background:transparent;")
            lbl.setWordWrap(True)
            lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            row_l.addWidget(lbl, 1)

            self._cowork_inner.addWidget(row_w)

        self._cowork_inner.addStretch()
        self._cowork_sep.show()
        self._cowork_scroll.show()
        self._cowork_handle.show()
        # 컨텐츠 높이에 맞게 자동 조절 (최대 160px, 최소 40px)
        QTimer.singleShot(0, self._auto_fit_cowork_height)

    def _prev(self):
        if self._month == 1: self._month = 12; self._year -= 1
        else: self._month -= 1
        self._selected = None   # 다른 달 이동 시 선택 초기화 → CoworkPanel 자동 숨김
        self._build()

    def _next(self):
        if self._month == 12: self._month = 1; self._year += 1
        else: self._month += 1
        self._selected = None   # 다른 달 이동 시 선택 초기화 → CoworkPanel 자동 숨김
        self._build()

    def _goto_today(self):
        t = date.today()
        self._year, self._month, self._selected = t.year, t.month, t
        self._build()

    def _click(self, d):
        self._selected = d
        self._build()
        self._update_cowork_panel(d)
        self.date_selected.emit(d)
        # 더블클릭 효과: 일정 추가는 ScheduleSection에서 버튼으로 처리
        # 단, 캘린더를 우클릭하면 일정 추가 메뉴 표시

    def contextMenuEvent(self, e):
        """달력 우클릭 → 커서 아래 날짜에 일정 또는 개인업무 추가"""
        # 커서 바로 아래 CalDayButton 탐지 (없으면 선택된 날짜로 fallback)
        child = self.childAt(e.pos())
        target_date = None
        if isinstance(child, CalDayButton):
            target_date = child._date
        elif child is not None and isinstance(child.parent(), CalDayButton):
            target_date = child.parent()._date
        if target_date is None:
            target_date = self._selected
        if target_date is None:
            return
        d_str = target_date.strftime("%m/%d")
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu{background:#313244;border:1px solid #45475a;border-radius:8px;padding:4px;}"
            "QMenu::item{padding:7px 18px;border-radius:6px;color:#cdd6f4;}"
            "QMenu::item:selected{background:#45475a;}"
        )
        a_sched    = menu.addAction(f"📅  {d_str} 단기 일정 추가")
        a_vacation = menu.addAction(f"🏖  {d_str} 휴가 / 교육 추가")
        menu.addSeparator()
        a_personal = menu.addAction(f"👤  {d_str} 개인업무 추가")
        ch = menu.exec(e.globalPos())
        if ch == a_sched:
            self.add_schedule_requested.emit(target_date)
        elif ch == a_vacation:
            self._vacation_req = target_date
            self.add_schedule_requested.emit(target_date)   # ScheduleSection이 vacation 타입 처리
        elif ch == a_personal:
            self.add_personal_task_requested.emit(target_date)

    def refresh(self):
        tasks = self.db.get_tasks()
        # set → dict 교체: hover 팝업에 태스크 상세 정보를 전달하기 위해
        self._deadline_map = {}
        for t in tasks:
            if t["due_date"] and not t["is_completed"] \
               and t["task_type"] in (TASK_TODO, TASK_URGENT):
                self._deadline_map.setdefault(t["due_date"], []).append(t)
        # 개인업무: 날짜별 task 리스트로 저장
        self._personal_map = {}
        for t in tasks:
            if t["due_date"] and not t["is_completed"] and t["task_type"] == TASK_PERSONAL:
                self._personal_map.setdefault(t["due_date"], []).append(t)
        self._sched_map = self.db.get_schedule_date_map()
        self._ical_map  = self.db.get_ical_date_map()
        self._build()
        self._update_cowork_panel(self._selected)


# ═══════════════════════════════════════════════════════════════════════════
# 7. TASK ITEM WIDGET (checkbox 있는 일반 태스크)
# ═══════════════════════════════════════════════════════════════════════════

class TaskItemWidget(QFrame):
    toggled              = Signal(int, bool)
    delete_requested     = Signal(int)
    log_requested        = Signal(int)
    edit_requested       = Signal(int)
    navigate_requested   = Signal(int)
    batch_select_changed = Signal(int, bool)

    def __init__(self, task_row, highlight: bool = False, parent=None,
                 linked_title: str | None = None, file_count: int = 0):
        super().__init__(parent)
        self._id           = task_row["id"]
        self._completed    = bool(task_row["is_completed"])
        self._file_path    = task_row["file_path"] if task_row["file_path"] else None
        self._linked_title = linked_title
        self._file_count   = file_count
        self._batch_chk: QCheckBox | None = None
        obj = "TaskItemCompleted" if self._completed else "TaskItem"
        self.setObjectName(obj)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(46)
        if highlight:
            self.setStyleSheet(
                f"QFrame#{obj}{{border:2px solid #89b4fa;border-radius:8px;}}"
            )
        self._build(task_row)

    def _build(self, r):
        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 8, 0)
        outer.setSpacing(0)

        # 왼쪽 컬러 바 (파스텔 색상 — ID 기반 순환 배정)
        bar = QFrame()
        bar.setFixedWidth(7)
        task_color = r["color"] if r["color"] else None
        if self._completed:
            clr = "#353548"
        elif task_color:
            clr = task_color
        else:
            # 모든 타입: ID 기반 파스텔 순환
            clr = ITEM_PASTEL_COLORS[r["id"] % len(ITEM_PASTEL_COLORS)]
        bar.setStyleSheet(f"background:{clr};border-radius:3px;margin:5px 0 5px 5px;")
        outer.addWidget(bar)

        inner = QHBoxLayout()
        inner.setContentsMargins(8, 6, 0, 6)
        inner.setSpacing(8)

        # 배치 선택 체크박스 (기본 숨김)
        self._batch_chk = QCheckBox()
        self._batch_chk.setObjectName("TaskCheck")
        self._batch_chk.setFixedSize(22, 22)
        self._batch_chk.setVisible(False)
        self._batch_chk.toggled.connect(lambda v: self.batch_select_changed.emit(self._id, v))
        inner.addWidget(self._batch_chk, 0, Qt.AlignmentFlag.AlignVCenter)

        # 체크박스
        self.chk = QCheckBox()
        self.chk.setObjectName("TaskCheck")
        self.chk.setChecked(self._completed)
        self.chk.setFixedSize(22, 22)
        if self._completed:
            self.chk.setToolTip("우클릭 → 미완료로 전환")
        self.chk.toggled.connect(lambda v: self.toggled.emit(self._id, v))
        inner.addWidget(self.chk, 0, Qt.AlignmentFlag.AlignVCenter)

        # 텍스트 열 (제목 + 목표 + 마감일)
        txt = QVBoxLayout()
        txt.setSpacing(1)
        txt.setContentsMargins(0, 0, 0, 0)

        # 연결 과제 표시 (긴급업무만)
        if self._linked_title:
            linked_lbl = QLabel(f"🔗 {self._linked_title}")
            linked_lbl.setObjectName("LinkedTodoLabel")
            linked_lbl.setStyleSheet("color:#89b4fa;font-size:9px;background:transparent;padding:0;")
            linked_lbl.setCursor(Qt.CursorShape.PointingHandCursor)
            linked_id = r.get("linked_todo_id")
            linked_lbl.mousePressEvent = lambda e, lid=linked_id: self.navigate_requested.emit(lid) if lid else None
            txt.addWidget(linked_lbl)

        title_text = r["title"]
        title_lbl = QLabel(f"<s>{title_text}</s>" if self._completed else title_text)
        title_lbl.setObjectName("TaskTitleDone" if self._completed else "TaskTitle")
        title_lbl.setWordWrap(True)
        txt.addWidget(title_lbl)

        # 목표 표시 (있을 때만)
        goal = r["goal"] if r["goal"] else ""
        if goal:
            g_lbl = QLabel(f"▸ {goal}")
            g_lbl.setObjectName("TaskGoal")
            g_lbl.setWordWrap(True)
            txt.addWidget(g_lbl)

        # 마감일 뱃지 (개인업무는 이벤트 날짜 단순 표기, 나머지는 카운트다운)
        if r["task_type"] == TASK_PERSONAL and r["due_date"]:
            try:
                d_ev = date.fromisoformat(r["due_date"])
                badge_lbl = QLabel(f"📅 {d_ev.strftime('%m/%d')}")
                badge_lbl.setObjectName("DueBadgeFuture")
                txt.addWidget(badge_lbl)
            except ValueError:
                pass
        else:
            b = due_badge(r["due_date"])
            if b:
                badge_lbl = QLabel(b[0])
                badge_lbl.setObjectName(b[1])
                txt.addWidget(badge_lbl)

        # 파일 출처 표시 (가져온 원본 표시 — 이후 수정/삭제는 사용자가 직접 관리)
        if r["source"] == SOURCE_FILE:
            src_lbl = QLabel("📄 파일 가져옴")
            src_lbl.setObjectName("SourceBadge")
            txt.addWidget(src_lbl)

        # 첨부 파일 뱃지
        if self._file_count > 0:
            file_badge = QLabel(f"📎 {self._file_count}개 파일")
            file_badge.setObjectName("SourceBadge")
            txt.addWidget(file_badge)

        inner.addLayout(txt, 1)

        # 파일 연결 버튼 (file_path 있을 때만, hover 시 표시)
        if r["file_path"]:
            btn_file = QPushButton("📁")
            btn_file.setObjectName("TaskFileBtn")
            btn_file.setFixedSize(24, 24)
            btn_file.setToolTip(f"파일 위치 열기:\n{r['file_path']}")
            fp = r["file_path"]
            btn_file.clicked.connect(lambda _=None, p=fp: open_file_path(p, self))
            inner.addWidget(btn_file, 0, Qt.AlignmentFlag.AlignVCenter)

        # 로그 버튼 (hover 시 표시)
        btn_log = QPushButton("📋")
        btn_log.setObjectName("TaskLogBtn")
        btn_log.setFixedSize(24, 24)
        btn_log.setToolTip("진행 로그 (더블클릭)")
        btn_log.clicked.connect(lambda: self.log_requested.emit(self._id))
        inner.addWidget(btn_log, 0, Qt.AlignmentFlag.AlignVCenter)

        # 편집 버튼 (hover 시 표시)
        btn_edit = QPushButton("✎")
        btn_edit.setObjectName("TaskEditBtn")
        btn_edit.setFixedSize(24, 24)
        btn_edit.setToolTip("편집 (더블클릭으로도 가능)")
        btn_edit.clicked.connect(lambda: self.edit_requested.emit(self._id))
        inner.addWidget(btn_edit, 0, Qt.AlignmentFlag.AlignVCenter)

        # 삭제 버튼 (hover 시 표시)
        btn_del = QPushButton("✕")
        btn_del.setObjectName("TaskDeleteBtn")
        btn_del.setFixedSize(24, 24)
        btn_del.setToolTip("삭제")
        btn_del.clicked.connect(lambda: self.delete_requested.emit(self._id))
        inner.addWidget(btn_del, 0, Qt.AlignmentFlag.AlignVCenter)

        outer.addLayout(inner, 1)

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.edit_requested.emit(self._id)

    def contextMenuEvent(self, e):
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu{background:#313244;border:1px solid #45475a;border-radius:8px;padding:4px;}"
            "QMenu::item{padding:7px 18px;border-radius:6px;color:#cdd6f4;}"
            "QMenu::item:selected{background:#45475a;}"
        )
        a_log  = menu.addAction("📋 진행 로그")
        a_edit = menu.addAction("✏  편집")
        menu.addSeparator()
        a_tog  = menu.addAction("🔲 미완료로" if self._completed else "✅ 완료 표시")
        menu.addSeparator()
        a_del  = menu.addAction("🗑  삭제")
        ch = menu.exec(e.globalPos())
        if ch == a_log:   self.log_requested.emit(self._id)
        elif ch == a_edit: self.edit_requested.emit(self._id)
        elif ch == a_tog:  self.chk.setChecked(not self._completed)
        elif ch == a_del:  self.delete_requested.emit(self._id)

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self._drag_start_pos = e.position().toPoint()
        super().mousePressEvent(e)

    def mouseMoveEvent(self, e):
        if not (e.buttons() & Qt.MouseButton.LeftButton):
            return
        if not hasattr(self, '_drag_start_pos'):
            return
        if (e.position().toPoint() - self._drag_start_pos).manhattanLength() < 12:
            return
        drag = QDrag(self)
        mime = QMimeData()
        mime.setData("application/x-task-id", str(self._id).encode())
        drag.setMimeData(mime)
        px = self.grab().scaledToWidth(
            min(self.width(), 280), Qt.TransformationMode.SmoothTransformation)
        drag.setPixmap(px)
        drag.setHotSpot(QPoint(px.width() // 2, 12))
        drag.exec(Qt.DropAction.MoveAction)

    def show_batch_mode(self, enabled: bool):
        if self._batch_chk:
            if not enabled:
                self._batch_chk.setChecked(False)
            self._batch_chk.setVisible(enabled)

    def is_batch_selected(self) -> bool:
        return bool(self._batch_chk and self._batch_chk.isChecked())


# ═══════════════════════════════════════════════════════════════════════════
# 8. MISC ITEM WIDGET (기타 — 확장/축소 가능 노트 카드)
# ═══════════════════════════════════════════════════════════════════════════

class MiscItemWidget(QFrame):
    delete_requested = Signal(int)
    edit_requested   = Signal(int)

    def __init__(self, task_row, parent=None):
        super().__init__(parent)
        self._id       = task_row["id"]
        self._expanded = False
        self.setObjectName("MiscItem")
        self._build(task_row)

    def _build(self, r):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 8, 12, 8)
        lay.setSpacing(4)

        # 헤더 행 (번호, 제목, 토글, 파일, 삭제)
        hdr = QHBoxLayout()
        hdr.setSpacing(8)

        self.btn_expand = QPushButton("▶")
        self.btn_expand.setObjectName("MiscExpandBtn")
        self.btn_expand.setFixedSize(20, 20)
        self.btn_expand.clicked.connect(self._toggle)
        hdr.addWidget(self.btn_expand)

        title_lbl = QLabel(r["title"])
        title_lbl.setObjectName("MiscTitle")
        title_lbl.setFont(QFont("맑은 고딕", 12, QFont.Weight.Bold))
        title_lbl.setWordWrap(True)
        title_lbl.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse |
            Qt.TextInteractionFlag.TextSelectableByKeyboard
        )
        hdr.addWidget(title_lbl, 1)

        # 파일 연결 버튼 (경로가 있으면 항상 표시)
        if r["file_path"]:
            btn_fp = QPushButton("📁")
            btn_fp.setObjectName("LogFileBtn")
            btn_fp.setFixedSize(22, 22)
            btn_fp.setToolTip(f"파일 열기:\n{r['file_path']}")
            fp = r["file_path"]
            btn_fp.clicked.connect(lambda _=None, p=fp: open_file_path(p, self))
            hdr.addWidget(btn_fp)

        btn_edit = QPushButton("✎")
        btn_edit.setObjectName("LogEditBtn")
        btn_edit.setFixedSize(22, 22)
        btn_edit.setToolTip("편집 (더블클릭으로도 가능)")
        btn_edit.clicked.connect(lambda: self.edit_requested.emit(self._id))
        hdr.addWidget(btn_edit)

        if r["source"] == SOURCE_MANUAL:
            btn_del = QPushButton("✕")
            btn_del.setObjectName("TaskDeleteBtn")
            btn_del.setFixedSize(22, 22)
            btn_del.setToolTip("삭제")
            btn_del.clicked.connect(lambda: self.delete_requested.emit(self._id))
            hdr.addWidget(btn_del)
        lay.addLayout(hdr)

        # 내용 (접힌 상태에서는 숨김)
        self.content_lbl = QLabel(r["description"] or "")
        self.content_lbl.setObjectName("MiscContent")
        self.content_lbl.setWordWrap(True)
        self.content_lbl.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse |
            Qt.TextInteractionFlag.TextSelectableByKeyboard
        )
        self.content_lbl.setVisible(False)
        lay.addWidget(self.content_lbl)

    def _toggle(self):
        self._expanded = not self._expanded
        self.btn_expand.setText("▼" if self._expanded else "▶")
        self.content_lbl.setVisible(self._expanded)

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.edit_requested.emit(self._id)


# ═══════════════════════════════════════════════════════════════════════════
# 9-A. MOVABLE FRAMELESS DIALOG BASE
# ═══════════════════════════════════════════════════════════════════════════

class _MovableDialog(QDialog):
    """Frameless QDialog — drag-to-move, Windows native edge-resize, size persistence."""

    _settings_key: str = ""
    _RESIZE_MARGIN = 6   # px — 테두리 리사이즈 감지 범위

    def __init__(self, parent=None):
        super().__init__(parent)
        self._drag_pos = None
        self._grip     = None   # 우하단 ⠿ 힌트 레이블 (시각용)

    # ── 크기/이동 저장·복원 ────────────────────────────────────────────────
    def showEvent(self, event):
        super().showEvent(event)
        if self._grip is None:
            self._grip = QLabel("⠿", self)
            self._grip.setFixedSize(20, 20)
            self._grip.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self._grip.setStyleSheet("color:#45475a;font-size:12px;background:transparent;")
            self._grip.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents)
        self._position_grip()
        key = self._settings_key or self.__class__.__name__
        settings = QSettings("CalendarTodoList", "MainWindowV2")
        saved = settings.value(f"dialog_size/{key}")
        if saved:
            try:
                w, h = int(saved.split("x")[0]), int(saved.split("x")[1])
                self.resize(w, h)
            except Exception:
                pass

    def closeEvent(self, event):
        self._save_size(); super().closeEvent(event)

    def accept(self):
        self._save_size(); super().accept()

    def reject(self):
        self._save_size(); super().reject()

    def _save_size(self):
        key = self._settings_key or self.__class__.__name__
        settings = QSettings("CalendarTodoList", "MainWindowV2")
        settings.setValue(f"dialog_size/{key}", f"{self.width()}x{self.height()}")

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._position_grip()

    def _position_grip(self):
        if self._grip:
            self._grip.move(self.width() - 20, self.height() - 20)
            self._grip.raise_()

    # ── Windows 네이티브 리사이즈 (WM_NCHITTEST) ──────────────────────────
    def nativeEvent(self, eventType, message):
        """테두리 드래그로 창 크기 조절 — Windows WM_NCHITTEST 활용"""
        if eventType == b"windows_generic_MSG":
            try:
                import ctypes
                from ctypes import wintypes
                msg = ctypes.cast(int(message), ctypes.POINTER(wintypes.MSG)).contents
                if msg.message == 0x0084:  # WM_NCHITTEST
                    # lParam = 화면 좌표 (물리 픽셀) → 위젯 로컬 좌표로 변환
                    sx = ctypes.c_int16(msg.lParam & 0xFFFF).value
                    sy = ctypes.c_int16((msg.lParam >> 16) & 0xFFFF).value
                    pos = self.mapFromGlobal(QPoint(sx, sy))
                    m = self._RESIZE_MARGIN
                    w, h = self.width(), self.height()
                    lft = pos.x() < m
                    rgt = pos.x() > w - m
                    top = pos.y() < m
                    bot = pos.y() > h - m
                    # 코너 → 변 순서로 우선순위 처리
                    if lft and top: return True, 13   # HTTOPLEFT
                    if rgt and top: return True, 14   # HTTOPRIGHT
                    if lft and bot: return True, 16   # HTBOTTOMLEFT
                    if rgt and bot: return True, 17   # HTBOTTOMRIGHT
                    if lft:         return True, 10   # HTLEFT
                    if rgt:         return True, 11   # HTRIGHT
                    if top:         return True, 12   # HTTOP
                    if bot:         return True, 15   # HTBOTTOM
            except Exception:
                pass
        return super().nativeEvent(eventType, message)

    # ── 드래그로 창 이동 ───────────────────────────────────────────────────
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_pos = (event.globalPosition().toPoint()
                              - self.frameGeometry().topLeft())
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if (self._drag_pos is not None
                and event.buttons() & Qt.MouseButton.LeftButton):
            self.move(event.globalPosition().toPoint() - self._drag_pos)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        self._drag_pos = None
        super().mouseReleaseEvent(event)


# ═══════════════════════════════════════════════════════════════════════════
# 9-A2. CUSTOM FILE PICKER DIALOG
# ═══════════════════════════════════════════════════════════════════════════

class _FilePickerDialog(_MovableDialog):
    """프레임리스 파일 선택 다이얼로그 — 테두리 드래그 리사이즈, 다중 선택 지원."""

    _settings_key = "file_picker"

    def __init__(self, title: str = "파일 선택", multi: bool = False, parent=None):
        super().__init__(parent)
        self._multi    = multi
        self._selected: list[str] = []
        self._cur_dir  = Path.home()
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setMinimumSize(460, 420)
        self._build(title)
        self._nav(self._cur_dir)
        QShortcut(QKeySequence("Escape"),    self, self.reject)
        QShortcut(QKeySequence("Return"),    self, self._try_accept)
        QShortcut(QKeySequence("Backspace"), self, self._go_up)
        QShortcut(QKeySequence("Alt+Up"),    self, self._go_up)

    def _build(self, title: str):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # ── 제목 바 (드래그 이동) ────────────────────────────────────────
        tb = QFrame(); tb.setObjectName("DialogTitleBar")
        tb.setFixedHeight(42)
        tb_lay = QHBoxLayout(tb)
        tb_lay.setContentsMargins(14, 0, 6, 0); tb_lay.setSpacing(8)
        lbl_title = QLabel(title)
        lbl_title.setStyleSheet("font-weight:bold;font-size:13px;color:#cdd6f4;")
        tb_lay.addWidget(lbl_title)
        tb_lay.addStretch()
        btn_close = QPushButton("✕"); btn_close.setObjectName("TitleBtnClose")
        btn_close.setFixedSize(34, 34); btn_close.clicked.connect(self.reject)
        tb_lay.addWidget(btn_close)
        lay.addWidget(tb)

        # ── 경로 네비게이션 바 ────────────────────────────────────────────
        nav = QHBoxLayout()
        nav.setContentsMargins(12, 8, 12, 4); nav.setSpacing(6)

        self._btn_up = QPushButton("↑"); self._btn_up.setObjectName("SecondaryBtn")
        self._btn_up.setFixedSize(30, 28); self._btn_up.setToolTip("상위 폴더 (Backspace)")
        self._btn_up.clicked.connect(self._go_up)
        nav.addWidget(self._btn_up)

        for label, path, tip in [
            ("🏠", Path.home(), "홈"),
            ("🖥", Path.home() / "Desktop", "바탕화면"),
            ("📥", Path.home() / "Downloads", "다운로드"),
        ]:
            b = QPushButton(label); b.setObjectName("SecondaryBtn")
            b.setFixedSize(30, 28); b.setToolTip(tip)
            b.clicked.connect(lambda _, p=path: self._nav(p))
            nav.addWidget(b)

        self._path_bar = QLineEdit()
        self._path_bar.setPlaceholderText("경로를 입력 후 Enter…")
        self._path_bar.setStyleSheet(
            "QLineEdit{background:#1e1e2e;border:1px solid #3d3d58;border-radius:6px;"
            "padding:3px 8px;color:#cdd6f4;font-size:11px;}"
            "QLineEdit:focus{border-color:#89b4fa;}"
        )
        self._path_bar.returnPressed.connect(self._nav_from_bar)
        nav.addWidget(self._path_bar, 1)
        lay.addLayout(nav)

        # ── 파일 리스트 ────────────────────────────────────────────────────
        from PySide6.QtWidgets import QListWidget, QListWidgetItem as _LWI
        self._list = QListWidget()
        self._list.setObjectName("FilePickerList")
        self._list.setStyleSheet(
            "QListWidget{background:#181825;border:none;padding:4px 6px;outline:none;}"
            "QListWidget::item{padding:5px 10px;border-radius:6px;color:#cdd6f4;min-height:24px;}"
            "QListWidget::item:hover{background:#313244;}"
            "QListWidget::item:selected{background:#45475a;color:#cdd6f4;}"
        )
        self._list.setSelectionMode(
            QListWidget.SelectionMode.ExtendedSelection if self._multi
            else QListWidget.SelectionMode.SingleSelection
        )
        self._list.itemDoubleClicked.connect(self._on_double)
        self._list.itemSelectionChanged.connect(self._on_sel_changed)
        self._list.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        lay.addWidget(self._list)
        lay.setStretchFactor(self._list, 1)

        # ── 선택 파일 표시 + 직접 입력 ────────────────────────────────────
        foot = QFrame()
        foot.setStyleSheet("QFrame{background:#13131d;border-top:1px solid #2e2e48;}")
        foot_lay = QVBoxLayout(foot)
        foot_lay.setContentsMargins(12, 8, 12, 8); foot_lay.setSpacing(4)

        sel_row = QHBoxLayout()
        lbl_sel = QLabel("선택:"); lbl_sel.setStyleSheet("color:#6c7086;font-size:11px;")
        sel_row.addWidget(lbl_sel)
        self._sel_lbl = QLabel("없음")
        self._sel_lbl.setStyleSheet("color:#a6adc8;font-size:11px;")
        self._sel_lbl.setWordWrap(True)
        sel_row.addWidget(self._sel_lbl, 1)
        btn_type = QPushButton("직접 입력")
        btn_type.setObjectName("SecondaryBtn"); btn_type.setFixedHeight(24)
        btn_type.setStyleSheet("font-size:10px;padding:0 8px;")
        btn_type.clicked.connect(self._manual_input)
        sel_row.addWidget(btn_type)
        foot_lay.addLayout(sel_row)

        btn_row = QHBoxLayout(); btn_row.setSpacing(8)
        btn_row.addStretch()
        btn_cancel = QPushButton("취소"); btn_cancel.setObjectName("SecondaryBtn")
        btn_cancel.setFixedHeight(34); btn_cancel.clicked.connect(self.reject)
        btn_row.addWidget(btn_cancel)
        self._btn_ok = QPushButton("선택"); self._btn_ok.setObjectName("PrimaryBtn")
        self._btn_ok.setFixedHeight(34); self._btn_ok.setEnabled(False)
        self._btn_ok.clicked.connect(self.accept)
        btn_row.addWidget(self._btn_ok)
        foot_lay.addLayout(btn_row)
        lay.addWidget(foot)

    def _nav(self, path: Path):
        try:
            items = sorted(path.iterdir(), key=lambda p: (p.is_file(), p.name.lower()))
        except (PermissionError, FileNotFoundError, OSError):
            return
        self._cur_dir = path
        self._path_bar.setText(str(path))
        self._list.clear()
        self._selected = []
        self._upd_sel_lbl()

        from PySide6.QtWidgets import QListWidgetItem as _LWI
        for item in items:
            if item.name.startswith('.'):
                continue  # 숨김 파일 제외
            ext = item.suffix.lower()
            if item.is_dir():
                icon = "📁"
            elif ext in ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.svg'):
                icon = "🖼"
            elif ext == '.pdf':
                icon = "📕"
            elif ext in ('.doc', '.docx', '.hwp', '.odt'):
                icon = "📝"
            elif ext in ('.xls', '.xlsx', '.csv', '.ods'):
                icon = "📊"
            elif ext in ('.ppt', '.pptx'):
                icon = "📊"
            elif ext in ('.zip', '.rar', '.7z', '.tar', '.gz'):
                icon = "📦"
            elif ext in ('.py', '.js', '.ts', '.java', '.c', '.cpp', '.rs', '.go'):
                icon = "💻"
            elif ext in ('.txt', '.md', '.log'):
                icon = "📄"
            elif ext in ('.mp3', '.wav', '.flac', '.aac', '.ogg'):
                icon = "🎵"
            elif ext in ('.mp4', '.avi', '.mov', '.mkv'):
                icon = "🎬"
            else:
                icon = "📄"

            size_str = ""
            if item.is_file():
                try:
                    sz = item.stat().st_size
                    if sz < 1024:
                        size_str = f"  {sz}B"
                    elif sz < 1024 ** 2:
                        size_str = f"  {sz // 1024}KB"
                    else:
                        size_str = f"  {sz // 1024 ** 2}MB"
                except OSError:
                    pass

            lw = _LWI(f"{icon}  {item.name}{size_str}")
            lw.setData(Qt.ItemDataRole.UserRole, item)
            if item.is_dir():
                lw.setForeground(QColor("#89b4fa"))
            self._list.addItem(lw)

    def _on_double(self, item):
        p: Path = item.data(Qt.ItemDataRole.UserRole)
        if p and p.is_dir():
            self._nav(p)
        elif p and p.is_file():
            self._selected = [str(p)]
            self._upd_sel_lbl()
            self._btn_ok.setEnabled(True)
            self.accept()

    def _on_sel_changed(self):
        sel = []
        for item in self._list.selectedItems():
            p: Path = item.data(Qt.ItemDataRole.UserRole)
            if p and p.is_file():
                sel.append(str(p))
        self._selected = sel
        self._upd_sel_lbl()
        self._btn_ok.setEnabled(bool(sel))

    def _upd_sel_lbl(self):
        if not self._selected:
            self._sel_lbl.setText("없음")
            self._sel_lbl.setToolTip("")
        elif len(self._selected) == 1:
            self._sel_lbl.setText(Path(self._selected[0]).name)
            self._sel_lbl.setToolTip(self._selected[0])  # 전체 경로는 툴팁으로
        else:
            names = ", ".join(Path(p).name for p in self._selected[:3])
            suffix = f" 외 {len(self._selected)-3}개" if len(self._selected) > 3 else ""
            self._sel_lbl.setText(f"{names}{suffix}")
            self._sel_lbl.setToolTip("\n".join(self._selected))

    def _go_up(self):
        parent = self._cur_dir.parent
        if parent != self._cur_dir:
            self._nav(parent)

    def _nav_from_bar(self):
        """경로 바에 직접 입력 후 Enter → 해당 경로로 이동"""
        text = self._path_bar.text().strip()
        if not text:
            return
        p = Path(text)
        if p.is_dir():
            self._nav(p)
        elif p.is_file():
            self._selected = [str(p)]
            self._upd_sel_lbl()
            self._btn_ok.setEnabled(True)
        else:
            self._path_bar.setStyleSheet(
                "QLineEdit{background:#2d1b1b;border:1px solid #f38ba8;border-radius:6px;"
                "padding:3px 8px;color:#f38ba8;font-size:11px;}"
            )
            QTimer.singleShot(1200, lambda: self._path_bar.setStyleSheet(
                "QLineEdit{background:#1e1e2e;border:1px solid #3d3d58;border-radius:6px;"
                "padding:3px 8px;color:#cdd6f4;font-size:11px;}"
                "QLineEdit:focus{border-color:#89b4fa;}"
            ))

    def _try_accept(self):
        if self._btn_ok.isEnabled():
            self.accept()

    def _manual_input(self):
        from PySide6.QtWidgets import QInputDialog
        text, ok = QInputDialog.getText(self, "직접 입력", "파일 경로:", text=str(self._cur_dir))
        if ok and text.strip():
            p = Path(text.strip())
            if p.is_file():
                self._selected = [str(p)]
                self._upd_sel_lbl()
                self._btn_ok.setEnabled(True)
            elif p.is_dir():
                self._nav(p)

    def get_paths(self) -> list[str]:
        return self._selected

    @staticmethod
    def pick_files(title: str = "파일 선택", multi: bool = True, parent=None) -> list[str]:
        """정적 헬퍼 — 선택된 파일 경로 목록 반환 (취소 시 빈 리스트)"""
        dlg = _FilePickerDialog(title=title, multi=multi, parent=parent)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            return dlg.get_paths()
        return []


# ═══════════════════════════════════════════════════════════════════════════
# 9-B. SCHEDULE DIALOG
# ═══════════════════════════════════════════════════════════════════════════

class ScheduleDialog(_MovableDialog):
    """단기 일정 / 휴가 / 교육 추가·편집 다이얼로그"""

    def __init__(self, parent=None, sched_data=None, preset_date: date = None):
        super().__init__(parent)
        self._data      = sched_data
        self._is_edit   = sched_data is not None
        self._preset    = preset_date or date.today()
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setMinimumWidth(460)
        self._build()
        if self._is_edit:
            self._load()
        QShortcut(QKeySequence("Escape"),       self, self.reject)
        QShortcut(QKeySequence("Ctrl+Return"),  self, self._ok)

    def _mk_lbl(self, text):
        l = QLabel(text); l.setObjectName("FormLabel"); return l

    def _mk_date(self, d: date) -> QDateEdit:
        de = QDateEdit()
        de.setCalendarPopup(True)
        de.setDisplayFormat("yyyy-MM-dd")
        de.setDate(QDate(d.year, d.month, d.day))
        return de

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        title = QLabel("일정 편집" if self._is_edit else "📅  일정 추가")
        title.setObjectName("DialogTitle")
        title.setFont(QFont("맑은 고딕", 14, QFont.Weight.Bold))
        lay.addWidget(title)

        # 유형 선택 (2행 레이아웃)
        lay.addWidget(self._mk_lbl("일정 유형"))
        type_row1 = QHBoxLayout(); type_row1.setSpacing(6)
        type_row2 = QHBoxLayout(); type_row2.setSpacing(6)
        self._type_btns: dict[str, QPushButton] = {}
        for i, etype in enumerate([SCHED_SINGLE, SCHED_VACATION, SCHED_TRAINING, SCHED_TRIP]):
            btn = QPushButton(f"{SCHED_ICONS[etype]}  {SCHED_LABELS[etype]}")
            btn.setCheckable(True)
            btn.setFixedHeight(32)
            clr = SCHED_COLORS[etype]
            btn.setStyleSheet(
                f"QPushButton{{background:transparent;border:2px solid #45475a;border-radius:8px;"
                f"font-size:12px;padding:0 8px;}}"
                f"QPushButton:checked{{background:rgba({self._hex_to_rgb(clr)},0.15);"
                f"border:2px solid {clr};color:{clr};font-weight:bold;}}"
                f"QPushButton:hover{{border-color:{clr};color:{clr};}}"
            )
            btn.clicked.connect(lambda _, t=etype: self._pick_type(t))
            self._type_btns[etype] = btn
            (type_row1 if i < 2 else type_row2).addWidget(btn)
        lay.addLayout(type_row1)
        lay.addLayout(type_row2)
        self._type_btns[SCHED_SINGLE].setChecked(True)
        self._cur_type = SCHED_SINGLE

        # 일정 이름
        lay.addWidget(self._mk_lbl("일정 이름 *"))
        self.ed_name = QLineEdit()
        self.ed_name.setPlaceholderText("일정 이름")
        lay.addWidget(self.ed_name)

        # 날짜 행
        date_row = QHBoxLayout(); date_row.setSpacing(12)

        col_start = QVBoxLayout()
        col_start.addWidget(self._mk_lbl("날짜 (시작일)"))
        self.de_start = self._mk_date(self._preset)
        col_start.addWidget(self.de_start)
        date_row.addLayout(col_start)

        col_end = QVBoxLayout()
        col_end.addWidget(self._mk_lbl("종료일 (선택)"))
        self.de_end = self._mk_date(self._preset)
        self.de_end.setEnabled(False)
        col_end.addWidget(self.de_end)
        date_row.addLayout(col_end)
        lay.addLayout(date_row)

        # 시간 (단기 일정만)
        lay.addWidget(self._mk_lbl("시간 (선택)"))
        self.ed_time = QLineEdit()
        self.ed_time.setPlaceholderText("예: 14:00 또는 14:00~16:00")
        lay.addWidget(self.ed_time)

        # 장소
        lay.addWidget(self._mk_lbl("장소 (선택)"))
        self.ed_loc = QLineEdit()
        self.ed_loc.setPlaceholderText("장소 입력...")
        lay.addWidget(self.ed_loc)

        # 내용
        lay.addWidget(self._mk_lbl("내용 (선택)"))
        self.ed_content = QTextEdit()
        self.ed_content.setPlaceholderText("상세 내용...")
        self.ed_content.setMinimumHeight(60)
        lay.addWidget(self.ed_content)
        lay.setStretchFactor(self.ed_content, 1)

        lay.addSpacing(4)
        br = QHBoxLayout(); br.addStretch()
        bc = QPushButton("취소"); bc.setObjectName("SecondaryBtn")
        bc.setFixedHeight(38); bc.clicked.connect(self.reject); br.addWidget(bc)
        bo = QPushButton("저장" if self._is_edit else "추가"); bo.setObjectName("PrimaryBtn")
        bo.setFixedHeight(38); bo.clicked.connect(self._ok); br.addWidget(bo)
        lay.addLayout(br)
        self.ed_name.setFocus()

    @staticmethod
    def _hex_to_rgb(hex_color: str) -> str:
        h = hex_color.lstrip("#")
        r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
        return f"{r},{g},{b}"

    def _pick_type(self, etype: str):
        self._cur_type = etype
        for k, btn in self._type_btns.items():
            btn.setChecked(k == etype)
        self.de_end.setEnabled(True)   # 모든 유형에서 종료일 설정 가능
        self.ed_time.setEnabled(etype == SCHED_SINGLE)

    def _load(self):
        d = self._data
        self._pick_type(d["event_type"])
        self.ed_name.setText(d["name"])
        try:
            sd = date.fromisoformat(d["event_date"])
            self.de_start.setDate(QDate(sd.year, sd.month, sd.day))
            if d["end_date"]:
                ed = date.fromisoformat(d["end_date"])
                self.de_end.setDate(QDate(ed.year, ed.month, ed.day))
        except Exception:
            pass
        self.ed_time.setText(d["start_time"] or "")
        self.ed_loc.setText(d["location"] or "")
        self.ed_content.setPlainText(d["content"] or "")

    def _ok(self):
        if not self.ed_name.text().strip():
            self.ed_name.setPlaceholderText("⚠ 이름을 입력하세요")
            self.ed_name.setFocus(); return
        self.accept()

    def values(self) -> dict:
        start_str = self.de_start.date().toString("yyyy-MM-dd")
        end_qdate = self.de_end.date()
        start_qdate = self.de_start.date()
        # 종료일이 시작일보다 이전이면 시작일로 강제 통일
        if end_qdate < start_qdate:
            end_qdate = start_qdate
        end_str = end_qdate.toString("yyyy-MM-dd")
        # 시작일과 동일하면 None (단일일 처리)
        end_str = None if end_str == start_str else end_str
        return {
            "name":       self.ed_name.text().strip(),
            "event_date": start_str,
            "end_date":   end_str,
            "start_time": self.ed_time.text().strip() or None,
            "location":   self.ed_loc.text().strip(),
            "content":    self.ed_content.toPlainText().strip(),
            "event_type": self._cur_type,
        }


# ═══════════════════════════════════════════════════════════════════════════
# 9. TASK DIALOG (추가/편집)
# ═══════════════════════════════════════════════════════════════════════════

class TaskDialog(_MovableDialog):
    def __init__(self, parent=None, task_data=None, task_type=None, preset_date: date = None, db=None):
        super().__init__(parent)
        self._data        = task_data
        self._is_edit     = task_data is not None
        self._task_type   = task_type or (task_data["task_type"] if task_data else None)
        self._preset_date = preset_date
        self._selected_color: str | None = None
        self._db          = db
        self._linked_todo_id: int | None = None
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setMinimumSize(460, 400)
        self._build()
        if self._is_edit:
            self._load()
        elif self._preset_date:
            # 달력 클릭으로 열린 경우: 날짜 자동 설정
            self.chk_due.setChecked(True)
            self.de_due.setDate(QDate(self._preset_date.year,
                                      self._preset_date.month,
                                      self._preset_date.day))
        QShortcut(QKeySequence("Escape"), self, self.reject)
        QShortcut(QKeySequence("Ctrl+Return"), self, self._ok)

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        _add_titles = {
            TASK_TODO:     "✨ 새 할 일 추가",
            TASK_URGENT:   "🚨 새 긴급업무 추가",
            TASK_PERSONAL: "👤 새 개인업무 추가",
            TASK_MISC:     "📌 새 기타 항목 추가",
        }
        t = QLabel("할 일 편집" if self._is_edit else _add_titles.get(self._task_type, "✨ 새 항목 추가"))
        t.setObjectName("DialogTitle")
        t.setFont(QFont("맑은 고딕", 14, QFont.Weight.Bold))
        lay.addWidget(t)

        def lbl(text):
            l = QLabel(text)
            l.setObjectName("FormLabel")
            return l

        lay.addWidget(lbl("제목 *"))
        self.ed_title = QLineEdit()
        self.ed_title.setPlaceholderText("제목을 입력하세요")
        lay.addWidget(self.ed_title)

        lay.addWidget(lbl("내용 (선택)"))
        self.ed_desc = QTextEdit()
        self.ed_desc.setPlaceholderText("상세 내용...")
        self.ed_desc.setMinimumHeight(60)
        lay.addWidget(self.ed_desc)
        lay.setStretchFactor(self.ed_desc, 1)

        lay.addWidget(lbl("목표 (선택)"))
        self.ed_goal = QLineEdit()
        self.ed_goal.setPlaceholderText("목표를 입력하세요 (선택)")
        lay.addWidget(self.ed_goal)

        row = QHBoxLayout()
        row.setSpacing(12)

        col_p = QVBoxLayout()
        col_p.addWidget(lbl("우선순위"))
        self.cb_prio = QComboBox()
        for k, v in PRIORITY_LABELS.items():
            self.cb_prio.addItem(v, k)
        self.cb_prio.setCurrentIndex(1)
        col_p.addWidget(self.cb_prio)
        row.addLayout(col_p)

        col_d = QVBoxLayout()
        dh = QHBoxLayout()
        dh.addWidget(lbl("마감일"))
        self.chk_due = QCheckBox("설정 (선택사항)")
        self.chk_due.setObjectName("TaskCheck")
        self.chk_due.toggled.connect(lambda v: self.de_due.setEnabled(v))
        dh.addWidget(self.chk_due)
        dh.addStretch()
        col_d.addLayout(dh)
        self.de_due = QDateEdit()
        self.de_due.setCalendarPopup(True)
        self.de_due.setDate(QDate.currentDate())
        self.de_due.setDisplayFormat("yyyy-MM-dd")
        self.de_due.setEnabled(False)
        col_d.addWidget(self.de_due)
        row.addLayout(col_d)
        lay.addLayout(row)

        # 색상 선택 (todo / personal 만)
        if self._task_type in (TASK_TODO, TASK_PERSONAL):
            lay.addWidget(lbl("항목 색상"))
            color_row = QHBoxLayout()
            color_row.setSpacing(6)
            self._color_btns: list[QPushButton] = []
            color_labels = ["기본", "빨강", "주황", "노랑", "초록", "하늘", "파랑", "보라", "분홍", "민트"]
            for i, (clr, name) in enumerate(zip(TASK_COLORS, color_labels)):
                btn = QPushButton()
                btn.setFixedSize(28, 28)
                btn.setToolTip(name)
                btn.setCheckable(True)
                if clr is None:
                    btn.setStyleSheet(
                        "QPushButton{background:#313244;border-radius:14px;border:2px solid #45475a;}"
                        "QPushButton:checked{border:2px solid #cdd6f4;}"
                    )
                    btn.setText("●")
                    btn.setStyleSheet(
                        "QPushButton{background:#313244;border-radius:14px;"
                        "border:2px solid #45475a;color:#6c7086;font-size:12px;}"
                        "QPushButton:checked{border:3px solid #cdd6f4;}"
                    )
                else:
                    btn.setStyleSheet(
                        f"QPushButton{{background:{clr};border-radius:14px;border:2px solid transparent;}}"
                        f"QPushButton:checked{{border:3px solid #cdd6f4;}}"
                    )
                btn.clicked.connect(lambda _, c=clr, b=btn: self._pick_color(c, b))
                color_row.addWidget(btn)
                self._color_btns.append(btn)
            color_row.addStretch()
            lay.addLayout(color_row)
            # 기본값 선택
            self._color_btns[0].setChecked(True)

        # 첨부 파일 (여러 개 — 버튼 선택 또는 드래그앤드롭)
        if self._task_type in (TASK_TODO, TASK_URGENT, TASK_MISC):
            # 헤더 행: "📎 첨부 파일" 레이블 + 우측 "＋ 파일 추가" 버튼
            file_hdr = QHBoxLayout()
            file_hdr.addWidget(lbl("📎  첨부 파일"))
            file_hdr.addStretch()
            btn_add_file = QPushButton("＋  파일 추가")
            btn_add_file.setObjectName("SecondaryBtn")
            btn_add_file.setFixedHeight(26)
            btn_add_file.setToolTip("파일 선택 (여러 개 가능) — 드래그앤드롭도 지원")
            btn_add_file.clicked.connect(self._add_file)
            file_hdr.addWidget(btn_add_file)
            lay.addLayout(file_hdr)

            # 파일 목록 드롭 영역
            self._drop_area = QFrame()
            self._drop_area.setObjectName("FileDropArea")
            self._drop_area.setStyleSheet(
                "QFrame#FileDropArea{border:1px dashed #3d3d58;border-radius:8px;"
                "background:#1a1a28;}"
            )
            self._drop_area.setAcceptDrops(True)
            self._drop_area.dragEnterEvent = self._file_drag_enter
            self._drop_area.dragLeaveEvent = self._file_drag_leave
            self._drop_area.dropEvent      = self._file_drop
            drop_lay = QVBoxLayout(self._drop_area)
            drop_lay.setContentsMargins(8, 6, 8, 6)
            drop_lay.setSpacing(2)

            self._file_empty_lbl = QLabel("파일을 드래그하거나 위 버튼으로 추가")
            self._file_empty_lbl.setStyleSheet(
                "color:#45475a;font-size:10px;background:transparent;padding:2px 0;"
            )
            self._file_empty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            drop_lay.addWidget(self._file_empty_lbl)

            self._file_list_lay = QVBoxLayout()
            self._file_list_lay.setSpacing(2)
            self._file_list_lay.setContentsMargins(0, 0, 0, 0)
            drop_lay.addLayout(self._file_list_lay)
            lay.addWidget(self._drop_area)

            self._pending_files: list[str] = []
            self._removed_file_ids: list[int] = []
            self._existing_files: list = []
            self.ed_fpath = None
        else:
            self.ed_fpath = None
            self._pending_files = []
            self._removed_file_ids = []
            self._existing_files = []
            self._file_list_lay = None

        # 긴급업무 연결 과제 선택
        if self._task_type == TASK_URGENT and self._db is not None:
            lay.addWidget(lbl("연결 과제 (선택)"))
            self.cb_linked = QComboBox()
            self.cb_linked.addItem("(연결 없음)", None)
            todo_tasks = self._db.get_tasks(TASK_TODO, completed=False)
            for tt in todo_tasks:
                self.cb_linked.addItem(tt["title"], tt["id"])
            lay.addWidget(self.cb_linked)
        else:
            self.cb_linked = None

        lay.addSpacing(4)
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        bc = QPushButton("취소")
        bc.setObjectName("SecondaryBtn")
        bc.setFixedHeight(38)
        bc.clicked.connect(self.reject)
        btn_row.addWidget(bc)
        bo = QPushButton("저장" if self._is_edit else "추가")
        bo.setObjectName("PrimaryBtn")
        bo.setFixedHeight(38)
        bo.clicked.connect(self._ok)
        btn_row.addWidget(bo)
        lay.addLayout(btn_row)
        self.ed_title.setFocus()

    def _browse_file(self):
        paths = _FilePickerDialog.pick_files("파일 선택", multi=False, parent=self)
        if paths:
            self.ed_fpath.setText(paths[0])

    def _browse_folder(self):
        path = QFileDialog.getExistingDirectory(
            self, "폴더 선택", options=QFileDialog.Option.DontUseNativeDialog
        )
        if path:
            self.ed_fpath.setText(path)

    def _file_drag_enter(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self._drop_area.setStyleSheet(
                "QFrame#FileDropArea{border:1px dashed #89b4fa;border-radius:8px;"
                "background:#1a1a34;}"
            )

    def _file_drag_leave(self, e):
        self._drop_area.setStyleSheet(
            "QFrame#FileDropArea{border:1px dashed #3d3d58;border-radius:8px;"
            "background:#1a1a28;}"
        )

    def _file_drop(self, e):
        self._drop_area.setStyleSheet(
            "QFrame#FileDropArea{border:1px dashed #3d3d58;border-radius:8px;"
            "background:#1a1a28;}"
        )
        for url in e.mimeData().urls():
            path = url.toLocalFile()
            if path and path not in self._pending_files:
                self._pending_files.append(path)
                self._add_file_row(path, is_existing=False, file_id=None)

    def _add_file(self):
        paths = _FilePickerDialog.pick_files("파일 추가 (Shift/Ctrl+클릭으로 여러 개)", multi=True, parent=self)
        for path in paths:
            if path and path not in self._pending_files:
                self._pending_files.append(path)
                self._add_file_row(path, is_existing=False, file_id=None)

    def _add_file_row(self, path: str, is_existing: bool, file_id: int | None):
        if self._file_list_lay is None:
            return
        # 파일이 추가되면 빈 상태 안내 숨김
        if hasattr(self, "_file_empty_lbl"):
            self._file_empty_lbl.hide()
        row_w = QFrame()
        row_w.setStyleSheet(
            "QFrame{background:#22223a;border-radius:6px;border:1px solid #2e2e48;}"
            "QFrame:hover{background:#2a2a46;border-color:#45475a;}"
        )
        row_lay = QHBoxLayout(row_w)
        row_lay.setContentsMargins(8, 4, 6, 4)
        row_lay.setSpacing(6)

        # 확장자로 아이콘 선택
        ext = Path(path).suffix.lower()
        if ext in ('.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp'):
            icon = "🖼"
        elif ext in ('.pdf',):
            icon = "📕"
        elif ext in ('.doc', '.docx', '.hwp', '.txt', '.md'):
            icon = "📝"
        elif ext in ('.xls', '.xlsx', '.csv'):
            icon = "📊"
        elif ext in ('.zip', '.rar', '.7z', '.tar', '.gz'):
            icon = "📦"
        elif ext in ('.py', '.js', '.ts', '.java', '.cpp', '.c', '.h'):
            icon = "💻"
        else:
            icon = "📄"

        fname = Path(path).name
        lbl_icon = QLabel(icon)
        lbl_icon.setStyleSheet("font-size:13px;background:transparent;")
        row_lay.addWidget(lbl_icon)

        lbl_name = QLabel(fname)
        lbl_name.setStyleSheet("color:#cdd6f4;font-size:11px;background:transparent;")
        lbl_name.setToolTip(path)
        lbl_name.setWordWrap(False)
        row_lay.addWidget(lbl_name, 1)

        # 파일 크기 표시
        try:
            sz = Path(path).stat().st_size
            if sz < 1024:
                sz_str = f"{sz}B"
            elif sz < 1024 * 1024:
                sz_str = f"{sz//1024}KB"
            else:
                sz_str = f"{sz//(1024*1024)}MB"
        except Exception:
            sz_str = ""
        if sz_str:
            lbl_sz = QLabel(sz_str)
            lbl_sz.setStyleSheet("color:#6c7086;font-size:10px;background:transparent;")
            row_lay.addWidget(lbl_sz)

        btn_rm = QPushButton("✕")
        btn_rm.setObjectName("TaskDeleteBtn")
        btn_rm.setFixedSize(20, 20)
        if is_existing and file_id is not None:
            fid = file_id
            btn_rm.clicked.connect(lambda _, w=row_w, fid=fid, p=path: self._remove_existing_file(w, fid, p))
        else:
            btn_rm.clicked.connect(lambda _, w=row_w, p=path: self._remove_pending_file(w, p))
        row_lay.addWidget(btn_rm)

        self._file_list_lay.addWidget(row_w)

    def _remove_pending_file(self, row_w, path: str):
        if path in self._pending_files:
            self._pending_files.remove(path)
        row_w.deleteLater()

    def _remove_existing_file(self, row_w, file_id: int, path: str):
        self._removed_file_ids.append(file_id)
        row_w.deleteLater()

    def _pick_color(self, color: str | None, clicked_btn: QPushButton):
        """색상 버튼 선택 처리 (단일 선택 라디오 동작)"""
        self._selected_color = color
        if hasattr(self, "_color_btns"):
            for b in self._color_btns:
                b.setChecked(False)
        clicked_btn.setChecked(True)

    def _load(self):
        d = self._data
        self.ed_title.setText(d["title"])
        self.ed_desc.setPlainText(d["description"] or "")
        self.ed_goal.setText(d["goal"] or "")
        for i in range(self.cb_prio.count()):
            if self.cb_prio.itemData(i) == d["priority"]:
                self.cb_prio.setCurrentIndex(i); break
        if d["due_date"]:
            self.chk_due.setChecked(True)
            parts = d["due_date"].split("-")
            self.de_due.setDate(QDate(int(parts[0]), int(parts[1]), int(parts[2])))
        # 파일 경로 복원 (legacy)
        if self.ed_fpath and d["file_path"]:
            self.ed_fpath.setText(d["file_path"])
        # 첨부 파일 목록 복원
        if self._db and self._is_edit and self._file_list_lay is not None:
            self._existing_files = self._db.get_task_files(d["id"])
            for f in self._existing_files:
                self._add_file_row(f["original_path"], is_existing=True, file_id=f["id"])
            # 레거시 file_path 처리
            if d["file_path"] and not self._existing_files:
                self._add_file_row(d["file_path"], is_existing=False, file_id=None)
        # 색상 복원
        if hasattr(self, "_color_btns"):
            saved_color = d["color"] if d["color"] else None
            self._selected_color = saved_color
            for btn, clr in zip(self._color_btns, TASK_COLORS):
                btn.setChecked(clr == saved_color)
        # 연결 과제 복원
        if self.cb_linked is not None:
            try:
                linked_id = d["linked_todo_id"]
            except (KeyError, IndexError):
                linked_id = None
            if linked_id:
                for i in range(self.cb_linked.count()):
                    if self.cb_linked.itemData(i) == linked_id:
                        self.cb_linked.setCurrentIndex(i)
                        break

    def _ok(self):
        if not self.ed_title.text().strip():
            self.ed_title.setPlaceholderText("⚠ 제목을 입력하세요")
            self.ed_title.setFocus(); return
        self.accept()

    def values(self) -> dict:
        return {
            "title":          self.ed_title.text().strip(),
            "description":    self.ed_desc.toPlainText().strip(),
            "goal":           self.ed_goal.text().strip(),
            "priority":       self.cb_prio.currentData(),
            "due_date":       self.de_due.date().toString("yyyy-MM-dd")
                              if self.chk_due.isChecked() else None,
            "color":          self._selected_color,
            "file_path":      self.ed_fpath.text().strip() if self.ed_fpath else None,
            "linked_todo_id": self.cb_linked.currentData() if self.cb_linked else None,
        }


# ═══════════════════════════════════════════════════════════════════════════
# 10. LOG DIALOG
# ═══════════════════════════════════════════════════════════════════════════

class LogItemWidget(QFrame):
    """개별 로그 항목 — 텍스트 선택/복사 가능, 더블클릭 또는 편집 버튼으로 인라인 수정"""
    delete_requested = Signal(int)
    edit_done        = Signal(int, str, str)   # log_id, new_content, file_path

    def __init__(self, log_row, parent=None):
        super().__init__(parent)
        self.log_id     = log_row["id"]
        self._content   = log_row["content"]
        self._file_path = log_row["file_path"] if log_row["file_path"] else ""
        self.setObjectName("LogItem")
        self._build(log_row)

    def _build(self, r):
        cl = QVBoxLayout(self)
        cl.setContentsMargins(12, 10, 10, 10)
        cl.setSpacing(4)

        # 헤더 행: 타임스탬프 + 파일 + 편집 + 삭제
        th = QHBoxLayout()
        ts = QLabel(f"[{r['created_at']}]")
        ts.setObjectName("LogTimestamp")
        th.addWidget(ts)
        th.addStretch()

        if r["file_path"]:
            btn_fp = QPushButton("📁")
            btn_fp.setObjectName("LogFileBtn")
            btn_fp.setFixedSize(22, 22)
            btn_fp.setToolTip(f"파일 열기:\n{r['file_path']}")
            fp = r["file_path"]
            btn_fp.clicked.connect(lambda _=None, p=fp: open_file_path(p, self))
            th.addWidget(btn_fp)

        btn_ed = QPushButton("✎")
        btn_ed.setObjectName("LogEditBtn")
        btn_ed.setFixedSize(22, 22)
        btn_ed.setToolTip("로그 편집")
        btn_ed.clicked.connect(self._enter_edit)
        th.addWidget(btn_ed)

        bd = QPushButton("✕")
        bd.setObjectName("LogDeleteBtn")
        bd.setFixedSize(22, 22)
        bd.setToolTip("삭제")
        bd.clicked.connect(lambda: self.delete_requested.emit(self.log_id))
        th.addWidget(bd)
        cl.addLayout(th)

        # 표시 모드: 텍스트 선택/복사 가능한 QLabel
        self.lbl_content = QLabel(r["content"])
        self.lbl_content.setObjectName("LogContent")
        self.lbl_content.setWordWrap(True)
        self.lbl_content.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse |
            Qt.TextInteractionFlag.TextSelectableByKeyboard
        )
        self.lbl_content.installEventFilter(self)
        cl.addWidget(self.lbl_content)

        # 편집 모드: 기본 숨김
        self.ed_content = QPlainTextEdit()
        self.ed_content.setMinimumHeight(60)
        self.ed_content.setVisible(False)
        self.ed_content.setPlaceholderText("내용 수정...")
        cl.addWidget(self.ed_content)
        cl.setStretchFactor(self.ed_content, 1)

        # 편집 모드: 파일 경로 변경
        self.ed_file_row = QWidget()
        efr_lay = QHBoxLayout(self.ed_file_row)
        efr_lay.setContentsMargins(0, 2, 0, 0)
        efr_lay.setSpacing(4)
        self.ed_file_path = QLineEdit()
        self.ed_file_path.setPlaceholderText("첨부 파일 경로 (선택 사항)...")
        efr_lay.addWidget(self.ed_file_path, 1)
        btn_fp_browse = QPushButton("📂")
        btn_fp_browse.setObjectName("LogFileBtn")
        btn_fp_browse.setFixedSize(28, 28)
        btn_fp_browse.setToolTip("파일 선택")
        btn_fp_browse.clicked.connect(self._browse_file)
        efr_lay.addWidget(btn_fp_browse)
        self.ed_file_row.setVisible(False)
        cl.addWidget(self.ed_file_row)

        # 저장/취소 버튼 행 (편집 모드에서만 표시)
        self.edit_btns = QWidget()
        eb_lay = QHBoxLayout(self.edit_btns)
        eb_lay.setContentsMargins(0, 0, 0, 0)
        eb_lay.setSpacing(6)
        eb_lay.addStretch()
        btn_save = QPushButton("저장")
        btn_save.setObjectName("PrimaryBtn")
        btn_save.setFixedHeight(32)
        btn_save.clicked.connect(self._save_edit)
        eb_lay.addWidget(btn_save)
        btn_cancel = QPushButton("취소")
        btn_cancel.setObjectName("SecondaryBtn")
        btn_cancel.setFixedHeight(32)
        btn_cancel.clicked.connect(self._cancel_edit)
        eb_lay.addWidget(btn_cancel)
        self.edit_btns.setVisible(False)
        cl.addWidget(self.edit_btns)

    def _enter_edit(self):
        self.ed_content.setPlainText(self._content)
        self.ed_file_path.setText(self._file_path)
        self.lbl_content.setVisible(False)
        self.ed_content.setVisible(True)
        self.ed_file_row.setVisible(True)
        self.edit_btns.setVisible(True)
        self.ed_content.setFocus()

    def _save_edit(self):
        new_text = self.ed_content.toPlainText().strip()
        if not new_text:
            return
        new_fp = self.ed_file_path.text().strip()
        self._content   = new_text
        self._file_path = new_fp
        self.lbl_content.setText(new_text)
        self.lbl_content.setVisible(True)
        self.ed_content.setVisible(False)
        self.ed_file_row.setVisible(False)
        self.edit_btns.setVisible(False)
        self.edit_done.emit(self.log_id, new_text, new_fp)

    def _cancel_edit(self):
        self.lbl_content.setVisible(True)
        self.ed_content.setVisible(False)
        self.ed_file_row.setVisible(False)
        self.edit_btns.setVisible(False)

    def _browse_file(self):
        paths = _FilePickerDialog.pick_files("파일 선택", multi=False, parent=self)
        if paths:
            self.ed_file_path.setText(paths[0])

    def eventFilter(self, obj, event):
        """내용 라벨 더블클릭 → 편집 모드 진입 (TextSelectable 이벤트 우회)"""
        if obj is self.lbl_content and event.type() == QEvent.Type.MouseButtonDblClick:
            self._enter_edit()
            return True
        return super().eventFilter(obj, event)

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self._enter_edit()


class ProgressEntryRow(QWidget):
    """과제진행상황 개별 항목 행 — 보기/편집 인라인 전환"""
    delete_requested = Signal(int)       # log_id
    edit_done        = Signal(int, str)  # log_id, new_content

    def __init__(self, entry, parent=None):
        super().__init__(parent)
        self._id      = entry["id"]
        self._content = entry["content"]
        self._ts      = (entry["created_at"] or "")[:16]
        self._editing = False
        self._build()

    def _build(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(2)

        # ── 뷰 행: 라벨 + 타임스탬프 + 버튼 ─────────────────────────────
        self._view_row = QWidget()
        lay = QHBoxLayout(self._view_row)
        lay.setContentsMargins(0, 3, 0, 3)
        lay.setSpacing(4)

        self._lbl = QLabel(f"• {self._content}")
        self._lbl.setObjectName("LogContent")
        self._lbl.setWordWrap(True)
        self._lbl.setStyleSheet("background:transparent;padding-left:4px;font-size:11px;")
        self._lbl.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse |
            Qt.TextInteractionFlag.TextSelectableByKeyboard
        )
        lay.addWidget(self._lbl, 1)

        self._ts_lbl = QLabel(self._ts)
        self._ts_lbl.setObjectName("LogTimestamp")
        self._ts_lbl.setStyleSheet("font-size:9px;color:#6c7086;background:transparent;")
        lay.addWidget(self._ts_lbl)

        self._btn_edit = QPushButton("✏")
        self._btn_edit.setObjectName("LogDeleteBtn")
        self._btn_edit.setFixedSize(22, 22)
        self._btn_edit.setToolTip("수정")
        self._btn_edit.clicked.connect(self._start_edit)
        lay.addWidget(self._btn_edit)

        self._btn_del = QPushButton("✕")
        self._btn_del.setObjectName("LogDeleteBtn")
        self._btn_del.setFixedSize(22, 22)
        self._btn_del.setToolTip("삭제")
        self._btn_del.clicked.connect(lambda: self.delete_requested.emit(self._id))
        lay.addWidget(self._btn_del)

        outer.addWidget(self._view_row)

        # ── 편집 행: QLineEdit + 저장/취소 버튼 (기본 숨김) ──────────────
        self._edit_row = QWidget()
        self._edit_row.setVisible(False)
        elay = QHBoxLayout(self._edit_row)
        elay.setContentsMargins(0, 2, 0, 2)
        elay.setSpacing(4)

        self._ed = QLineEdit(self._content)
        self._ed.setMinimumHeight(32)
        self._ed.setStyleSheet(
            "QLineEdit{background:#1e1e2e;border:1px solid #89b4fa;"
            "border-radius:4px;padding:2px 6px;font-size:11px;color:#cdd6f4;}"
        )
        self._ed.returnPressed.connect(self._save)
        elay.addWidget(self._ed, 1)

        self._btn_save = QPushButton("저장")
        self._btn_save.setObjectName("PrimaryBtn")
        self._btn_save.setFixedHeight(34)
        self._btn_save.setFixedWidth(46)
        self._btn_save.setToolTip("저장 (Enter)")
        self._btn_save.clicked.connect(self._save)
        elay.addWidget(self._btn_save)

        self._btn_cancel = QPushButton("취소")
        self._btn_cancel.setObjectName("SecondaryBtn")
        self._btn_cancel.setFixedHeight(34)
        self._btn_cancel.setFixedWidth(46)
        self._btn_cancel.setToolTip("취소")
        self._btn_cancel.clicked.connect(self._cancel)
        elay.addWidget(self._btn_cancel)

        outer.addWidget(self._edit_row)

    def _start_edit(self):
        self._editing = True
        self._ed.setText(self._content)
        self._view_row.setVisible(False)
        self._edit_row.setVisible(True)
        self._ed.setFocus()
        self._ed.selectAll()

    def _cancel(self):
        self._editing = False
        self._edit_row.setVisible(False)
        self._view_row.setVisible(True)

    def _save(self):
        new_text = self._ed.text().strip()
        if not new_text:
            return
        self._content = new_text
        self._lbl.setText(f"• {new_text}")
        self.edit_done.emit(self._id, new_text)
        self._cancel()


class LogDialog(_MovableDialog):
    """진행 로그 다이얼로그 — 좌측 사이드바 탭 (일반사항 / 과제진행상황)"""

    def __init__(self, db: Database, task_id: int, parent=None):
        super().__init__(parent)
        self.db, self.task_id = db, task_id
        self._task = db.get_task(task_id)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setMinimumSize(560, 480)
        self._cur_tab = "general"
        self._log_attach_path = ""
        self._build()
        self._switch_tab("general")
        self._auto_size()
        QShortcut(QKeySequence("Escape"), self, self.accept)
        QShortcut(QKeySequence("Ctrl+Return"), self, self._add_general)

    def _build(self):
        root = QHBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ── 좌측 사이드바 ────────────────────────────────────────────────
        sidebar = QWidget()
        sidebar.setFixedWidth(150)
        sidebar.setObjectName("LogSidebar")
        sidebar.setStyleSheet(
            "QWidget#LogSidebar{background:#181825;border-right:1px solid #313244;}"
        )
        sb_lay = QVBoxLayout(sidebar)
        sb_lay.setContentsMargins(0, 0, 0, 0)
        sb_lay.setSpacing(0)

        # 태스크 제목 (사이드바 헤더)
        task_hdr = QLabel(self._task["title"])
        task_hdr.setObjectName("SectionTitle")
        task_hdr.setWordWrap(True)
        task_hdr.setStyleSheet(
            "color:#cdd6f4;font-weight:bold;font-size:11px;"
            "padding:14px 10px 10px 10px;background:transparent;"
        )
        sb_lay.addWidget(task_hdr)

        sep0 = QFrame(); sep0.setFrameShape(QFrame.Shape.HLine)
        sep0.setMaximumHeight(1)
        sep0.setStyleSheet("background:#313244;")
        sb_lay.addWidget(sep0)

        def make_tab_btn(icon, text, key):
            btn = QPushButton(f"  {icon}  {text}")
            btn.setCheckable(True)
            btn.setObjectName("LogTabBtn")
            btn.setStyleSheet(
                "QPushButton#LogTabBtn{"
                "  text-align:left;padding:12px 10px;"
                "  border:none;border-radius:0;"
                "  color:#a6adc8;font-size:11px;background:transparent;"
                "}"
                "QPushButton#LogTabBtn:checked{"
                "  color:#cdd6f4;background:#24243e;"
                "  border-left:3px solid #89b4fa;"
                "}"
                "QPushButton#LogTabBtn:hover:!checked{"
                "  background:#1e1e2e;"
                "}"
            )
            btn.clicked.connect(lambda: self._switch_tab(key))
            return btn

        self.btn_general  = make_tab_btn("📝", "일반사항", "general")
        self.btn_progress = make_tab_btn("📊", "과제진행상황", "progress")
        sb_lay.addWidget(self.btn_general)
        sb_lay.addWidget(self.btn_progress)
        sb_lay.addStretch()

        # 닫기 버튼 (사이드바 하단)
        btn_close = QPushButton("✕  닫기")
        btn_close.setObjectName("SecondaryBtn")
        btn_close.setStyleSheet(
            "QPushButton{border:none;border-radius:0;padding:10px;color:#6c7086;background:transparent;}"
            "QPushButton:hover{color:#cdd6f4;background:#1e1e2e;}"
        )
        btn_close.clicked.connect(self.accept)
        sb_lay.addWidget(btn_close)

        root.addWidget(sidebar)

        # ── 우측 콘텐츠 영역 ─────────────────────────────────────────────
        self.content_area = QWidget()
        content_lay = QVBoxLayout(self.content_area)
        content_lay.setContentsMargins(0, 0, 0, 0)
        content_lay.setSpacing(0)

        # 태스크 정보 박스
        info = QFrame()
        info.setObjectName("TaskInfoBox")
        info.setStyleSheet(
            "QFrame#TaskInfoBox{background:#1e1e2e;border:none;"
            "border-bottom:1px solid #313244;}"
        )
        info_l = QVBoxLayout(info)
        info_l.setContentsMargins(16, 12, 16, 12)
        info_l.setSpacing(4)

        t_lbl = QLabel(self._task["title"])
        t_lbl.setObjectName("TaskInfoTitle")
        t_lbl.setFont(QFont("맑은 고딕", 12, QFont.Weight.Bold))
        t_lbl.setWordWrap(True)
        info_l.addWidget(t_lbl)

        goal = self._task["goal"]
        if goal:
            g_lbl = QLabel(f"▸ {goal}")
            g_lbl.setObjectName("TaskInfoDesc")
            g_lbl.setWordWrap(True)
            info_l.addWidget(g_lbl)

        content_lay.addWidget(info)

        # 탭 콘텐츠 스택 (QWidget swap 방식)
        self.stack = QWidget()
        self.stack_lay = QVBoxLayout(self.stack)
        self.stack_lay.setContentsMargins(0, 0, 0, 0)
        self.stack_lay.setSpacing(0)
        content_lay.addWidget(self.stack, 1)

        root.addWidget(self.content_area, 1)

        # ── 일반사항 패널 ─────────────────────────────────────────────────
        self.panel_general = QWidget()
        pg_lay = QVBoxLayout(self.panel_general)
        pg_lay.setContentsMargins(16, 14, 16, 14)
        pg_lay.setSpacing(10)

        hdr_g = QHBoxLayout()
        lbl_h = QLabel("📝 일반사항")
        lbl_h.setObjectName("DialogTitle")
        lbl_h.setFont(QFont("맑은 고딕", 12, QFont.Weight.Bold))
        hdr_g.addWidget(lbl_h)
        hdr_g.addStretch()
        self.lbl_count = QLabel("0개")
        self.lbl_count.setObjectName("FormLabel")
        hdr_g.addWidget(self.lbl_count)
        pg_lay.addLayout(hdr_g)

        # ── QSplitter: 위=로그 목록, 아래=입력창 (드래그로 비율 조절) ──
        splitter_g = QSplitter(Qt.Orientation.Vertical)
        splitter_g.setHandleWidth(6)
        splitter_g.setStyleSheet(
            "QSplitter::handle{background:#313244;border-radius:2px;margin:1px 4px;}"
            "QSplitter::handle:hover{background:#45475a;}"
        )

        # 위쪽: 로그 목록 스크롤 영역
        self.scroll_general = QScrollArea()
        self.scroll_general.setWidgetResizable(True)
        self.scroll_general.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.log_cont = QWidget()
        self.log_lay = QVBoxLayout(self.log_cont)
        self.log_lay.setContentsMargins(0, 0, 0, 0)
        self.log_lay.setSpacing(6)
        self.log_lay.addStretch()
        self.scroll_general.setWidget(self.log_cont)
        self.scroll_general.setMinimumHeight(80)
        splitter_g.addWidget(self.scroll_general)

        # 아래쪽: 입력 패널 (splitter로 사용자가 비율 조절 가능)
        input_panel = QWidget()
        inp_lay = QVBoxLayout(input_panel)
        inp_lay.setContentsMargins(0, 6, 0, 0)
        inp_lay.setSpacing(5)

        inp_hdr = QHBoxLayout()
        inp_hdr.addWidget(self._mk_lbl("✏  새 메모 작성", "FormLabel"))
        inp_hdr.addStretch()
        inp_hdr.addWidget(self._mk_lbl("Ctrl+Enter 저장", "LogTimestamp"))
        inp_lay.addLayout(inp_hdr)

        self.ed = QPlainTextEdit()
        self.ed.setPlaceholderText("내용을 입력하세요... (여러 줄 가능)")
        self.ed.setMinimumHeight(60)
        inp_lay.addWidget(self.ed, 1)

        fa = QHBoxLayout(); fa.setSpacing(6)
        self.lbl_attach = QLabel("📎 파일 없음")
        self.lbl_attach.setObjectName("LogTimestamp")
        fa.addWidget(self.lbl_attach, 1)
        btn_attach = QPushButton("📂 첨부")
        btn_attach.setObjectName("SecondaryBtn")
        btn_attach.setFixedHeight(30)
        btn_attach.clicked.connect(self._browse_attach)
        fa.addWidget(btn_attach)
        btn_detach = QPushButton("✕")
        btn_detach.setObjectName("LogDeleteBtn")
        btn_detach.setFixedSize(26, 26)
        btn_detach.clicked.connect(self._clear_attach)
        fa.addWidget(btn_detach)

        ba_g = QPushButton("메모 추가")
        ba_g.setObjectName("PrimaryBtn")
        ba_g.setFixedHeight(30)
        ba_g.clicked.connect(self._add_general)
        fa.addWidget(ba_g)
        inp_lay.addLayout(fa)

        splitter_g.addWidget(input_panel)
        splitter_g.setStretchFactor(0, 1)   # 로그 목록이 남은 공간 전부 차지
        splitter_g.setStretchFactor(1, 0)   # 입력 패널은 최소 크기 유지
        pg_lay.addWidget(splitter_g, 1)

        # ── 과제진행상황 패널 ─────────────────────────────────────────────
        self.panel_progress = QWidget()
        pp_lay = QVBoxLayout(self.panel_progress)
        pp_lay.setContentsMargins(16, 14, 16, 14)
        pp_lay.setSpacing(10)

        hdr_p = QHBoxLayout()
        lbl_ph = QLabel("📊 과제진행상황")
        lbl_ph.setObjectName("DialogTitle")
        lbl_ph.setFont(QFont("맑은 고딕", 12, QFont.Weight.Bold))
        hdr_p.addWidget(lbl_ph)
        hdr_p.addStretch()
        btn_new_grp = QPushButton("＋ 새 그룹")
        btn_new_grp.setObjectName("PrimaryBtn")
        btn_new_grp.setFixedHeight(30)
        btn_new_grp.clicked.connect(self._add_progress_group)
        hdr_p.addWidget(btn_new_grp)
        pp_lay.addLayout(hdr_p)

        self.scroll_progress = QScrollArea()
        self.scroll_progress.setWidgetResizable(True)
        self.scroll_progress.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.prog_cont = QWidget()
        self.prog_lay  = QVBoxLayout(self.prog_cont)
        self.prog_lay.setContentsMargins(0, 0, 0, 0)
        self.prog_lay.setSpacing(8)
        self.prog_lay.addStretch()
        self.scroll_progress.setWidget(self.prog_cont)
        self.scroll_progress.setMinimumHeight(200)
        pp_lay.addWidget(self.scroll_progress, 1)

    def _mk_lbl(self, text, obj, font=None):
        l = QLabel(text); l.setObjectName(obj)
        if font: l.setFont(font)
        return l

    def _auto_size(self):
        """로그 내용 양에 맞게 초기 창 크기 자동 설정."""
        from PySide6.QtGui import QFontMetrics
        logs = self.db.get_general_logs(self.task_id)

        # ── 폭: 가장 긴 줄 기준 ──────────────────────────────────────────
        fm = QFontMetrics(self.font())
        char_w = fm.horizontalAdvance("가")   # 한글 한 글자 폭
        max_px = 0
        for lg in logs:
            for line in lg["content"].split("\n"):
                px = fm.horizontalAdvance(line)
                if px > max_px:
                    max_px = px
        SIDEBAR = 150
        PADDING = 80   # 좌우 여백 + 스크롤바
        content_w = max_px + PADDING
        target_w = SIDEBAR + max(content_w, 360)   # 최소 콘텐츠 360px
        target_w = max(560, min(target_w, 960))    # 전체 560~960px

        # ── 높이: 로그 수 × 추정 행 높이 ────────────────────────────────
        ROW_H    = max(60, char_w * 3)   # 폰트 기반 행 높이 추정 (최소 60)
        INPUT_H  = 200                    # 입력 패널 + 헤더
        HEADER_H = 120                    # 태스크 정보 + 탭 헤더
        log_area = len(logs) * ROW_H
        target_h = HEADER_H + log_area + INPUT_H
        target_h = max(480, min(target_h, 880))    # 480~880px

        self.resize(target_w, target_h)

    def _switch_tab(self, key: str):
        self._cur_tab = key
        self.btn_general.setChecked(key == "general")
        self.btn_progress.setChecked(key == "progress")

        # 스택에서 기존 패널 제거
        while self.stack_lay.count():
            item = self.stack_lay.takeAt(0)
            if item.widget():
                item.widget().setParent(None)

        if key == "general":
            self.stack_lay.addWidget(self.panel_general)
            self.panel_general.setVisible(True)
            self._load_general()
            self.ed.setFocus()
        else:
            self.stack_lay.addWidget(self.panel_progress)
            self.panel_progress.setVisible(True)
            self._load_progress()

    def _load_general(self):
        while self.log_lay.count() > 1:
            item = self.log_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        logs = self.db.get_general_logs(self.task_id)
        self.lbl_count.setText(f"{len(logs)}개")
        if not logs:
            emp = QLabel("아직 메모가 없습니다.")
            emp.setObjectName("TaskInfoDesc")
            emp.setAlignment(Qt.AlignmentFlag.AlignCenter)
            emp.setStyleSheet("color:#6c7086;padding:20px 0;")
            self.log_lay.insertWidget(0, emp)
        else:
            for i, lg in enumerate(logs):
                card = LogItemWidget(lg)
                card.delete_requested.connect(self._del_log)
                card.edit_done.connect(self._edit_log)
                self.log_lay.insertWidget(i, card)
        self.scroll_general.verticalScrollBar().setValue(
            self.scroll_general.verticalScrollBar().maximum()
        )

    def _load_progress(self):
        while self.prog_lay.count() > 1:
            item = self.prog_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        groups = self.db.get_progress_groups(self.task_id)
        if not groups:
            emp = QLabel("진행 그룹이 없습니다.\n'＋ 새 그룹' 버튼으로 추가하세요.")
            emp.setObjectName("TaskInfoDesc")
            emp.setAlignment(Qt.AlignmentFlag.AlignCenter)
            emp.setWordWrap(True)
            emp.setStyleSheet("color:#6c7086;padding:20px 0;")
            self.prog_lay.insertWidget(0, emp)
        else:
            for i, grp in enumerate(groups):
                card = self._make_group_card(grp)
                self.prog_lay.insertWidget(i, card)
        self.scroll_progress.verticalScrollBar().setValue(
            self.scroll_progress.verticalScrollBar().maximum()
        )

    def _make_group_card(self, grp) -> QFrame:
        card = QFrame()
        card.setObjectName("LogItem")
        card.setStyleSheet(
            "QFrame#LogItem{background:#24243e;border:1px solid #313244;border-radius:8px;}"
        )
        cl = QVBoxLayout(card)
        cl.setContentsMargins(12, 10, 12, 10)
        cl.setSpacing(6)

        gid = grp["id"]

        # ── 헤더 행 ──────────────────────────────────────────────────────
        hdr = QHBoxLayout(); hdr.setSpacing(4)

        # 그룹 제목 (뷰/편집 인라인 전환)
        title_lbl = QLabel(f"▶ {grp['title']}")
        title_lbl.setObjectName("TaskTitle")
        title_lbl.setStyleSheet("font-weight:bold;font-size:11px;background:transparent;")
        hdr.addWidget(title_lbl, 1)

        title_ed = QLineEdit(grp["title"])
        title_ed.setMaximumHeight(24)
        title_ed.setVisible(False)
        hdr.addWidget(title_ed, 1)

        date_lbl = QLabel(grp["created_at"][:10])
        date_lbl.setObjectName("LogTimestamp")
        date_lbl.setStyleSheet("font-size:10px;background:transparent;color:#6c7086;")
        hdr.addWidget(date_lbl)

        # ✏ 제목 편집 버튼
        btn_edit_title = QPushButton("✏")
        btn_edit_title.setObjectName("LogDeleteBtn")
        btn_edit_title.setFixedSize(20, 20)
        btn_edit_title.setToolTip("그룹 제목 수정")
        hdr.addWidget(btn_edit_title)

        # ✔ 제목 저장 버튼 (편집 모드)
        btn_save_title = QPushButton("✔")
        btn_save_title.setObjectName("PrimaryBtn")
        btn_save_title.setFixedSize(20, 20)
        btn_save_title.setVisible(False)
        btn_save_title.setToolTip("저장 (Enter)")
        hdr.addWidget(btn_save_title)

        # ✖ 제목 취소 버튼 (편집 모드)
        btn_cancel_title = QPushButton("✖")
        btn_cancel_title.setObjectName("LogDeleteBtn")
        btn_cancel_title.setFixedSize(20, 20)
        btn_cancel_title.setVisible(False)
        btn_cancel_title.setToolTip("취소")
        hdr.addWidget(btn_cancel_title)

        # ✕ 그룹 삭제 버튼
        btn_del_grp = QPushButton("✕")
        btn_del_grp.setObjectName("LogDeleteBtn")
        btn_del_grp.setFixedSize(20, 20)
        btn_del_grp.setToolTip("그룹 삭제")
        btn_del_grp.clicked.connect(lambda _=None, g=gid: self._del_group(g))
        hdr.addWidget(btn_del_grp)

        cl.addLayout(hdr)

        # ── 그룹 제목 편집 로직 ──────────────────────────────────────────
        def _enter_title_edit():
            title_lbl.setVisible(False)
            title_ed.setVisible(True)
            btn_edit_title.setVisible(False)
            btn_save_title.setVisible(True)
            btn_cancel_title.setVisible(True)
            date_lbl.setVisible(False)
            btn_del_grp.setVisible(False)
            title_ed.setFocus(); title_ed.selectAll()

        def _cancel_title_edit():
            title_ed.setVisible(False)
            btn_save_title.setVisible(False)
            btn_cancel_title.setVisible(False)
            title_lbl.setVisible(True)
            date_lbl.setVisible(True)
            btn_edit_title.setVisible(True)
            btn_del_grp.setVisible(True)

        def _save_title():
            new_t = title_ed.text().strip()
            if new_t:
                self.db.update_progress_group_title(gid, new_t)
                title_lbl.setText(f"▶ {new_t}")
            _cancel_title_edit()

        btn_edit_title.clicked.connect(_enter_title_edit)
        btn_save_title.clicked.connect(_save_title)
        btn_cancel_title.clicked.connect(_cancel_title_edit)
        title_ed.returnPressed.connect(_save_title)

        # 구분선
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setMaximumHeight(1)
        sep.setStyleSheet("background:#313244;margin:0 0;")
        cl.addWidget(sep)

        # ── 기존 항목들 (ProgressEntryRow) ──────────────────────────────
        entries = self.db.get_progress_logs(gid)
        for entry in entries:
            row = ProgressEntryRow(entry, card)
            row.delete_requested.connect(self._del_progress_entry)
            row.edit_done.connect(self._edit_progress_entry)
            cl.addWidget(row)

        if not entries:
            emp_lbl = QLabel("항목이 없습니다. 아래에서 추가하세요.")
            emp_lbl.setStyleSheet("color:#6c7086;font-size:10px;padding:2px 4px;")
            cl.addWidget(emp_lbl)

        # ── 새 항목 추가 행 ──────────────────────────────────────────────
        add_sep = QFrame(); add_sep.setFrameShape(QFrame.Shape.HLine)
        add_sep.setMaximumHeight(1)
        add_sep.setStyleSheet("background:#313244;margin:2px 0;")
        cl.addWidget(add_sep)

        add_row = QHBoxLayout(); add_row.setSpacing(6)
        ed_entry = QLineEdit()
        ed_entry.setPlaceholderText("새 항목 입력 후 Enter 또는 ＋ 버튼...")
        ed_entry.setMinimumHeight(34)
        ed_entry.setStyleSheet(
            "QLineEdit{background:#1e1e2e;border:1px solid #45475a;"
            "border-radius:4px;padding:2px 8px;font-size:11px;color:#cdd6f4;}"
            "QLineEdit:focus{border:1px solid #89b4fa;}"
        )
        add_row.addWidget(ed_entry, 1)
        btn_add_entry = QPushButton("＋ 추가")
        btn_add_entry.setObjectName("PrimaryBtn")
        btn_add_entry.setFixedHeight(34)
        btn_add_entry.clicked.connect(
            lambda _=None, e=ed_entry, g=gid: self._add_progress_entry(e, g)
        )
        ed_entry.returnPressed.connect(
            lambda e=ed_entry, g=gid: self._add_progress_entry(e, g)
        )
        add_row.addWidget(btn_add_entry)
        cl.addLayout(add_row)

        return card

    def _add_progress_group(self):
        from PySide6.QtWidgets import QInputDialog
        title, ok = QInputDialog.getText(
            self, "새 진행 그룹", "그룹 제목 (예: 1차 미팅, 현장 방문):"
        )
        if ok and title.strip():
            self.db.add_progress_group(self.task_id, title.strip())
            self._load_progress()

    def _add_progress_entry(self, ed: QLineEdit, group_id: int):
        txt = ed.text().strip()
        if not txt:
            return
        self.db.add_progress_log(self.task_id, group_id, txt)
        ed.clear()
        self._load_progress()

    def _del_group(self, group_id: int):
        r = QMessageBox.question(
            self, "그룹 삭제", "이 진행 그룹과 모든 항목을 삭제하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if r == QMessageBox.StandardButton.Yes:
            self.db.delete_progress_group(group_id)
            self._load_progress()

    def _del_progress_entry(self, log_id: int):
        self.db.delete_progress_log(log_id)
        self._load_progress()

    def _edit_progress_entry(self, log_id: int, new_content: str):
        self.db.update_progress_log(log_id, new_content)
        # ProgressEntryRow 가 자체 라벨을 즉시 업데이트하므로 reload 불필요

    def _browse_attach(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "파일 선택", "", "모든 파일 (*.*)",
            options=QFileDialog.Option.DontUseNativeDialog,
        )
        if path:
            self._log_attach_path = path
            self.lbl_attach.setText(f"📎 {Path(path).name}")

    def _clear_attach(self):
        self._log_attach_path = ""
        self.lbl_attach.setText("📎 파일 없음")

    def _add_general(self):
        txt = self.ed.toPlainText().strip()
        if not txt:
            self.ed.setPlaceholderText("⚠ 내용을 입력하세요"); return
        self.db.add_log(self.task_id, txt, self._log_attach_path or None)
        self._log_attach_path = ""
        self.lbl_attach.setText("📎 파일 없음")
        self.ed.clear()
        self._load_general()

    def _edit_log(self, log_id: int, new_content: str, file_path: str):
        self.db.update_log(log_id, new_content, file_path or None)
        self._load_general()

    def _del_log(self, log_id):
        r = QMessageBox.question(self, "메모 삭제", "이 메모를 삭제하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if r == QMessageBox.StandardButton.Yes:
            self.db.delete_log(log_id)
            self._load_general()


# ═══════════════════════════════════════════════════════════════════════════
# 11. TASK SECTION (todo / urgent / personal)
# ═══════════════════════════════════════════════════════════════════════════

class TaskSection(QWidget):
    completion_changed = Signal()   # 완료/미완료 변경 시 CompletedSection 갱신용
    navigate_to        = Signal(int)  # 긴급업무→과제 스크롤 요청

    def __init__(self, db: Database, task_type: str, title: str,
                 header_color: str = "#89b4fa", parent=None):
        super().__init__(parent)
        self.db, self.task_type, self.title_str = db, task_type, title
        self._header_color  = header_color
        self._collapsed     = False
        self._highlight_date: date | None = None
        self._sort_mode     = "default"
        self._batch_mode    = False
        self._batch_selected: set[int] = set()
        self.setObjectName("SectionWidget")
        self._build()
        self.refresh()
        sc = QShortcut(QKeySequence("Ctrl+N"), self)
        sc.setContext(Qt.ShortcutContext.WidgetWithChildrenShortcut)
        sc.activated.connect(self._add)

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)

        # 헤더
        hdr_w = QWidget(); hdr_w.setObjectName("SectionHeader")
        hdr_w.setStyleSheet(
            f"QWidget#SectionHeader{{border-left:4px solid {self._header_color};"
            f"border-radius:8px 8px 0 0;}}"
        )
        hdr_l = QVBoxLayout(hdr_w)
        hdr_l.setContentsMargins(12,10,10,8); hdr_l.setSpacing(6)

        title_row = QHBoxLayout()
        tl = QLabel(self.title_str); tl.setObjectName("SectionTitle")
        tl.setFont(QFont("맑은 고딕", 12, QFont.Weight.Bold))
        if self.task_type == TASK_URGENT:
            tl.setToolTip("이번 주 처리해야 하는 단기 업무.\n과제/할 일(장기 관리)과 구분해서 사용하세요.")
        elif self.task_type == TASK_TODO:
            tl.setToolTip("마감일 기준으로 관리하는 장기 과제·할 일.\n긴급업무(단기)와 구분해서 사용하세요.")
        elif self.task_type == TASK_PERSONAL:
            tl.setToolTip("나만 보는 메모·약속·개인 일정.\n팀 공유 없이 혼자 관리하는 항목을 넣으세요.")
        title_row.addWidget(tl); title_row.addStretch()
        lbl_sort = QLabel("정렬:"); lbl_sort.setObjectName("SectionStats")
        title_row.addWidget(lbl_sort)
        self.cb_sort = QComboBox()
        self.cb_sort.setFixedHeight(24)
        self.cb_sort.setFixedWidth(108)
        self.cb_sort.setObjectName("SortCombo")
        self.cb_sort.addItem("기본", "default")
        self.cb_sort.addItem("마감일↑", "due_asc")
        self.cb_sort.addItem("마감일↓", "due_desc")
        self.cb_sort.addItem("우선순위", "priority")
        self.cb_sort.addItem("생성일↑", "created_asc")
        self.cb_sort.addItem("생성일↓", "created_desc")
        self.cb_sort.addItem("제목", "title")
        self.cb_sort.currentIndexChanged.connect(self._on_sort_changed)
        title_row.addWidget(self.cb_sort)
        self.lbl_stats = QLabel("0/0 완료"); self.lbl_stats.setObjectName("SectionStats")
        title_row.addWidget(self.lbl_stats)
        self.btn_batch = QPushButton("☐"); self.btn_batch.setObjectName("SectionCollapseBtn")
        self.btn_batch.setFixedSize(26, 26); self.btn_batch.setToolTip("선택 모드 — 체크 후 일괄 완료/삭제")
        self.btn_batch.clicked.connect(self._toggle_batch_mode)
        title_row.addWidget(self.btn_batch)
        self.btn_col = QPushButton("▼"); self.btn_col.setObjectName("SectionCollapseBtn")
        self.btn_col.setFixedSize(26,26); self.btn_col.setToolTip("섹션 접기/펼치기")
        self.btn_col.clicked.connect(self._toggle)
        title_row.addWidget(self.btn_col)
        hdr_l.addLayout(title_row)

        self.prog = QProgressBar()
        self.prog.setMaximumHeight(5); self.prog.setTextVisible(False)
        self.prog.setRange(0,100); self.prog.setValue(0)
        hdr_l.addWidget(self.prog)
        lay.addWidget(hdr_w)

        # 콘텐츠
        self.body = QWidget()
        self.body.setAcceptDrops(True)
        self.body.dragEnterEvent = self._body_drag_enter
        self.body.dragMoveEvent  = self._body_drag_move
        self.body.dragLeaveEvent = self._body_drag_leave
        self.body.dropEvent      = self._body_drop

        b_lay = QVBoxLayout(self.body)
        b_lay.setContentsMargins(8,6,8,8); b_lay.setSpacing(4)
        # 배치 작업 바 (기본 숨김)
        self.batch_bar = QWidget()
        self.batch_bar.setVisible(False)
        bb_lay = QHBoxLayout(self.batch_bar)
        bb_lay.setContentsMargins(8, 4, 8, 4); bb_lay.setSpacing(6)
        self.lbl_batch_count = QLabel("0개 선택")
        self.lbl_batch_count.setObjectName("FormLabel")
        bb_lay.addWidget(self.lbl_batch_count)
        bb_lay.addStretch()
        btn_ba_all = QPushButton("전체 선택")
        btn_ba_all.setObjectName("SecondaryBtn")
        btn_ba_all.setFixedHeight(30)
        btn_ba_all.clicked.connect(self._batch_select_all)
        bb_lay.addWidget(btn_ba_all)
        btn_ba_done = QPushButton("✅ 완료")
        btn_ba_done.setObjectName("PrimaryBtn")
        btn_ba_done.setFixedHeight(30)
        btn_ba_done.clicked.connect(self._batch_complete)
        bb_lay.addWidget(btn_ba_done)
        btn_ba_del = QPushButton("🗑 삭제")
        btn_ba_del.setObjectName("TaskDeleteBtn")
        btn_ba_del.setFixedHeight(30)
        btn_ba_del.clicked.connect(self._batch_delete)
        bb_lay.addWidget(btn_ba_del)
        b_lay.addWidget(self.batch_bar)
        self.items_lay = QVBoxLayout()
        self.items_lay.setContentsMargins(0,0,0,0); self.items_lay.setSpacing(4)
        b_lay.addLayout(self.items_lay)
        _empty_text = {TASK_TODO: "할 일이 없습니다.", TASK_URGENT: "긴급업무가 없습니다.",
                       TASK_PERSONAL: "개인업무가 없습니다."}.get(self.task_type, "항목이 없습니다.")
        self.empty_lbl = QLabel(_empty_text)
        self.empty_lbl.setObjectName("TaskInfoDesc")
        self.empty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_lbl.setStyleSheet("color:#7f849c;padding:14px 0;font-size:12px;")
        self.empty_lbl.hide()
        b_lay.addWidget(self.empty_lbl)
        self.btn_add = QPushButton("＋  새 항목 추가  (Ctrl+N)")
        self.btn_add.setObjectName("AddTaskBtn"); self.btn_add.setMinimumHeight(38)
        self.btn_add.clicked.connect(self._add)
        b_lay.addWidget(self.btn_add)
        lay.addWidget(self.body)

        # 드롭 인디케이터 (파란 선)
        self._indicator = QFrame(self.body)
        self._indicator.setFixedHeight(2)
        self._indicator.setStyleSheet("background:#89b4fa;border-radius:1px;")
        self._indicator.hide()

    def refresh(self):
        while self.items_lay.count():
            item = self.items_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        tasks = self.db.get_tasks(self.task_type, completed=False)  # 미완료만
        total, done = self.db.get_task_stats(self.task_type)
        self.lbl_stats.setText(f"{done}/{total} 완료")
        self.prog.setValue(int(done/total*100) if total else 0)
        tasks = self._apply_sort(tasks)
        if not tasks:
            self.empty_lbl.show()
        else:
            self.empty_lbl.hide()
            for t in tasks:
                w = TaskItemWidget(t, highlight=(
                    self._highlight_date is not None and
                    t["due_date"] == self._highlight_date.isoformat()
                ))
                w.toggled.connect(self._toggled)
                w.delete_requested.connect(self._delete)
                w.log_requested.connect(self._log)
                w.edit_requested.connect(self._edit)
                if hasattr(w, "batch_select_changed"):
                    w.batch_select_changed.connect(self._on_batch_select)
                if self._batch_mode:
                    w.show_batch_mode(True)
                self.items_lay.addWidget(w)

    def _add(self, preset_date: date = None):
        dlg = TaskDialog(self, task_type=self.task_type, preset_date=preset_date)
        if dlg.exec():
            v = dlg.values()
            self.db.add_task(v["title"], v["description"], v["goal"],
                             self.task_type, v["priority"], v["due_date"], SOURCE_MANUAL,
                             v.get("color"), v.get("file_path"))
            self.refresh()

    def add_for_date(self, d: date):
        """달력 날짜 클릭 시 호출 — 해당 날짜 미리 설정"""
        self._add(preset_date=d)

    def _edit(self, tid):
        t = self.db.get_task(tid)
        if not t: return
        dlg = TaskDialog(self, t, task_type=self.task_type)
        if dlg.exec():
            v = dlg.values()
            self.db.update_task(tid, title=v["title"], description=v["description"],
                                goal=v["goal"], priority=v["priority"], due_date=v["due_date"],
                                color=v.get("color"), file_path=v.get("file_path"))
            self.refresh()

    def _delete(self, tid):
        t = self.db.get_task(tid)
        title = t["title"] if t else "이 태스크"
        r = QMessageBox.question(self, "삭제 확인",
            f"'{title}'\n을(를) 삭제하시겠습니까? 관련 로그도 함께 삭제됩니다.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if r == QMessageBox.StandardButton.Yes:
            self.db.delete_task(tid); self.refresh()

    def _toggled(self, tid, checked):
        self.db.toggle_complete(tid, checked)
        if checked and self.task_type == TASK_URGENT:
            dlg = _UrgentLinkDialog(self.db, tid, self)
            dlg.exec()
        self.refresh()
        self.completion_changed.emit()

    def _log(self, tid):
        LogDialog(self.db, tid, self).exec()

    def highlight_date(self, d: date | None):
        """달력에서 날짜 선택 시 해당 마감일 태스크 강조"""
        self._highlight_date = d
        self.refresh()

    def _toggle(self):
        self._collapsed = not self._collapsed
        self.body.setVisible(not self._collapsed)
        self.btn_col.setText("▶" if self._collapsed else "▼")

    # ── 드래그 앤 드롭 ─────────────────────────────────────────────────────
    def _body_drag_enter(self, e):
        if e.mimeData().hasFormat("application/x-task-id"):
            e.acceptProposedAction()

    def _body_drag_move(self, e):
        if e.mimeData().hasFormat("application/x-task-id"):
            idx = self._drop_index(e.position().toPoint())
            self._show_indicator(idx)
            e.acceptProposedAction()

    def _body_drag_leave(self, e):
        self._indicator.hide()

    def _body_drop(self, e):
        self._indicator.hide()
        if not e.mimeData().hasFormat("application/x-task-id"):
            return
        task_id = int(e.mimeData().data("application/x-task-id").data().decode())
        task = self.db.get_task(task_id)
        if not task:
            return
        # 완료된 항목이 드래그 온 경우 → 미완료로 복원
        if task["is_completed"]:
            self.db.toggle_complete(task_id, False)
            self.completion_changed.emit()
            self.refresh()
            e.acceptProposedAction()
            return
        # 다른 타입 섹션에서 온 경우 무시
        if task["task_type"] != self.task_type:
            return
        # 같은 섹션 내 재정렬
        drop_idx = self._drop_index(e.position().toPoint())
        current_ids = [
            self.items_lay.itemAt(i).widget()._id
            for i in range(self.items_lay.count())
            if self.items_lay.itemAt(i).widget()
        ]
        if task_id in current_ids:
            current_ids.remove(task_id)
        drop_idx = min(drop_idx, len(current_ids))
        current_ids.insert(drop_idx, task_id)
        self.db.update_sort_order(current_ids)
        self.refresh()
        e.acceptProposedAction()

    def _drop_index(self, pos) -> int:
        for i in range(self.items_lay.count()):
            item = self.items_lay.itemAt(i)
            if item and item.widget():
                w = item.widget()
                w_pos = w.mapTo(self.body, QPoint(0, w.height() // 2))
                if pos.y() < w_pos.y():
                    return i
        return self.items_lay.count()

    def _show_indicator(self, idx: int):
        count = self.items_lay.count()
        if count == 0:
            self._indicator.hide()
            return
        if idx < count:
            w = self.items_lay.itemAt(idx).widget()
            if w:
                y = w.mapTo(self.body, QPoint(0, 0)).y() - 3
                self._indicator.setGeometry(8, y, self.body.width() - 16, 2)
        else:
            w = self.items_lay.itemAt(count - 1).widget()
            if w:
                y = w.mapTo(self.body, QPoint(0, w.height())).y() + 1
                self._indicator.setGeometry(8, y, self.body.width() - 16, 2)
        self._indicator.show()
        self._indicator.raise_()

    # ── 정렬 ──────────────────────────────────────────────────────────────────
    def _on_sort_changed(self):
        self._sort_mode = self.cb_sort.currentData()
        self.refresh()

    def _apply_sort(self, tasks: list) -> list:
        mode = self._sort_mode
        if mode == "due_asc":
            return sorted(tasks, key=lambda t: t["due_date"] or "9999-99-99")
        if mode == "due_desc":
            return sorted(tasks, key=lambda t: t["due_date"] or "0000-00-00", reverse=True)
        if mode == "priority":
            return sorted(tasks, key=lambda t: t["priority"])
        if mode == "created_asc":
            return sorted(tasks, key=lambda t: t["created_at"])
        if mode == "created_desc":
            return sorted(tasks, key=lambda t: t["created_at"], reverse=True)
        if mode == "title":
            return sorted(tasks, key=lambda t: (t["title"] or "").lower())
        return list(tasks)   # "default" → DB 순서 유지

    # ── 배치 선택 모드 ────────────────────────────────────────────────────────
    def _toggle_batch_mode(self):
        self._batch_mode = not self._batch_mode
        self._batch_selected.clear()
        self.batch_bar.setVisible(self._batch_mode)
        self.btn_batch.setText("☑" if self._batch_mode else "☐")
        self.lbl_batch_count.setText("0개 선택")
        for i in range(self.items_lay.count()):
            w = self.items_lay.itemAt(i).widget()
            if w and hasattr(w, "show_batch_mode"):
                w.show_batch_mode(self._batch_mode)

    def _on_batch_select(self, tid: int, checked: bool):
        if checked:
            self._batch_selected.add(tid)
        else:
            self._batch_selected.discard(tid)
        self.lbl_batch_count.setText(f"{len(self._batch_selected)}개 선택")

    def _batch_select_all(self):
        self._batch_selected.clear()
        for i in range(self.items_lay.count()):
            w = self.items_lay.itemAt(i).widget()
            if w and hasattr(w, "_id") and w.isVisible():
                self._batch_selected.add(w._id)
                if hasattr(w, "_batch_chk"):
                    w._batch_chk.setChecked(True)
        self.lbl_batch_count.setText(f"{len(self._batch_selected)}개 선택")

    def _batch_complete(self):
        if not self._batch_selected:
            return
        for tid in list(self._batch_selected):
            self.db.toggle_complete(tid, True)
        self._batch_selected.clear()
        self._batch_mode = False
        self.batch_bar.setVisible(False)
        self.btn_batch.setText("☐")
        self.refresh()
        self.completion_changed.emit()

    def _batch_delete(self):
        if not self._batch_selected:
            return
        count = len(self._batch_selected)
        msg = QMessageBox(self)
        msg.setWindowTitle("삭제 확인")
        msg.setText(f"선택한 {count}개 항목을 삭제하시겠습니까?")
        msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        msg.setDefaultButton(QMessageBox.StandardButton.No)
        if msg.exec() != QMessageBox.StandardButton.Yes:
            return
        for tid in list(self._batch_selected):
            self.db.delete_task(tid)
        self._batch_selected.clear()
        self._batch_mode = False
        self.batch_bar.setVisible(False)
        self.btn_batch.setText("☐")
        self.refresh()

    def set_filter(self, query: str) -> int:
        """검색어로 태스크 필터링. 표시된 항목 수 반환."""
        shown = 0
        for i in range(self.items_lay.count()):
            item = self.items_lay.itemAt(i)
            if item and item.widget():
                w = item.widget()
                if query:
                    task = self.db.get_task(w._id)
                    if task:
                        match = any(
                            query in (task[f] or "").lower()
                            for f in ("title", "description", "goal")
                        )
                        w.setVisible(match)
                        if match:
                            shown += 1
                    else:
                        w.setVisible(False)
                else:
                    w.setVisible(True)
                    shown += 1
        return shown


# ═══════════════════════════════════════════════════════════════════════════
# 11b. URGENT LINK DIALOG + COMPLETED SECTION
# ═══════════════════════════════════════════════════════════════════════════

class _UrgentLinkDialog(_MovableDialog):
    """긴급업무 완료 시 연결 과제의 진행상황에 기록 (또는 일반 로그 기록)"""

    def __init__(self, db: Database, urgent_task_id: int, parent=None):
        super().__init__(parent)
        self.db = db
        self._urgent = db.get_task(urgent_task_id)
        self._linked_id: int | None = None
        try:
            self._linked_id = self._urgent["linked_todo_id"]
        except (KeyError, TypeError):
            pass
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setMinimumWidth(440)
        self._build()
        QShortcut(QKeySequence("Escape"), self, self.reject)

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 18); lay.setSpacing(12)

        t = QLabel("긴급업무 완료")
        t.setObjectName("DialogTitle")
        t.setFont(QFont("맑은 고딕", 13, QFont.Weight.Bold))
        lay.addWidget(t)

        urgent_lbl = QLabel(f"완료: {self._urgent['title']}")
        urgent_lbl.setObjectName("TaskInfoDesc")
        urgent_lbl.setWordWrap(True)
        lay.addWidget(urgent_lbl)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setMaximumHeight(1); sep.setObjectName("Separator")
        lay.addWidget(sep)

        if self._linked_id:
            # 연결된 과제가 있으면 → 진행상황 탭에 기록
            linked_task = self.db.get_task(self._linked_id)
            linked_name = linked_task["title"] if linked_task else f"(ID {self._linked_id})"

            info_lbl = QLabel(
                f"연결 과제 '{linked_name}'의\n과제진행상황에 자동 기록됩니다."
            )
            info_lbl.setObjectName("FormLabel")
            info_lbl.setWordWrap(True)
            info_lbl.setStyleSheet("color:#a6e3a1;background:transparent;")
            lay.addWidget(info_lbl)

            note_lbl = QLabel("진행 내용 (선택)")
            note_lbl.setObjectName("FormLabel")
            lay.addWidget(note_lbl)
            self.ed = QPlainTextEdit()
            self.ed.setPlaceholderText("이번 완료에서 달성한 내용, 결과 등...")
            self.ed.setMinimumHeight(60)
            lay.addWidget(self.ed)
            lay.setStretchFactor(self.ed, 1)

            br = QHBoxLayout(); br.addStretch()
            bc = QPushButton("건너뛰기"); bc.setObjectName("SecondaryBtn")
            bc.setFixedHeight(34); bc.clicked.connect(self.reject); br.addWidget(bc)
            ba = QPushButton("진행상황 기록"); ba.setObjectName("PrimaryBtn")
            ba.setFixedHeight(34); ba.clicked.connect(self._save_linked); br.addWidget(ba)
            lay.addLayout(br)

        else:
            # 연결 없음 → 일반 로그 기록 (구 방식)
            lbl = QLabel("완료 내용을 기록할 과제/할일을 선택하세요 (선택 사항)")
            lbl.setObjectName("FormLabel"); lbl.setWordWrap(True)
            lay.addWidget(lbl)

            self.combo = QComboBox()
            todo_tasks = self.db.get_tasks(TASK_TODO, completed=False)
            self.combo.addItem("(기록 안 함)", None)
            for task in todo_tasks:
                self.combo.addItem(task["title"], task["id"])
            lay.addWidget(self.combo)

            note_lbl = QLabel("추가 메모 (선택)")
            note_lbl.setObjectName("FormLabel")
            lay.addWidget(note_lbl)
            self.ed = QPlainTextEdit()
            self.ed.setPlaceholderText("선택 사항...")
            self.ed.setMaximumHeight(60)
            lay.addWidget(self.ed)

            br = QHBoxLayout(); br.addStretch()
            bc = QPushButton("건너뛰기"); bc.setObjectName("SecondaryBtn")
            bc.setFixedHeight(34); bc.clicked.connect(self.reject); br.addWidget(bc)
            ba = QPushButton("로그 기록"); ba.setObjectName("PrimaryBtn")
            ba.setFixedHeight(34); ba.clicked.connect(self._save_general); br.addWidget(ba)
            lay.addLayout(br)

    def _save_linked(self):
        """연결 과제의 과제진행상황 탭에 progress_group 생성 + 항목 추가"""
        note = self.ed.toPlainText().strip()
        today = date.today().strftime("%m/%d")
        grp_title = f"{self._urgent['title']} ({today})"
        grp_id = self.db.add_progress_group(
            self._linked_id, grp_title,
            source_urgent_id=self._urgent["id"]
        )
        # 기본 항목: 완료 표시
        base_content = f"긴급업무 완료"
        if note:
            base_content += f": {note}"
        self.db.add_progress_log(self._linked_id, grp_id, base_content)
        self.accept()

    def _save_general(self):
        """연결 없음: 선택한 과제의 과제진행상황에 progress_group으로 기록"""
        todo_id = self.combo.currentData()
        if todo_id is not None:
            note = self.ed.toPlainText().strip()
            today = date.today().strftime("%m/%d")
            grp_title = f"{self._urgent['title']} ({today})"
            grp_id = self.db.add_progress_group(
                todo_id, grp_title,
                source_urgent_id=self._urgent["id"]
            )
            base_content = "긴급업무 완료"
            if note:
                base_content += f": {note}"
            self.db.add_progress_log(todo_id, grp_id, base_content)
        self.accept()


class _CompletedItem(QFrame):
    """완료업무 섹션의 단일 항목"""
    restore_requested = Signal(int)
    delete_requested  = Signal(int)
    edit_requested    = Signal(int)

    def __init__(self, task_row, db=None, parent=None):
        super().__init__(parent)
        self._id = task_row["id"]
        self._db = db
        self._detail_visible = False
        self.setObjectName("TaskItemCompleted")
        self.setMinimumHeight(38)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._build(task_row)

    def _build(self, r):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(10, 4, 8, 4); outer.setSpacing(0)

        # ── 메인 행 ──────────────────────────────────────────────────────────
        lay = QHBoxLayout(); lay.setSpacing(8)

        # 타입 뱃지
        badge_color = "#89b4fa" if r["task_type"] == TASK_TODO else "#f38ba8"
        badge = QLabel("과제" if r["task_type"] == TASK_TODO else "긴급")
        badge.setStyleSheet(
            f"color:{badge_color};font-size:9px;font-weight:bold;"
            f"background:transparent;padding:0 2px;"
        )
        lay.addWidget(badge)

        ic = QLabel("✅")
        ic.setStyleSheet("font-size:12px;background:transparent;")
        lay.addWidget(ic)

        title = QLabel(f"<s>{r['title']}</s>")
        title.setObjectName("TaskTitleDone")
        title.setWordWrap(True)
        lay.addWidget(title, 1)

        if r["completed_at"]:
            try:
                dt_lbl = QLabel(r["completed_at"][:10])
                dt_lbl.setStyleSheet("color:#6c7086;font-size:10px;background:transparent;")
                lay.addWidget(dt_lbl)
            except Exception:
                pass

        # 상세 보기 버튼 (내용·파일이 있을 때만)
        _desc = r["description"] or ""
        _goal = r["goal"] or ""
        _fpath = r["file_path"] or ""
        _task_files = []
        if self._db is not None:
            try:
                _task_files = self._db.get_task_files(self._id)
            except Exception:
                _task_files = []
        _has_detail = bool(_desc or _goal or _fpath or _task_files)
        if _has_detail:
            self._btn_detail = QPushButton("▸")
            self._btn_detail.setObjectName("SectionCollapseBtn")
            self._btn_detail.setFixedSize(22, 22)
            self._btn_detail.setToolTip("상세 내용 보기/접기")
            self._btn_detail.clicked.connect(self._toggle_detail)
            lay.addWidget(self._btn_detail)

        btn_r = QPushButton("↩")
        btn_r.setObjectName("SecondaryBtn")
        btn_r.setFixedSize(26, 26)
        btn_r.setToolTip("미완료로 복원")
        btn_r.clicked.connect(lambda: self.restore_requested.emit(self._id))
        lay.addWidget(btn_r)

        btn_d = QPushButton("✕")
        btn_d.setObjectName("TaskDeleteBtn")
        btn_d.setFixedSize(26, 26)
        btn_d.setToolTip("삭제")
        btn_d.clicked.connect(lambda: self.delete_requested.emit(self._id))
        lay.addWidget(btn_d)

        outer.addLayout(lay)

        # ── 상세 영역 (기본 숨김) ─────────────────────────────────────────────
        if _has_detail:
            self._detail_w = QWidget()
            self._detail_w.setStyleSheet(
                "QWidget{background:#1a1a2e;border-radius:6px;"
                "border-left:2px solid #45475a;margin:2px 0 2px 20px;padding:4px 8px;}"
            )
            dl = QVBoxLayout(self._detail_w)
            dl.setContentsMargins(8, 4, 8, 4); dl.setSpacing(3)
            if _goal:
                g = QLabel(f"🎯 {_goal}")
                g.setStyleSheet("color:#a6adc8;font-size:11px;background:transparent;")
                g.setWordWrap(True)
                dl.addWidget(g)
            if _desc:
                d = QLabel(_desc)
                d.setStyleSheet("color:#a6adc8;font-size:11px;background:transparent;")
                d.setWordWrap(True)
                dl.addWidget(d)
            # 레거시 단일 파일
            if _fpath:
                fp = _fpath
                f_lbl = QLabel(f"📎 {fp}")
                f_lbl.setStyleSheet(
                    "color:#89b4fa;font-size:10px;background:transparent;"
                    "text-decoration:underline;"
                )
                f_lbl.setWordWrap(True)
                f_lbl.setToolTip(fp)
                f_lbl.setCursor(Qt.CursorShape.PointingHandCursor)
                f_lbl.mousePressEvent = lambda _e, p=fp: open_file_path(p, self)
                dl.addWidget(f_lbl)
            # 다중 첨부 파일 (task_files 테이블)
            for tf in _task_files:
                disp_path = tf["copy_path"] or tf["original_path"] or ""
                fname = tf["filename"] or Path(disp_path).name if disp_path else "파일"
                f_row = QHBoxLayout()
                f_icon = QLabel("📎")
                f_icon.setStyleSheet("font-size:11px;background:transparent;color:#89b4fa;")
                f_row.addWidget(f_icon)
                f_lbl2 = QLabel(fname)
                f_lbl2.setStyleSheet(
                    "color:#89b4fa;font-size:10px;background:transparent;"
                    "text-decoration:underline;"
                )
                f_lbl2.setToolTip(disp_path)
                f_lbl2.setCursor(Qt.CursorShape.PointingHandCursor)
                f_lbl2.mousePressEvent = lambda _e, p=disp_path: open_file_path(p, self)
                f_row.addWidget(f_lbl2, 1)
                # 경로 복사 버튼
                btn_cp = QPushButton("⎘")
                btn_cp.setFixedSize(18, 18)
                btn_cp.setToolTip(f"경로 복사: {disp_path}")
                btn_cp.setStyleSheet(
                    "QPushButton{background:transparent;border:none;color:#6c7086;font-size:11px;}"
                    "QPushButton:hover{color:#89b4fa;}"
                )
                btn_cp.clicked.connect(lambda _, p=disp_path: QApplication.clipboard().setText(p))
                f_row.addWidget(btn_cp)
                dl.addLayout(f_row)
            self._detail_w.hide()
            outer.addWidget(self._detail_w)

    def _toggle_detail(self):
        self._detail_visible = not self._detail_visible
        self._detail_w.setVisible(self._detail_visible)
        self._btn_detail.setText("▾" if self._detail_visible else "▸")

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self._drag_start = e.position().toPoint()
        super().mousePressEvent(e)

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.edit_requested.emit(self._id)
        super().mouseDoubleClickEvent(e)

    def mouseMoveEvent(self, e):
        if not (e.buttons() & Qt.MouseButton.LeftButton):
            return
        if not hasattr(self, '_drag_start'):
            return
        if (e.position().toPoint() - self._drag_start).manhattanLength() < 12:
            return
        drag = QDrag(self)
        mime = QMimeData()
        mime.setData("application/x-task-id", str(self._id).encode())
        drag.setMimeData(mime)
        drag.exec(Qt.DropAction.MoveAction)

    def contextMenuEvent(self, e):
        menu = QMenu(self)
        menu.setStyleSheet(
            "QMenu{background:#313244;border:1px solid #45475a;border-radius:8px;padding:4px;}"
            "QMenu::item{padding:7px 18px;border-radius:6px;color:#cdd6f4;}"
            "QMenu::item:selected{background:#45475a;}"
        )
        a_r = menu.addAction("↩  미완료로 복원")
        menu.addSeparator()
        a_d = menu.addAction("🗑  삭제")
        ch = menu.exec(e.globalPos())
        if ch == a_r: self.restore_requested.emit(self._id)
        elif ch == a_d: self.delete_requested.emit(self._id)


class _YearGroup(QWidget):
    """완료업무 연도별 그룹 위젯"""
    def __init__(self, year: str, tasks: list, db, parent=None):
        super().__init__(parent)
        self.db = db
        self._year = year
        self._tasks = tasks
        self._collapsed = True
        self._build()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 4)
        lay.setSpacing(0)

        # 연도 헤더
        hdr = QWidget()
        hdr.setStyleSheet(
            "QWidget{background:#2a2a3e;border-radius:6px;margin:0px;}"
        )
        hdr_l = QHBoxLayout(hdr)
        hdr_l.setContentsMargins(12, 6, 8, 6)

        tl = QLabel(f"📁  {self._year}년  ({len(self._tasks)}건)")
        tl.setObjectName("SectionTitle")
        tl.setFont(QFont("맑은 고딕", 11, QFont.Weight.Bold))
        hdr_l.addWidget(tl)
        hdr_l.addStretch()

        self.btn_col = QPushButton("▶")
        self.btn_col.setObjectName("SectionCollapseBtn")
        self.btn_col.setFixedSize(24, 24)
        self.btn_col.setToolTip("펼치기 / 접기")
        self.btn_col.clicked.connect(self._toggle)
        hdr_l.addWidget(self.btn_col)
        lay.addWidget(hdr)

        # 항목 컨테이너
        self.body = QWidget()
        body_l = QVBoxLayout(self.body)
        body_l.setContentsMargins(4, 4, 0, 0)
        body_l.setSpacing(3)
        for t in self._tasks:
            w = _CompletedItem(t, db=self.db)
            w.restore_requested.connect(self._restore)
            w.delete_requested.connect(self._delete)
            w.edit_requested.connect(self._edit_item)
            body_l.addWidget(w)
        lay.addWidget(self.body)
        self.body.hide()

    def _toggle(self):
        self._collapsed = not self._collapsed
        self.body.setVisible(not self._collapsed)
        self.btn_col.setText("▶" if self._collapsed else "▼")

    def _restore(self, tid):
        task = self.db.get_task(tid)
        if task and task["task_type"] == TASK_URGENT:
            # 연결된 과제의 진행 그룹이 있는지 확인 → 경고
            urgent_groups = self.db.get_urgent_progress_groups(tid)
            if urgent_groups:
                grp_titles = "\n".join(f"  • {g['title']}" for g in urgent_groups[:3])
                msg = (
                    f"이 긴급업무를 미완료로 복원하면\n"
                    f"연결된 과제의 진행 그룹이 삭제됩니다:\n\n"
                    f"{grp_titles}\n\n"
                    f"계속하시겠습니까?"
                )
                r = QMessageBox.warning(
                    self, "진행 그룹 삭제 경고", msg,
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                    QMessageBox.StandardButton.No,
                )
                if r != QMessageBox.StandardButton.Yes:
                    return
                for g in urgent_groups:
                    self.db.delete_progress_group(g["id"])
        self.db.toggle_complete(tid, False)
        # 부모 CompletedSection에 refresh 요청
        p = self.parent()
        while p:
            if isinstance(p, CompletedSection):
                p.refresh(); break
            p = p.parent()

    def _delete(self, tid):
        task = self.db.get_task(tid)
        title = task["title"] if task else "이 항목"
        r = QMessageBox.question(self, "삭제 확인",
            f"'{title}'\n을(를) 영구 삭제하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if r == QMessageBox.StandardButton.Yes:
            self.db.delete_task(tid)
            p = self.parent()
            while p:
                if isinstance(p, CompletedSection):
                    p.refresh(); break
                p = p.parent()

    def _edit_item(self, tid):
        r = QMessageBox.question(self, "완료 항목 편집",
            "완료된 항목을 편집하시겠습니까?\n(완료 상태는 유지됩니다)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if r != QMessageBox.StandardButton.Yes:
            return
        task = self.db.get_task(tid)
        if not task:
            return
        dlg = TaskDialog(
            self.db, task["task_type"],
            is_edit=True, task_id=tid,
            parent=self.window()
        )
        if dlg.exec() == QDialog.DialogCode.Accepted:
            p = self.parent()
            while p:
                if isinstance(p, CompletedSection):
                    p.refresh(); break
                p = p.parent()


class CompletedSection(QWidget):
    """완료업무 섹션 — 완료된 todo/urgent 태스크 연도별 그룹화 표시"""

    def __init__(self, db: Database, parent=None):
        super().__init__(parent)
        self.db = db
        self._collapsed = True
        self.setObjectName("SectionWidget")
        self._build()
        self.refresh()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0); lay.setSpacing(0)

        hdr_w = QWidget(); hdr_w.setObjectName("SectionHeader")
        hdr_w.setStyleSheet(
            "QWidget#SectionHeader{border-left:4px solid #585b70;border-radius:8px 8px 0 0;}"
        )
        hdr_l = QVBoxLayout(hdr_w)
        hdr_l.setContentsMargins(12, 10, 10, 8)

        title_row = QHBoxLayout()
        tl = QLabel("✅  완료업무"); tl.setObjectName("SectionTitle")
        tl.setFont(QFont("맑은 고딕", 12, QFont.Weight.Bold))
        tl.setToolTip("펼치면 완료업무 검색 가능 (제목·내용·목표)")
        title_row.addWidget(tl); title_row.addStretch()
        self.lbl_count = QLabel("0건"); self.lbl_count.setObjectName("SectionStats")
        title_row.addWidget(self.lbl_count)
        self.btn_col = QPushButton("▶"); self.btn_col.setObjectName("SectionCollapseBtn")
        self.btn_col.setFixedSize(26, 26); self.btn_col.setToolTip("섹션 접기/펼치기 — 완료업무 검색창 포함")
        self.btn_col.clicked.connect(self._toggle)
        title_row.addWidget(self.btn_col)
        hdr_l.addLayout(title_row)
        lay.addWidget(hdr_w)

        self.body = QWidget()
        self.body.setAcceptDrops(True)
        self.body.dragEnterEvent = self._drag_enter
        self.body.dragMoveEvent  = self._drag_move
        self.body.dragLeaveEvent = lambda e: None
        self.body.dropEvent      = self._drop

        b_lay = QVBoxLayout(self.body)
        b_lay.setContentsMargins(8, 6, 8, 8); b_lay.setSpacing(6)

        # 검색바
        self.ed_search = QLineEdit()
        self.ed_search.setPlaceholderText("🔍  완료업무 검색... (Esc: 초기화)")
        self.ed_search.setObjectName("SearchBar")
        self.ed_search.textChanged.connect(self._search_filter)
        self.ed_search.installEventFilter(self)
        b_lay.addWidget(self.ed_search)

        self.items_lay = QVBoxLayout()
        self.items_lay.setContentsMargins(0, 0, 0, 0); self.items_lay.setSpacing(4)
        b_lay.addLayout(self.items_lay)
        self.empty_lbl = QLabel("완료된 항목이 없습니다.")
        self.empty_lbl.setObjectName("TaskInfoDesc")
        self.empty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_lbl.setStyleSheet("color:#7f849c;padding:14px 0;font-size:12px;")
        b_lay.addWidget(self.empty_lbl)
        self.no_result_lbl = QLabel("검색 결과가 없습니다.")
        self.no_result_lbl.setObjectName("TaskInfoDesc")
        self.no_result_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.no_result_lbl.setStyleSheet("color:#6c7086;padding:14px 0;font-size:12px;")
        self.no_result_lbl.hide()
        b_lay.addWidget(self.no_result_lbl)
        lay.addWidget(self.body)
        self.body.hide()

    def refresh(self):
        q = self.ed_search.text().strip().lower() if hasattr(self, "ed_search") else ""
        self._render(q)

    def _search_filter(self, text: str):
        self._render(text.strip().lower())

    def _render(self, q: str = ""):
        while self.items_lay.count():
            item = self.items_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()

        all_tasks = self.db.get_tasks(completed=True)
        tasks = [t for t in all_tasks if t["task_type"] in (TASK_TODO, TASK_URGENT)]
        self.lbl_count.setText(f"{len(tasks)}건")

        if not tasks:
            self.empty_lbl.show()
            self.no_result_lbl.hide()
            return
        self.empty_lbl.hide()

        if q:
            tasks = [t for t in tasks if
                     q in (t["title"] or "").lower() or
                     q in (t["description"] or "").lower() or
                     q in (t["goal"] or "").lower()]

        if not tasks:
            self.no_result_lbl.show()
            return
        self.no_result_lbl.hide()

        # 연도별 그룹화 (최신 연도 먼저)
        by_year: dict[str, list] = {}
        for t in tasks:
            try:
                year = (t["completed_at"] or t["created_at"] or "0000")[:4]
            except Exception:
                year = "기타"
            by_year.setdefault(year, []).append(t)

        for year in sorted(by_year.keys(), reverse=True):
            group = _YearGroup(year, by_year[year], self.db, parent=self.body)
            self.items_lay.addWidget(group)

    def eventFilter(self, obj, event):
        if obj is self.ed_search and event.type() == event.Type.KeyPress:
            if event.key() == Qt.Key.Key_Escape:
                self.ed_search.clear()
        return super().eventFilter(obj, event)

    def _toggle(self):
        self._collapsed = not self._collapsed
        self.body.setVisible(not self._collapsed)
        self.btn_col.setText("▶" if self._collapsed else "▼")

    def _drag_enter(self, e):
        if e.mimeData().hasFormat("application/x-task-id"):
            e.acceptProposedAction()

    def _drag_move(self, e):
        if e.mimeData().hasFormat("application/x-task-id"):
            e.acceptProposedAction()

    def _drop(self, e):
        if not e.mimeData().hasFormat("application/x-task-id"):
            return
        task_id = int(e.mimeData().data("application/x-task-id").data().decode())
        task = self.db.get_task(task_id)
        if task and not task["is_completed"]:
            self.db.toggle_complete(task_id, True)
            self.refresh()
        e.acceptProposedAction()


# ═══════════════════════════════════════════════════════════════════════════
# 12. MISC SECTION (기타 — 노트 카드 목록)
# ═══════════════════════════════════════════════════════════════════════════

class MiscSection(QWidget):
    def __init__(self, db: Database, header_color: str = "#585b70", parent=None):
        super().__init__(parent)
        self.db = db
        self._header_color = header_color
        self._collapsed = False
        self.setObjectName("SectionWidget")
        self._build()
        self.refresh()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0,0,0,0); lay.setSpacing(0)

        hdr_w = QWidget(); hdr_w.setObjectName("SectionHeader")
        hdr_w.setStyleSheet(
            f"QWidget#SectionHeader{{border-left:4px solid {self._header_color};"
            f"border-radius:8px 8px 0 0;}}"
        )
        hdr_l = QHBoxLayout(hdr_w)
        hdr_l.setContentsMargins(12,10,10,10)
        tl = QLabel("📌  기타"); tl.setObjectName("SectionTitle")
        tl.setFont(QFont("맑은 고딕", 12, QFont.Weight.Bold))
        hdr_l.addWidget(tl); hdr_l.addStretch()
        self.lbl_cnt = QLabel("0개"); self.lbl_cnt.setObjectName("SectionStats")
        hdr_l.addWidget(self.lbl_cnt)
        btn_col = QPushButton("▼"); btn_col.setObjectName("SectionCollapseBtn")
        btn_col.setFixedSize(26,26); btn_col.setToolTip("섹션 접기/펼치기")
        btn_col.clicked.connect(self._toggle)
        hdr_l.addWidget(btn_col)
        self.btn_col_ref = btn_col
        lay.addWidget(hdr_w)

        self.body = QWidget()
        b_lay = QVBoxLayout(self.body)
        b_lay.setContentsMargins(8,6,8,8); b_lay.setSpacing(6)
        self.items_lay = QVBoxLayout()
        self.items_lay.setContentsMargins(0,0,0,0); self.items_lay.setSpacing(6)
        b_lay.addLayout(self.items_lay)
        self.empty_lbl = QLabel("기타 항목이 없습니다.")
        self.empty_lbl.setObjectName("TaskInfoDesc")
        self.empty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_lbl.setStyleSheet("color:#7f849c;padding:14px 0;font-size:12px;")
        self.empty_lbl.hide()
        b_lay.addWidget(self.empty_lbl)
        btn_add = QPushButton("＋  기타 항목 추가")
        btn_add.setObjectName("AddTaskBtn")
        btn_add.setMinimumHeight(34)
        btn_add.setToolTip("새 기타 항목 추가")
        btn_add.clicked.connect(self._add)
        b_lay.addWidget(btn_add)
        lay.addWidget(self.body)

    def refresh(self):
        while self.items_lay.count():
            item = self.items_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        tasks = self.db.get_tasks(TASK_MISC)
        self.lbl_cnt.setText(f"{len(tasks)}개")
        if not tasks:
            self.empty_lbl.show()
        else:
            self.empty_lbl.hide()
            for t in tasks:
                w = MiscItemWidget(t)
                w.delete_requested.connect(self._delete)
                w.edit_requested.connect(self._edit)
                self.items_lay.addWidget(w)

    def _add(self):
        dlg = TaskDialog(self, task_type=TASK_MISC)
        if dlg.exec():
            v = dlg.values()
            self.db.add_task(v["title"], v["description"], v["goal"],
                             TASK_MISC, v["priority"], file_path=v.get("file_path"))
            self.refresh()

    def _edit(self, tid):
        t = self.db.get_task(tid)
        if not t: return
        dlg = TaskDialog(self, t, task_type=TASK_MISC)
        if dlg.exec():
            v = dlg.values()
            self.db.update_task(tid, title=v["title"], description=v["description"],
                                goal=v["goal"], priority=v["priority"],
                                file_path=v["file_path"])
            self.refresh()

    def _delete(self, tid):
        t = self.db.get_task(tid)
        title = t["title"] if t else "이 항목"
        r = QMessageBox.question(self, "삭제 확인", f"'{title}' 을(를) 삭제하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if r == QMessageBox.StandardButton.Yes:
            self.db.delete_task(tid); self.refresh()

    def _toggle(self):
        self._collapsed = not self._collapsed
        self.body.setVisible(not self._collapsed)
        self.btn_col_ref.setText("▶" if self._collapsed else "▼")


# ═══════════════════════════════════════════════════════════════════════════
# 12-B. SCHEDULE SECTION (단기 일정 + 휴가/교육)
# ═══════════════════════════════════════════════════════════════════════════

class ScheduleItemWidget(QFrame):
    """일정 아이템 위젯"""
    edit_requested   = Signal(int)
    delete_requested = Signal(int)

    def __init__(self, sched_row, parent=None):
        super().__init__(parent)
        self._id = sched_row["id"]
        self.setObjectName("ScheduleItem")
        self._build(sched_row)

    def _build(self, r):
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 8, 0)
        lay.setSpacing(0)

        etype = r["event_type"]
        color = SCHED_COLORS.get(etype, "#89b4fa")
        icon  = SCHED_ICONS.get(etype, "📅")

        # 컬러 바
        bar = QFrame(); bar.setFixedWidth(4)
        bar.setStyleSheet(f"background:{color};border-radius:2px;margin:6px 0 6px 6px;")
        lay.addWidget(bar)

        inner = QVBoxLayout()
        inner.setContentsMargins(10, 7, 0, 7)
        inner.setSpacing(2)

        # 제목 행
        name_row = QHBoxLayout(); name_row.setSpacing(6)
        name_lbl = QLabel(f"{icon} {r['name']}")
        name_lbl.setObjectName("ScheduleItemName")
        name_lbl.setWordWrap(True)
        name_row.addWidget(name_lbl, 1)

        btn_edit = QPushButton("✎"); btn_edit.setObjectName("LogEditBtn")
        btn_edit.setFixedSize(22, 22)
        btn_edit.setToolTip("편집 (더블클릭으로도 가능)")
        btn_edit.clicked.connect(lambda: self.edit_requested.emit(self._id))
        name_row.addWidget(btn_edit)

        btn_del = QPushButton("✕"); btn_del.setObjectName("TaskDeleteBtn")
        btn_del.setFixedSize(22, 22)
        btn_del.clicked.connect(lambda: self.delete_requested.emit(self._id))
        name_row.addWidget(btn_del)
        inner.addLayout(name_row)

        # 날짜/시간/장소 메타
        meta_parts = []
        try:
            sd = date.fromisoformat(r["event_date"])
            if r["end_date"] and r["end_date"] != r["event_date"]:
                ed = date.fromisoformat(r["end_date"])
                meta_parts.append(f"📆 {sd.strftime('%m/%d')} ~ {ed.strftime('%m/%d')}")
            else:
                meta_parts.append(f"📆 {sd.strftime('%Y.%m.%d')}")
        except Exception:
            pass
        if r["start_time"]: meta_parts.append(f"🕐 {r['start_time']}")
        if r["location"]:   meta_parts.append(f"📍 {r['location']}")

        if meta_parts:
            meta_lbl = QLabel("  ".join(meta_parts))
            meta_lbl.setObjectName("ScheduleItemMeta")
            inner.addWidget(meta_lbl)

        if r["content"]:
            cnt = QLabel(r["content"]); cnt.setObjectName("ScheduleItemContent")
            cnt.setWordWrap(True)
            cnt.setTextInteractionFlags(
                Qt.TextInteractionFlag.TextSelectableByMouse |
                Qt.TextInteractionFlag.TextSelectableByKeyboard
            )
            inner.addWidget(cnt)

        lay.addLayout(inner, 1)

    def mouseDoubleClickEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            self.edit_requested.emit(self._id)


class ScheduleSection(QWidget):
    """단기 일정 + 휴가/교육 섹션"""

    def __init__(self, db: Database, calendar_widget=None,
                 header_color: str = "#f9e2af", parent=None):
        super().__init__(parent)
        self.db  = db
        self.cal = calendar_widget
        self._header_color = header_color
        self._collapsed = False
        self.setObjectName("SectionWidget")
        self._build()
        self.refresh()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0); lay.setSpacing(0)

        # 헤더
        hdr_w = QWidget(); hdr_w.setObjectName("SectionHeader")
        hdr_w.setStyleSheet(
            f"QWidget#SectionHeader{{border-left:4px solid {self._header_color};"
            f"border-radius:8px 8px 0 0;}}"
        )
        hdr_l = QHBoxLayout(hdr_w)
        hdr_l.setContentsMargins(12, 10, 10, 10)

        tl = QLabel("📅  일정 관리"); tl.setObjectName("SectionTitle")
        tl.setFont(QFont("맑은 고딕", 12, QFont.Weight.Bold))
        hdr_l.addWidget(tl); hdr_l.addStretch()

        self.lbl_cnt = QLabel("0개"); self.lbl_cnt.setObjectName("SectionStats")
        hdr_l.addWidget(self.lbl_cnt)

        btn_col = QPushButton("▼"); btn_col.setObjectName("SectionCollapseBtn")
        btn_col.setFixedSize(26, 26); btn_col.setToolTip("섹션 접기/펼치기")
        btn_col.clicked.connect(self._toggle)
        hdr_l.addWidget(btn_col)
        self.btn_col_ref = btn_col
        lay.addWidget(hdr_w)

        # 바디
        self.body = QWidget()
        b_lay = QVBoxLayout(self.body)
        b_lay.setContentsMargins(8, 6, 8, 8); b_lay.setSpacing(4)

        self.items_lay = QVBoxLayout()
        self.items_lay.setContentsMargins(0, 0, 0, 0); self.items_lay.setSpacing(4)
        b_lay.addLayout(self.items_lay)

        self.empty_lbl = QLabel("등록된 일정이 없습니다.")
        self.empty_lbl.setObjectName("TaskInfoDesc")
        self.empty_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.empty_lbl.setStyleSheet("color:#6c7086;padding:12px 0;font-size:12px;")
        b_lay.addWidget(self.empty_lbl)

        # 추가 버튼 (단기일정 / 휴가 / 교육 통합)
        btn_add = QPushButton("＋  일정 추가")
        btn_add.setObjectName("AddTaskBtn"); btn_add.setMinimumHeight(34)
        btn_add.setToolTip("새 일정 추가")
        btn_add.clicked.connect(lambda: self._add())
        b_lay.addWidget(btn_add)

        lay.addWidget(self.body)

    def refresh(self):
        while self.items_lay.count():
            item = self.items_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()

        rows = self.db.get_schedules()
        today_str = date.today().isoformat()

        # 지난 단기 일정 제외, 진행 중인 휴가/교육 포함
        upcoming = []
        for r in rows:
            end = r["end_date"] or r["event_date"]
            if end >= today_str:
                upcoming.append(r)

        self.lbl_cnt.setText(f"{len(upcoming)}개")

        if not upcoming:
            self.empty_lbl.show()
        else:
            self.empty_lbl.hide()
            for r in upcoming:
                w = ScheduleItemWidget(r)
                w.edit_requested.connect(self._edit)
                w.delete_requested.connect(self._delete)
                self.items_lay.addWidget(w)

    def set_filter(self, query: str) -> int:
        """검색어로 일정 필터링. 표시된 항목 수 반환."""
        shown = 0
        for i in range(self.items_lay.count()):
            item = self.items_lay.itemAt(i)
            if item and item.widget():
                w = item.widget()
                if query:
                    sched = self.db.get_schedule_by_id(w._id)
                    if sched:
                        match = any(
                            query in (sched[f] or "").lower()
                            for f in ("name", "location", "content")
                        )
                        w.setVisible(match)
                        if match:
                            shown += 1
                    else:
                        w.setVisible(False)
                else:
                    w.setVisible(True)
                    shown += 1
        return shown

    def _add(self, default_type: str = SCHED_SINGLE, preset_date: date = None):
        dlg = ScheduleDialog(self, preset_date=preset_date)
        # preset type
        dlg._pick_type(default_type)
        if dlg.exec():
            v = dlg.values()
            self.db.add_schedule(**v)
            self.refresh()
            if self.cal: self.cal.refresh()

    def add_for_date(self, d: date):
        """달력에서 날짜 클릭 시 호출"""
        self._add(SCHED_SINGLE, preset_date=d)

    def _edit(self, sid):
        rows = self.db.get_schedules()
        row  = next((r for r in rows if r["id"] == sid), None)
        if not row: return
        dlg = ScheduleDialog(self, sched_data=row)
        if dlg.exec():
            v = dlg.values()
            self.db.update_schedule(sid, **v)
            self.refresh()
            if self.cal: self.cal.refresh()

    def _delete(self, sid):
        rows = self.db.get_schedules()
        row  = next((r for r in rows if r["id"] == sid), None)
        name = row["name"] if row else "이 일정"
        r = QMessageBox.question(self, "삭제 확인", f"'{name}' 일정을 삭제하시겠습니까?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No)
        if r == QMessageBox.StandardButton.Yes:
            self.db.delete_schedule(sid)
            self.refresh()
            if self.cal: self.cal.refresh()

    def _toggle(self):
        self._collapsed = not self._collapsed
        self.body.setVisible(not self._collapsed)
        self.btn_col_ref.setText("▶" if self._collapsed else "▼")


# ═══════════════════════════════════════════════════════════════════════════
# 12-B. COWORK TODAY SECTION (오늘의 팀 일정 상시 표시)
# ═══════════════════════════════════════════════════════════════════════════

class CoworkTodaySection(QWidget):
    """오늘 + 내일 카카오워크 팀 일정을 항상 노출하는 고정 섹션"""

    def __init__(self, db: Database, parent=None):
        super().__init__(parent)
        self.db             = db
        self._collapsed     = False
        self._tmr_expanded  = False
        self.setObjectName("SectionWidget")
        self._build()
        self.refresh()

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # ── 헤더 ────────────────────────────────────────────────────────
        hdr = QWidget(); hdr.setObjectName("SectionHeader")
        hdr.setMinimumHeight(36)
        hl  = QHBoxLayout(hdr)
        hl.setContentsMargins(12, 4, 12, 4); hl.setSpacing(8)

        self._col_btn = QPushButton("▼")
        self._col_btn.setFixedSize(18, 18)
        self._col_btn.setObjectName("CollapseBtn")
        self._col_btn.setToolTip("섹션 접기/펼치기")
        self._col_btn.clicked.connect(self._toggle_collapse)
        hl.addWidget(self._col_btn)

        ico = QLabel("🏢")
        ico.setStyleSheet("background:transparent;font-size:13px;")
        hl.addWidget(ico)

        title = QLabel("오늘의 연구실 일정")
        title.setObjectName("SectionTitle")
        title.setFont(QFont("맑은 고딕", 11, QFont.Weight.Bold))
        hl.addWidget(title)
        hl.addStretch()

        self._count_lbl = QLabel("")
        self._count_lbl.setObjectName("SectionStats")
        hl.addWidget(self._count_lbl)

        lay.addWidget(hdr)

        # ── 컨텐츠 ──────────────────────────────────────────────────────
        self._content = QWidget()
        cl = QVBoxLayout(self._content)
        cl.setContentsMargins(8, 4, 8, 8); cl.setSpacing(2)

        # 오늘 이벤트 컨테이너
        self._today_cont = QWidget()
        self._today_lay  = QVBoxLayout(self._today_cont)
        self._today_lay.setContentsMargins(0, 0, 0, 0)
        self._today_lay.setSpacing(2)
        cl.addWidget(self._today_cont)

        # 내일 토글 버튼
        self._tmr_btn = QPushButton("▶  내일")
        self._tmr_btn.setObjectName("SecondaryBtn")
        self._tmr_btn.setMinimumHeight(28)
        self._tmr_btn.setMinimumWidth(200)
        self._tmr_btn.setStyleSheet(
            "text-align:left;padding:4px 8px;font-size:11px;"
            "background:transparent;border:none;color:#7f849c;"
        )
        self._tmr_btn.clicked.connect(self._toggle_tomorrow)
        cl.addWidget(self._tmr_btn)

        # 내일 이벤트 컨테이너 (기본 숨김)
        self._tmr_cont = QWidget()
        self._tmr_cont.hide()
        self._tmr_lay  = QVBoxLayout(self._tmr_cont)
        self._tmr_lay.setContentsMargins(0, 0, 0, 0)
        self._tmr_lay.setSpacing(2)
        cl.addWidget(self._tmr_cont)

        cl.addStretch()
        lay.addWidget(self._content)

    # ── 토글 ────────────────────────────────────────────────────────────
    def _toggle_collapse(self):
        self._collapsed = not self._collapsed
        self._content.setVisible(not self._collapsed)
        self._col_btn.setText("▶" if self._collapsed else "▼")

    def _toggle_tomorrow(self):
        self._tmr_expanded = not self._tmr_expanded
        self._tmr_cont.setVisible(self._tmr_expanded)
        prefix = "▼  내일" if self._tmr_expanded else "▶  내일"
        self._tmr_btn.setText(prefix + self._tmr_btn.text().split("내일", 1)[1])

    # ── 데이터 갱신 ─────────────────────────────────────────────────────
    @staticmethod
    def _clear(layout: QVBoxLayout):
        while layout.count():
            item = layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def refresh(self):
        today    = date.today()
        tomorrow = today + timedelta(days=1)
        ical_map = self.db.get_ical_date_map()

        today_ev = ical_map.get(today.isoformat(), [])
        tmr_ev   = ical_map.get(tomorrow.isoformat(), [])

        self._clear(self._today_lay)
        self._clear(self._tmr_lay)

        if not today_ev and not tmr_ev:
            ph = QLabel("연구실 일정 없음  (iCal 미연결 시 설정 ⚙ 에서 연결)")
            ph.setStyleSheet(
                "color:#6c7086;font-size:10px;padding:4px 2px;background:transparent;")
            self._today_lay.addWidget(ph)
            self._count_lbl.setText("없음")
            self._tmr_btn.hide()
        else:
            self._tmr_btn.show()
            self._populate(self._today_lay, today_ev)
            self._populate(self._tmr_lay,   tmr_ev)
            self._count_lbl.setText(f"오늘 {len(today_ev)}건")
            prefix = "▼  내일" if self._tmr_expanded else "▶  내일"
            self._tmr_btn.setText(f"{prefix} ({len(tmr_ev)}건)" if tmr_ev else f"{prefix} (없음)")

    def _populate(self, layout: QVBoxLayout, events: list):
        def sort_key(ev):
            pri, _ = _ical_classify(ev["summary"], ev["start_time_str"])
            return (pri, ev["start_time_str"] or "")

        for ev in sorted(events, key=sort_key):
            summary    = ev["summary"]
            pri, color = _ical_classify(summary, ev["start_time_str"])
            time_label = _ical_time_label(ev)   # "" for all-day

            row_w = QWidget()
            row_w.setStyleSheet("QWidget{background:transparent;margin:1px 0;}")
            row_l = QHBoxLayout(row_w)
            row_l.setContentsMargins(2, 1, 4, 1); row_l.setSpacing(6)

            bar = QFrame()
            bar.setFrameShape(QFrame.Shape.VLine)
            bar.setFixedWidth(3)
            bar.setStyleSheet(f"background:{color};border:none;")
            row_l.addWidget(bar)

            txt = summary if not time_label else f"{time_label}  {summary}"
            lbl = QLabel(txt)
            lbl.setStyleSheet(
                f"color:{color};font-size:11px;font-weight:bold;background:transparent;")
            lbl.setWordWrap(True)
            lbl.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            row_l.addWidget(lbl, 1)
            layout.addWidget(row_w)


# ═══════════════════════════════════════════════════════════════════════════
# 13. UPDATE PANEL (파일 연동 컨트롤 바)
# ═══════════════════════════════════════════════════════════════════════════

# 13-C. EXPORT DIALOG (업무 보고용 내보내기)
# ═══════════════════════════════════════════════════════════════════════════

class ExportDialog(_MovableDialog):
    """과제/할 일/긴급 업무/출장·교육 일정을 보고 양식으로 내보내기"""

    def __init__(self, db: Database, settings: QSettings, parent=None):
        super().__init__(parent)
        self.db       = db
        self.settings = settings
        self._items: list[dict] = []
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setMinimumWidth(640)
        self.setMinimumHeight(480)
        self.resize(720, 600)
        self._build()
        self._load_items()
        QShortcut(QKeySequence("Escape"), self, self.reject)

    def _lbl(self, text):
        l = QLabel(text); l.setObjectName("FormLabel"); return l

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        title = QLabel("📤  업무 내보내기 (보고용)")
        title.setObjectName("DialogTitle")
        title.setFont(QFont("맑은 고딕", 14, QFont.Weight.Bold))
        lay.addWidget(title)

        note = QLabel(
            "✦ 각 항목의 제목·내용·날짜를 직접 수정 가능 / 체크 해제 시 제외\n"
            "✦ 이전 Export와 동일한 항목에는 '(이후 진행사항 없음)'이 자동 표시됩니다"
        )
        note.setObjectName("TaskInfoDesc")
        note.setStyleSheet("color:#7f849c;font-size:11px;padding:2px 0 4px 0;")
        lay.addWidget(note)

        # 그룹사명 일괄 입력
        g_row = QHBoxLayout(); g_row.setSpacing(8)
        g_row.addWidget(self._lbl("기본 그룹사:"))
        self.ed_group = QLineEdit()
        self.ed_group.setPlaceholderText("예: A그룹사  →  전체 적용 클릭")
        self.ed_group.setText(self.settings.value("export_group", ""))
        g_row.addWidget(self.ed_group, 1)
        btn_fill = QPushButton("전체 적용")
        btn_fill.setObjectName("SecondaryBtn")
        btn_fill.setFixedHeight(30)
        btn_fill.setToolTip("위 그룹사명을 모든 항목에 일괄 적용")
        btn_fill.clicked.connect(self._fill_all_groups)
        g_row.addWidget(btn_fill)
        lay.addLayout(g_row)

        # 항목 목록 (스크롤)
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setMinimumHeight(300)
        self._cont = QWidget()
        self._items_lay = QVBoxLayout(self._cont)
        self._items_lay.setContentsMargins(4, 4, 4, 4)
        self._items_lay.setSpacing(4)
        self._scroll.setWidget(self._cont)
        lay.addWidget(self._scroll, 1)

        # 버튼 행
        btn_row = QHBoxLayout(); btn_row.addStretch()

        btn_prev = QPushButton("👁  미리보기")
        btn_prev.setObjectName("SecondaryBtn")
        btn_prev.setFixedHeight(36)
        btn_prev.clicked.connect(self._show_preview)
        btn_row.addWidget(btn_prev)

        btn_exp = QPushButton("📤  내보내기")
        btn_exp.setObjectName("PrimaryBtn")
        btn_exp.setFixedHeight(36)
        btn_exp.clicked.connect(self._export)
        btn_row.addWidget(btn_exp)

        btn_close = QPushButton("닫기")
        btn_close.setObjectName("SecondaryBtn")
        btn_close.setFixedHeight(36)
        btn_close.clicked.connect(self.reject)
        btn_row.addWidget(btn_close)

        lay.addLayout(btn_row)

    def _fill_all_groups(self):
        """기본 그룹사명을 모든 항목의 그룹 필드에 일괄 적용"""
        group = self.ed_group.text().strip()
        for item in self._items:
            item["ed_group"].setText(group)

    # ── 항목 로드 ───────────────────────────────────────────────────────────

    def _load_items(self):
        while self._items_lay.count():
            item = self._items_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self._items.clear()

        prev_titles = self._load_prev_titles()

        self._add_sec_header("📝  과제 / 할 일 목록")
        for t in self.db.get_tasks(TASK_TODO):
            self._add_task_item(t, prev_titles)

        self._add_sec_header("🚨  이번주 긴급 업무")
        for t in self.db.get_tasks(TASK_URGENT):
            self._add_task_item(t, prev_titles)

        self._add_sec_header("📅  일정 관리 (출장 / 교육)")
        for s in self.db.get_schedules():
            if s["event_type"] in (SCHED_TRIP, SCHED_TRAINING):
                self._add_sched_item(s, prev_titles)

        self._items_lay.addStretch()

    def _add_sec_header(self, text: str):
        lbl = QLabel(text)
        lbl.setObjectName("SectionTitle")
        lbl.setFont(QFont("맑은 고딕", 11, QFont.Weight.Bold))
        lbl.setStyleSheet("padding: 8px 0 2px 2px;")
        self._items_lay.addWidget(lbl)

    def _add_task_item(self, t, prev_titles: set):
        frame = QFrame(); frame.setObjectName("TaskItem")
        fl = QVBoxLayout(frame)
        fl.setContentsMargins(8, 6, 8, 6); fl.setSpacing(4)

        r1 = QHBoxLayout(); r1.setSpacing(6)
        chk = QCheckBox(); chk.setObjectName("TaskCheck"); chk.setChecked(True)
        r1.addWidget(chk)
        ed_group = QLineEdit(self.settings.value("export_group", ""))
        ed_group.setPlaceholderText("그룹사")
        ed_group.setFixedWidth(100)
        r1.addWidget(ed_group)
        ed_title = QLineEdit(t["title"])
        r1.addWidget(ed_title, 1)
        fl.addLayout(r1)

        r2 = QHBoxLayout(); r2.setSpacing(6)
        desc = t["description"] or ""
        desc_placeholder = "(업무 내용 미입력 — 직접 기재하세요)"
        unchanged = t["title"] in prev_titles
        content_text = desc + (" (이후 진행사항 없음)" if unchanged and desc else "")
        ed_content = QLineEdit(content_text)
        ed_content.setPlaceholderText(desc_placeholder if not desc else "")
        r2.addWidget(ed_content, 1)

        due_str = ""
        if t["due_date"]:
            try:
                d = date.fromisoformat(t["due_date"])
                if t["is_completed"]:
                    due_str = f"기완료: {d.strftime('%Y.%m.%d')}"
                else:
                    due_str = f"마감예정: {d.strftime('%Y.%m.%d')}"
            except Exception:
                pass
        ed_due = QLineEdit(due_str)
        ed_due.setPlaceholderText("마감예정: YYYY.MM.DD")
        ed_due.setFixedWidth(150)
        r2.addWidget(ed_due)
        fl.addLayout(r2)

        self._items_lay.addWidget(frame)
        self._items.append({"chk": chk, "ed_group": ed_group, "ed_title": ed_title,
                             "ed_content": ed_content, "ed_due": ed_due})

    def _add_sched_item(self, s, prev_titles: set):
        frame = QFrame(); frame.setObjectName("TaskItem")
        fl = QVBoxLayout(frame)
        fl.setContentsMargins(8, 6, 8, 6); fl.setSpacing(4)

        r1 = QHBoxLayout(); r1.setSpacing(6)
        chk = QCheckBox(); chk.setObjectName("TaskCheck"); chk.setChecked(True)
        r1.addWidget(chk)
        ed_group = QLineEdit(self.settings.value("export_group", ""))
        ed_group.setPlaceholderText("그룹사")
        ed_group.setFixedWidth(100)
        r1.addWidget(ed_group)
        type_lbl = SCHED_LABELS.get(s["event_type"], "")
        ed_title = QLineEdit(f"[{type_lbl}] {s['name']}" if type_lbl else s["name"])
        r1.addWidget(ed_title, 1)
        fl.addLayout(r1)

        r2 = QHBoxLayout(); r2.setSpacing(6)
        content = s["content"] or ""
        desc_placeholder = "(일정 내용 미입력 — 직접 기재하세요)"
        unchanged = s["name"] in prev_titles
        content_text = content + (" (이후 진행사항 없음)" if unchanged and content else "")
        ed_content = QLineEdit(content_text)
        ed_content.setPlaceholderText(desc_placeholder if not content else "")
        r2.addWidget(ed_content, 1)

        due_str = ""
        try:
            sd = date.fromisoformat(s["event_date"])
            if s["end_date"] and s["end_date"] != s["event_date"]:
                ed_d = date.fromisoformat(s["end_date"])
                due_str = f"{sd.strftime('%m/%d')}~{ed_d.strftime('%m/%d')}"
            else:
                due_str = sd.strftime("%Y.%m.%d")
        except Exception:
            pass
        ed_due = QLineEdit(due_str)
        ed_due.setFixedWidth(150)
        r2.addWidget(ed_due)
        fl.addLayout(r2)

        self._items_lay.addWidget(frame)
        self._items.append({"chk": chk, "ed_group": ed_group, "ed_title": ed_title,
                             "ed_content": ed_content, "ed_due": ed_due})

    def _load_prev_titles(self) -> set:
        """가장 최근 Export 파일에서 업무 제목 추출 (이후 진행사항 없음 판별용)"""
        if not EXPORT_DIR.exists():
            return set()
        files = sorted(EXPORT_DIR.glob("*.txt"), reverse=True)
        if not files:
            return set()
        titles = set()
        try:
            with open(files[0], encoding="utf-8-sig", errors="replace") as f:
                for line in f:
                    line = line.strip()
                    if line.startswith("[") and "]" in line:
                        parts = line.split("]", 1)
                        if len(parts) == 2:
                            titles.add(parts[1].strip())
        except Exception:
            pass
        return titles

    # ── 미리보기 / 내보내기 ─────────────────────────────────────────────────

    def _build_text(self) -> str:
        lines = []
        for item in self._items:
            if not item["chk"].isChecked():
                continue
            group   = item["ed_group"].text().strip() or "그룹사"
            title   = item["ed_title"].text().strip()
            content = item["ed_content"].text().strip()
            due     = item["ed_due"].text().strip()
            if not title:
                continue
            lines.append(f"[{group}] {title}")
            body = f"- {content}" if content else "-"
            if due:
                body += f" ({due})"
            lines.append(body)
            lines.append("")
        return "\n".join(lines).rstrip()

    def _show_preview(self):
        text = self._build_text()
        dlg = _MovableDialog(self)
        dlg.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        dlg.setModal(True)
        dlg.setMinimumWidth(500)
        dlg.setMinimumHeight(400)
        v = QVBoxLayout(dlg)
        v.setContentsMargins(20, 16, 20, 16); v.setSpacing(10)
        t = QLabel("📄  내보내기 미리보기")
        t.setObjectName("DialogTitle")
        t.setFont(QFont("맑은 고딕", 13, QFont.Weight.Bold))
        v.addWidget(t)
        te = QPlainTextEdit(text if text else "(내보낼 항목 없음)")
        te.setReadOnly(True)
        v.addWidget(te)
        btn = QPushButton("닫기")
        btn.setObjectName("SecondaryBtn")
        btn.setFixedHeight(36)
        btn.clicked.connect(dlg.accept)
        v.addWidget(btn, alignment=Qt.AlignmentFlag.AlignRight)
        dlg.exec()

    def _export(self):
        group = self.ed_group.text().strip()
        if group:
            self.settings.setValue("export_group", group)

        text = self._build_text()
        if not text:
            QMessageBox.information(self, "내보내기", "내보낼 항목이 없습니다.\n항목을 선택하거나 추가하세요.")
            return

        EXPORT_DIR.mkdir(exist_ok=True)
        filename = date.today().strftime("%Y.%m.%d") + ".txt"
        filepath = EXPORT_DIR / filename
        try:
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(text)
            QMessageBox.information(
                self, "내보내기 완료",
                f"파일이 저장되었습니다:\n{filepath}\n\n"
                f"폴더: {EXPORT_DIR}"
            )
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "저장 오류", f"저장 실패:\n{e}")


# ═══════════════════════════════════════════════════════════════════════════
# 14. OPTIONS DIALOG
# ═══════════════════════════════════════════════════════════════════════════

class OptionsDialog(_MovableDialog):
    """설정 다이얼로그 — 테마/화면, 섹션설정, 알림 3개 탭"""

    theme_changed          = Signal(str)   # theme key
    opacity_changed        = Signal(int)   # 0-100
    section_changed        = Signal()      # 섹션 가시성 변경
    notif_changed          = Signal(bool)  # 알림 on/off
    font_size_changed      = Signal(int)   # pt
    window_height_changed  = Signal(int)   # px (0=자동)

    def __init__(self, settings: QSettings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setMinimumWidth(520)
        self._build()
        self._load_settings()
        QShortcut(QKeySequence("Escape"), self, self.reject)

    # ── UI 구성 ─────────────────────────────────────────────────────────────
    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        # 제목
        title = QLabel("⚙  설정")
        title.setObjectName("DialogTitle")
        title.setFont(QFont("맑은 고딕", 14, QFont.Weight.Bold))
        lay.addWidget(title)

        # 탭 위젯
        self.tabs = QTabWidget()
        self.tabs.setObjectName("OptionsTab")
        lay.addWidget(self.tabs)

        self._build_tab_theme()
        self._build_tab_sections()
        self._build_tab_notif()
        self._build_tab_cowork()

        # 버튼
        br = QHBoxLayout(); br.addStretch()
        bc = QPushButton("닫기"); bc.setObjectName("SecondaryBtn")
        bc.setFixedHeight(38); bc.clicked.connect(self.accept)
        br.addWidget(bc)
        lay.addLayout(br)

    def _lbl(self, text):
        l = QLabel(text); l.setObjectName("FormLabel"); return l

    def _build_tab_theme(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(16, 16, 16, 8)
        lay.setSpacing(14)

        # ── 투명도 ────────────────────────────────────────────
        lay.addWidget(self._lbl("창 투명도"))
        op_row = QHBoxLayout()
        self.slider_opacity = QSlider(Qt.Orientation.Horizontal)
        self.slider_opacity.setRange(20, 100)
        self.slider_opacity.setSingleStep(5)
        self.slider_opacity.setTickInterval(10)
        self.slider_opacity.setTickPosition(QSlider.TickPosition.TicksBelow)
        op_row.addWidget(self.slider_opacity, 1)
        self.lbl_opacity = QLabel("100%")
        self.lbl_opacity.setFixedWidth(42)
        self.lbl_opacity.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        op_row.addWidget(self.lbl_opacity)
        self.slider_opacity.valueChanged.connect(self._on_opacity)
        lay.addLayout(op_row)

        # ── 글자 크기 ───────────────────────────────────────────
        lay.addWidget(self._lbl("글자 크기"))
        fs_row = QHBoxLayout(); fs_row.setSpacing(8)
        self.spin_fontsize = QSpinBox()
        self.spin_fontsize.setRange(8, 18)
        self.spin_fontsize.setSingleStep(1)
        self.spin_fontsize.setValue(10)
        self.spin_fontsize.setSuffix(" pt")
        self.spin_fontsize.setFixedWidth(90)
        self.spin_fontsize.valueChanged.connect(self._on_fontsize)
        fs_row.addWidget(self.spin_fontsize)
        fs_rec = QLabel("권장: 10–11pt  (캘린더 깨짐 방지)")
        fs_rec.setStyleSheet("color:#6c7086;font-size:11px;")
        fs_row.addWidget(fs_rec)
        fs_row.addStretch()
        lay.addLayout(fs_row)

        # ── 글씨체 ─────────────────────────────────────────────
        lay.addWidget(self._lbl("글씨체"))
        self.combo_font = QComboBox()
        from PySide6.QtGui import QFontDatabase
        families = sorted(QFontDatabase.families())
        self.combo_font.addItems(families)
        cur_ff = self.settings.value("font_family", "맑은 고딕")
        if cur_ff in families:
            self.combo_font.setCurrentText(cur_ff)
        self.combo_font.currentTextChanged.connect(self._on_font_family)
        lay.addWidget(self.combo_font)

        # ── 창 너비 ────────────────────────────────────────────
        lay.addWidget(self._lbl("창 너비 (px)"))
        w_row = QHBoxLayout()
        self.spin_width = QSpinBox()
        self.spin_width.setRange(380, 1400)
        self.spin_width.setSingleStep(20)
        self.spin_width.setValue(WINDOW_WIDTH)
        self.spin_width.setFixedWidth(100)
        self.spin_width.valueChanged.connect(lambda v: self.settings.setValue("window_width", v))
        w_row.addWidget(self.spin_width)
        w_rec = QLabel("권장: 680–780px  (캘린더 셀 충분 확보)")
        w_rec.setStyleSheet("color:#6c7086;font-size:11px;")
        w_row.addWidget(w_rec)
        w_row.addStretch()
        lay.addLayout(w_row)

        # ── 창 높이 ────────────────────────────────────────────
        lay.addWidget(self._lbl("창 높이 (px)"))
        h_row = QHBoxLayout()
        self.spin_height = QSpinBox()
        self.spin_height.setRange(0, 2160)
        self.spin_height.setSingleStep(20)
        self.spin_height.setSpecialValueText("자동 (화면 96%)")
        self.spin_height.setFixedWidth(150)
        # 0(자동)↔400(최소 수동) 사이 죽은 구간 방지
        self._prev_height_val = 0
        self._height_fixing = False
        def _height_step_fix(val):
            if self._height_fixing:
                return
            if 0 < val < 400:
                self._height_fixing = True
                # ▲ 눌렀으면 400으로, ▼ 눌렀으면 0으로
                self.spin_height.setValue(400 if val > self._prev_height_val else 0)
                self._height_fixing = False
                return
            self._prev_height_val = val
            self.settings.setValue("window_height", val)
            self.window_height_changed.emit(val)
        self.spin_height.valueChanged.connect(_height_step_fix)
        h_row.addWidget(self.spin_height)
        h_rec = QLabel("0=자동  |  수동 지정 시 즉시 적용")
        h_rec.setStyleSheet("color:#6c7086;font-size:11px;")
        h_row.addWidget(h_rec)
        h_row.addStretch()
        lay.addLayout(h_row)

        # ── 배치 모니터 ────────────────────────────────────────
        lay.addWidget(self._lbl("배치 모니터"))
        from PySide6.QtGui import QGuiApplication
        self.combo_monitor = QComboBox()
        self.combo_monitor.addItem("오른쪽 모니터 (기본)", "right")
        self.combo_monitor.addItem("왼쪽 모니터", "left")
        self.combo_monitor.addItem("기본(Primary) 모니터", "primary")
        screens = QGuiApplication.screens()
        for idx, s in enumerate(screens):
            geo = s.geometry()
            self.combo_monitor.addItem(
                f"모니터 {idx+1} ({geo.width()}×{geo.height()})", f"index:{idx}"
            )
        cur_mon = self.settings.value("monitor_placement", "right")
        for i in range(self.combo_monitor.count()):
            if self.combo_monitor.itemData(i) == cur_mon:
                self.combo_monitor.setCurrentIndex(i); break
        self.combo_monitor.currentIndexChanged.connect(
            lambda: self.settings.setValue("monitor_placement",
                                           self.combo_monitor.currentData())
        )
        lay.addWidget(self.combo_monitor)
        mon_note = QLabel("위젯은 선택한 모니터의 오른쪽 끝 상단에 배치됩니다.")
        mon_note.setStyleSheet("color:#6c7086;font-size:11px;")
        lay.addWidget(mon_note)

        # ── UI 테마 ────────────────────────────────────────────
        lay.addWidget(self._lbl("UI 색상 테마"))
        self._theme_btns: dict[str, QPushButton] = {}
        themes_row1 = QHBoxLayout(); themes_row1.setSpacing(8)
        themes_row2 = QHBoxLayout(); themes_row2.setSpacing(8)
        theme_keys = list(THEMES.keys())
        for i, key in enumerate(theme_keys):
            t = THEMES[key]
            btn = QPushButton(t["name"])
            btn.setCheckable(True)
            btn.setFixedHeight(36)
            btn.setStyleSheet(
                f"QPushButton{{background:{t['base']};color:{t['text']};"
                f"border:2px solid {t['surface0']};border-radius:8px;font-size:12px;padding:0 10px;}}"
                f"QPushButton:checked{{border:2px solid {t['blue']};font-weight:bold;}}"
                f"QPushButton:hover{{border:2px solid {t['blue']};}}"
            )
            btn.clicked.connect(lambda _, k=key: self._pick_theme(k))
            self._theme_btns[key] = btn
            (themes_row1 if i < 3 else themes_row2).addWidget(btn)
        themes_row1.addStretch(); themes_row2.addStretch()
        lay.addLayout(themes_row1)
        lay.addLayout(themes_row2)

        lay.addSpacing(16)
        sep_up = QFrame(); sep_up.setFrameShape(QFrame.Shape.HLine)
        sep_up.setMaximumHeight(1); lay.addWidget(sep_up)
        lay.addSpacing(8)
        lay.addWidget(self._lbl("앱 업데이트"))
        up_row = QHBoxLayout(); up_row.setSpacing(8)
        self.btn_check_update = QPushButton("⬆  버전 선택 / 업데이트")
        self.btn_check_update.setObjectName("SecondaryBtn")
        self.btn_check_update.setFixedHeight(34)
        self.btn_check_update.clicked.connect(self._check_github_version)
        up_row.addWidget(self.btn_check_update)
        self.lbl_update_status = QLabel(f"현재: {APP_VERSION}")
        self.lbl_update_status.setStyleSheet("color:#6c7086;font-size:11px;")
        up_row.addWidget(self.lbl_update_status, 1)
        lay.addLayout(up_row)

        lay.addStretch()
        self.tabs.addTab(w, "🎨  테마 & 화면")

    def _build_tab_sections(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(16, 16, 16, 8)
        lay.setSpacing(10)

        lay.addWidget(self._lbl("표시할 섹션 선택"))

        self._sec_checks: dict[str, QCheckBox] = {}
        sections = [
            ("show_calendar", "📅  달력"),
            ("show_todo",     "📝  과제 / 할 일 목록"),
            ("show_urgent",   "🚨  이번주 긴급 업무"),
            ("show_schedule", "📅  일정 관리"),
            ("show_misc",     "📌  기타"),
            ("show_personal", "👤  개인업무"),
        ]
        for key, label in sections:
            chk = QCheckBox(label)
            chk.setObjectName("TaskCheck")
            chk.setChecked(True)
            chk.toggled.connect(lambda _: self.section_changed.emit())
            self._sec_checks[key] = chk
            lay.addWidget(chk)

        lay.addSpacing(8)
        lay.addWidget(self._lbl("※ 달력 숨김 시 달력 하이라이트 기능이 비활성화됩니다."))

        lay.addSpacing(16)
        sep_lnk = QFrame(); sep_lnk.setFrameShape(QFrame.Shape.HLine)
        sep_lnk.setMaximumHeight(1); lay.addWidget(sep_lnk)
        lay.addSpacing(8)
        lay.addWidget(self._lbl("바탕화면 바로가기"))
        btn_lnk = QPushButton("🖥  바탕화면에 바로가기 만들기")
        btn_lnk.setObjectName("SecondaryBtn")
        btn_lnk.setFixedHeight(34)
        btn_lnk.clicked.connect(self._create_desktop_shortcut)
        lay.addWidget(btn_lnk)
        self.lbl_lnk_status = QLabel("")
        self.lbl_lnk_status.setObjectName("FormLabel")
        self.lbl_lnk_status.setWordWrap(True)
        lay.addWidget(self.lbl_lnk_status)

        lay.addSpacing(16)
        sep_fb = QFrame(); sep_fb.setFrameShape(QFrame.Shape.HLine)
        sep_fb.setMaximumHeight(1); lay.addWidget(sep_fb)
        lay.addSpacing(8)
        lay.addWidget(self._lbl("피드백 / 문의"))
        btn_feedback = QPushButton("✉  개발자에게 피드백 보내기")
        btn_feedback.setObjectName("SecondaryBtn")
        btn_feedback.setFixedHeight(34)
        btn_feedback.clicked.connect(self._send_feedback)
        lay.addWidget(btn_feedback)

        lay.addStretch()
        self.tabs.addTab(w, "📋  섹션 설정")

    def _build_tab_notif(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(16, 16, 16, 8)
        lay.setSpacing(12)

        lay.addWidget(self._lbl("마감일 알림"))
        self.chk_notif = QCheckBox("마감일 알림 활성화 (매 시간 체크)")
        self.chk_notif.setObjectName("TaskCheck")
        self.chk_notif.setChecked(True)
        self.chk_notif.toggled.connect(lambda v: (
            self.settings.setValue("notif_enabled", v),
            self.notif_changed.emit(v)
        ))
        lay.addWidget(self.chk_notif)

        lay.addSpacing(4)
        lay.addWidget(self._lbl("자동 시작"))
        note_autostart = QLabel(
            "자동 시작은 Windows 작업 스케줄러 'CalendarTodoList' 태스크로 관리됩니다.\n"
            "활성/비활성화는 작업 스케줄러에서 직접 변경하세요."
        )
        note_autostart.setObjectName("TaskInfoDesc")
        note_autostart.setWordWrap(True)
        note_autostart.setStyleSheet("color:#6c7086;font-size:11px;padding-top:4px;")
        lay.addWidget(note_autostart)

        lay.addStretch()
        self.tabs.addTab(w, "🔔  알림")

    def _build_tab_cowork(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(16, 16, 16, 8)
        lay.setSpacing(12)

        lay.addWidget(self._lbl("카카오워크 캘린더 iCal URL"))
        note = QLabel(
            "카카오워크 캘린더 → 설정 → 'iCal 구독 링크 복사' 에서 확인 가능합니다.\n"
            "비공개 URL이므로 외부에 공유하지 마세요."
        )
        note.setObjectName("TaskInfoDesc")
        note.setWordWrap(True)
        note.setStyleSheet("color:#7f849c;font-size:11px;")
        lay.addWidget(note)

        self.ed_ical_url = QLineEdit()
        self.ed_ical_url.setPlaceholderText(
            "https://calendar.kakaowork.com/ical/calendars/.../basic.ics"
        )
        self.ed_ical_url.textChanged.connect(
            lambda v: self.settings.setValue("ical_url", _ical_url_encode(v.strip()))
        )
        lay.addWidget(self.ed_ical_url)

        sync_row = QHBoxLayout(); sync_row.setSpacing(8)
        self.lbl_ical_status = QLabel("—")
        self.lbl_ical_status.setObjectName("UpdateTime")
        self.lbl_ical_status.setStyleSheet("color:#6c7086;font-size:11px;")
        sync_row.addWidget(self.lbl_ical_status, 1)

        btn_sync = QPushButton("🔄  지금 동기화")
        btn_sync.setObjectName("RefreshBtn")
        btn_sync.setFixedHeight(30)
        btn_sync.clicked.connect(self._sync_ical_now)
        sync_row.addWidget(btn_sync)
        lay.addLayout(sync_row)

        lay.addWidget(self._lbl("자동 동기화 주기"))
        interval_row = QHBoxLayout(); interval_row.setSpacing(8)
        self.cb_ical_interval = QComboBox()
        for label, val in [("수동만", 0), ("30분마다", 30), ("1시간마다", 60),
                           ("3시간마다", 180), ("매일 1회", 1440)]:
            self.cb_ical_interval.addItem(label, val)
        self.cb_ical_interval.currentIndexChanged.connect(self._on_ical_interval)
        interval_row.addWidget(self.cb_ical_interval)
        interval_row.addStretch()
        lay.addLayout(interval_row)

        # ── 구분선 ────────────────────────────────────────────────────────
        sep1 = QFrame(); sep1.setFrameShape(QFrame.Shape.HLine)
        sep1.setStyleSheet("background:#2e2e48;border:none;max-height:1px;margin:4px 0;")
        lay.addWidget(sep1)

        # ── 아침 브리핑 알림 ──────────────────────────────────────────────
        lay.addWidget(self._lbl("📢  아침 브리핑 시간"))
        brief_note = QLabel(
            "매일 지정 시간에 오늘의 팀 일정 전체를 트레이 알림으로 요약합니다."
        )
        brief_note.setObjectName("TaskInfoDesc")
        brief_note.setWordWrap(True)
        brief_note.setStyleSheet("color:#7f849c;font-size:11px;")
        lay.addWidget(brief_note)

        brief_row = QHBoxLayout(); brief_row.setSpacing(8)
        self.ed_brief_time = QLineEdit()
        self.ed_brief_time.setPlaceholderText("HH:MM  예: 08:30")
        self.ed_brief_time.setFixedWidth(100)
        self.ed_brief_time.textChanged.connect(
            lambda v: self.settings.setValue("ical_brief_time", v.strip()))
        brief_row.addWidget(self.ed_brief_time)
        brief_row.addStretch()
        lay.addLayout(brief_row)

        # ── 구분선 ────────────────────────────────────────────────────────
        sep2 = QFrame(); sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setStyleSheet("background:#2e2e48;border:none;max-height:1px;margin:4px 0;")
        lay.addWidget(sep2)

        # ── 관심 인원 명단 ────────────────────────────────────────────────
        lay.addWidget(self._lbl("👤  개인 일정 알림 대상 (15분 전 알림)"))
        watch_note = QLabel(
            "아래 이름이 이벤트 제목에 포함된 경우에만 15분 전 개인 알림이 발송됩니다.\n"
            "한 줄에 한 명씩 입력하세요."
        )
        watch_note.setObjectName("TaskInfoDesc")
        watch_note.setWordWrap(True)
        watch_note.setStyleSheet("color:#7f849c;font-size:11px;")
        lay.addWidget(watch_note)

        self.te_watch = QTextEdit()
        self.te_watch.setPlaceholderText("김성희\n김현표\n김별\n최재영")
        self.te_watch.setMaximumHeight(100)
        self.te_watch.textChanged.connect(
            lambda: self.settings.setValue(
                "ical_watch_persons", self.te_watch.toPlainText().strip()))
        lay.addWidget(self.te_watch)

        lay.addStretch()
        self.tabs.addTab(w, "🏢  Co-work")

    # ── 로직 ────────────────────────────────────────────────────────────────
    def _on_opacity(self, v: int):
        self.lbl_opacity.setText(f"{v}%")
        self.settings.setValue("opacity", v)
        self.opacity_changed.emit(v)

    def _on_fontsize(self, v: int):
        self.settings.setValue("font_size", v)
        self.font_size_changed.emit(v)

    def _on_font_family(self, family: str):
        self.settings.setValue("font_family", family)
        self.font_size_changed.emit(self.spin_fontsize.value())  # 테마 재적용 트리거

    def _pick_theme(self, key: str):
        for k, btn in self._theme_btns.items():
            btn.setChecked(k == key)
        self.settings.setValue("theme", key)
        self.theme_changed.emit(key)

    def _load_settings(self):
        # 투명도
        op = self.settings.value("opacity", 100, type=int)
        self.slider_opacity.setValue(op)
        self.lbl_opacity.setText(f"{op}%")

        # 글자 크기
        fs = self.settings.value("font_size", 10, type=int)
        self.spin_fontsize.setValue(fs)

        # 창 너비
        w = self.settings.value("window_width", WINDOW_WIDTH, type=int)
        self.spin_width.setValue(w)

        # 창 높이 (0=자동)
        h = self.settings.value("window_height", 0, type=int)
        self.spin_height.setValue(h)

        # 테마
        theme = self.settings.value("theme", "dark")
        for k, btn in self._theme_btns.items():
            btn.setChecked(k == theme)

        # 섹션 가시성
        for key, chk in self._sec_checks.items():
            chk.setChecked(self.settings.value(key, True, type=bool))

        # 알림
        self.chk_notif.setChecked(self.settings.value("notif_enabled", True, type=bool))

        # Co-work iCal
        self.ed_ical_url.setText(_ical_url_decode(self.settings.value("ical_url", "")))
        interval = self.settings.value("ical_interval", 60, type=int)
        for i in range(self.cb_ical_interval.count()):
            if self.cb_ical_interval.itemData(i) == interval:
                self.cb_ical_interval.setCurrentIndex(i)
                break

        # 아침 브리핑 시간
        self.ed_brief_time.setText(self.settings.value("ical_brief_time", "08:30"))

        # 관심 인원 명단
        default_watch = "김성희\n김현표\n김별\n최재영"
        self.te_watch.setPlainText(
            self.settings.value("ical_watch_persons", default_watch))

    def get_section_visibility(self) -> dict[str, bool]:
        return {k: chk.isChecked() for k, chk in self._sec_checks.items()}

    def _on_ical_interval(self, _):
        v = self.cb_ical_interval.currentData()
        self.settings.setValue("ical_interval", v)

    def _sync_ical_now(self):
        """설정 탭에서 즉시 동기화 요청 — MainWindow에서 처리"""
        url = self.ed_ical_url.text().strip()
        if not url:
            self.lbl_ical_status.setText("⚠ URL을 먼저 입력하세요")
            return
        self.lbl_ical_status.setText("동기화 중...")
        QApplication.processEvents()
        # 신호 대신 직접 부모에서 fetch 실행 (동기 호출)
        parent = self.parent()
        if parent and hasattr(parent, "_fetch_ical"):
            ok, msg = parent._fetch_ical(url)
            self.lbl_ical_status.setText(msg)
        else:
            self.lbl_ical_status.setText("⚠ 메인 윈도우를 찾을 수 없음")

    def _create_desktop_shortcut(self):
        """PowerShell로 바탕화면에 .lnk 바로가기 생성"""
        import subprocess, shlex
        desktop = Path.home() / "Desktop"
        lnk_path = desktop / "Calendar and Todo.lnk"
        # 현재 실행 환경 감지
        if getattr(sys, "frozen", False):
            target = sys.executable
            args_str = ""
        else:
            target = sys.executable
            script  = str(Path(__file__).resolve())
            args_str = f'"{script}"'
        work_dir = str(Path(target).parent)
        ps = (
            "$ws = New-Object -ComObject WScript.Shell; "
            f"$s = $ws.CreateShortcut('{lnk_path}'); "
            f"$s.TargetPath = '{target}'; "
            f"$s.Arguments = '{args_str}'; "
            f"$s.WorkingDirectory = '{work_dir}'; "
            "$s.Save()"
        )
        try:
            subprocess.run(
                ["powershell", "-NoProfile", "-Command", ps],
                capture_output=True, timeout=10, check=True,
            )
            self.lbl_lnk_status.setText(f"✅ 생성됨: {lnk_path}")
            self.lbl_lnk_status.setStyleSheet("color:#a6e3a1;")
        except Exception as e:
            self.lbl_lnk_status.setText(f"⚠ 실패: {e}")
            self.lbl_lnk_status.setStyleSheet("color:#f38ba8;")

    def _send_feedback(self):
        """기본 메일 클라이언트로 피드백 메일 작성 창 열기"""
        subject = f"[Calendar & ToDo] 피드백 ({APP_VERSION})"
        body = (
            "안녕하세요,\n\n"
            "피드백/건의 내용을 작성해 주세요:\n\n\n"
            "---\n"
            f"앱 버전: {APP_VERSION} ({APP_VERSION_DATE})\n"
            f"OS: {sys.platform}\n"
        )
        import urllib.parse
        mailto = (
            "mailto:hyunjun.kwak@hd.com"
            f"?subject={urllib.parse.quote(subject)}"
            f"&body={urllib.parse.quote(body)}"
        )
        QDesktopServices.openUrl(QUrl(mailto))

    def _check_github_version(self):
        """GitHub Releases 업데이트 다이얼로그 열기"""
        dlg = AppUpdateDialog(parent=self)
        dlg.exec()


# ═══════════════════════════════════════════════════════════════════════════
# 14-A2. 앱 업데이트 다이얼로그 (버전 선택 + 다운로드)
# ═══════════════════════════════════════════════════════════════════════════

class AppUpdateDialog(_MovableDialog):
    """GitHub Releases에서 버전 목록을 불러와 선택·다운로드하는 다이얼로그"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setMinimumWidth(560)
        self.resize(560, 520)
        self._releases: list[dict] = []
        self._build()
        QShortcut(QKeySequence("Escape"), self, self.accept)
        # 다이얼로그 표시 후 바로 릴리즈 목록 로드
        QTimer.singleShot(100, self._fetch_releases)

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(12)

        # ── 타이틀 ─────────────────────────────────────────────────────────
        title_row = QHBoxLayout()
        title = QLabel("⬆  버전 업데이트")
        title.setObjectName("DialogTitle")
        title.setFont(QFont("맑은 고딕", 14, QFont.Weight.Bold))
        title_row.addWidget(title, 1)
        btn_close = QPushButton("✕")
        btn_close.setObjectName("TitleBarBtn")
        btn_close.setFixedSize(28, 28)
        btn_close.clicked.connect(self.accept)
        title_row.addWidget(btn_close)
        lay.addLayout(title_row)

        # 현재 버전 표시
        cur_lbl = QLabel(f"현재 버전: {APP_VERSION}  |  업데이트 날짜: {APP_VERSION_DATE}")
        cur_lbl.setStyleSheet("color:#6c7086;font-size:11px;")
        lay.addWidget(cur_lbl)

        # ── 상태 레이블 ────────────────────────────────────────────────────
        self.lbl_status = QLabel("GitHub에서 버전 목록을 불러오는 중...")
        self.lbl_status.setStyleSheet("color:#a6adc8;font-size:12px;")
        lay.addWidget(self.lbl_status)

        # ── 버전 목록 + 릴리즈 노트 스플리터 ──────────────────────────────
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(4)
        splitter.setChildrenCollapsible(False)

        # 왼쪽: 버전 목록
        left_w = QWidget()
        left_lay = QVBoxLayout(left_w)
        left_lay.setContentsMargins(0, 0, 4, 0)
        left_lay.setSpacing(4)
        lbl_list = QLabel("버전 선택")
        lbl_list.setStyleSheet("color:#89b4fa;font-size:11px;font-weight:bold;")
        left_lay.addWidget(lbl_list)
        self.list_versions = QListWidget()
        self.list_versions.setObjectName("LogList")
        self.list_versions.setAlternatingRowColors(False)
        self.list_versions.currentRowChanged.connect(self._on_version_selected)
        left_lay.addWidget(self.list_versions)
        splitter.addWidget(left_w)

        # 오른쪽: 릴리즈 노트
        right_w = QWidget()
        right_lay = QVBoxLayout(right_w)
        right_lay.setContentsMargins(4, 0, 0, 0)
        right_lay.setSpacing(4)
        lbl_notes = QLabel("변경 내역")
        lbl_notes.setStyleSheet("color:#89b4fa;font-size:11px;font-weight:bold;")
        right_lay.addWidget(lbl_notes)
        self.txt_notes = QPlainTextEdit()
        self.txt_notes.setReadOnly(True)
        self.txt_notes.setObjectName("LogInput")
        self.txt_notes.setPlaceholderText("버전을 선택하면 변경 내역이 표시됩니다")
        right_lay.addWidget(self.txt_notes)
        splitter.addWidget(right_w)
        splitter.setSizes([200, 340])
        lay.addWidget(splitter, 1)

        # ── 다운로드 진행 바 ────────────────────────────────────────────────
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFixedHeight(18)
        self.progress_bar.setObjectName("UpdateProgressBar")
        self.progress_bar.hide()
        lay.addWidget(self.progress_bar)

        self.lbl_download = QLabel("")
        self.lbl_download.setStyleSheet("color:#a6e3a1;font-size:11px;")
        self.lbl_download.hide()
        lay.addWidget(self.lbl_download)

        # ── 버튼 행 ────────────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_gh = QPushButton("🌐 GitHub 릴리즈 페이지")
        btn_gh.setObjectName("SecondaryBtn")
        btn_gh.setFixedHeight(34)
        btn_gh.clicked.connect(lambda: QDesktopServices.openUrl(
            QUrl("https://github.com/Hyunjun15/calendar-todo-widget/releases")
        ))
        btn_row.addWidget(btn_gh)
        self.btn_download = QPushButton("⬇  다운로드")
        self.btn_download.setObjectName("PrimaryBtn")
        self.btn_download.setFixedHeight(34)
        self.btn_download.setEnabled(False)
        self.btn_download.clicked.connect(self._download_selected)
        btn_row.addWidget(self.btn_download)
        lay.addLayout(btn_row)

        # ── 안내 문구 ───────────────────────────────────────────────────────
        note = QLabel("💡 다운로드 후 압축 해제하여 main.py를 교체하거나 사용_안내.txt를 참고하세요.")
        note.setStyleSheet("color:#6c7086;font-size:10px;")
        note.setWordWrap(True)
        lay.addWidget(note)

    def _fetch_releases(self):
        """GitHub API로 모든 릴리즈 목록 가져오기"""
        try:
            import json as _json
            import ssl
            ctx = ssl.create_default_context()
            url = "https://api.github.com/repos/Hyunjun15/calendar-todo-widget/releases?per_page=30"
            req = urllib.request.Request(url, headers={"User-Agent": "CalendarTodoWidget"})
            with urllib.request.urlopen(req, context=ctx, timeout=10) as resp:
                self._releases = _json.loads(resp.read().decode())
            self._populate_list()
        except Exception as ex:
            self.lbl_status.setText(f"⚠ 연결 실패: {ex}")
            self.lbl_status.setStyleSheet("color:#f38ba8;font-size:12px;")

    def _populate_list(self):
        """릴리즈 목록을 QListWidget에 채우기"""
        self.list_versions.clear()
        if not self._releases:
            self.lbl_status.setText("릴리즈 정보를 찾을 수 없습니다.")
            return
        for rel in self._releases:
            tag = rel.get("tag_name", "")
            name = rel.get("name", tag)
            published = rel.get("published_at", "")[:10]
            is_current = (tag == APP_VERSION)
            is_latest = (rel == self._releases[0])
            label = name
            if is_current:
                label = f"✅ {name}  (현재)"
            elif is_latest:
                label = f"🆕 {name}  (최신)"
            item = QListWidgetItem(label)
            item.setData(Qt.ItemDataRole.UserRole, rel)
            if is_current:
                item.setForeground(QColor("#a6e3a1"))
            elif is_latest:
                item.setForeground(QColor("#89b4fa"))
            self.list_versions.addItem(item)
        self.lbl_status.setText(f"총 {len(self._releases)}개 버전 — 버전을 선택하세요")
        self.lbl_status.setStyleSheet("color:#a6adc8;font-size:12px;")
        # 최신 버전 또는 현재 버전 자동 선택
        self.list_versions.setCurrentRow(0)

    def _on_version_selected(self, row: int):
        if row < 0 or row >= len(self._releases):
            self.btn_download.setEnabled(False)
            return
        rel = self._releases[row]
        body = rel.get("body", "") or "(변경 내역 없음)"
        self.txt_notes.setPlainText(body)
        tag = rel.get("tag_name", "")
        is_current = (tag == APP_VERSION)
        self.btn_download.setEnabled(not is_current)
        if is_current:
            self.btn_download.setText("현재 버전")
        else:
            self.btn_download.setText(f"⬇  {tag} 다운로드")

    def _download_selected(self):
        row = self.list_versions.currentRow()
        if row < 0 or row >= len(self._releases):
            return
        rel = self._releases[row]
        tag = rel.get("tag_name", "")
        if not tag:
            return

        # 다운로드 URL (소스 zip)
        zip_url = f"https://github.com/Hyunjun15/calendar-todo-widget/archive/refs/tags/{tag}.zip"
        downloads_dir = Path.home() / "Downloads"
        downloads_dir.mkdir(exist_ok=True)
        save_path = downloads_dir / f"calendar-todo-widget-{tag}.zip"

        # 이미 존재하면 확인
        if save_path.exists():
            r = QMessageBox.question(self, "파일 존재",
                f"이미 다운로드된 파일이 있습니다:\n{save_path}\n\n덮어쓰시겠습니까?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No)
            if r != QMessageBox.StandardButton.Yes:
                return

        self.btn_download.setEnabled(False)
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        self.lbl_download.hide()

        try:
            import ssl, json as _json
            ctx = ssl.create_default_context()
            req = urllib.request.Request(zip_url, headers={"User-Agent": "CalendarTodoWidget"})
            with urllib.request.urlopen(req, context=ctx, timeout=30) as resp:
                total = int(resp.headers.get("Content-Length", 0))
                downloaded = 0
                chunk = 8192
                with open(save_path, "wb") as f:
                    while True:
                        buf = resp.read(chunk)
                        if not buf:
                            break
                        f.write(buf)
                        downloaded += len(buf)
                        if total > 0:
                            pct = int(downloaded * 100 / total)
                            self.progress_bar.setValue(pct)
                            QApplication.processEvents()
            self.progress_bar.setValue(100)
            self.lbl_download.setText(f"✅ 저장 완료: {save_path}")
            self.lbl_download.show()
            r2 = QMessageBox.information(self, "다운로드 완료",
                f"버전 {tag} 다운로드 완료!\n\n저장 위치:\n{save_path}\n\n"
                "압축을 해제하고 main.py와 assets/ 폴더를 현재 폴더에 교체한 후\n"
                "앱을 재시작하세요.",
                QMessageBox.StandardButton.Open | QMessageBox.StandardButton.Ok,
                QMessageBox.StandardButton.Open)
            if r2 == QMessageBox.StandardButton.Open:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(downloads_dir)))
        except Exception as ex:
            self.progress_bar.hide()
            QMessageBox.warning(self, "다운로드 실패",
                f"다운로드 중 오류가 발생했습니다:\n\n{ex}",
                QMessageBox.StandardButton.Ok)
        finally:
            self.btn_download.setEnabled(True)
            tag2 = rel.get("tag_name", "")
            is_current2 = (tag2 == APP_VERSION)
            if is_current2:
                self.btn_download.setText("현재 버전")
            else:
                self.btn_download.setText(f"⬇  {tag2} 다운로드")


# ═══════════════════════════════════════════════════════════════════════════
# 14-A3. 사용 안내서 다이얼로그 (HTML)
# ═══════════════════════════════════════════════════════════════════════════

class HelpDialog(_MovableDialog):
    """앱 사용 안내서 — HTML 렌더링 + What's New 탭"""

    _SECTIONS = [
        ("whatsnew", "🆕 새 기능"),
        ("shortcuts", "⌨ 단축키"),
        ("tasks", "📝 태스크 관리"),
        ("search", "🔍 검색"),
        ("schedule", "📅 일정 관리"),
        ("calendar", "📆 캘린더"),
        ("ical", "🏢 연구실 일정"),
        ("settings", "⚙ 설정"),
        ("backup", "💾 백업 / 복원"),
        ("export", "📤 내보내기"),
        ("update", "⬆ 업데이트"),
    ]

    def __init__(self, parent=None, show_tab: str = "whatsnew"):
        super().__init__(parent)
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setMinimumSize(680, 520)
        self.resize(740, 620)
        self._initial_tab = show_tab
        self._build()
        QShortcut(QKeySequence("Escape"), self, self.accept)

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 16, 20, 16)
        lay.setSpacing(10)

        # 타이틀
        title_row = QHBoxLayout()
        title = QLabel(f"📖  사용 안내서  —  {APP_VERSION}")
        title.setObjectName("DialogTitle")
        title.setFont(QFont("맑은 고딕", 14, QFont.Weight.Bold))
        title_row.addWidget(title, 1)
        btn_close = QPushButton("✕")
        btn_close.setObjectName("TitleBarBtn")
        btn_close.setFixedSize(28, 28)
        btn_close.clicked.connect(self.accept)
        title_row.addWidget(btn_close)
        lay.addLayout(title_row)

        # 스플리터: 좌측 목차 + 우측 콘텐츠
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(3)

        # 좌측 목차
        self.list_sections = QListWidget()
        self.list_sections.setFixedWidth(150)
        self.list_sections.setStyleSheet(
            "QListWidget{background:#1e1e2e;border:1px solid #313244;border-radius:6px;}"
            "QListWidget::item{padding:8px 10px;color:#cdd6f4;}"
            "QListWidget::item:selected{background:#313244;color:#89b4fa;border-radius:4px;}"
        )
        for key, label in self._SECTIONS:
            self.list_sections.addItem(label)
        self.list_sections.currentRowChanged.connect(self._on_section)
        splitter.addWidget(self.list_sections)

        # 우측 HTML 뷰어
        self.browser = QTextBrowser()
        self.browser.setOpenExternalLinks(True)
        self.browser.setStyleSheet(
            "QTextBrowser{background:#181825;border:1px solid #313244;border-radius:6px;"
            "padding:12px;color:#cdd6f4;font-size:13px;}"
        )
        splitter.addWidget(self.browser)
        splitter.setSizes([150, 560])
        lay.addWidget(splitter, 1)

        # 초기 탭 선택
        idx = next((i for i, (k, _) in enumerate(self._SECTIONS) if k == self._initial_tab), 0)
        self.list_sections.setCurrentRow(idx)

    def _on_section(self, row: int):
        if row < 0 or row >= len(self._SECTIONS):
            return
        key = self._SECTIONS[row][0]
        html = self._render(key)
        self.browser.setHtml(html)

    def _css(self) -> str:
        return """
        body { font-family: '맑은 고딕', sans-serif; color: #cdd6f4; line-height: 1.7; }
        h2 { color: #89b4fa; border-bottom: 1px solid #313244; padding-bottom: 6px; margin-top: 18px; }
        h3 { color: #f9e2af; margin-top: 14px; }
        kbd { background: #313244; color: #a6e3a1; padding: 2px 7px; border-radius: 4px;
              font-family: 'Consolas', monospace; font-size: 12px; border: 1px solid #45475a; }
        .loc { background: rgba(137,180,250,0.12); color: #89b4fa; padding: 2px 6px;
               border-radius: 4px; font-size: 12px; }
        .new { background: rgba(166,227,161,0.15); color: #a6e3a1; padding: 1px 6px;
               border-radius: 3px; font-size: 11px; }
        ul { margin-left: 6px; padding-left: 14px; }
        li { margin-bottom: 4px; }
        .note { color: #fab387; font-size: 12px; }
        """

    def _render(self, key: str) -> str:
        css = self._css()
        body = getattr(self, f"_html_{key}", lambda: "<p>준비 중...</p>")()
        return f"<html><head><style>{css}</style></head><body>{body}</body></html>"

    def _html_whatsnew(self) -> str:
        return """
        <h2>🆕 v3.27 새 기능 및 개선</h2>

        <h3>🔧 창 크기 설정 개선</h3>
        <ul>
          <li><span class="loc">설정 → 테마 &amp; 화면</span> SpinBox 상하 버튼에 <b>화살표(▲▼)</b> 표시 추가</li>
          <li>창 높이: <b>0(자동) ↔ 400(최소)</b> 사이 죽은 구간 제거 — 버튼 한 번에 점프</li>
          <li>창 너비 최소값을 실제 최소 크기(380px)와 일치시킴</li>
        </ul>

        <h3>🎨 색상 대비 전면 개선</h3>
        <ul>
          <li>할 일 카드의 편집/삭제/로그 버튼이 더 잘 보이도록 대비 향상</li>
          <li>캘린더 마감일 셀 배경 강화 — 마감일 있는 날짜를 더 쉽게 구분</li>
          <li>섹션 통계, 빈 상태 메시지, 완료 항목 글씨 가독성 개선</li>
        </ul>

        <h3>💾 백업 호환성 강화</h3>
        <ul>
          <li>JSON 내보내기에 <b>첨부파일 정보</b> 포함</li>
          <li>불러오기 시 <b>완료 상태, 생성일, 정렬 순서</b> 완전 복원</li>
        </ul>

        <h3>🛡 안정성 개선</h3>
        <ul>
          <li>앱 비정상 종료 후에도 30초 뒤 자동으로 재실행 가능</li>
          <li>태스크 삭제 시 데이터 안전성 강화</li>
        </ul>
        """

    def _html_shortcuts(self) -> str:
        return """
        <h2>⌨ 키보드 단축키</h2>
        <table cellspacing="8">
        <tr><td><kbd>Ctrl</kbd>+<kbd>N</kbd></td><td>새 태스크 추가</td></tr>
        <tr><td><kbd>Ctrl</kbd>+<kbd>F</kbd></td><td>전체 검색</td></tr>
        <tr><td><kbd>Ctrl</kbd>+<kbd>T</kbd></td><td>항상 위에 표시 토글</td></tr>
        <tr><td><kbd>Ctrl</kbd>+<kbd>M</kbd></td><td>위젯 접기 / 펼치기</td></tr>
        <tr><td><kbd>Ctrl</kbd>+<kbd>,</kbd></td><td>설정 열기</td></tr>
        <tr><td><kbd>Ctrl</kbd>+<kbd>Enter</kbd></td><td>다이얼로그 저장</td></tr>
        <tr><td><kbd>Ctrl</kbd>+<kbd>←</kbd> / <kbd>→</kbd></td><td>캘린더 이전/다음 달</td></tr>
        <tr><td><kbd>Home</kbd></td><td>캘린더 오늘로 이동</td></tr>
        <tr><td><kbd>Esc</kbd></td><td>다이얼로그/검색 닫기</td></tr>
        </table>

        <h3>마우스 조작</h3>
        <ul>
          <li><b>더블클릭</b> — 항목 편집 다이얼로그 열기</li>
          <li><b>우클릭</b> — 컨텍스트 메뉴 (편집/삭제/완료)</li>
          <li><b>드래그</b> — 태스크 순서 변경</li>
        </ul>
        """

    def _html_tasks(self) -> str:
        return """
        <h2>📝 태스크 관리</h2>

        <h3>태스크 종류</h3>
        <ul>
          <li>📝 <b>과제 / 할 일</b> — 마감일 기준 장기 과제</li>
          <li>🚨 <b>긴급업무</b> — 이번 주 단기 처리</li>
          <li>📌 <b>기타</b> — 자유 메모</li>
          <li>👤 <b>개인업무</b> — 개인 일정/업무</li>
        </ul>

        <h3>태스크 추가</h3>
        <p><kbd>Ctrl</kbd>+<kbd>N</kbd> 또는 <span class="loc">섹션 하단 [＋ 새 항목 추가]</span> 버튼</p>
        <ul>
          <li>제목, 설명, 목표, 우선순위(높음/보통/낮음), 마감일, 색상 설정</li>
          <li>마감일은 <b>선택사항</b> — 체크박스로 활성화</li>
        </ul>

        <h3>파일 첨부 <span class="new">v3.20+</span></h3>
        <p><span class="loc">태스크 편집 → ＋ 파일 추가</span></p>
        <ul>
          <li>커스텀 파일 선택 창: 폴더 트리, 즐겨찾기, 다중 선택</li>
          <li>드래그 &amp; 드롭 지원, 확장자별 아이콘 + 파일 크기 표시</li>
        </ul>

        <h3>완료 처리</h3>
        <ul>
          <li>좌측 체크박스 클릭 → 완료 (취소선 + 흐림)</li>
          <li>완료 항목은 <span class="loc">✅ 완료업무</span> 섹션에서 확인</li>
          <li>완료 취소: 우클릭 → "미완료로"</li>
          <li>완료된 항목 <b>더블클릭 → 편집 가능</b> <span class="new">v3.20+</span></li>
          <li>접기/펼치기(▸/▾)로 상세 내용 확인 <span class="new">v3.13+</span></li>
        </ul>

        <h3>정렬 · 일괄 처리</h3>
        <ul>
          <li>정렬 드롭다운: 기본/마감일/우선순위/생성일/제목</li>
          <li>☐ 버튼 → 배치 모드에서 일괄 완료/삭제</li>
        </ul>
        """

    def _html_search(self) -> str:
        return """
        <h2>🔍 검색</h2>
        <p><span class="loc">타이틀바 🔍 검색 버튼</span> 또는 <kbd>Ctrl</kbd>+<kbd>F</kbd></p>
        <ul>
          <li>검색 범위: 과제, 긴급업무, 개인업무, 일정</li>
          <li>제목 / 내용 / 목표 / 일정명 / 장소 실시간 필터링</li>
          <li>검색 결과 수 자동 표시</li>
        </ul>
        <h3>완료업무 검색</h3>
        <p><span class="loc">✅ 완료업무 섹션</span> 펼치면 전용 검색창 표시</p>
        """

    def _html_schedule(self) -> str:
        return """
        <h2>📅 일정 관리</h2>
        <p><span class="loc">일정 섹션 하단 [＋ 일정 추가]</span> 또는 캘린더 날짜 우클릭</p>
        <h3>일정 종류</h3>
        <ul>
          <li>📅 단기 일정</li>
          <li>🏖 연차 / 휴가</li>
          <li>📚 교육</li>
          <li>✈ 출장</li>
        </ul>
        <p>필드: 제목, 시작일~종료일, 시간, 장소, 내용</p>
        <p>편집: 더블클릭 | 삭제: 우클릭</p>
        """

    def _html_calendar(self) -> str:
        return """
        <h2>📆 캘린더</h2>
        <ul>
          <li><b>파란 배경</b> = 오늘</li>
          <li><b>분홍 배경</b> = 마감일이 있는 날</li>
          <li><b>색상 점</b> = 연구실 일정(iCal)</li>
        </ul>
        <h3>조작</h3>
        <ul>
          <li><span class="loc">◀ ▶ 버튼</span> 또는 <kbd>Ctrl</kbd>+<kbd>←</kbd>/<kbd>→</kbd> — 월 이동</li>
          <li><span class="loc">[오늘] 버튼</span> 또는 <kbd>Home</kbd> — 오늘로 이동</li>
          <li><b>날짜 클릭</b> — 해당일 태스크/일정 팝업</li>
          <li><b>날짜 우클릭</b> — 일정 추가 / 개인업무 추가</li>
        </ul>
        """

    def _html_ical(self) -> str:
        return """
        <h2>🏢 연구실 일정 (카카오워크 iCal)</h2>
        <p>카카오워크 팀 캘린더를 iCal 형식으로 연동합니다.</p>
        <h3>색상 분류</h3>
        <ul>
          <li>🔵 <b>파란색</b> — 소장님일정 (최상단)</li>
          <li>🟢 <b>초록색</b> — 연차, 예비군, 반차 등</li>
          <li>🟠 <b>주황색</b> — 종일 일정</li>
          <li>🟡 <b>노란색</b> — 시간 지정 일정</li>
        </ul>
        <h3>설정</h3>
        <p><span class="loc">설정(⚙) → Co-work 탭</span> → iCal URL 입력 → 🔄 동기화</p>
        """

    def _html_settings(self) -> str:
        return """
        <h2>⚙ 설정</h2>
        <p><span class="loc">타이틀바 ⚙ 버튼</span> 또는 <kbd>Ctrl</kbd>+<kbd>,</kbd></p>

        <h3>탭 1 — 테마 &amp; 화면</h3>
        <ul>
          <li>테마: 다크 / 블랙 / 라떼 / 네이비 / 그루박스 / 도쿄 나이트</li>
          <li>투명도, 글자 크기, 글씨체</li>
          <li>창 너비 (380~1400px), 창 높이 (0=자동 또는 400~2160px)</li>
          <li>배치 모니터 선택</li>
          <li><span class="loc">버전 선택 / 업데이트</span> 버튼 <span class="new">v3.20+</span></li>
        </ul>

        <h3>탭 2 — 섹션 설정</h3>
        <ul>
          <li>각 섹션 표시/숨기기</li>
          <li>바탕화면 바로가기, 피드백 보내기</li>
        </ul>

        <h3>탭 3 — 알림</h3>
        <p>마감일 알림 On/Off (매 시간 체크)</p>

        <h3>탭 4 — Co-work &amp; iCal</h3>
        <p>카카오워크 iCal URL 입력 + 동기화 주기 설정</p>
        """

    def _html_backup(self) -> str:
        return """
        <h2>💾 백업 / 복원</h2>
        <p><span class="loc">타이틀바 💾 버튼</span></p>

        <h3>내보내기</h3>
        <ul>
          <li>할 일 + 로그 + 진행 그룹 + 첨부파일 정보를 JSON으로 저장</li>
          <li>일정(schedules)도 선택 내보내기 가능</li>
        </ul>

        <h3>가져오기</h3>
        <ul>
          <li>JSON 파일에서 항목별 선택 복원</li>
          <li><b>완료 상태, 생성일, 정렬 순서 완전 복원</b> <span class="new">v3.27+</span></li>
          <li>첨부파일: 원본 경로에 파일이 있으면 자동 재첨부</li>
        </ul>

        <h3>자동 백업</h3>
        <p>7일마다 앱 시작 시 자동 DB 백업<br>
        <span class="note">저장 위치: %USERPROFILE%\\.productivity_widget\\backups\\</span></p>
        """

    def _html_export(self) -> str:
        return """
        <h2>📤 내보내기 (Export)</h2>
        <p><span class="loc">타이틀바 📤 버튼</span></p>
        <ul>
          <li>완료/미완료 필터 선택</li>
          <li>항목별 그룹명 설정</li>
          <li>미리보기 → 클립보드 복사 또는 파일 저장</li>
        </ul>
        """

    def _html_update(self) -> str:
        return """
        <h2>⬆ 업데이트</h2>
        <p><span class="loc">설정(⚙) → 테마 &amp; 화면 → 버전 선택 / 업데이트</span></p>

        <h3>기능 <span class="new">v3.20+</span></h3>
        <ul>
          <li>GitHub 전체 릴리즈 목록 조회</li>
          <li>각 버전의 변경 내역(릴리즈 노트) 미리보기</li>
          <li>원하는 버전 선택 후 직접 다운로드 (진행률 표시)</li>
        </ul>

        <h3>수동 업데이트</h3>
        <p>다운로드한 zip에서 <b>main.py</b>와 <b>assets/</b> 폴더를 현재 위치에 덮어쓰고 재시작</p>
        <p class="note">기존 데이터(tasks.db)는 자동으로 새 버전에 맞게 변환됩니다.</p>
        """


# 14-B. JSON 백업 / 복원 다이얼로그
# ═══════════════════════════════════════════════════════════════════════════

class JsonBackupDialog(_MovableDialog):
    """데이터 백업(JSON 내보내기) + 복원(선택적 가져오기)"""

    def __init__(self, db: Database, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.FramelessWindowHint)
        self.setModal(True)
        self.setMinimumWidth(560)
        self.resize(560, 600)
        self._import_tasks: list[dict] = []   # 가져오기 후 파싱된 태스크 목록
        self._import_chks: list[QCheckBox] = []
        self._build()
        QShortcut(QKeySequence("Escape"), self, self.accept)

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(24, 20, 24, 20)
        lay.setSpacing(14)

        title = QLabel("💾  데이터 백업 / 복원")
        title.setObjectName("DialogTitle")
        title.setFont(QFont("맑은 고딕", 14, QFont.Weight.Bold))
        lay.addWidget(title)

        self.tabs = QTabWidget()
        self.tabs.setObjectName("OptionsTab")
        lay.addWidget(self.tabs)

        self._build_export_tab()
        self._build_import_tab()

        br = QHBoxLayout(); br.addStretch()
        bc = QPushButton("닫기"); bc.setObjectName("SecondaryBtn")
        bc.setFixedHeight(36); bc.clicked.connect(self.accept)
        br.addWidget(bc)
        lay.addLayout(br)

    def _build_export_tab(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(16, 16, 16, 8)
        lay.setSpacing(12)

        info = QLabel(
            "모든 할 일, 로그, 진행 그룹, 일정 데이터를\n"
            "JSON 파일로 내보냅니다."
        )
        info.setObjectName("TaskInfoDesc")
        info.setWordWrap(True)
        lay.addWidget(info)

        # 내보낼 항목 선택
        self.chk_exp_tasks = QCheckBox("할 일 + 로그 + 진행 그룹")
        self.chk_exp_tasks.setObjectName("TaskCheck")
        self.chk_exp_tasks.setChecked(True)
        lay.addWidget(self.chk_exp_tasks)

        self.chk_exp_sched = QCheckBox("개인 일정 (schedules)")
        self.chk_exp_sched.setObjectName("TaskCheck")
        self.chk_exp_sched.setChecked(True)
        lay.addWidget(self.chk_exp_sched)

        lay.addStretch()

        self.lbl_exp_status = QLabel("")
        self.lbl_exp_status.setObjectName("FormLabel")
        self.lbl_exp_status.setWordWrap(True)
        lay.addWidget(self.lbl_exp_status)

        btn_exp = QPushButton("📤  JSON 파일로 내보내기")
        btn_exp.setObjectName("PrimaryBtn")
        btn_exp.setFixedHeight(38)
        btn_exp.clicked.connect(self._do_export)
        lay.addWidget(btn_exp)

        self.tabs.addTab(w, "📤  내보내기")

    def _build_import_tab(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(16, 16, 16, 8)
        lay.setSpacing(10)

        info = QLabel(
            "JSON 백업 파일을 불러온 뒤 가져올 항목을 선택하세요.\n"
            "이미 존재하는 제목의 할 일은 건너뜁니다."
        )
        info.setObjectName("TaskInfoDesc")
        info.setWordWrap(True)
        lay.addWidget(info)

        btn_load = QPushButton("📂  파일 선택")
        btn_load.setObjectName("SecondaryBtn")
        btn_load.setFixedHeight(34)
        btn_load.clicked.connect(self._load_import_file)
        lay.addWidget(btn_load)

        # 항목 목록
        self.scroll_imp = QScrollArea()
        self.scroll_imp.setWidgetResizable(True)
        self.scroll_imp.setMinimumHeight(200)
        self._imp_cont = QWidget()
        self._imp_lay  = QVBoxLayout(self._imp_cont)
        self._imp_lay.setContentsMargins(4, 4, 4, 4)
        self._imp_lay.setSpacing(4)
        self._imp_lay.addStretch()
        self.scroll_imp.setWidget(self._imp_cont)
        lay.addWidget(self.scroll_imp, 1)

        imp_btns = QHBoxLayout()
        btn_all = QPushButton("전체 선택")
        btn_all.setObjectName("SecondaryBtn")
        btn_all.setFixedHeight(30)
        btn_all.clicked.connect(lambda: [c.setChecked(True) for c in self._import_chks])
        imp_btns.addWidget(btn_all)
        btn_none = QPushButton("전체 해제")
        btn_none.setObjectName("SecondaryBtn")
        btn_none.setFixedHeight(30)
        btn_none.clicked.connect(lambda: [c.setChecked(False) for c in self._import_chks])
        imp_btns.addWidget(btn_none)
        imp_btns.addStretch()
        lay.addLayout(imp_btns)

        self.lbl_imp_status = QLabel("")
        self.lbl_imp_status.setObjectName("FormLabel")
        self.lbl_imp_status.setWordWrap(True)
        lay.addWidget(self.lbl_imp_status)

        self.btn_do_import = QPushButton("📥  선택한 항목 가져오기")
        self.btn_do_import.setObjectName("PrimaryBtn")
        self.btn_do_import.setFixedHeight(38)
        self.btn_do_import.setEnabled(False)
        self.btn_do_import.clicked.connect(self._do_import)
        lay.addWidget(self.btn_do_import)

        self.tabs.addTab(w, "📥  가져오기")

    # ── 내보내기 ──────────────────────────────────────────────────────────
    def _do_export(self):
        import json as _json
        path, _ = QFileDialog.getSaveFileName(
            self, "백업 파일 저장", str(Path.home() / "calendar_backup.json"),
            "JSON 파일 (*.json)"
        )
        if not path:
            return
        data: dict = {
            "version": APP_VERSION,
            "exported_at": datetime.now().isoformat(),
        }
        if self.chk_exp_tasks.isChecked():
            data["tasks"] = self._collect_tasks()
        if self.chk_exp_sched.isChecked():
            data["schedules"] = self._collect_schedules()
        try:
            with open(path, "w", encoding="utf-8") as f:
                _json.dump(data, f, ensure_ascii=False, indent=2)
            task_cnt = len(data.get("tasks", []))
            sched_cnt = len(data.get("schedules", []))
            self.lbl_exp_status.setText(
                f"✅ 저장 완료: 할 일 {task_cnt}개, 일정 {sched_cnt}개\n{path}"
            )
            self.lbl_exp_status.setStyleSheet("color:#a6e3a1;")
        except Exception as e:
            self.lbl_exp_status.setText(f"⚠ 저장 실패: {e}")
            self.lbl_exp_status.setStyleSheet("color:#f38ba8;")

    def _collect_tasks(self) -> list:
        tasks = []
        for t in self.db.get_tasks():
            td = dict(t)
            td["logs"] = [dict(l) for l in self.db.get_general_logs(t["id"])]
            td["progress_groups"] = []
            for g in self.db.get_progress_groups(t["id"]):
                gd = dict(g)
                gd["entries"] = [dict(e) for e in self.db.get_progress_logs(g["id"])]
                td["progress_groups"].append(gd)
            td["task_files"] = [
                {"filename": f["filename"], "original_path": f["original_path"]}
                for f in self.db.get_task_files(t["id"])
            ]
            tasks.append(td)
        # 완료된 태스크도 포함
        for t in self.db.get_tasks(completed=True):
            td = dict(t)
            td["logs"] = [dict(l) for l in self.db.get_general_logs(t["id"])]
            td["progress_groups"] = []
            td["task_files"] = [
                {"filename": f["filename"], "original_path": f["original_path"]}
                for f in self.db.get_task_files(t["id"])
            ]
            tasks.append(td)
        return tasks

    def _collect_schedules(self) -> list:
        return [dict(s) for s in self.db.get_schedules()]

    # ── 가져오기 ──────────────────────────────────────────────────────────
    def _load_import_file(self):
        import json as _json
        path, _ = QFileDialog.getOpenFileName(
            self, "백업 파일 선택", str(Path.home()), "JSON 파일 (*.json)"
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = _json.load(f)
        except Exception as e:
            self.lbl_imp_status.setText(f"⚠ 파일 오류: {e}")
            self.lbl_imp_status.setStyleSheet("color:#f38ba8;")
            return
        self._import_tasks = data.get("tasks", [])
        self._import_sched = data.get("schedules", [])
        # 목록 렌더링
        while self._imp_lay.count() > 1:
            item = self._imp_lay.takeAt(0)
            if item.widget(): item.widget().deleteLater()
        self._import_chks = []
        type_labels = {
            "todo": "📝 과제", "urgent": "🚨 긴급",
            "misc": "📌 기타", "personal": "👤 개인",
        }
        for t in self._import_tasks:
            lbl = type_labels.get(t.get("task_type", ""), "?")
            chk = QCheckBox(f"[{lbl}]  {t.get('title','?')}")
            chk.setObjectName("TaskCheck")
            chk.setChecked(True)
            self._imp_lay.insertWidget(self._imp_lay.count() - 1, chk)
            self._import_chks.append(chk)
        if self._import_sched:
            sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine); sep.setMaximumHeight(1)
            self._imp_lay.insertWidget(self._imp_lay.count() - 1, sep)
            self.chk_imp_sched = QCheckBox(f"일정 {len(self._import_sched)}개 전체")
            self.chk_imp_sched.setObjectName("TaskCheck")
            self.chk_imp_sched.setChecked(True)
            self._imp_lay.insertWidget(self._imp_lay.count() - 1, self.chk_imp_sched)
        else:
            self.chk_imp_sched = None
        src_ver = data.get("version", "?")
        src_date = data.get("exported_at", "?")[:16]
        self.lbl_imp_status.setText(
            f"파일: {Path(path).name}  ({src_ver}, {src_date})\n"
            f"할 일 {len(self._import_tasks)}개, 일정 {len(self._import_sched)}개"
        )
        self.lbl_imp_status.setStyleSheet("color:#a6adc8;")
        self.btn_do_import.setEnabled(bool(self._import_tasks or self._import_sched))

    def _do_import(self):
        added_t = added_s = 0
        for chk, task in zip(self._import_chks, self._import_tasks):
            if not chk.isChecked():
                continue
            try:
                new_id = self.db.add_task(
                    task.get("title",""),
                    task.get("description",""),
                    task.get("goal",""),
                    task.get("task_type", "todo"),
                    task.get("priority", 2),
                    task.get("due_date"),
                    source=task.get("source", SOURCE_MANUAL),
                    color=task.get("color"),
                    file_path=task.get("file_path"),
                )
                # 완료 상태 · 생성일 · 정렬순서 원본 복원
                patches = {}
                if task.get("is_completed"):
                    patches["is_completed"] = 1
                    patches["completed_at"] = task.get("completed_at")
                if task.get("created_at"):
                    patches["created_at"] = task["created_at"]
                if task.get("sort_order"):
                    patches["sort_order"] = task["sort_order"]
                if patches:
                    self.db.update_task(new_id, **patches)
                # 로그 복원
                for lg in task.get("logs", []):
                    self.db.add_log(new_id, lg.get("content",""), lg.get("file_path"))
                # 진행 그룹 복원
                for g in task.get("progress_groups", []):
                    grp_id = self.db.add_progress_group(new_id, g.get("title","복원된 그룹"))
                    for e in g.get("entries", []):
                        self.db.add_progress_log(new_id, grp_id, e.get("content",""))
                # 첨부파일 복원 (원본 경로에 파일이 존재할 경우)
                for tf in task.get("task_files", []):
                    orig = tf.get("original_path", "")
                    if orig and Path(orig).exists():
                        try:
                            self.db.add_task_file(new_id, orig)
                        except Exception:
                            pass
                added_t += 1
            except Exception:
                pass
        if self.chk_imp_sched and self.chk_imp_sched.isChecked():
            for s in self._import_sched:
                try:
                    self.db.add_schedule(
                        s.get("name",""), s.get("event_date",""),
                        s.get("end_date"), s.get("start_time"),
                        s.get("location",""), s.get("content",""),
                        s.get("event_type","schedule"),
                    )
                    added_s += 1
                except Exception:
                    pass
        self.lbl_imp_status.setText(f"✅ 완료: 할 일 {added_t}개, 일정 {added_s}개 가져옴")
        self.lbl_imp_status.setStyleSheet("color:#a6e3a1;")
        self.btn_do_import.setEnabled(False)


# ═══════════════════════════════════════════════════════════════════════════
# 15. TITLE BAR
# ═══════════════════════════════════════════════════════════════════════════

class _ElidedLabel(QLabel):
    """공간이 부족할 때 '...' 말줄임표로 텍스트를 자르는 QLabel."""
    def paintEvent(self, event):
        from PySide6.QtGui import QPainter
        painter = QPainter(self)
        metrics = self.fontMetrics()
        elided = metrics.elidedText(
            self.text(), Qt.TextElideMode.ElideRight, self.width()
        )
        painter.setFont(self.font())
        painter.setPen(self.palette().windowText().color())
        painter.drawText(self.rect(), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, elided)


class TitleBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("TitleBar")
        self.setFixedHeight(48)
        self._drag: QPoint | None = None
        self._build()

    def _build(self):
        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 0, 8, 0)
        lay.setSpacing(4)

        nm = _ElidedLabel("📋  ToDo & Calendar")
        nm.setObjectName("TitleLabel")
        nm.setFont(QFont("맑은 고딕", 12, QFont.Weight.Bold))
        nm.setToolTip(f"버전 {APP_VERSION}  ({APP_VERSION_DATE})")
        nm.setSizePolicy(QSizePolicy.Policy.Ignored, QSizePolicy.Policy.Preferred)
        nm.setMinimumWidth(60)
        lay.addWidget(nm, 1)

        ver_lbl = QLabel(APP_VERSION)
        ver_lbl.setObjectName("VersionLabel")
        ver_lbl.setToolTip(f"업데이트: {APP_VERSION_DATE}")
        ver_lbl.setMinimumWidth(56)
        lay.addWidget(ver_lbl)

        self.btn_backup = QPushButton("💾")
        self.btn_backup.setObjectName("TitleBtn")
        self.btn_backup.setFixedSize(34, 34); self.btn_backup.setToolTip("데이터 백업 / 복원 (JSON)")
        lay.addWidget(self.btn_backup)

        self.btn_export = QPushButton("📤")
        self.btn_export.setObjectName("TitleBtn")
        self.btn_export.setFixedSize(34, 34); self.btn_export.setToolTip("업무 내보내기 (보고용)")
        lay.addWidget(self.btn_export)

        self.btn_help = QPushButton("❓")
        self.btn_help.setObjectName("TitleBtn")
        self.btn_help.setFixedSize(34, 34); self.btn_help.setToolTip("사용 안내서")
        lay.addWidget(self.btn_help)

        self.btn_search = QPushButton("🔍 검색")
        self.btn_search.setObjectName("TitleBtn")
        self.btn_search.setFixedHeight(34)
        self.btn_search.setMinimumWidth(52)
        self.btn_search.setToolTip("검색 (Ctrl+F)")
        lay.addWidget(self.btn_search)

        self.btn_options = QPushButton("⚙")
        self.btn_options.setObjectName("TitleBtn")
        self.btn_options.setFixedSize(34, 34); self.btn_options.setToolTip("설정 (Ctrl+,)")
        lay.addWidget(self.btn_options)

        self.btn_pin = QPushButton("📌")
        self.btn_pin.setObjectName("TitleBtn"); self.btn_pin.setCheckable(True)
        self.btn_pin.setFixedSize(34, 34); self.btn_pin.setToolTip("항상 위에 표시 (Ctrl+T)")
        lay.addWidget(self.btn_pin)

        self.btn_col = QPushButton("▲")
        self.btn_col.setObjectName("TitleBtn")
        self.btn_col.setFixedSize(34, 34); self.btn_col.setToolTip("섹션 접기 (Ctrl+M)")
        lay.addWidget(self.btn_col)

        self.btn_close = QPushButton("✕")
        self.btn_close.setObjectName("TitleBtnClose")
        self.btn_close.setFixedSize(34, 34); self.btn_close.setToolTip("닫기")
        lay.addWidget(self.btn_close)

    def mousePressEvent(self, e):
        if e.button() == Qt.MouseButton.LeftButton:
            # 핀 고정 중이면 드래그 불가 (위치 잠금)
            if not self.btn_pin.isChecked():
                self._drag = e.globalPosition().toPoint() - self.window().frameGeometry().topLeft()

    def mouseMoveEvent(self, e):
        if e.buttons() == Qt.MouseButton.LeftButton and self._drag:
            self.window().move(e.globalPosition().toPoint() - self._drag)

    def mouseReleaseEvent(self, e):
        self._drag = None


# ═══════════════════════════════════════════════════════════════════════════
# 15. MAIN WINDOW
# ═══════════════════════════════════════════════════════════════════════════

class MainWindow(QWidget):
    def __init__(self, db: Database):
        super().__init__()
        self.db               = db
        self._collapse_state  = 0
        self._saved_full_height = WINDOW_HEIGHT
        self.settings = QSettings("CalendarTodoList", "MainWindowV2")
        self._setup_win()
        self._load_qss()
        self._build()
        self._connect()
        self._shortcuts()
        self._setup_tray()
        self._setup_deadline_timer()
        self._setup_ical_timer()
        self._setup_ical_notif_timer()
        self._position()
        self._restore()
        # 시작 1분 후 첫 마감일 알림 체크
        QTimer.singleShot(60_000, self._check_deadlines)
        # 시작 3초 후 첨부 파일 경로 누락 체크
        QTimer.singleShot(3_000, self._check_missing_files)
        # 시작 10초 후 7일 주기 자동 DB 백업 체크
        QTimer.singleShot(10_000, self._auto_backup_db)

    def _compute_window_size(self) -> tuple[int, int]:
        """현재 배치될 스크린 해상도에 비례해 창 크기 계산 (기준 비율 유지)"""
        tgt = self._pick_monitor()
        geo = tgt.availableGeometry()
        # 기준: WINDOW_WIDTH/WINDOW_HEIGHT = 0.5 (aspect ratio)
        aspect = WINDOW_WIDTH / WINDOW_HEIGHT
        target_h = int(min(geo.height() * 0.96, WINDOW_HEIGHT))
        target_w = int(target_h * aspect)
        return max(target_w, 380), max(target_h, 500)

    def _setup_win(self):
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.Window)
        self.setObjectName("MainWindow")
        w, h = self._compute_window_size()
        self.resize(w, h)
        self.setMinimumSize(380, 400)

    def _load_qss(self):
        qss = ASSETS_DIR / "style.qss"
        base = ""
        try:
            with open(qss, encoding="utf-8") as f:
                base = f.read()
        except FileNotFoundError:
            pass
        # 저장된 테마 기준으로 QSS 적용 (dark도 포함 — 테마별 오버라이드 항상 적용)
        theme = self.settings.value("theme", "dark")
        fs = int(self.settings.value("font_size", 10))
        self.setStyleSheet(base + EXTRA_QSS + build_theme_qss(theme, fs))

    def _build(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(1, 1, 1, 1)
        lay.setSpacing(0)

        self.tb = TitleBar(self)
        lay.addWidget(self.tb)

        sep = QFrame(); sep.setObjectName("Separator")
        sep.setFrameShape(QFrame.Shape.HLine); sep.setMaximumHeight(1)
        lay.addWidget(sep)

        # 검색 바 (Ctrl+F, 기본 숨김)
        self._search_bar = self._make_search_bar()
        self._search_bar.setVisible(False)
        lay.addWidget(self._search_bar)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        cont = QWidget(); cont.setObjectName("ContentWidget")
        cl = QVBoxLayout(cont)
        cl.setContentsMargins(10, 10, 10, 10); cl.setSpacing(10)

        self.calendar = CalendarWidget(self.db)
        cl.addWidget(self.calendar)

        self.sec_cowork = CoworkTodaySection(self.db)
        cl.addWidget(self.sec_cowork)

        self.sec_todo = TaskSection(self.db, TASK_TODO, "📝  과제 / 할 일 목록",
                                    header_color="#89b4fa")
        cl.addWidget(self.sec_todo)

        self.sec_urgent = TaskSection(self.db, TASK_URGENT, "🚨  이번주 긴급 업무",
                                      header_color="#f38ba8")
        cl.addWidget(self.sec_urgent)

        self.sec_completed = CompletedSection(self.db)
        cl.addWidget(self.sec_completed)

        self.sec_schedule = ScheduleSection(self.db, self.calendar,
                                             header_color="#f9e2af")
        cl.addWidget(self.sec_schedule)

        self.sec_misc = MiscSection(self.db, header_color="#94e2d5")
        cl.addWidget(self.sec_misc)

        self.sec_personal = TaskSection(self.db, TASK_PERSONAL, "👤  개인업무",
                                        header_color="#cba6f7")
        cl.addWidget(self.sec_personal)

        cl.addStretch()
        self.scroll.setWidget(cont)
        lay.addWidget(self.scroll)

    def _connect(self):
        self.tb.btn_close.clicked.connect(self._on_close_btn)
        self.tb.btn_col.clicked.connect(self._on_collapse)
        self.tb.btn_pin.toggled.connect(self._on_pin)
        self.tb.btn_help.clicked.connect(self._open_help)
        self.tb.btn_backup.clicked.connect(self._open_backup)
        self.tb.btn_export.clicked.connect(self._open_export)
        self.tb.btn_options.clicked.connect(self._open_options)
        self.tb.btn_search.clicked.connect(self._toggle_search)
        self.sec_todo.completion_changed.connect(self.sec_completed.refresh)
        self.sec_urgent.completion_changed.connect(self.sec_completed.refresh)
        # 달력 날짜 선택 → 태스크 하이라이트 연동
        self.calendar.date_selected.connect(self._on_date_selected)
        self.calendar.add_schedule_requested.connect(self.sec_schedule.add_for_date)
        self.calendar.add_personal_task_requested.connect(self.sec_personal.add_for_date)
        # 긴급업무 연결 과제 네비게이션
        self.sec_urgent.navigate_to.connect(self._scroll_to_task)
        # 모니터 구성 변경 감지
        from PySide6.QtGui import QGuiApplication
        QGuiApplication.instance().screenAdded.connect(lambda _: self._on_screen_changed())
        QGuiApplication.instance().screenRemoved.connect(lambda _: self._on_screen_changed())

    def _shortcuts(self):
        QShortcut(QKeySequence("Ctrl+T"), self,
                  lambda: self.tb.btn_pin.setChecked(not self.tb.btn_pin.isChecked()))
        QShortcut(QKeySequence("Ctrl+M"), self, self._on_collapse)
        QShortcut(QKeySequence("Ctrl+,"), self, self._open_options)
        QShortcut(QKeySequence("Ctrl+F"), self, self._toggle_search)
        # 포커스 위치에 관계없이 과제/할 일 추가 (WindowShortcut = 창에 포커스만 있으면 동작)
        QShortcut(QKeySequence("Ctrl+N"), self, self.sec_todo._add)

    # ── 검색 ─────────────────────────────────────────────────────────────────
    def _make_search_bar(self) -> QWidget:
        w = QWidget()
        w.setObjectName("SearchBar")
        w.setStyleSheet(
            "QWidget#SearchBar{background:#181825;border-bottom:1px solid #313244;}"
        )
        lay = QHBoxLayout(w)
        lay.setContentsMargins(10, 6, 10, 6); lay.setSpacing(8)

        icon_lbl = QLabel("🔍")
        icon_lbl.setStyleSheet("background:transparent;font-size:14px;")
        lay.addWidget(icon_lbl)

        self._search_edit = QLineEdit()
        self._search_edit.setPlaceholderText("할 일·긴급·개인업무·일정 검색... (Esc: 닫기)")
        self._search_edit.setObjectName("SearchEdit")
        self._search_edit.textChanged.connect(self._on_search)
        lay.addWidget(self._search_edit, 1)

        self._search_count = QLabel("")
        self._search_count.setStyleSheet(
            "color:#6c7086;font-size:11px;background:transparent;"
        )
        lay.addWidget(self._search_count)

        btn_close = QPushButton("✕")
        btn_close.setObjectName("TaskDeleteBtn")
        btn_close.setFixedSize(24, 24)
        btn_close.clicked.connect(self._close_search)
        lay.addWidget(btn_close)

        esc = QShortcut(QKeySequence("Escape"), self._search_edit)
        esc.setContext(Qt.ShortcutContext.WidgetShortcut)
        esc.activated.connect(self._close_search)

        return w

    def _toggle_search(self):
        visible = not self._search_bar.isVisible()
        self._search_bar.setVisible(visible)
        if visible:
            self._search_edit.setFocus()
            self._search_edit.selectAll()
        else:
            self._close_search()

    def _close_search(self):
        self._search_bar.setVisible(False)
        self._search_edit.clear()
        self._on_search("")

    def _on_search(self, query: str):
        q = query.strip().lower()
        total = 0
        for sec in (self.sec_todo, self.sec_urgent, self.sec_personal):
            total += sec.set_filter(q)
        total += self.sec_schedule.set_filter(q)
        self._search_count.setText(f"{total}개 결과" if q else "")

    def _scroll_to_task(self, task_id: int):
        """긴급업무 연결 과제로 스크롤 + 하이라이트"""
        if self.sec_todo._collapsed:
            self.sec_todo._toggle()
        for i in range(self.sec_todo.items_lay.count()):
            item = self.sec_todo.items_lay.itemAt(i)
            if item and item.widget():
                w = item.widget()
                if getattr(w, '_id', None) == task_id:
                    self.scroll.ensureWidgetVisible(w, 20, 20)
                    orig = w.styleSheet()
                    w.setStyleSheet(
                        orig + "QFrame#TaskItem{border:2px solid #f38ba8;border-radius:8px;}"
                    )
                    QTimer.singleShot(1800, lambda ww=w, os=orig: ww.setStyleSheet(os))
                    return

    def _check_missing_files(self):
        """첨부 파일 원본 경로 누락 여부 확인 후 트레이 알림"""
        try:
            missing = self.db.get_missing_file_tasks()
        except Exception:
            return
        if missing:
            count = len(missing)
            names = ", ".join(m["filename"] for m in missing[:3])
            if count > 3:
                names += f" 외 {count - 3}개"
            self._tray.showMessage(
                "첨부 파일 경로 변경",
                f"원본 위치를 찾을 수 없는 파일 {count}개:\n{names}",
                QSystemTrayIcon.MessageIcon.Warning,
                4000,
            )

    # ── 자동 DB 백업 ─────────────────────────────────────────────────────────
    def _auto_backup_db(self):
        """7일 주기 자동 DB 백업 — ~/.productivity_widget/backups/ 에 날짜별 저장"""
        try:
            last_str = self.settings.value("last_db_backup", "", type=str)
            today = date.today()

            if last_str:
                last_date = datetime.strptime(last_str, "%Y-%m-%d").date()
                if (today - last_date).days < 7:
                    return   # 아직 7일 미경과

            BACKUPS_DIR.mkdir(parents=True, exist_ok=True)
            src = self.db.db_path
            dst = BACKUPS_DIR / f"tasks_{today.strftime('%Y-%m-%d')}.db"
            shutil.copy2(str(src), str(dst))
            self.settings.setValue("last_db_backup", today.strftime("%Y-%m-%d"))
            _log.info("DB auto-backup → %s", dst)
        except Exception as e:
            _log.warning("DB auto-backup failed: %s", e)

    # ── 시스템 트레이 ────────────────────────────────────────────────────────
    def _setup_tray(self):
        """시스템 트레이 아이콘 설정 — 닫기 시 트레이로 최소화"""
        # 아이콘: 간단한 파란 원 (별도 아이콘 파일 없이 생성)
        icon = self._make_tray_icon()
        self._tray = QSystemTrayIcon(icon, self)
        self._tray.setToolTip("Calendar and To do list")

        menu = QMenu()
        menu.setStyleSheet(
            "QMenu{background:#313244;border:1px solid #45475a;border-radius:8px;padding:4px;}"
            "QMenu::item{padding:7px 18px;border-radius:6px;color:#cdd6f4;font-size:12px;}"
            "QMenu::item:selected{background:#45475a;}"
        )
        act_show  = QAction("📋  위젯 열기", self)
        act_check = QAction("🔔  마감일 확인", self)
        act_quit  = QAction("✕  종료", self)
        act_show.triggered.connect(self._show_window)
        act_check.triggered.connect(self._check_deadlines)
        act_quit.triggered.connect(self._quit_app)
        menu.addAction(act_show)
        menu.addAction(act_check)
        menu.addSeparator()
        menu.addAction(act_quit)

        self._tray.setContextMenu(menu)
        self._tray.activated.connect(self._on_tray_activated)
        self._tray.show()

    def _make_tray_icon(self) -> QIcon:
        """프로그래매틱으로 트레이 아이콘 생성"""
        px = QPixmap(32, 32)
        px.fill(Qt.GlobalColor.transparent)
        painter = QPainter(px)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setBrush(QColor("#89b4fa"))
        painter.setPen(Qt.PenStyle.NoPen)
        painter.drawEllipse(2, 2, 28, 28)
        painter.setBrush(QColor("#1e1e2e"))
        # 체크마크 느낌의 간단한 사각형들
        painter.drawRect(8, 14, 5, 2)
        painter.drawRect(8, 18, 12, 2)
        painter.drawRect(8, 22, 9, 2)
        painter.end()
        return QIcon(px)

    def _on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self._show_window()

    def _show_window(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def _on_close_btn(self):
        """닫기 버튼 → 트레이로 최소화 (종료 아님)"""
        self.hide()
        self._tray.showMessage(
            "Calendar and To do list",
            "트레이에서 실행 중입니다. 더블클릭으로 열기.",
            QSystemTrayIcon.MessageIcon.Information,
            2000
        )

    def _quit_app(self):
        """트레이 메뉴 '종료' → 완전 종료"""
        self.settings.setValue("window_pos", self.pos())
        self._tray.hide()
        QApplication.quit()

    # ── 카카오워크 iCal 동기화 ───────────────────────────────────────────────
    def _setup_ical_timer(self):
        self._ical_timer = QTimer(self)
        self._ical_timer.timeout.connect(self._auto_fetch_ical)
        self._restart_ical_timer()

    def _restart_ical_timer(self):
        interval_min = self.settings.value("ical_interval", 60, type=int)
        self._ical_timer.stop()
        if interval_min > 0:
            self._ical_timer.start(interval_min * 60 * 1000)

    def _auto_fetch_ical(self):
        url = _ical_url_decode(self.settings.value("ical_url", "")).strip()
        if url:
            ok, _ = self._fetch_ical(url)
            if ok:
                self.sec_cowork.refresh()

    def _setup_ical_notif_timer(self):
        """iCal 시간 지정 이벤트 15분 전 트레이 알림 (1분마다 체크)"""
        self._notified_ical: set[tuple] = set()   # (uid, date_str)
        self._notified_ical_date = date.today()
        self._ical_notif_timer = QTimer(self)
        self._ical_notif_timer.setInterval(60_000)  # 1분
        self._ical_notif_timer.timeout.connect(self._check_ical_upcoming)
        self._ical_notif_timer.start()

    def _check_ical_upcoming(self):
        """1분마다: ① 아침 브리핑 ② 관심 인원 15분 전 알림"""
        if not self.settings.value("notif_enabled", True, type=bool):
            return

        self._check_ical_morning_briefing()
        self._check_ical_watch_persons()

    def _check_ical_morning_briefing(self):
        """설정 시간에 오늘 팀 일정 전체 요약을 트레이 알림으로 발송 (하루 1회)"""
        brief_time = self.settings.value("ical_brief_time", "08:30").strip()
        try:
            bh, bm = map(int, brief_time.split(":"))
        except Exception:
            return

        now = datetime.now()
        if now.hour != bh or now.minute != bm:
            return

        today_str  = date.today().isoformat()
        last_brief = self.settings.value("last_ical_brief_date", "")
        if last_brief == today_str:
            return  # 오늘 이미 발송

        self.settings.setValue("last_ical_brief_date", today_str)

        ical_map = self.db.get_ical_date_map()
        events   = ical_map.get(today_str, [])
        if not events:
            return  # 일정 없으면 알림 생략

        def sort_key(ev):
            t = ev["start_time_str"] or ""
            return (1 if t else 0, t)

        lines = []
        for ev in sorted(events, key=sort_key)[:8]:
            t = ev["start_time_str"] or "종일"
            lines.append(f"{t}  {ev['summary']}")
        if len(events) > 8:
            lines.append(f"... 외 {len(events) - 8}건")

        self._tray.showMessage(
            f"🏢 오늘의 연구실 일정  ({len(events)}건)",
            "\n".join(lines),
            QSystemTrayIcon.MessageIcon.Information,
            10000
        )

    def _check_ical_watch_persons(self):
        """관심 인원 명단에 있는 사람의 이벤트가 15분 이내 시작 시 트레이 알림"""
        watch_str  = self.settings.value(
            "ical_watch_persons", "김성희\n김현표\n김별\n최재영")
        watch_names = [n.strip() for n in watch_str.splitlines() if n.strip()]
        if not watch_names:
            return

        today = date.today()
        if today != self._notified_ical_date:
            self._notified_ical.clear()
            self._notified_ical_date = today

        now      = datetime.now()
        ical_map = self.db.get_ical_date_map()
        events   = ical_map.get(today.isoformat(), [])

        for ev in events:
            t_str   = ev["start_time_str"]
            summary = ev["summary"]
            if not t_str:
                continue  # 종일 이벤트 제외
            if not any(name in summary for name in watch_names):
                continue  # 관심 인원 미포함

            key = (ev["uid"], today.isoformat())
            if key in self._notified_ical:
                continue

            try:
                ev_dt = datetime.strptime(
                    f"{today.isoformat()} {t_str}", "%Y-%m-%d %H:%M")
            except ValueError:
                continue

            diff_min = (ev_dt - now).total_seconds() / 60
            if 0 <= diff_min <= 15:
                matched = next(n for n in watch_names if n in summary)
                loc = f"  @ {ev['location']}" if ev["location"] else ""
                self._tray.showMessage(
                    f"🏢 {matched} 일정 알림",
                    f"곧 시작: {t_str}  {summary}{loc}\n({int(diff_min)}분 후)",
                    QSystemTrayIcon.MessageIcon.Information,
                    8000
                )
                self._notified_ical.add(key)

    def _fetch_ical(self, url: str) -> tuple[bool, str]:
        """iCal URL에서 이벤트를 다운로드·파싱·DB 저장 후 달력 갱신.
        반환: (성공여부, 상태 메시지)"""
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "CalendarTodoList/2.13"})
            with urllib.request.urlopen(req, timeout=15) as resp:
                raw = resp.read().decode("utf-8", errors="replace")
        except Exception as e:
            last_ok = self.settings.value("ical_last_sync", "")
            hint = f"  (마지막 성공: {last_ok})" if last_ok else ""
            return False, f"⚠ 연결 실패: {e}{hint}"

        try:
            parser = ICalParser()
            events = parser.parse(raw)
        except Exception as e:
            last_ok = self.settings.value("ical_last_sync", "")
            hint = f"  (마지막 성공: {last_ok})" if last_ok else ""
            return False, f"⚠ 파싱 실패: {e}{hint}"

        self.db.sync_ical_events(events)
        self.calendar.refresh()
        self.sec_cowork.refresh()
        self.settings.setValue("ical_last_sync", now_str)
        now_disp = datetime.now().strftime("%H:%M")
        return True, f"✅ {now_disp}  {len(events)}개 이벤트 동기화"

    # ── 마감일 알림 ──────────────────────────────────────────────────────────
    def _setup_deadline_timer(self):
        """매 시간 마감일 체크 타이머"""
        self._deadline_timer = QTimer(self)
        self._deadline_timer.setInterval(3_600_000)   # 1시간
        self._deadline_timer.timeout.connect(self._check_deadlines)
        self._deadline_timer.start()

    def _check_deadlines(self):
        """D-3 / D-1 / D-day / 마감 초과 트레이 알림"""
        today  = date.today()
        d1     = today + timedelta(days=1)
        d3     = today + timedelta(days=3)

        overdue, due_today, due_d1, due_d3 = [], [], [], []

        for t in self.db.get_tasks():
            if t["is_completed"] or not t["due_date"]:
                continue
            try:
                d = date.fromisoformat(t["due_date"])
            except ValueError:
                continue
            if d < today:
                overdue.append(t["title"])
            elif d == today:
                due_today.append(t["title"])
            elif d == d1:
                due_d1.append(t["title"])
            elif d == d3:
                due_d3.append(t["title"])

        # 하루 1회만 알림 (마감 초과 제외)
        last_notif = self.settings.value("last_notif_date", "")
        today_str  = today.isoformat()
        if last_notif == today_str and not overdue:
            return

        def _fmt(lst, n=2):
            s = ", ".join(lst[:n])
            return s + ("..." if len(lst) > n else "")

        lines = []
        if overdue:
            lines.append(f"🔴 마감 초과 {len(overdue)}건: {_fmt(overdue)}")
        if due_today:
            lines.append(f"🟠 D-day {len(due_today)}건: {_fmt(due_today)}")
        if due_d1:
            lines.append(f"🟡 D-1 {len(due_d1)}건: {_fmt(due_d1)}")
        if due_d3:
            lines.append(f"🟢 D-3 {len(due_d3)}건: {_fmt(due_d3)}")

        if lines:
            self._tray.showMessage(
                "📋 Calendar and To do list — 마감일 알림",
                "\n".join(lines),
                QSystemTrayIcon.MessageIcon.Warning,
                7000
            )
            self.settings.setValue("last_notif_date", today_str)

    # ── 달력 날짜 선택 → 태스크 하이라이트 ──────────────────────────────────
    def _on_date_selected(self, d: date):
        """달력에서 날짜 클릭 → 해당 날짜 마감 태스크 강조 표시"""
        # 같은 날짜 재클릭 시 하이라이트 해제 (토글)
        current = getattr(self, "_cal_filter", None)
        target  = None if (current == d) else d
        self._cal_filter = target

        self.sec_todo.highlight_date(target)
        self.sec_urgent.highlight_date(target)
        self.sec_personal.highlight_date(target)

    # ── 모니터 변경 감지 ──────────────────────────────────────────────────────
    def _on_screen_changed(self, _=None):
        """모니터 추가/제거 시 저장된 위치가 유효한지 확인 후 재배치"""
        from PySide6.QtGui import QGuiApplication
        # 현재 창 위치가 어떤 스크린에도 속하지 않으면 재배치
        pos = self.frameGeometry().center()
        on_screen = any(
            s.geometry().contains(pos) for s in QGuiApplication.screens()
        )
        if not on_screen:
            self.settings.remove("window_pos")
            self._position()

    def _apply_collapse(self, state: int):
        """
        state 0: 전체 표시
        state 1: 캘린더만 (섹션 숨김)
        state 2: 타이틀바만 (얇은 탭 — 최소화)
        """
        prev_state = self._collapse_state
        self._collapse_state = state
        sections = [self.sec_todo, self.sec_urgent,
                    self.sec_schedule, self.sec_misc, self.sec_personal]
        if state == 0:
            for w in sections: w.setVisible(True)
            self.calendar.setVisible(True)
            self.scroll.setVisible(True)
            self.tb.btn_col.setText("▲")
            self.tb.btn_col.setToolTip("캘린더만 보기 (Ctrl+M)")
            self.setMinimumHeight(400)
            self.setMaximumHeight(16777215)
            self.resize(self.width(), self._saved_full_height)
        elif state == 1:  # 캘린더만
            if prev_state == 0:
                self._saved_full_height = self.height()
            for w in sections: w.setVisible(False)
            self.calendar.setVisible(True)
            self.scroll.setVisible(True)
            self.tb.btn_col.setText("▼")
            self.tb.btn_col.setToolTip("최소화 (Ctrl+M)")
            self.setMinimumHeight(200)
            self.setMaximumHeight(16777215)
            cal_h = self.calendar.sizeHint().height()
            tb_h  = self.tb.height()
            self.resize(self.width(), max(cal_h + tb_h + 24, 280))
        else:  # state == 2: 타이틀바만
            for w in sections: w.setVisible(False)
            self.calendar.setVisible(False)
            self.scroll.setVisible(False)
            self.tb.btn_col.setText("▽")
            self.tb.btn_col.setToolTip("전체 표시 (Ctrl+M)")
            tb_h = self.tb.height()
            self.setMinimumHeight(tb_h)
            self.setMaximumHeight(tb_h)
            self.resize(self.width(), tb_h)
        self.settings.setValue("collapse_state", state)

    def _on_collapse(self):
        next_state = (self._collapse_state + 1) % 3
        self._apply_collapse(next_state)

    def _on_pin(self, pinned):
        pos = self.pos()
        f = self.windowFlags()
        if pinned: f |= Qt.WindowType.WindowStaysOnTopHint
        else:      f &= ~Qt.WindowType.WindowStaysOnTopHint
        self.setWindowFlags(f)
        self.move(pos); self.show()
        self.settings.setValue("pinned", pinned)

    def _pick_monitor(self):
        """설정의 monitor_placement 값으로 배치할 스크린 선택"""
        from PySide6.QtGui import QGuiApplication
        screens = QGuiApplication.screens()
        placement = self.settings.value("monitor_placement", "right")
        if placement == "primary":
            return QGuiApplication.primaryScreen()
        if placement == "left":
            return min(screens, key=lambda s: s.geometry().x())
        if placement.startswith("index:"):
            try:
                idx = int(placement.split(":")[1])
                if 0 <= idx < len(screens):
                    return screens[idx]
            except (ValueError, IndexError):
                pass
        # default "right": 가장 오른쪽 모니터
        return max(screens, key=lambda s: s.geometry().x())

    def _position(self):
        saved = self.settings.value("window_pos")
        if saved:
            self.move(saved); return
        tgt = self._pick_monitor()
        geo = tgt.availableGeometry()
        self.move(geo.right() - self.width(), geo.top())

    def _resize_anchored_right(self, new_width: int):
        """너비 변경 시 오른쪽 가장자리 고정 → 왼쪽으로 확장"""
        # monitor_placement 설정 기준 스크린 사용 (_pick_monitor와 일치)
        tgt = self._pick_monitor()
        geo = tgt.availableGeometry()
        # 오른쪽 끝 = 현재 x + 현재 width, 단 스크린 right를 넘지 않도록 clamp
        right_edge = min(self.x() + self.width(), geo.right())
        self.resize(new_width, self.height())
        new_x = max(geo.left(), right_edge - new_width)
        self.move(new_x, self.y())
        # window_pos 저장하지 않음 — 다음 실행 시 _position()이 스크린 기준 재계산

    def _on_screen_changed(self):
        """모니터 구성 변경 시 — 창 크기를 현재 스크린에 맞게 재계산"""
        self.settings.remove("window_pos")  # 기존 저장 위치 초기화
        w, h = self._compute_window_size()
        self.resize(w, h)
        self._position()

    def _restore(self):
        if self.settings.value("pinned", False, type=bool):
            self.tb.btn_pin.setChecked(True)
        saved_state = self.settings.value("collapse_state", 0, type=int)
        if saved_state in (1, 2):
            self._apply_collapse(saved_state)
        # 저장된 테마 / 투명도 / 창 너비 복원
        theme = self.settings.value("theme", "dark")
        if theme != "dark":
            self._apply_theme(theme)
        op = self.settings.value("opacity", 100, type=int)
        if op != 100:
            self._apply_opacity(op)
        w = max(380, min(1400, self.settings.value("window_width", WINDOW_WIDTH, type=int)))
        if w != WINDOW_WIDTH:
            self._resize_anchored_right(w)
        h = max(0, min(2160, self.settings.value("window_height", 0, type=int)))
        if h >= 400:
            self.resize(self.width(), h)
        # 섹션 가시성 복원
        self._apply_section_visibility({
            "show_calendar": self.settings.value("show_calendar", True, type=bool),
            "show_todo":     self.settings.value("show_todo",     True, type=bool),
            "show_urgent":   self.settings.value("show_urgent",   True, type=bool),
            "show_schedule": self.settings.value("show_schedule", True, type=bool),
            "show_misc":     self.settings.value("show_misc",     True, type=bool),
            "show_personal": self.settings.value("show_personal", True, type=bool),
        })
        # 글자 크기 복원
        fs = self.settings.value("font_size", 10, type=int)
        if fs != 10:
            self._apply_font_size(fs)

        # iCal URL이 설정돼 있으면 시작 후 2초 뒤 자동 동기화
        if _ical_url_decode(self.settings.value("ical_url", "")).strip():
            QTimer.singleShot(2000, self._auto_fetch_ical)

        # 알림 활성화 여부
        notif_on = self.settings.value("notif_enabled", True, type=bool)
        if not notif_on:
            self._deadline_timer.stop()

        # 첫 실행 온보딩
        if self.settings.value("first_run", True, type=bool):
            self.settings.setValue("first_run", False)
            self.settings.setValue("last_seen_version", APP_VERSION)
            QTimer.singleShot(400, self._show_onboarding)
        else:
            # 버전 업데이트 후 "새 기능" 자동 표시
            last_ver = self.settings.value("last_seen_version", "")
            if last_ver != APP_VERSION:
                self.settings.setValue("last_seen_version", APP_VERSION)
                QTimer.singleShot(600, lambda: self._open_help("whatsnew"))

    def _show_onboarding(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("📋 처음 오셨군요! — 간단 사용 가이드")
        msg.setText(
            "<b>Calendar and To do list</b>에 오신 것을 환영합니다!<br><br>"
            "이 앱에는 4가지 섹션이 있습니다:<br><br>"
            "📝 <b>과제 / 할 일</b> — 마감일 기준으로 관리하는 장기 과제·할 일<br>"
            "🚨 <b>긴급업무</b> — 이번 주 안에 처리해야 하는 단기 업무<br>"
            "👤 <b>개인업무</b> — 나만 보는 메모·약속·개인 일정<br>"
            "📅 <b>일정</b> — 날짜·시간이 정해진 이벤트 (캘린더와 연동)<br><br>"
            "<b>💡 팁</b>: 캘린더 날짜를 <b>우클릭</b>하면 일정을 바로 추가할 수 있어요.<br>"
            "상단 🔍 <b>검색</b> 버튼으로 모든 항목을 검색할 수 있습니다."
        )
        msg.setStandardButtons(QMessageBox.StandardButton.Ok)
        msg.button(QMessageBox.StandardButton.Ok).setText("시작하기")
        msg.exec()

    # ── 사용 안내서 다이얼로그 ──────────────────────────────────────────────
    def _open_help(self, show_tab: str = "whatsnew"):
        dlg = HelpDialog(self, show_tab=show_tab)
        dlg.exec()

    # ── 백업/복원 다이얼로그 ────────────────────────────────────────────────
    def _open_backup(self):
        dlg = JsonBackupDialog(self.db, self)
        dlg.exec()
        self._refresh_all()

    # ── 내보내기 다이얼로그 ─────────────────────────────────────────────────
    def _open_export(self):
        dlg = ExportDialog(self.db, self.settings, self)
        dlg.exec()

    # ── 설정 다이얼로그 ──────────────────────────────────────────────────────
    def _apply_window_height(self, h: int):
        """창 높이 즉시 적용. h=0 이면 화면 96% 자동 계산."""
        if h <= 0:
            _, h = self._compute_window_size()
        self.resize(self.width(), h)

    def _open_options(self):
        dlg = OptionsDialog(self.settings, self)
        dlg.theme_changed.connect(self._apply_theme)
        dlg.opacity_changed.connect(self._apply_opacity)
        dlg.font_size_changed.connect(self._apply_font_size)
        dlg.window_height_changed.connect(self._apply_window_height)
        dlg.section_changed.connect(lambda: self._apply_section_visibility(
            dlg.get_section_visibility()
        ))
        dlg.notif_changed.connect(self._on_notif_toggle)
        dlg.exec()
        # 창 너비 변경 반영 (오른쪽 고정, 왼쪽으로 확장)
        w = self.settings.value("window_width", WINDOW_WIDTH, type=int)
        self._resize_anchored_right(w)
        # 섹션 가시성 최종 저장
        vis = dlg.get_section_visibility()
        for k, v in vis.items():
            self.settings.setValue(k, v)
        # iCal 타이머 갱신
        self._restart_ical_timer()

    def _apply_theme(self, theme_key: str, font_size: int = None):
        """기본 QSS + EXTRA_QSS + 테마 오버라이드 순서로 적용"""
        if font_size is None:
            font_size = int(self.settings.value("font_size", 10))
        qss = ASSETS_DIR / "style.qss"
        base_qss = ""
        try:
            with open(qss, encoding="utf-8") as f:
                base_qss = f.read()
        except FileNotFoundError:
            pass
        ff = self.settings.value("font_family", "맑은 고딕")
        self.setStyleSheet(base_qss + EXTRA_QSS + build_theme_qss(theme_key, font_size, ff))
        # EventPopup도 테마 적용 (수정 5)
        self.calendar._popup.apply_theme(theme_key)

    def _apply_opacity(self, value: int):
        """투명도 설정 (0~100)"""
        self.setWindowOpacity(max(0.1, value / 100.0))

    def _apply_font_size(self, size: int):
        """앱 전체 글자 크기 변경"""
        fs = max(8, min(18, size))
        app = QApplication.instance()
        f = app.font()
        f.setPointSize(fs)
        app.setFont(f)
        # QSS 재적용 — font_size 명시적으로 전달
        self._apply_theme(self.settings.value("theme", "dark"), font_size=fs)

    def _apply_section_visibility(self, vis: dict[str, bool]):
        """섹션 가시성 일괄 적용"""
        self.calendar.setVisible(vis.get("show_calendar", True))
        self.sec_todo.setVisible(vis.get("show_todo", True))
        self.sec_urgent.setVisible(vis.get("show_urgent", True))
        self.sec_schedule.setVisible(vis.get("show_schedule", True))
        self.sec_misc.setVisible(vis.get("show_misc", True))
        self.sec_personal.setVisible(vis.get("show_personal", True))

    def _on_notif_toggle(self, enabled: bool):
        if enabled:
            self._deadline_timer.start()
            self._ical_notif_timer.start()
        else:
            self._deadline_timer.stop()
            self._ical_notif_timer.stop()

    def _refresh_all(self):
        self.calendar.refresh()
        self.sec_cowork.refresh()
        self.sec_todo.refresh()
        self.sec_urgent.refresh()
        self.sec_completed.refresh()
        self.sec_schedule.refresh()
        self.sec_misc.refresh()
        self.sec_personal.refresh()

    def closeEvent(self, e):
        """X 버튼(시스템) 클릭 시에도 트레이로"""
        e.ignore()
        self.hide()


# ═══════════════════════════════════════════════════════════════════════════
# 16. STYLE SUPPLEMENT (QSS에 없는 새 objectName 스타일)
# ═══════════════════════════════════════════════════════════════════════════

EXTRA_QSS = """
/* ── 파일 드롭 영역 ─────────────────────────────────────── */
QFrame#FileDropArea {
    border: 2px dashed #45475a;
    border-radius: 10px;
    background: #181825;
    min-height: 90px;
}

/* ── 업데이트 진행 바 ─────────────────────────────────────── */
QProgressBar#UpdateProgressBar {
    background: #27273a;
    border-radius: 4px;
    border: 1px solid #45475a;
    text-align: center;
    color: #cdd6f4;
    font-size: 10px;
}
QProgressBar#UpdateProgressBar::chunk {
    background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #89b4fa,stop:1 #a6e3a1);
    border-radius: 4px;
}

/* ── 툴팁 ───────────────────────────────────────────────── */
QToolTip {
    background-color: #27273a;
    color: #cdd6f4;
    border: 1px solid #45475a;
    border-radius: 6px;
    padding: 5px 8px;
    font-size: 12px;
}

/* ── 정렬 드롭다운 (작은 높이에 맞게 padding 축소) ─────── */
QComboBox#SortCombo {
    background: #27273a;
    border: 1px solid #3d3d58;
    border-radius: 6px;
    padding: 2px 8px;
    color: #cdd6f4;
    font-size: 12px;
}

QLabel#TaskGoal { font-size: 11px; color: #cdd6f4; font-style: italic; }
QLabel#SourceBadge { font-size: 10px; color: #45475a; background: #22223a; border-radius: 3px; padding: 0 4px; }

/* ── 기타 아이템 ─────────────────────────────────────────── */
QFrame#MiscItem {
    background-color: #22223a;
    border-radius: 8px;
    border: 1px solid #2e2e48;
    border-left: 3px solid #94e2d5;
}
QFrame#MiscItem:hover { background-color: #2a2a46; border-color: #3d3d60; border-left-color: #94e2d5; }

QLabel#MiscTitle { font-size: 12px; color: #cdd6f4; font-weight: bold; }
QLabel#MiscContent {
    font-size: 12px; color: #a6adc8;
    padding: 6px 4px 2px 4px;
    line-height: 1.5;
}

QPushButton#MiscExpandBtn {
    background: #313244; border-radius: 4px;
    color: #6c7086; font-size: 10px; font-weight: bold;
}
QPushButton#MiscExpandBtn:hover { background: #45475a; color: #94e2d5; }

/* ── 로그 편집 버튼 ──────────────────────────────────────── */
QPushButton#LogEditBtn {
    background: transparent; border-radius: 4px;
    color: #45475a; font-size: 13px;
    min-width: 22px; max-width: 22px;
}
QPushButton#LogEditBtn:hover { background: rgba(137,180,250,0.15); color: #89b4fa; }

QPushButton#LogFileBtn {
    background: transparent; border-radius: 4px;
    color: #f9e2af; font-size: 13px;
    min-width: 22px; max-width: 22px;
}
QPushButton#LogFileBtn:hover { background: rgba(249,226,175,0.2); color: #f9e2af; }

/* ── 옵션 탭 위젯 ────────────────────────────────────────── */
QTabWidget#OptionsTab::pane {
    border: 1px solid #2a2a3d; border-radius: 0 10px 10px 10px;
    background: #1e1e2e;
}
QTabWidget#OptionsTab > QTabBar::tab {
    background: #11111b; color: #6c7086;
    border-radius: 7px 7px 0 0;
    padding: 7px 16px; margin-right: 2px; font-size: 12px;
}
QTabWidget#OptionsTab > QTabBar::tab:selected {
    background: #27273a; color: #cdd6f4; font-weight: bold;
    border-bottom: 2px solid #89b4fa;
}
QTabWidget#OptionsTab > QTabBar::tab:hover:!selected {
    background: #1a1a2e; color: #a6adc8;
}

/* QSlider */
QSlider::groove:horizontal { background: #27273a; height: 4px; border-radius: 2px; }
QSlider::handle:horizontal {
    background: #89b4fa; width: 16px; height: 16px;
    border-radius: 8px; margin: -6px 0; border: 2px solid #11111b;
}
QSlider::sub-page:horizontal {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #89b4fa, stop:1 #b4befe);
    border-radius: 2px;
}

/* QSpinBox */
QSpinBox {
    background: #27273a; border: 1px solid #3d3d58;
    border-radius: 7px; padding: 6px 10px;
    color: #cdd6f4; font-size: 13px;
}
QSpinBox::up-button, QSpinBox::down-button {
    width: 18px; background: #313244; border-radius: 4px; margin: 2px;
}
QSpinBox::up-button:hover, QSpinBox::down-button:hover { background: #45475a; }
QSpinBox::up-arrow   { image: none; border-left: 4px solid transparent; border-right: 4px solid transparent; border-bottom: 5px solid #a6adc8; width: 0; height: 0; }
QSpinBox::down-arrow { image: none; border-left: 4px solid transparent; border-right: 4px solid transparent; border-top: 5px solid #a6adc8; width: 0; height: 0; }

/* ── 일정 아이템 ─────────────────────────────────────────── */
QFrame#ScheduleItem {
    background-color: #22223a;
    border-radius: 8px;
    border: 1px solid #2e2e48;
}
QFrame#ScheduleItem:hover { background-color: #2a2a46; border-color: #3d3d60; }

QLabel#ScheduleItemName    { font-size: 12px; color: #cdd6f4; font-weight: bold; }
QLabel#ScheduleItemMeta    { font-size: 11px; color: #a6adc8; }  /* 수정 7: #7f849c → #a6adc8 */
QLabel#ScheduleItemContent { font-size: 11px; color: #7f849c; }  /* 수정 7: #585b70 → #7f849c */

/* ── 이벤트 팝업 (스타일은 EventPopup 클래스 내부 setStyleSheet로 관리) ── */
QLabel#EventPopupDate { font-size: 12px; font-weight: bold; color: #89b4fa; }

/* ── 버전 레이블 ─────────────────────────────────────────── */
QLabel#VersionLabel {
    color: #45475a;
    font-size: 10px;
    font-weight: bold;
    background: #1a1a2e;
    border: 1px solid #2e2e48;
    border-radius: 4px;
    padding: 1px 6px;
    letter-spacing: 0.5px;
}
"""


# ═══════════════════════════════════════════════════════════════════════════
# 17. ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════════

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("Calendar and To do list")
    app.setOrganizationName("CalendarTodoList")

    # ── 단일 인스턴스 잠금 ────────────────────────────────────────────────
    # 첫 실행 시 디렉터리가 없으면 QLockFile이 tryLock 실패 → 앱이 조용히 종료되는 문제 방지
    _lock_dir = Path.home() / ".productivity_widget"
    _lock_dir.mkdir(exist_ok=True)
    lock_path = str(_lock_dir / "app.lock")
    lock = QLockFile(lock_path)
    lock.setStaleLockTime(30000)  # 30초 — 크래시 후 잠금 자동 해제
    if not lock.tryLock(100):
        # 이미 실행 중 → 조용히 종료
        sys.exit(0)

    font = QFont("맑은 고딕", 10)
    font.setHintingPreference(QFont.HintingPreference.PreferFullHinting)
    app.setFont(font)
    app.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )

    _log.info("App starting — %s (%s)", APP_VERSION, APP_VERSION_DATE)

    db  = Database()
    win = MainWindow(db)

    # ── 모니터 변경 감지 ──────────────────────────────────────────────────
    app.screenAdded.connect(win._on_screen_changed)
    app.screenRemoved.connect(win._on_screen_changed)

    win.show()

    code = app.exec()
    lock.unlock()
    db.close()
    _log.info("App exited (code=%d)", code)
    sys.exit(code)


if __name__ == "__main__":
    main()
