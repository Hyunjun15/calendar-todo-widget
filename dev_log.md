# 생산성 위젯 개발 로그

> 최종 업데이트: 2026-03-27
> 현재 버전: **v2.10**
> 실행: `pythonw main.py` (백그라운드) / `python main.py` (콘솔)

---

## 버전 히스토리 요약

| 버전 | 주요 변경 |
|------|-----------|
| v2.7 | calendar_analysis.md 7가지 수정, CalDayHasTasks hover 대비 개선, SCHED_TRIP 버그 수정 |
| v2.8 | 로그 파일 첨부, 더블클릭→편집 변경, TaskGoal 흰색, 수정/삭제 버튼 가시성, 배포 스크립트 |
| v2.9 | CalDayHasTasks hover 직접 stylesheet 방식으로 수정, PyInstaller BASE_DIR 처리 |
| v2.10 | 달력 마감 태스크 hover EventPopup 표시, _deadline_map 교체, 시스템 툴팁 제거 |
| v2.11 | EventPopup z-order 수정(parent 전달·raise_), enterEvent/leaveEvent 추가, QLockFile 디렉터리 선생성 |
| v2.12 | sync_from_file → insert-only 가져오기 (삭제·덮어쓰기 제거), 버튼·배지 텍스트 수정 |

---

## v2.8 작업 내역 (2026-03-27)

### 1. 로그 파일 첨부 기능
- **DB**: `task_logs` 테이블에 `file_path TEXT DEFAULT NULL` 컬럼 추가
- **마이그레이션**: `_migrate()`에 `task_logs` ALTER TABLE 추가
- **`add_log(task_id, content, file_path=None)`**: file_path 파라미터 추가
- **`update_log(log_id, content, file_path=None)`**: file_path 파라미터 추가
- **`LogItemWidget`**:
  - `edit_done = Signal(int, str, str)` (log_id, content, file_path)
  - 헤더에 `📁` 파일 버튼 추가 (file_path 있을 때)
  - 편집 모드에 파일 경로 QLineEdit + `📂` 탐색 버튼 추가
  - `_browse_file()` 메서드 추가
- **`LogDialog`**:
  - 새 로그 입력 영역에 `📂 첨부` / `✕` 버튼 추가
  - `_browse_attach()`, `_clear_attach()` 메서드 추가
  - `_edit_log()` 시그니처 변경 → `_load()` 재호출로 파일 버튼 갱신

### 2. 더블클릭 동작 변경
- **변경 전**: 파일 있으면 탐색기, 없으면 로그 다이얼로그
- **변경 후**: 항상 편집 다이얼로그 열기
- 파일 열기는 `📁` 버튼 또는 우클릭 메뉴 사용

### 3. TaskGoal 색상 흰색
- `EXTRA_QSS`: `QLabel#TaskGoal` → `color: #cdd6f4`
- `build_theme_qss()`: `color: {t['text']}`

### 4. 수정/삭제 버튼 가시성 개선
- `style.qss`: hover 시 색상 `#45475a` → `#7f849c`
- 대상: `TaskEditBtn`, `TaskDeleteBtn`

### 5. LogFileBtn QSS 추가
```css
QPushButton#LogFileBtn {
    background: transparent; border-radius: 4px;
    color: #f9e2af; font-size: 13px;
    min-width: 22px; max-width: 22px;
}
QPushButton#LogFileBtn:hover { background: rgba(249,226,175,0.2); }
```

### 6. 배포 스크립트 (`build_dist.py`)
- PyInstaller `--onedir` 빌드 자동화
- ASCII 이름으로 빌드 후 `생산성위젯`으로 rename (Windows cmd 인코딩 이슈 우회)
- 포함 파일: exe, assets/, Update works/, 바탕화면_바로가기.bat, 사용_안내.txt

---

## v2.9 작업 내역 (2026-03-27)

### CalDayHasTasks hover 최종 수정
- **문제**: QSS `:hover` 수도 클래스가 Windows 플랫폼에서 `* { color }` cascade와 충돌
- **해결**: `CalDayButton.enterEvent/leaveEvent`에서 직접 `setStyleSheet()` 적용

```python
_DEADLINE_BTN_HOVER_QSS = (
    "QPushButton {"
    " background: #6b1529; color: #ffffff;"
    " border: 1px solid #a02040; border-radius: 17px;"
    " font-size: 12px; font-weight: bold; }"
)

def enterEvent(self, e):
    super().enterEvent(e)
    if self._has_deadline and self.objectName() == "CalDayHasTasks":
        self.setStyleSheet(_DEADLINE_BTN_HOVER_QSS)
    self.hovered.emit(...)

def leaveEvent(self, e):
    super().leaveEvent(e)
    if self._has_deadline and self.objectName() == "CalDayHasTasks":
        self.setStyleSheet("")  # 앱 전역 QSS 복원
    self.unhovered.emit()
```

### PyInstaller BASE_DIR 처리
```python
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent  # 배포 exe 기준
else:
    BASE_DIR = Path(__file__).parent        # 개발 환경 기준
```

---

## v2.10 작업 내역 (2026-03-27)

### 달력 마감 태스크 hover 팝업 표시

#### 수정 1: `_deadline_dates(set)` → `_deadline_map(dict)`
```python
# 변경 전
self._deadline_dates = {
    t["due_date"] for t in tasks
    if t["due_date"] and not t["is_completed"]
    and t["task_type"] in (TASK_TODO, TASK_URGENT)
}

# 변경 후 — hover 시 태스크 상세 전달을 위해 dict로 교체
self._deadline_map = {}
for t in tasks:
    if t["due_date"] and not t["is_completed"] \
       and t["task_type"] in (TASK_TODO, TASK_URGENT):
        self._deadline_map.setdefault(t["due_date"], []).append(t)
```
- `__init__`: `_deadline_dates: set` → `_deadline_map: dict`
- `_build()`: `ds in self._deadline_dates` → `ds in self._deadline_map`

#### 수정 2: `_on_hover()` deadline 포함
```python
def _on_hover(self, d: date, global_pos):
    ds       = d.isoformat()
    events   = self._sched_map.get(ds, [])
    personal = self._personal_map.get(ds, [])
    deadline = self._deadline_map.get(ds, [])
    if events or personal or deadline:
        self._popup.show_for(d, events, personal, global_pos,
                             deadline_tasks=deadline)
    else:
        self._popup.schedule_hide()
```

#### 수정 3: `EventPopup.show_for()` deadline 카드 렌더링
```python
def show_for(self, d, events, personal_tasks, btn_global_pos, deadline_tasks=None):
    ...
    # personal 카드 이후 deadline 카드 추가
    for task in (deadline_tasks or []):
        card = QFrame(); card.setObjectName("EventPopupCard")
        prio_colors = {1: "#f38ba8", 2: "#fab387", 3: "#a6e3a1"}
        pcolor = prio_colors.get(task["priority"], "#f38ba8")
        # 📌 제목 (우선순위 색상), [섹션] 마감일, ▸ 목표 표시
```

#### 수정 4: 시스템 툴팁 제거
- `_build()`에서 `btn.setToolTip("📌 마감 예정 태스크")` 3줄 삭제
- 이유: EventPopup이 직접 렌더링하므로 불필요, 흰 배경 충돌 원인

---

## v2.11 작업 내역 (2026-03-30)

### 1. EventPopup hover 팝업 z-order 수정 (마감 태스크 hover 미표시 버그 수정)

**문제**: `EventPopup()`에 parent 미전달 → Windows에서 Tool 창이 메인 윈도우 뒤에 가려짐.
마우스를 팝업 위로 이동 시 버튼 `leaveEvent` → `schedule_hide` 150ms 후 팝업 소멸.

**수정 내용**:
- `CalendarWidget.__init__`: `EventPopup()` → `EventPopup(self)` (parent 전달)
- `EventPopup.show_for()`: `self.show()` 후 `self.raise_()` 추가 (z-order 강제 상승)
- `EventPopup`에 `enterEvent`/`leaveEvent` 추가:
  - `enterEvent`: `_hide_timer.stop()` (팝업 위에 있는 동안 사라지지 않음)
  - `leaveEvent`: `_hide_timer.start()` (팝업 벗어나면 150ms 후 숨김)

### 2. 배포 첫 실행 시 앱 무응답 버그 수정

**문제**: `main()`에서 `QLockFile` 생성 전 `~/.productivity_widget/` 디렉터리가 없으면
`tryLock()` 실패 → `sys.exit(0)` 조용히 종료. 신규 유저 첫 실행 시 항상 발생.

**수정 내용** (`main()`):
```python
# 변경 전
lock_path = str(Path.home() / ".productivity_widget" / "app.lock")
lock = QLockFile(lock_path)

# 변경 후
_lock_dir = Path.home() / ".productivity_widget"
_lock_dir.mkdir(exist_ok=True)          # 첫 실행 시 디렉터리 선생성
lock_path = str(_lock_dir / "app.lock")
lock = QLockFile(lock_path)
```
`Database.__init__`의 `app_dir.mkdir()`보다 먼저 실행되므로 별도 생성 필요.

---

## v2.12 작업 내역 (2026-03-31)

### txt 파일 가져오기 방식 변경 (동기화 → 단방향 추가)

**문제**: `sync_from_file()`이 매 새로고침마다 source='file' 태스크를 파일 기준으로
삭제·덮어씌움 → 사용자가 위젯에서 편집/추가한 내용이 롤백됨.

**수정**: `sync_from_file()`을 "신규 항목만 INSERT" 방식으로 변경.
- 기존 DB 항목은 source 구분 없이 절대 삭제·수정하지 않음
- 파일에서 가져온 이후 모든 수정/삭제는 사용자가 직접 관리
- INSERT 대상: 동일 제목이 DB(같은 task_type)에 없는 항목만
- 반환값: 추가된 항목 수 (상태 메시지에 표시)

**변경 파일/함수**:
- `Database.sync_from_file()`: 삭제 루프·UPDATE 블록 제거, insert-only + count 반환
- `UpdatePanel.do_update()`: 총 추가 수 집계 후 상태 메시지 표시
- `UpdatePanel._build()`: "지금 갱신" → "신규 가져오기"
- `TaskItemWidget._build()`: "📄 파일 연동" → "📄 파일 가져옴"

---

## 현재 파일 구조

```
main.py                  # 진입점 + 전체 UI (단일 파일)
database.py              # SQLite 래퍼 (참고용 — 실제는 main.py 내장)
build_dist.py            # 배포 패키지 빌드 스크립트
assets/style.qss         # QSS 다크 테마
Update works/            # 업무 내역 txt 파일
배포패키지/생산성위젯/   # 빌드 출력 (zip 압축 후 배포)
```

---

## 알려진 이슈 / 미완료 항목

현재 없음. 다음 세션에서 사용자 요청 사항에 따라 진행.

---

## 주요 상수 / 구조 참고

```python
# 태스크 타입
TASK_TODO     = "todo"
TASK_URGENT   = "urgent"
TASK_MISC     = "misc"
TASK_PERSONAL = "personal"

# 일정 타입
SCHED_SINGLE   = "schedule"
SCHED_VACATION = "vacation"
SCHED_TRAINING = "training"
SCHED_TRIP     = "trip"

# DB 위치
~/.productivity_widget/tasks.db

# QSS 적용 순서 (나중이 우선)
style.qss → EXTRA_QSS → build_theme_qss(theme)
```

### build_theme_qss() 구조
- `* { color }`, `QWidget { background }` — 전역 기본값
- 카드 배경: `QFrame#TaskItem`, `LogItem`, `ScheduleItem`
- 텍스트: `QLabel#TaskTitle`, `TaskGoal`, `TaskTitleDone` 등
- `cal_deadline_css`: 라이트/다크 테마별 CalDayHasTasks 스타일
- **주의**: CalDayHasTasks hover는 QSS가 아닌 `enterEvent/leaveEvent`에서 직접 처리

### CalDayButton hover 방식
QSS `:hover`를 쓰지 않고 `enterEvent`에서 `self.setStyleSheet(...)`, `leaveEvent`에서 `self.setStyleSheet("")` 직접 호출. Windows 플랫폼 cascade 문제 우회.
