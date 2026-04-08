# ui-test — 앱 직접 실행 UI 자동화 테스트

너는 **QA 자동화 엔지니어**다.  
`ui_test_lib.py` + `ui_test.py`를 사용해 앱을 직접 조작하고 결과를 보고한다.

---

## 역할 정의

`/ui-verify`(정적 분석)가 할 수 없는 것을 직접 수행한다:

| 항목 | 방법 |
|------|------|
| 다이얼로그 실제 열림/닫힘 | `find_dialogs()` + `dialog_opened()` |
| 창 리사이즈 후 실제 크기 | `resize_window()` + `win_size()` |
| 버튼 클릭 반응 | `click()` + `screenshot()` 비교 |
| 키보드 단축키 동작 | `hotkey()` + `key()` |
| 파일 선택 다이얼로그 탐색 | `click()` 시퀀스 |
| 캘린더 달 전환 | `hotkey(Ctrl, Left/Right)` |

---

## 실행 방법

### Step 0 — 앱 상태 확인

```python
# ui_test_lib.py 임포트 후
import ui_test_lib as T
hwnd = T.find_app_hwnd()
```

- `hwnd != 0`: 앱 실행 중 → 테스트 시작
- `hwnd == 0`: `T.launch_app()` 으로 실행 후 대기

---

### Step 1 — 테스트 실행

#### 전체 실행:
```bash
python ui_test.py
```

#### 특정 섹션만:
```bash
python ui_test.py --section s4,s6
```

#### 섹션 목록:
```bash
python ui_test.py --list
```

**섹션 목록:**
| ID | 내용 |
|----|------|
| s1 | 앱 기동 및 기본 창 |
| s2 | 섹션 헤더 표시 |
| s3 | 창 크기 조절 |
| s4 | 태스크 추가 다이얼로그 (열기/닫기/리사이즈) |
| s5 | 일정 우클릭 컨텍스트 메뉴 |
| s6 | 캘린더 키보드 단축키 (Ctrl+Left/Right, Home) |
| s7 | 검색 바 열기/닫기 |
| s8 | 파일 선택 다이얼로그 (_FilePickerDialog) |
| s9 | 설정 다이얼로그 열기/닫기/리사이즈 |
| s10 | 완료 항목 섹션 |

---

### Step 2 — 스크린샷 분석

테스트 중 캡처된 스크린샷은 `troubleshooting_screenshots/` 에 저장된다.

각 FAIL 항목에 대해:
1. `Read` 로 스크린샷 파일 열기 (이미지 직접 확인)
2. 문제 원인 파악
3. `/dev-fix` 에 전달할 문제 설명 작성

---

### Step 3 — 보고서 작성

```
# UI 런타임 테스트 보고서
생성일시: YYYY-MM-DD HH:MM
대상 버전: vX.XX

## 총평
✅ PASS / ❌ FAIL

---

## 결과 요약
총 N건 | PASS M건 | FAIL K건

---

## FAIL 항목

### [T-01] 섹션 이름 — 테스트 항목명
- **증상**: 무엇이 잘못됐는가
- **스크린샷**: `troubleshooting_screenshots/파일명.png`
- **예상 원인**: 코드의 어느 부분이 문제인가
- **dev-fix 전달 내용**: 수정 요청 문구

---

## 수동 확인 필요 항목
(자동화로 확인 불가한 항목)
□ 텍스트 렌더링이 선명한가 (HiDPI)
□ 다크/라이트 테마 전환 시 색상이 올바른가
□ 긴 텍스트 입력 시 레이아웃 무너지지 않는가
□ 마우스 오버 팝업이 화면 밖으로 나가지 않는가
```

보고서를 다음 경로에 저장:
```
C:\Users\Admin\Desktop\Claude\To_do_list_and_calender\.claude\ui-test-latest.md
```

---

### Step 4 — 판정 후 안내

#### PASS 시:
```
✅ UI 런타임 테스트 통과

다음 단계: /ui-verify (정적 분석) → 릴리즈
```

#### FAIL 시:
```
❌ UI 런타임 테스트 실패

FAIL 항목을 /dev-fix 에 전달합니다.
→ /dev-fix 실행 시 ui-test-latest.md 를 자동으로 읽어 수정합니다.
```

---

## 주의사항

- 테스트 중 마우스/키보드를 조작하지 않는다 (자동화 간섭)
- 앱이 포그라운드에 있어야 키보드 이벤트가 정상 전달된다
- 다이얼로그가 열린 상태에서 다음 테스트로 넘어가지 않는다 — 항상 ESC로 닫은 후 진행
- 테스트 실패 시 앱이 비정상 상태일 수 있으므로 `T.escape()` 여러 번 호출 후 계속

---

## 워크플로우 내 위치

```
/ux-check       → 사용자 관점 문제 발견
/dev-fix        → 코드 수정 구현
/ui-verify      → 정적 코드 분석
/ui-test        → 런타임 직접 동작 검증 ← 지금 여기
  ├─ PASS → 릴리즈
  └─ FAIL → /dev-fix 재실행 → /ui-verify + /ui-test 재확인 → 릴리즈
```
