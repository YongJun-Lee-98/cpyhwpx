# pyhwpx → cpyhwpx 함수 매핑 개요

## 프로젝트 목표
pyhwpx(Python)를 cpyhwpx(C++ Python 확장 모듈)로 포팅

## 소스 파일 구조

| 파일 | 줄 수 | 함수/메서드 | 역할 |
|------|-------|------------|------|
| core.py | 9,455 | ~311개 | 메인 Hwp 클래스, Ctrl, XHwpDocuments, XHwpDocument |
| run_methods.py | 5,077 | 684개 | HAction.Run() 액션 래퍼 (RunMethods 클래스) |
| param_helpers.py | 510 | 90개 | 파라미터 변환 헬퍼 (ParamHelpers 클래스) |
| fonts.py | 1,779 | 111개 프리셋 | 폰트 정의 딕셔너리 |
| **총계** | **~16,821** | **~1,196개** | |

---

## 클래스 구조

### 상속 관계
```
Hwp(ParamHelpers, RunMethods)
├── ParamHelpers  ← param_helpers.py (90개 헬퍼 메서드)
└── RunMethods    ← run_methods.py (678개 Run 액션)
```

### 주요 클래스

#### 1. Hwp (core.py:933~)
메인 클래스. HWP 자동화의 최상위 인터페이스.
- ParamHelpers, RunMethods를 상속
- COM 객체(HWPFrame.HwpObject)를 래핑
- 200+ 고수준 메서드 제공

#### 2. Ctrl (core.py:398~671)
컨트롤(표, 그림, 글상자 등) 래퍼 클래스.
- `CtrlID`: 컨트롤 종류 (tbl, gso, eqed 등)
- `CtrlCh`: 컨트롤 문자 (1~31)
- `Next`, `Prev`: 연결 리스트 탐색
- `Properties`: 속성 파라미터셋

#### 3. XHwpDocuments (core.py:672~771)
열린 문서 컬렉션 래퍼.
- `__getitem__`, `__iter__`, `__len__` 지원
- `Add()`, `Close()`, `FindItem()` 메서드

#### 4. XHwpDocument (core.py:774~930)
단일 문서 래퍼.
- `Open()`, `Save()`, `SaveAs()`, `Clear()`, `Close()`
- `DocumentID`, `FullName`, `Path`, `Modified` 등 속성

#### 5. ParamHelpers (param_helpers.py:4~510)
파라미터 변환 헬퍼 메서드 모음.
- 선 스타일, 색상, 정렬, 테이블 등 상수 변환
- `HwpLineType()`, `RGBColor()`, `NumberFormat()` 등

#### 6. RunMethods (run_methods.py:28~)
HAction.Run() 액션 래퍼 클래스.
- 678개 Run 액션 메서드
- `BreakPara()`, `Cancel()`, `SelectAll()` 등

---

## 매핑 문서 목록

| 문서 | 내용 | 함수 수 | 상태 |
|------|------|--------|------|
| 00_overview.md | 전체 요약 및 진행상황 | - | ✅ 완료 |
| 01_core_classes.md | Hwp, Ctrl, XHwpDocument 클래스 | 4개 클래스 | ✅ 완료 |
| 02_properties.md | 속성(Properties) 매핑 | 35개 | ✅ 완료 |
| 03_file_io.md | 파일 입출력 | ~25개 | ✅ 완료 |
| 04_text_editing.md | 텍스트 편집 | ~35개 | ✅ 완료 |
| 05_table_operations.md | 표 관련 | ~65개 | ✅ 완료 |
| 06_field_metatag.md | 필드/메타태그 | ~25개 | ✅ 완료 |
| 07_shape_objects.md | 개체(그림, 도형) | ~60개 | ✅ 완료 |
| 08_style_formatting.md | 스타일, 글자/문단 모양 | ~70개 | ✅ 완료 |
| 09_run_actions.md | HAction.Run() 액션 | 684개 | ✅ 완료 |
| 10_param_helpers.md | 파라미터 헬퍼 | 90개 | ✅ 완료 |
| 11_fonts.md | 폰트 정의 | 111개 프리셋 | ✅ 완료 |
| 12_utility.md | 유틸리티 함수 | ~60개 | ✅ 완료 |

---

## 유틸리티 함수 (core.py 상단)

| # | 함수명 | 시그니처 | 역할 |
|---|--------|---------|------|
| 1 | `com_initialized` | (func) → wrapper | COM 초기화 데코레이터 |
| 2 | `log_error` | (method) → wrapper | 에러 로깅 데코레이터 |
| 3 | `addr_to_tuple` | (cell_address: str) → Tuple[int, int] | 엑셀 주소("A1") → 튜플(1, 1) |
| 4 | `tuple_to_addr` | (row: int, col: int) → str | 튜플(1, 1) → 엑셀 주소("A1") |
| 5 | `crop_data_from_selection` | (data, selection) → List[str] | 선택 영역 데이터 추출 |
| 6 | `check_registry_key` | (key_name: str) → bool | 보안모듈 레지스트리 확인 |
| 7 | `rename_duplicates_in_list` | (file_list: List[str]) → List[str] | 중복 파일명 처리 |
| 8 | `check_tuple_of_ints` | (var: tuple) → bool | 정수 튜플 검증 |
| 9 | `excel_address_to_tuple_zero_based` | (address: str) → Tuple | 0-based 엑셀 주소 변환 |

---

## 주요 의존성

### Python 패키지
- `pywin32`: COM 인터페이스 (win32com, pythoncom)
- `pandas`, `numpy`: 데이터 처리
- `PIL (Pillow)`: 이미지 처리
- `pyperclip`: 클립보드 처리

### Windows API
- `win32api`, `win32gui`, `win32con`: 윈도우 제어
- `pythoncom`: COM 스레드 관리

### COM 인터페이스
- `HWPFrame.HwpObject`: 한/글 메인 객체
- Type Library GUID: `{7D2B6F3C-1D95-4E0C-BF5A-5EE564186FBC}`

---

## C++ 포팅 시 고려사항

### 1. COM 초기화
```cpp
// Python: pythoncom.CoInitialize()
// C++:
CoInitialize(NULL);  // 또는 CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
```

### 2. 타입 매핑
| Python | C++ |
|--------|-----|
| str | std::wstring |
| bool | bool |
| int | long / int |
| float | double |
| List[str] | std::vector<std::wstring> |
| dict | std::map / std::unordered_map |
| Tuple[int, int, int] | struct { int list; int para; int pos; } |

### 3. 필수 구현 순서
1. COM 초기화 및 HwpObject 생성
2. RegisterModule (보안 모듈)
3. 파일 I/O (Open, Save, SaveAs)
4. 기본 편집 (InsertText, GetText)
5. 커서 이동 (SetPos, GetPos, MovePos)
6. 표 제어 (CreateTable, 셀 이동)
7. 필드/메타태그
8. Run 액션 (678개)
9. 파라미터 헬퍼 (90개)

---

## 진행 상황

### Part A: Mapping 문서 작성 (100% 완료)

- [x] 프로젝트 구조 분석
- [x] 00_overview.md 작성
- [x] 01_core_classes.md 작성 (4개 클래스)
- [x] 02_properties.md 작성 (35개 속성)
- [x] 03_file_io.md 작성 (~25개 함수)
- [x] 04_text_editing.md 작성 (~35개 함수)
- [x] 05_table_operations.md 작성 (~65개 함수)
- [x] 06_field_metatag.md 작성 (~25개 함수)
- [x] 07_shape_objects.md 작성 (~60개 함수)
- [x] 08_style_formatting.md 작성 (~70개 함수)
- [x] 09_run_actions.md 작성 (684개 Run 액션)
- [x] 10_param_helpers.md 작성 (90개 헬퍼)
- [x] 11_fonts.md 작성 (111개 폰트 프리셋)
- [x] 12_utility.md 작성 (~60개 함수)

### Part B: C++ 구현 (대기)

- [ ] src/ 폴더 구조 생성
- [ ] CMakeLists.txt 작성
- [ ] pybind11 설정
- [ ] HwpWrapper 기본 클래스 구현
- [ ] 파일 I/O 구현
- [ ] 텍스트 편집 구현
- [ ] 표 제어 구현
- [ ] Run 액션 (684개) 구현
- [ ] 파라미터 헬퍼 (90개) 구현
- [ ] 테스트 및 검증
