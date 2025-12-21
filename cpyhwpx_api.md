# cpyhwpx API Reference

cpyhwpx는 한/글(HWP) 문서 자동화를 위한 C++ 기반 Python 라이브러리입니다.

## 목차

- [설치 및 사용법](#설치-및-사용법)
- [Hwp 클래스](#hwp-클래스)
  - [생성자](#생성자)
  - [초기화/종료](#초기화종료)
  - [파일 I/O](#파일-io)
  - [텍스트 편집](#텍스트-편집)
  - [위치 관리](#위치-관리)
  - [UI 제어](#ui-제어)
  - [문서 상태](#문서-상태)
  - [검색/치환](#검색치환)
  - [액션](#액션)
  - [속성](#속성)
- [유틸리티 함수](#유틸리티-함수)
- [열거형/타입](#열거형타입)

---

## 설치 및 사용법

```python
import cpyhwpx

# Hwp 인스턴스 생성 및 초기화
hwp = cpyhwpx.Hwp()
hwp.initialize()

# 문서 작업
hwp.insert_text("Hello, HWP!")
hwp.save_as("output.hwp")

# 종료
hwp.quit()
```

### 아키텍처 정보

```python
info = cpyhwpx.get_architecture_info()
# {'python_bits': 64, 'version': '1.1.0'}
```

---

## Hwp 클래스

### 생성자

```python
Hwp(visible=True, new_instance=True)
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `visible` | bool | True | HWP 창 표시 여부 |
| `new_instance` | bool | True | 새 인스턴스 생성 여부 |

---

### 초기화/종료

#### `initialize()`
HWP COM 객체를 초기화합니다.

```python
hwp.initialize()
```

**반환값:** `bool` - 성공 여부

---

#### `register_module(module_type, module_data)`
보안 모듈을 등록합니다.

```python
hwp.register_module("FilePathCheckDLL", "security.dll")
```

| 파라미터 | 타입 | 설명 |
|----------|------|------|
| `module_type` | str | 모듈 타입 |
| `module_data` | str | 모듈 데이터 |

**반환값:** `bool` - 성공 여부

---

#### `quit(save=False)`
HWP를 종료합니다.

```python
hwp.quit()          # 저장하지 않고 종료
hwp.quit(save=True) # 저장하고 종료
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `save` | bool | False | 종료 전 저장 여부 |

**반환값:** `bool` - 성공 여부

---

#### `is_initialized()`
초기화 완료 여부를 확인합니다.

```python
if hwp.is_initialized():
    print("HWP 초기화 완료")
```

**반환값:** `bool` - 초기화 완료 여부

---

### 파일 I/O

#### `open(filename, format_="", arg="")`
파일을 엽니다.

```python
hwp.open("document.hwp")
hwp.open("document.hwpx", format_="HWPX")
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `filename` | str | - | 파일 경로 |
| `format_` | str | "" | 파일 포맷 (HWP, HWPX, HTML 등) |
| `arg` | str | "" | 추가 인수 |

**반환값:** `bool` - 성공 여부

---

#### `save(save_if_dirty=True)`
현재 문서를 저장합니다.

```python
hwp.save()
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `save_if_dirty` | bool | True | 수정된 경우에만 저장 |

**반환값:** `bool` - 성공 여부

---

#### `save_as(filename, format_="HWP", arg="")`
다른 이름으로 저장합니다.

```python
hwp.save_as("output.hwp")
hwp.save_as("output.pdf", format_="PDF")
hwp.save_as("output.docx", format_="DOCX")
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `filename` | str | - | 저장할 파일 경로 |
| `format_` | str | "HWP" | 파일 포맷 |
| `arg` | str | "" | 추가 인수 |

**반환값:** `bool` - 성공 여부

---

#### `clear(option=1)`
문서를 초기화합니다.

```python
hwp.clear()
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `option` | int | 1 | 초기화 옵션 |

**반환값:** `bool` - 성공 여부

---

#### `close(is_dirty=False)`
현재 문서를 닫습니다.

```python
hwp.close()
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `is_dirty` | bool | False | 수정 여부 무시 |

**반환값:** `bool` - 성공 여부

---

### 텍스트 편집

#### `insert_text(text)`
현재 위치에 텍스트를 삽입합니다.

```python
hwp.insert_text("Hello, World!")
hwp.insert_text("줄바꿈\n포함 텍스트")
```

| 파라미터 | 타입 | 설명 |
|----------|------|------|
| `text` | str | 삽입할 텍스트 |

**반환값:** `bool` - 성공 여부

---

#### `get_text()`
현재 위치의 텍스트를 가져옵니다.

```python
text = hwp.get_text()
print(text)
```

**반환값:** `str` - 텍스트 내용

---

#### `get_selected_text(keep_select=False)`
선택된 텍스트를 가져옵니다.

```python
selected = hwp.get_selected_text()
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `keep_select` | bool | False | 선택 상태 유지 여부 |

**반환값:** `str` - 선택된 텍스트

---

### 위치 관리

#### `get_pos()`
현재 커서 위치를 가져옵니다.

```python
list_id, para, pos = hwp.get_pos()
```

**반환값:** `tuple` - (list_id, para, pos)

---

#### `set_pos(list_id, para, pos)`
커서 위치를 설정합니다.

```python
hwp.set_pos(list_id, para, pos)
```

| 파라미터 | 타입 | 설명 |
|----------|------|------|
| `list_id` | int | 리스트 ID |
| `para` | int | 문단 번호 |
| `pos` | int | 문단 내 위치 |

**반환값:** `bool` - 성공 여부

---

#### `move_pos(move_id, para=0, pos=0)`
커서를 이동합니다.

```python
hwp.move_pos(1)  # 문서 시작으로 이동
hwp.move_pos(2)  # 문서 끝으로 이동
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `move_id` | int | - | 이동 ID (MoveID 열거형 참조) |
| `para` | int | 0 | 문단 번호 |
| `pos` | int | 0 | 문단 내 위치 |

**반환값:** `bool` - 성공 여부

---

### UI 제어

#### `set_visible(visible)`
HWP 창의 표시 여부를 설정합니다.

```python
hwp.set_visible(True)   # 표시
hwp.set_visible(False)  # 숨김
```

| 파라미터 | 타입 | 설명 |
|----------|------|------|
| `visible` | bool | 표시 여부 |

**반환값:** `bool` - 성공 여부

---

#### `maximize_window()`
창을 최대화합니다.

```python
hwp.maximize_window()
```

**반환값:** `bool` - 성공 여부

---

#### `minimize_window()`
창을 최소화합니다.

```python
hwp.minimize_window()
```

**반환값:** `bool` - 성공 여부

---

### 문서 상태

#### `is_empty()`
문서가 비어있는지 확인합니다.

```python
if hwp.is_empty():
    print("빈 문서입니다")
```

**반환값:** `bool` - 빈 문서 여부

---

#### `is_modified()`
문서가 수정되었는지 확인합니다.

```python
if hwp.is_modified():
    print("저장되지 않은 변경사항이 있습니다")
```

**반환값:** `bool` - 수정 여부

---

#### `is_cell()`
현재 위치가 표 셀 안인지 확인합니다.

```python
if hwp.is_cell():
    print("표 셀 안에 있습니다")
```

**반환값:** `bool` - 셀 안 여부

---

### 검색/치환

#### `find(text, forward=True, match_case=False, regex=False, replace_mode=False)`
텍스트를 검색합니다.

```python
found = hwp.find("검색어")
found = hwp.find("패턴.*", regex=True)
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `text` | str | - | 검색할 텍스트 |
| `forward` | bool | True | 앞으로 검색 |
| `match_case` | bool | False | 대소문자 구분 |
| `regex` | bool | False | 정규식 사용 |
| `replace_mode` | bool | False | 치환 모드 |

**반환값:** `bool` - 검색 성공 여부

---

#### `replace(find_text, replace_text, forward=True, match_case=False, regex=False)`
텍스트를 치환합니다.

```python
hwp.replace("찾을텍스트", "바꿀텍스트")
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `find_text` | str | - | 찾을 텍스트 |
| `replace_text` | str | - | 바꿀 텍스트 |
| `forward` | bool | True | 앞으로 검색 |
| `match_case` | bool | False | 대소문자 구분 |
| `regex` | bool | False | 정규식 사용 |

**반환값:** `bool` - 치환 성공 여부

---

#### `replace_all(find_text, replace_text, match_case=False, regex=False)`
모든 텍스트를 치환합니다.

```python
count = hwp.replace_all("찾을텍스트", "바꿀텍스트")
print(f"{count}개 치환됨")
```

| 파라미터 | 타입 | 기본값 | 설명 |
|----------|------|--------|------|
| `find_text` | str | - | 찾을 텍스트 |
| `replace_text` | str | - | 바꿀 텍스트 |
| `match_case` | bool | False | 대소문자 구분 |
| `regex` | bool | False | 정규식 사용 |

**반환값:** `int` - 치환된 개수

---

### 액션

#### `run(action_name)`
HWP 액션을 실행합니다.

```python
hwp.run("Copy")      # 복사
hwp.run("Paste")     # 붙여넣기
hwp.run("SelectAll") # 전체 선택
```

| 파라미터 | 타입 | 설명 |
|----------|------|------|
| `action_name` | str | 액션 이름 |

**반환값:** `bool` - 실행 성공 여부

---

### 속성

#### `version` (읽기 전용)
HWP 버전을 반환합니다.

```python
print(hwp.version)  # "11.0.0.0"
```

**타입:** `str`

---

#### `build_number` (읽기 전용)
빌드 번호를 반환합니다.

```python
print(hwp.build_number)  # "1234"
```

**타입:** `str`

---

#### `current_page` (읽기 전용)
현재 페이지 번호를 반환합니다.

```python
print(hwp.current_page)  # 1
```

**타입:** `int`

---

#### `page_count` (읽기 전용)
전체 페이지 수를 반환합니다.

```python
print(hwp.page_count)  # 10
```

**타입:** `int`

---

#### `edit_mode` (읽기/쓰기)
편집 모드를 가져오거나 설정합니다.

```python
mode = hwp.edit_mode
hwp.edit_mode = 1
```

**타입:** `int`

---

## 유틸리티 함수

### `get_architecture_info()`
현재 아키텍처 정보를 반환합니다.

```python
import cpyhwpx
info = cpyhwpx.get_architecture_info()
# {'python_bits': 64, 'version': '1.1.0'}
```

**반환값:** `dict`
- `python_bits`: Python 비트 수 (64)
- `version`: 라이브러리 버전

---

## 열거형/타입


### ViewState
뷰 상태 열거형

### MoveID
이동 ID 열거형

### CtrlType
컨트롤 타입 열거형

### FileFormat
파일 포맷 열거형

### LineStyle
선 스타일 열거형

### HAlign
가로 정렬 열거형

### VAlign
세로 정렬 열거형

### HwpPos
위치 구조체

### CharShape
글자 모양 구조체

### ParaShape
문단 모양 구조체

### FontPreset
글꼴 프리셋

### FontDefs
글꼴 정의

---

## 요구사항

- Python 3.8+ (64-bit)
- 한/글 (HWP) 설치
- Windows 10/11

---

## 라이선스

MIT License
