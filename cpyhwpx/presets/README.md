# cpyhwpx 프리셋 시스템

Claude Code 및 LLM이 HWP 자동화 작업을 수행할 때 참조할 수 있는 프리셋 정의입니다.

## 파일 구조

```
cpyhwpx/presets/
├── _schema.yaml      # 프리셋 스키마 정의
├── _common.yaml      # 공통 에러 처리, 상태 체크
├── table.yaml        # 표 작업 프리셋 (25개)
├── text.yaml         # 텍스트 편집 프리셋 (30개)
├── field.yaml        # 필드/메타태그 프리셋 (15개)
├── shape.yaml        # 개체 작업 프리셋 (25개)
├── document.yaml     # 문서 구조 프리셋 (20개)
├── navigation.yaml   # 탐색/이동 프리셋 (20개)
└── README.md         # 이 파일
```

## 프리셋 사용법

### 1. 프리셋 구조 이해

각 프리셋은 다음 구조를 따릅니다:

```yaml
id: table_goto_nth              # 고유 식별자
name: "n번째 표로 이동"          # 한글 이름
description: "설명..."          # 상세 설명
category: table                 # 카테고리

preconditions:                  # 사전 조건 (커서/선택 상태)
  cursor_in_table: true
  check_method: "IsCell"

actions:                        # 실행할 액션 시퀀스
  - type: method
    action: "GetIntoNthTable"
    params:
      n: "{{index}}"
    description: "n번째 표로 이동"

postconditions:                 # 완료 후 상태
  cursor_position: "표의 첫 번째 셀"

examples:                       # Python 코드 예시
  - scenario: "첫 번째 표로 이동"
    code: |
      hwp.GetIntoNthTable(0)

error_handling:                 # 에러 처리 가이드
  common_errors:
    - error: "표를 찾을 수 없음"
      solution: "GetCtrlList()로 확인"
```

### 2. 사전 조건 (preconditions)

작업 전 확인해야 할 커서/선택 상태입니다:

| 조건 | 설명 | 확인 메서드 |
|------|------|------------|
| `cursor_in_table` | 커서가 표 안에 있어야 함 | `IsCell()` |
| `cursor_in_cell` | 커서가 셀 안에 있어야 함 | `IsCell()` |
| `has_selection` | 선택 영역이 있어야 함 | `GetSelectedText()` |
| `control_selected` | 특정 컨트롤 선택 필요 | `GetCurSelectedCtrl()` |

### 3. 액션 타입

| 타입 | 용도 | 예시 |
|------|------|------|
| `run` | HAction.Run() 액션 실행 | `TableLeftCell`, `CharShapeBold` |
| `method` | HwpWrapper 메서드 호출 | `InsertText`, `SetFont` |
| `conditional` | 조건부 실행 | if/else 분기 |
| `loop` | 반복 실행 | 데이터 순회 |

### 4. 변수 문법

```yaml
params:
  text: "{{content}}"              # 필수 변수
  font_name: "{{font|맑은 고딕}}"   # 기본값 있는 변수
  size: "{{size|1000}}"            # 숫자 기본값
```

---

## 프리셋 수정 가이드

### 1. 새 프리셋 추가하기

#### 1.1 기존 카테고리에 추가

해당 카테고리 YAML 파일 끝에 추가:

```yaml
# table.yaml 끝에 추가
- id: table_my_custom_preset
  name: "새로운 표 작업"
  description: "커스텀 표 작업 설명"
  category: table

  preconditions:
    cursor_in_table: true

  actions:
    - type: run
      action: "TableCellBlock"
      description: "셀 블록 모드"
    - type: method
      action: "CellFill"
      params:
        r: 255
        g: 255
        b: 0
      description: "노란색 배경"

  examples:
    - scenario: "기본 사용"
      code: |
        hwp.run("TableCellBlock")
        hwp.CellFill(255, 255, 0)
```

#### 1.2 새 카테고리 추가

1. `cpyhwpx/presets/`에 새 YAML 파일 생성 (예: `custom.yaml`)
2. `_schema.yaml`의 categories에 새 카테고리 추가:

```yaml
categories:
  custom:
    description: "커스텀 프리셋"
    file: "custom.yaml"
```

3. 새 파일에 프리셋 정의 작성

### 2. 액션 추가/수정

#### 2.1 액션 추가

```yaml
actions:
  # 기존 액션들
  - type: run
    action: "TableLeftCell"

  # 새 액션 추가 (원하는 위치에)
  - type: method
    action: "CellFill"
    params:
      r: 200
      g: 200
      b: 200
    description: "셀 배경색 설정"
```

#### 2.2 조건부 액션

```yaml
actions:
  - type: conditional
    condition: "IsCell()"
    then:
      - type: run
        action: "TableLeftCell"
    else:
      - type: method
        action: "GetIntoNthTable"
        params:
          n: 0
```

#### 2.3 반복 액션

```yaml
actions:
  - type: loop
    count: "{{repeat_count}}"
    actions:
      - type: run
        action: "TableRightCell"
```

### 3. 사전 조건 수정

```yaml
preconditions:
  # 단순 조건
  cursor_in_table: true
  has_selection: true

  # 특정 컨트롤 타입 필요
  control_selected: "gso"    # tbl, gso, pic, eqed 등

  # 커스텀 체크 메서드
  check_method: "IsCell"
```

### 4. 파라미터 변수화

```yaml
actions:
  - type: method
    action: "SetFont"
    params:
      face_name: "{{font_name}}"        # 필수 변수
      height: "{{font_size|1000}}"      # 기본값 1000
      bold: "{{is_bold|0}}"             # 기본값 0
```

---

## 사용 가능한 참조

### Run 액션 목록

`src/HwpAction.h` 또는 `docs/mapping/09_run_actions.md` 참조:

**표 작업:**
- `TableLeftCell`, `TableRightCell`, `TableUpperCell`, `TableLowerCell`
- `TableCellBlock`, `TableCellBlockRow`, `TableCellBlockCol`
- `TableMergeCell`, `TableSplitCell`, `TableAppendRow`
- `TableFormulaSumAuto`, `TableFormulaSumHor`, `TableFormulaSumVer`

**커서 이동:**
- `MoveDocBegin`, `MoveDocEnd`, `MoveParaBegin`, `MoveParaEnd`
- `MoveLineBegin`, `MoveLineEnd`, `MoveNextWord`, `MovePrevWord`
- `MoveSelDocEnd`, `MoveSelLineEnd` (선택하며 이동)

**글자 서식:**
- `CharShapeBold`, `CharShapeItalic`, `CharShapeUnderline`
- `CharShapeStrikeout`, `CharShapeSuperscript`, `CharShapeSubscript`

**문단 서식:**
- `ParagraphShapeAlignLeft`, `ParagraphShapeAlignCenter`
- `ParagraphShapeAlignRight`, `ParagraphShapeAlignJustify`

**편집:**
- `Delete`, `DeleteBack`, `DeleteWord`, `DeleteLine`
- `Copy`, `Cut`, `Paste`, `Undo`, `Redo`

**나누기:**
- `BreakPara`, `BreakLine`, `BreakPage`, `BreakSection`

### Method 목록

`src/HwpWrapper.h` 또는 `docs/mapping/` 문서 참조:

**파일 I/O:**
- `Open(filename)`, `Save()`, `SaveAs(filename, format)`
- `Close()`, `Clear()`, `InsertFile(filename)`

**텍스트:**
- `InsertText(text)`, `GetText()`, `GetSelectedText()`
- `Find(text)`, `FindForward(text)`, `ReplaceAll(find, replace)`

**표:**
- `CreateTable(rows, cols)`, `GetIntoNthTable(n)`
- `GetTableRowCount()`, `GetTableColCount()`, `CellFill(r, g, b)`

**필드:**
- `CreateField(name)`, `GetFieldList()`, `GetFieldText(field)`
- `PutFieldText(field, text)`, `MoveToField(field)`

**위치:**
- `GetPos()`, `SetPos(list, para, pos)`, `MovePos(move_id)`

**개체:**
- `InsertPicture(path)`, `GetCtrlList()`, `SelectCtrl(ctrl)`

---

## 검증 체크리스트

새 프리셋 추가/수정 시 확인:

- [ ] `id`가 고유하고 영문_소문자_언더스코어 형식인가?
- [ ] `category`가 올바른 값인가? (table/text/field/shape/document/navigation)
- [ ] `actions`의 각 항목에 `type`과 `action`이 있는가?
- [ ] `run` 타입 액션이 `HwpAction.h`에 존재하는가?
- [ ] `method` 타입 액션이 `HwpWrapper.h`에 존재하는가?
- [ ] 변수(`{{...}}`)가 examples에서 올바르게 사용되는가?
- [ ] 사전 조건이 적절히 정의되어 있는가?
- [ ] 예시 코드가 실행 가능한가?

---

## 테스트 방법

프리셋이 제대로 동작하는지 확인:

```python
import cpyhwpx

# HWP 초기화
hwp = cpyhwpx.Hwp(visible=True)
hwp.initialize()

# 테스트 문서 열기
hwp.Open("test.hwp")

# 프리셋의 사전 조건 확인
if hwp.IsCell():
    print("표 안에 있음 - 표 작업 가능")

# 프리셋의 액션 순차 실행
hwp.GetIntoNthTable(0)      # action 1
hwp.run("TableCellBlock")   # action 2
hwp.CellFill(217, 217, 217) # action 3

# 사후 조건 확인
print(f"현재 표 안: {hwp.IsCell()}")

# 종료
hwp.Close()
hwp.quit()
```

---

## 에러 처리

`_common.yaml`의 `error_catalog` 참조:

| 코드 | 이름 | 해결 방법 |
|------|------|----------|
| E001 | DocumentNotOpen | `hwp.Open()` 또는 `hwp.Clear()` 실행 |
| E002 | NotInTable | `hwp.GetIntoNthTable()` 실행 |
| E003 | FieldNotFound | `GetFieldList()`로 필드 목록 확인 |
| E004 | ControlNotSelected | `hwp.SelectCtrl()` 실행 |
| E005 | InvalidIndex | 유효한 인덱스 범위 확인 |
| E006 | NoSelection | 텍스트/개체 먼저 선택 |
