# 표 관련 함수 매핑

## 요약
- 총 함수: 약 65개 (core.py ~25개 + Run 액션 ~40개)
- 포팅 완료: 0개 (0%)
- 우선순위: High (Phase 4)

---

## 표 생성/이동 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 1 | `create_table()` | (rows, cols, treat_as_char, ...) | bool | 표 생성 | `CreateTable()` | TODO |
| 2 | `get_into_nth_table()` | (n, select_cell) | Ctrl/bool | n번째 표로 이동 | `GetIntoNthTable()` | TODO |
| 3 | `get_into_table_caption()` | () | bool | 표 캡션으로 이동 | `GetIntoTableCaption()` | TODO |
| 4 | `table_from_data()` | (data, transpose, ...) | None | 데이터를 표로 변환 | `TableFromData()` | TODO |
| 5 | `is_cell()` | () | bool | 표 안 여부 확인 | `IsCell()` | TODO |

## 표 정보 조회 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 6 | `get_row_num()` | () | int | 표의 행 개수 | `GetRowNum()` | TODO |
| 7 | `get_col_num()` | () | int | 표의 열 개수 | `GetColNum()` | TODO |
| 8 | `get_row_height()` | (as_) | float | 행 높이 조회 | `GetRowHeight()` | TODO |
| 9 | `get_col_width()` | (as_) | float | 열 너비 조회 | `GetColWidth()` | TODO |
| 10 | `get_table_height()` | (as_) | float | 표 높이 조회 | `GetTableHeight()` | TODO |

## 표 크기 조절 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 11 | `set_row_height()` | (height, as_) | bool | 행 높이 설정 | `SetRowHeight()` | TODO |
| 12 | `set_col_width()` | (width, as_) | bool | 열 너비 설정 | `SetColWidth()` | TODO |
| 13 | `set_table_outside_margin_left()` | (val, as_) | bool | 표 왼쪽 바깥 여백 | `SetTableOutsideMarginLeft()` | TODO |
| 14 | `set_table_outside_margin_right()` | (val, as_) | bool | 표 오른쪽 바깥 여백 | `SetTableOutsideMarginRight()` | TODO |
| 15 | `set_table_outside_margin_top()` | (val, as_) | bool | 표 위쪽 바깥 여백 | `SetTableOutsideMarginTop()` | TODO |
| 16 | `set_table_outside_margin_bottom()` | (val, as_) | bool | 표 아래쪽 바깥 여백 | `SetTableOutsideMarginBottom()` | TODO |

## 셀 서식 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 17 | `cell_fill()` | (face_color) | bool | 셀 배경색 채우기 | `CellFill()` | TODO |

## 표 변환 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 18 | `table_to_string()` | (rowsep, colsep) | str | 표를 문자열로 | `TableToString()` | TODO |
| 19 | `table_to_csv()` | (n, filename, encoding, startrow) | None | 표를 CSV로 저장 | `TableToCsv()` | TODO |
| 20 | `table_to_df()` | (n, header, index) | DataFrame | 표를 DataFrame으로 | `TableToDf()` | TODO |
| 21 | `table_to_df_q()` | (n, header, index) | DataFrame | 빠른 표→DataFrame | `TableToDfQ()` | TODO |
| 22 | `table_to_bottom()` | (offset) | bool | 표를 페이지 하단으로 | `TableToBottom()` | TODO |

---

## 표 관련 Run 액션 (HAction.Run)

### 셀 이동

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 1 | `TableLeftCell()` | 왼쪽 셀로 이동 | `TableLeftCell()` | TODO |
| 2 | `TableRightCell()` | 오른쪽 셀로 이동 | `TableRightCell()` | TODO |
| 3 | `TableUpperCell()` | 위쪽 셀로 이동 | `TableUpperCell()` | TODO |
| 4 | `TableLowerCell()` | 아래쪽 셀로 이동 | `TableLowerCell()` | TODO |
| 5 | `TableRightCellAppend()` | 오른쪽 셀 (없으면 다음 행) | `TableRightCellAppend()` | TODO |
| 6 | `TableColBegin()` | 열 처음으로 이동 | `TableColBegin()` | TODO |
| 7 | `TableColEnd()` | 열 끝으로 이동 | `TableColEnd()` | TODO |
| 8 | `TableColPageUp()` | 열 맨 위로 이동 | `TableColPageUp()` | TODO |
| 9 | `TableColPageDown()` | 열 맨 아래로 이동 | `TableColPageDown()` | TODO |

### 행/열 추가/삭제

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 10 | `TableAppendRow()` | 행 추가 | `TableAppendRow()` | TODO |
| 11 | `TableSubtractRow()` | 행 삭제 | `TableSubtractRow()` | TODO |
| 12 | `TableDeleteCell()` | 셀 삭제 | `TableDeleteCell()` | TODO |

### 셀 선택

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 13 | `TableCellBlock()` | 현재 셀 블록 선택 | `TableCellBlock()` | TODO |
| 14 | `TableCellBlockRow()` | 행 전체 블록 선택 | `TableCellBlockRow()` | TODO |
| 15 | `TableCellBlockCol()` | 열 전체 블록 선택 | `TableCellBlockCol()` | TODO |
| 16 | `TableCellBlockExtend()` | 셀 블록 확장 | `TableCellBlockExtend()` | TODO |
| 17 | `TableCellBlockExtendAbs()` | 셀 블록 절대 확장 | `TableCellBlockExtendAbs()` | TODO |

### 셀 병합/분할

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 18 | `TableMergeCell()` | 셀 병합 | `TableMergeCell()` | TODO |
| 19 | `TableSplitCell()` | 셀 분할 | `TableSplitCell()` | TODO |
| 20 | `TableMergeTable()` | 표 병합 | `TableMergeTable()` | TODO |
| 21 | `TableSplitTable()` | 표 분할 | `TableSplitTable()` | TODO |

### 셀 정렬

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 22 | `TableCellAlignLeftTop()` | 왼쪽 위 정렬 | `TableCellAlignLeftTop()` | TODO |
| 23 | `TableCellAlignLeftCenter()` | 왼쪽 가운데 정렬 | `TableCellAlignLeftCenter()` | TODO |
| 24 | `TableCellAlignLeftBottom()` | 왼쪽 아래 정렬 | `TableCellAlignLeftBottom()` | TODO |
| 25 | `TableCellAlignCenterTop()` | 가운데 위 정렬 | `TableCellAlignCenterTop()` | TODO |
| 26 | `TableCellAlignCenterCenter()` | 가운데 가운데 정렬 | `TableCellAlignCenterCenter()` | TODO |
| 27 | `TableCellAlignCenterBottom()` | 가운데 아래 정렬 | `TableCellAlignCenterBottom()` | TODO |
| 28 | `TableCellAlignRightTop()` | 오른쪽 위 정렬 | `TableCellAlignRightTop()` | TODO |
| 29 | `TableCellAlignRightCenter()` | 오른쪽 가운데 정렬 | `TableCellAlignRightCenter()` | TODO |
| 30 | `TableCellAlignRightBottom()` | 오른쪽 아래 정렬 | `TableCellAlignRightBottom()` | TODO |
| 31 | `TableVAlignTop()` | 수직 위 정렬 | `TableVAlignTop()` | TODO |
| 32 | `TableVAlignCenter()` | 수직 가운데 정렬 | `TableVAlignCenter()` | TODO |
| 33 | `TableVAlignBottom()` | 수직 아래 정렬 | `TableVAlignBottom()` | TODO |

### 셀 테두리

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 34 | `TableCellBorderAll()` | 모든 테두리 | `TableCellBorderAll()` | TODO |
| 35 | `TableCellBorderNo()` | 테두리 없음 | `TableCellBorderNo()` | TODO |
| 36 | `TableCellBorderOutside()` | 바깥 테두리 | `TableCellBorderOutside()` | TODO |
| 37 | `TableCellBorderInside()` | 안쪽 테두리 | `TableCellBorderInside()` | TODO |
| 38 | `TableCellBorderInsideHorz()` | 가로 안쪽 테두리 | `TableCellBorderInsideHorz()` | TODO |
| 39 | `TableCellBorderInsideVert()` | 세로 안쪽 테두리 | `TableCellBorderInsideVert()` | TODO |
| 40 | `TableCellBorderTop()` | 위쪽 테두리 | `TableCellBorderTop()` | TODO |
| 41 | `TableCellBorderBottom()` | 아래쪽 테두리 | `TableCellBorderBottom()` | TODO |
| 42 | `TableCellBorderLeft()` | 왼쪽 테두리 | `TableCellBorderLeft()` | TODO |
| 43 | `TableCellBorderRight()` | 오른쪽 테두리 | `TableCellBorderRight()` | TODO |
| 44 | `TableCellBorderDiagonalDown()` | 대각선 (↘) | `TableCellBorderDiagonalDown()` | TODO |
| 45 | `TableCellBorderDiagonalUp()` | 대각선 (↗) | `TableCellBorderDiagonalUp()` | TODO |

### 표 크기 조절

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 46 | `TableResizeUp()` | 표 위쪽으로 크기 조절 | `TableResizeUp()` | TODO |
| 47 | `TableResizeDown()` | 표 아래쪽으로 크기 조절 | `TableResizeDown()` | TODO |
| 48 | `TableResizeLeft()` | 표 왼쪽으로 크기 조절 | `TableResizeLeft()` | TODO |
| 49 | `TableResizeRight()` | 표 오른쪽으로 크기 조절 | `TableResizeRight()` | TODO |
| 50 | `TableResizeCellUp()` | 셀 위쪽으로 크기 조절 | `TableResizeCellUp()` | TODO |
| 51 | `TableResizeCellDown()` | 셀 아래쪽으로 크기 조절 | `TableResizeCellDown()` | TODO |
| 52 | `TableResizeCellLeft()` | 셀 왼쪽으로 크기 조절 | `TableResizeCellLeft()` | TODO |
| 53 | `TableResizeCellRight()` | 셀 오른쪽으로 크기 조절 | `TableResizeCellRight()` | TODO |
| 54 | `TableDistributeCellHeight()` | 행 높이 균등 분배 | `TableDistributeCellHeight()` | TODO |
| 55 | `TableDistributeCellWidth()` | 열 너비 균등 분배 | `TableDistributeCellWidth()` | TODO |

### 표 자동 기능

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 56 | `TableAutoFill()` | 자동 채우기 | `TableAutoFill()` | TODO |
| 57 | `TableAutoFillDlg()` | 자동 채우기 대화상자 | `TableAutoFillDlg()` | TODO |
| 58 | `TableFormulaSumAuto()` | 자동 합계 | `TableFormulaSumAuto()` | TODO |
| 59 | `TableFormulaSumHor()` | 가로 합계 | `TableFormulaSumHor()` | TODO |
| 60 | `TableFormulaSumVer()` | 세로 합계 | `TableFormulaSumVer()` | TODO |
| 61 | `TableFormulaAvgAuto()` | 자동 평균 | `TableFormulaAvgAuto()` | TODO |
| 62 | `TableFormulaAvgHor()` | 가로 평균 | `TableFormulaAvgHor()` | TODO |
| 63 | `TableFormulaAvgVer()` | 세로 평균 | `TableFormulaAvgVer()` | TODO |
| 64 | `TableFormulaProAuto()` | 자동 곱 | `TableFormulaProAuto()` | TODO |
| 65 | `TableFormulaProHor()` | 가로 곱 | `TableFormulaProHor()` | TODO |
| 66 | `TableFormulaProVer()` | 세로 곱 | `TableFormulaProVer()` | TODO |

---

## 상세 설명

### create_table()
표를 생성한다.

```python
def create_table(self,
    rows: int = 1,
    cols: int = 1,
    treat_as_char: bool = True,
    width_type: int = 0,
    height_type: int = 0,
    header: bool = True,
    height: int = 0
) -> bool:
    """
    Args:
        rows: 행 수
        cols: 열 수
        treat_as_char: 글자처럼 취급 여부
        width_type: 너비 지정 방식
            - 0: 단에 맞춤
            - 1: 문단에 맞춤
            - 2: 임의값
        height_type: 높이 지정 방식
            - 0: 자동
            - 1: 임의값
        header: 1행을 제목행으로 설정
        height: 표 높이 (height_type=1일 때)

    Returns:
        성공시 True
    """
```

**C++ 구현:**
```cpp
bool Hwp::CreateTable(int rows, int cols, bool treatAsChar,
                      int widthType, int heightType,
                      bool header, int height) {
    // HParameterSet.HTableCreation 획득
    IDispatch* pParamSet = GetHParameterSet(L"HTableCreation");

    // GetDefault 호출
    InvokeMethod(m_pHAction, L"GetDefault",
                 L"TableCreate", GetHSet(pParamSet));

    // 속성 설정
    SetProperty(pParamSet, L"Rows", rows);
    SetProperty(pParamSet, L"Cols", cols);
    SetProperty(pParamSet, L"WidthType", widthType);
    SetProperty(pParamSet, L"HeightType", heightType);

    // 용지 여백 계산하여 표 너비 설정
    IDispatch* pSecDef = GetHParameterSet(L"HSecDef");
    InvokeMethod(m_pHAction, L"GetDefault", L"PageSetup", GetHSet(pSecDef));

    IDispatch* pPageDef = GetProperty(pSecDef, L"PageDef");
    int paperWidth = GetIntProperty(pPageDef, L"PaperWidth");
    int leftMargin = GetIntProperty(pPageDef, L"LeftMargin");
    int rightMargin = GetIntProperty(pPageDef, L"RightMargin");
    int gutterLen = GetIntProperty(pPageDef, L"GutterLen");

    int totalWidth = paperWidth - leftMargin - rightMargin - gutterLen - MiliToHwpUnit(2);

    // ColWidth 배열 생성
    InvokeMethod(pParamSet, L"CreateItemArray", L"ColWidth", cols);
    int eachColWidth = (totalWidth - MiliToHwpUnit(3.6 * cols)) / cols;
    // ... 각 열 너비 설정

    // TableProperties 설정
    IDispatch* pTableProps = GetProperty(pParamSet, L"TableProperties");
    SetProperty(pTableProps, L"TreatAsChar", treatAsChar);
    SetProperty(pTableProps, L"Width", totalWidth);

    // Execute
    return InvokeMethod(m_pHAction, L"Execute",
                        L"TableCreate", GetHSet(pParamSet));
}
```

### get_into_nth_table()
n번째 표로 이동한다.

```python
def get_into_nth_table(self, n=0, select_cell=False):
    """
    Args:
        n: 표 인덱스 (0부터 시작, 음수 가능)
        select_cell: 셀 선택 상태로 이동할지 여부

    Returns:
        성공시 Ctrl 객체, 실패시 False
    """
```

### is_cell()
캐럿이 표 안에 있는지 확인한다.

```python
def is_cell(self) -> bool:
    """
    Returns:
        표 안에 있으면 True, 아니면 False
    """
    return self.key_indicator()[-1].startswith("(")
```

### get_row_height() / set_row_height()
행 높이 조회/설정

```python
def get_row_height(self,
    as_: Literal["mm", "hwpunit", "point", "inch"] = "mm"
) -> Union[float, int]:
    """
    Args:
        as_: 반환 단위 (mm, hwpunit, point, inch)

    Returns:
        행 높이
    """

def set_row_height(self,
    height: Union[int, float],
    as_: Literal["mm", "hwpunit"] = "mm"
) -> bool:
    """
    Args:
        height: 행 높이 값
        as_: 단위 (mm 또는 hwpunit)

    Returns:
        성공시 True
    """
```

### set_col_width()
열 너비를 설정한다.

```python
def set_col_width(self,
    width: Union[int, float, list, tuple],
    as_: Literal["mm", "ratio"] = "ratio"
) -> bool:
    """
    Args:
        width: 열 너비
            - int/float: 현재 열의 너비 (mm 단위)
            - list/tuple: 모든 열의 비율 (예: [1, 2, 3])
        as_: 단위
            - "mm": mm 단위
            - "ratio": 비율 (width가 리스트일 때)

    Returns:
        성공시 True

    Examples:
        # 열 너비를 1:2:3 비율로 설정
        hwp.set_col_width([1, 2, 3])
    """
```

### cell_fill()
셀 배경색을 채운다.

```python
def cell_fill(self, face_color: Tuple[int, int, int] = (217, 217, 217)):
    """
    Args:
        face_color: RGB 색상 튜플 (0~255)

    Returns:
        성공시 True
    """
    pset = self.hwp.HParameterSet.HCellBorderFill
    self.hwp.HAction.GetDefault("CellFill", pset.HSet)
    pset.FillAttr.type = self.hwp.BrushType("NullBrush|WinBrush")
    pset.FillAttr.WinBrushFaceColor = self.hwp.RGBColor(*face_color)
    return self.hwp.HAction.Execute("CellFill", pset.HSet)
```

### table_from_data()
데이터를 표로 변환한다.

```python
def table_from_data(self,
    data: Union[pd.DataFrame, dict, list, str],
    transpose: bool = False,
    header0: str = "",
    treat_as_char: bool = False,
    header: bool = True,
    index: bool = True,
    cell_fill: Union[bool, Tuple[int, int, int]] = False,
    header_bold: bool = True
) -> None:
    """
    Args:
        data: 표로 변환할 데이터
            - DataFrame, dict, list
            - 파일 경로 (csv, xlsx, json)
        transpose: 행/열 전환
        header0: 인덱스 포함시 (1,1) 셀 텍스트
        treat_as_char: 글자처럼 취급
        header: 제목행 설정
        index: 인덱스 포함
        cell_fill: 헤더 배경색 (True면 기본 회색)
        header_bold: 헤더 굵게
    """
```

### TableSplitCell()
셀을 분할한다. (Run 액션)

```python
def TableSplitCell(self,
    div_horizontal: int = 2,
    div_vertical: int = 1,
    split_type: int = 1
) -> bool:
    """
    Args:
        div_horizontal: 가로 분할 수
        div_vertical: 세로 분할 수
        split_type:
            - 0: 병합 안함
            - 1: 셀로 나누기 (기본값)
            - 2: 열로 나누기
            - 3: 행으로 나누기
    """
    pset = self.hwp.HParameterSet.HTableSplit
    self.hwp.HAction.GetDefault("TableSplitCell", pset.HSet)
    pset.DivHor = div_horizontal
    pset.DivVer = div_vertical
    pset.SplitType = split_type
    return self.hwp.HAction.Execute("TableSplitCell", pset.HSet)
```

---

## C++ 클래스 구조

```cpp
// Hwp.h - 표 관련 메서드
class Hwp : public ParamHelpers, public RunMethods {
public:
    // 표 생성/이동
    bool CreateTable(int rows = 1, int cols = 1,
                     bool treatAsChar = true,
                     int widthType = 0, int heightType = 0,
                     bool header = true, int height = 0);
    Ctrl* GetIntoNthTable(int n = 0, bool selectCell = false);
    bool GetIntoTableCaption();
    void TableFromData(const std::wstring& data,
                       bool transpose = false,
                       bool treatAsChar = false);
    bool IsCell();

    // 표 정보 조회
    int GetRowNum();
    int GetColNum();
    double GetRowHeight(const std::wstring& unit = L"mm");
    double GetColWidth(const std::wstring& unit = L"mm");
    double GetTableHeight(const std::wstring& unit = L"mm");

    // 표 크기 조절
    bool SetRowHeight(double height, const std::wstring& unit = L"mm");
    bool SetColWidth(double width, const std::wstring& unit = L"mm");
    bool SetColWidthRatio(const std::vector<double>& ratio);
    bool SetTableOutsideMargin(double left, double right,
                                double top, double bottom,
                                const std::wstring& unit = L"mm");

    // 셀 서식
    bool CellFill(int r = 217, int g = 217, int b = 217);

    // Run 액션 - 셀 이동
    bool TableLeftCell();
    bool TableRightCell();
    bool TableUpperCell();
    bool TableLowerCell();
    bool TableRightCellAppend();
    bool TableColBegin();
    bool TableColEnd();
    bool TableColPageUp();
    bool TableColPageDown();

    // Run 액션 - 행/열 추가/삭제
    bool TableAppendRow();
    bool TableSubtractRow();
    bool TableDeleteCell(bool remainCell = false);

    // Run 액션 - 셀 선택
    bool TableCellBlock();
    bool TableCellBlockRow();
    bool TableCellBlockCol();
    bool TableCellBlockExtend();
    bool TableCellBlockExtendAbs();

    // Run 액션 - 셀 병합/분할
    bool TableMergeCell();
    bool TableSplitCell(int divHor = 2, int divVer = 1, int splitType = 1);
    bool TableMergeTable();
    bool TableSplitTable();

    // Run 액션 - 셀 정렬
    bool TableCellAlignLeftTop();
    bool TableCellAlignCenterCenter();
    // ... 기타 정렬

    // Run 액션 - 셀 테두리
    bool TableCellBorderAll();
    bool TableCellBorderNo();
    bool TableCellBorderOutside();
    // ... 기타 테두리

    // Run 액션 - 표 크기 조절
    bool TableDistributeCellHeight();
    bool TableDistributeCellWidth();

    // Run 액션 - 표 자동 기능
    bool TableAutoFill();
    bool TableFormulaSumAuto();
    bool TableFormulaSumHor();
    bool TableFormulaSumVer();
    bool TableFormulaAvgAuto();
};
```

---

## 구현 우선순위

1. **Critical**: CreateTable(), IsCell() - 표 생성/확인 필수
2. **High**: GetIntoNthTable() - 표 이동
3. **High**: TableCellBlock(), TableMergeCell(), TableSplitCell() - 셀 조작
4. **High**: TableRightCell(), TableLeftCell() 등 - 셀 이동
5. **Medium**: SetRowHeight(), SetColWidth() - 크기 조절
6. **Medium**: CellFill(), TableCellBorder* - 셀 서식
7. **Low**: TableFromData(), table_to_df() 등 - 데이터 변환 (pandas 의존)
