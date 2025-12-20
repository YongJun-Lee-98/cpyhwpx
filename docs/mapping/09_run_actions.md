# Run 액션 함수 매핑

## 개요

| 항목 | 내용 |
|------|------|
| 소스 파일 | `pyhwpx/run_methods.py` |
| 총 메서드 수 | **684개** |
| 클래스 | `RunMethods` |
| 포팅 완료 | 0개 (0%) |
| 우선순위 | Medium |

## 설명

`RunMethods` 클래스는 HWP의 `HAction.Run()` 메서드를 래핑한 액션들의 모음입니다.
대부분의 메서드는 단순히 `self.hwp.HAction.Run("ActionName")`을 호출하는 형태입니다.

### 기본 패턴
```python
def ActionName(self) -> bool:
    """액션 설명"""
    return self.hwp.HAction.Run("ActionName")
```

### C++ 구현 패턴
```cpp
bool HwpWrapper::ActionName() {
    return InvokeMethod(L"HAction.Run", L"ActionName");
}
```

---

## 카테고리별 액션 목록

### 1. 자동 기능 (Auto*) - 22개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `ASendBrowserText()` | `ASendBrowserText` | 웹브라우저로 보내기 |
| 2 | `AutoChangeHangul()` | `AutoChangeHangul` | 낱자모 우선입력 토글 (구버전) |
| 3 | `AutoChangeRun()` | `AutoChangeRun` | 글자판 자동 변경 토글 |
| 4 | `AutoSpellRun()` | `AutoSpell Run` | 맞춤법 도우미 토글 |
| 5-21 | `AutoSpellSelect0()` ~ `AutoSpellSelect16()` | `AutoSpellSelect0` ~ `AutoSpellSelect16` | 맞춤법 도우미 어휘 선택 (0~16) |

---

### 2. 줄바꿈/나누기 (Break*) - 6개

| # | Python 메서드 | HAction 문자열 | 설명 | 단축키 |
|---|--------------|----------------|------|--------|
| 1 | `BreakColDef()` | `BreakColDef` | 단 정의 삽입 | Ctrl+Alt+Enter |
| 2 | `BreakColumn()` | `BreakColumn` | 단 나누기 (배분다단) | Ctrl+Shift+Enter |
| 3 | `BreakLine()` | `BreakLine` | 줄 나누기 | Shift+Enter |
| 4 | `BreakPage()` | `BreakPage` | 쪽 나누기 | Ctrl+Enter |
| 5 | `BreakPara()` | `BreakPara` | 문단 나누기 (엔터) | Enter |
| 6 | `BreakSection()` | `BreakSection` | 구역 나누기 | Shift+Alt+Enter |

---

### 3. 글자 모양 (CharShape*) - 32개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `CharShapeBold()` | `CharShapeBold` | 진하게 토글 |
| 2 | `CharShapeCenterline()` | `CharShapeCenterline` | 취소선 토글 |
| 3 | `CharShapeEmboss()` | `CharShapeEmboss` | 양각 토글 |
| 4 | `CharShapeEngrave()` | `CharShapeEngrave` | 음각 토글 |
| 5 | `CharShapeHeight()` | `CharShapeHeight` | 글자 크기 대화상자 |
| 6 | `CharShapeHeightDecrease()` | `CharShapeHeightDecrease` | 글자 크기 1pt 감소 |
| 7 | `CharShapeHeightIncrease()` | `CharShapeHeightIncrease` | 글자 크기 1pt 증가 |
| 8 | `CharShapeItalic()` | `CharShapeItalic` | 이탤릭 토글 |
| 9 | `CharShapeLang()` | `CharShapeLang` | 글자 언어 |
| 10 | `CharShapeNextFaceName()` | `CharShapeNextFaceName` | 다음 글꼴 |
| 11 | `CharShapeNormal()` | `CharShapeNormal` | 글자 모양 초기화 |
| 12 | `CharShapeOutline()` | `CharShapeOutline` | 외곽선 토글 |
| 13 | `CharShapePrevFaceName()` | `CharShapePrevFaceName` | 이전 글꼴 |
| 14 | `CharShapeShadow()` | `CharShapeShadow` | 그림자 토글 |
| 15 | `CharShapeSpacing()` | `CharShapeSpacing` | 자간 대화상자 |
| 16 | `CharShapeSpacingDecrease()` | `CharShapeSpacingDecrease` | 자간 1% 감소 |
| 17 | `CharShapeSpacingIncrease()` | `CharShapeSpacingIncrease` | 자간 1% 증가 |
| 18 | `CharShapeSubscript()` | `CharShapeSubscript` | 아래첨자 토글 |
| 19 | `CharShapeSuperscript()` | `CharShapeSuperscript` | 위첨자 토글 |
| 20 | `CharShapeSuperSubscript()` | `CharShapeSuperSubscript` | 첨자 순환 토글 |
| 21 | `CharShapeTextColorBlack()` | `CharShapeTextColorBlack` | 글자색 검정 |
| 22 | `CharShapeTextColorBlue()` | `CharShapeTextColorBlue` | 글자색 파랑 |
| 23 | `CharShapeTextColorBluish()` | `CharShapeTextColorBluish` | 글자색 청록 |
| 24 | `CharShapeTextColorGreen()` | `CharShapeTextColorGreen` | 글자색 초록 |
| 25 | `CharShapeTextColorRed()` | `CharShapeTextColorRed` | 글자색 빨강 |
| 26 | `CharShapeTextColorViolet()` | `CharShapeTextColorViolet` | 글자색 보라 |
| 27 | `CharShapeTextColorWhite()` | `CharShapeTextColorWhite` | 글자색 흰색 |
| 28 | `CharShapeTextColorYellow()` | `CharShapeTextColorYellow` | 글자색 노랑 |
| 29 | `CharShapeTypeface()` | `CharShapeTypeface` | 글꼴 이름 대화상자 |
| 30 | `CharShapeUnderline()` | `CharShapeUnderline` | 밑줄 토글 |
| 31 | `CharShapeWidth()` | `CharShapeWidth` | 장평 대화상자 |
| 32 | `CharShapeWidthDecrease()` | `CharShapeWidthDecrease` | 장평 1% 감소 |
| 33 | `CharShapeWidthIncrease()` | `CharShapeWidthIncrease` | 장평 1% 증가 |

---

### 4. 문단 모양 (ParagraphShape*) - 18개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `ParagraphShapeAlignCenter()` | `ParagraphShapeAlignCenter` | 가운데 정렬 |
| 2 | `ParagraphShapeAlignDistribute()` | `ParagraphShapeAlignDistribute` | 배분 정렬 |
| 3 | `ParagraphShapeAlignDivision()` | `ParagraphShapeAlignDivision` | 나눔 정렬 |
| 4 | `ParagraphShapeAlignJustify()` | `ParagraphShapeAlignJustify` | 양쪽 정렬 |
| 5 | `ParagraphShapeAlignLeft()` | `ParagraphShapeAlignLeft` | 왼쪽 정렬 |
| 6 | `ParagraphShapeAlignRight()` | `ParagraphShapeAlignRight` | 오른쪽 정렬 |
| 7 | `ParagraphShapeDecreaseLeftMargin()` | `ParagraphShapeDecreaseLeftMargin` | 왼쪽 여백 감소 |
| 8 | `ParagraphShapeDecreaseLineSpacing()` | `ParagraphShapeDecreaseLineSpacing` | 줄 간격 감소 |
| 9 | `ParagraphShapeDecreaseMargin()` | `ParagraphShapeDecreaseMargin` | 여백 감소 |
| 10 | `ParagraphShapeDecreaseRightMargin()` | `ParagraphShapeDecreaseRightMargin` | 오른쪽 여백 감소 |
| 11 | `ParagraphShapeIncreaseLeftMargin()` | `ParagraphShapeIncreaseLeftMargin` | 왼쪽 여백 증가 |
| 12 | `ParagraphShapeIncreaseLineSpacing()` | `ParagraphShapeIncreaseLineSpacing` | 줄 간격 증가 |
| 13 | `ParagraphShapeIncreaseMargin()` | `ParagraphShapeIncreaseMargin` | 여백 증가 |
| 14 | `ParagraphShapeIncreaseRightMargin()` | `ParagraphShapeIncreaseRightMargin` | 오른쪽 여백 증가 |
| 15 | `ParagraphShapeIndentAtCaret()` | `ParagraphShapeIndentAtCaret` | 캐럿 위치에서 들여쓰기 |
| 16 | `ParagraphShapeIndentNegative()` | `ParagraphShapeIndentNegative` | 내어쓰기 |
| 17 | `ParagraphShapeIndentPositive()` | `ParagraphShapeIndentPositive` | 들여쓰기 |
| 18 | `ParagraphShapeProtect()` | `ParagraphShapeProtect` | 문단 보호 |
| 19 | `ParagraphShapeSingleRow()` | `ParagraphShapeSingleRow` | 한 줄로 입력 |
| 20 | `ParagraphShapeWithNext()` | `ParagraphShapeWithNext` | 다음 문단과 함께 |

---

### 5. 이동 (Move*) - 85개

#### 5.1 기본 이동

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `MoveDown()` | `MoveDown` | 아래로 이동 |
| 2 | `MoveLeft()` | `MoveLeft` | 왼쪽으로 이동 |
| 3 | `MoveRight()` | `MoveRight` | 오른쪽으로 이동 |
| 4 | `MoveUp()` | `MoveUp` | 위로 이동 |
| 5 | `MoveLineBegin()` | `MoveLineBegin` | 줄 시작으로 |
| 6 | `MoveLineEnd()` | `MoveLineEnd` | 줄 끝으로 |
| 7 | `MoveLineDown()` | `MoveLineDown` | 다음 줄로 |
| 8 | `MoveLineUp()` | `MoveLineUp` | 이전 줄로 |

#### 5.2 문단/단어 이동

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `MoveParaBegin()` | `MoveParaBegin` | 문단 시작으로 |
| 2 | `MoveParaEnd()` | `MoveParaEnd` | 문단 끝으로 |
| 3 | `MoveWordBegin()` | `MoveWordBegin` | 단어 시작으로 |
| 4 | `MoveWordEnd()` | `MoveWordEnd` | 단어 끝으로 |
| 5 | `MoveNextWord()` | `MoveNextWord` | 다음 단어로 |
| 6 | `MovePrevWord()` | `MovePrevWord` | 이전 단어로 |
| 7 | `MoveNextChar()` | `MoveNextChar` | 다음 문자로 |
| 8 | `MovePrevChar()` | `MovePrevChar` | 이전 문자로 |

#### 5.3 페이지/문서 이동

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `MoveDocBegin()` | `MoveDocBegin` | 문서 시작으로 |
| 2 | `MoveDocEnd()` | `MoveDocEnd` | 문서 끝으로 |
| 3 | `MovePageBegin()` | `MovePageBegin` | 페이지 시작으로 |
| 4 | `MovePageEnd()` | `MovePageEnd` | 페이지 끝으로 |
| 5 | `MovePageDown()` | `MovePageDown` | 다음 페이지로 |
| 6 | `MovePageUp()` | `MovePageUp` | 이전 페이지로 |

#### 5.4 선택하며 이동 (MoveSel*)

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `MoveSelDown()` | `MoveSelDown` | 아래로 선택 이동 |
| 2 | `MoveSelLeft()` | `MoveSelLeft` | 왼쪽으로 선택 이동 |
| 3 | `MoveSelRight()` | `MoveSelRight` | 오른쪽으로 선택 이동 |
| 4 | `MoveSelUp()` | `MoveSelUp` | 위로 선택 이동 |
| 5 | `MoveSelDocBegin()` | `MoveSelDocBegin` | 문서 시작까지 선택 |
| 6 | `MoveSelDocEnd()` | `MoveSelDocEnd` | 문서 끝까지 선택 |
| 7 | `MoveSelLineBegin()` | `MoveSelLineBegin` | 줄 시작까지 선택 |
| 8 | `MoveSelLineEnd()` | `MoveSelLineEnd` | 줄 끝까지 선택 |
| 9 | `MoveSelLineDown()` | `MoveSelLineDown` | 다음 줄까지 선택 |
| 10 | `MoveSelLineUp()` | `MoveSelLineUp` | 이전 줄까지 선택 |
| 11 | `MoveSelNextChar()` | `MoveSelNextChar` | 다음 문자까지 선택 |
| 12 | `MoveSelPrevChar()` | `MoveSelPrevChar` | 이전 문자까지 선택 |
| 13 | `MoveSelNextWord()` | `MoveSelNextWord` | 다음 단어까지 선택 |
| 14 | `MoveSelPrevWord()` | `MoveSelPrevWord` | 이전 단어까지 선택 |
| 15 | `MoveSelPageDown()` | `MoveSelPageDown` | 다음 페이지까지 선택 |
| 16 | `MoveSelPageUp()` | `MoveSelPageUp` | 이전 페이지까지 선택 |
| 17 | `MoveSelParaBegin()` | `MoveSelParaBegin` | 문단 시작까지 선택 |
| 18 | `MoveSelParaEnd()` | `MoveSelParaEnd` | 문단 끝까지 선택 |
| 19 | `MoveSelWordBegin()` | `MoveSelWordBegin` | 단어 시작까지 선택 |
| 20 | `MoveSelWordEnd()` | `MoveSelWordEnd` | 단어 끝까지 선택 |

#### 5.5 기타 이동

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `MoveColumnBegin()` | `MoveColumnBegin` | 열 시작으로 |
| 2 | `MoveColumnEnd()` | `MoveColumnEnd` | 열 끝으로 |
| 3 | `MoveListBegin()` | `MoveListBegin` | 리스트 시작으로 |
| 4 | `MoveListEnd()` | `MoveListEnd` | 리스트 끝으로 |
| 5 | `MoveParentList()` | `MoveParentList` | 상위 리스트로 |
| 6 | `MoveRootList()` | `MoveRootList` | 루트 리스트로 |
| 7 | `MoveTopLevelBegin()` | `MoveTopLevelBegin` | 최상위 시작으로 |
| 8 | `MoveTopLevelEnd()` | `MoveTopLevelEnd` | 최상위 끝으로 |
| 9 | `MoveTopLevelList()` | `MoveTopLevelList` | 최상위 리스트로 |
| 10 | `MoveNextColumn()` | `MoveNextColumn` | 다음 열로 |
| 11 | `MovePrevColumn()` | `MovePrevColumn` | 이전 열로 |
| 12 | `MoveNextPos()` | `MoveNextPos` | 다음 위치로 |
| 13 | `MovePrevPos()` | `MovePrevPos` | 이전 위치로 |
| 14 | `MoveNextPosEx()` | `MoveNextPosEx` | 다음 위치로 (확장) |
| 15 | `MovePrevPosEx()` | `MovePrevPosEx` | 이전 위치로 (확장) |
| 16 | `MoveNextParaBegin()` | `MoveNextParaBegin` | 다음 문단 시작으로 |
| 17 | `MovePrevParaBegin()` | `MovePrevParaBegin` | 이전 문단 시작으로 |
| 18 | `MovePrevParaEnd()` | `MovePrevParaEnd` | 이전 문단 끝으로 |
| 19 | `MoveSectionDown()` | `MoveSectionDown` | 다음 섹션으로 |
| 20 | `MoveSectionUp()` | `MoveSectionUp` | 이전 섹션으로 |
| 21 | `MoveScrollDown()` | `MoveScrollDown` | 스크롤 아래로 |
| 22 | `MoveScrollUp()` | `MoveScrollUp` | 스크롤 위로 |
| 23 | `MoveScrollNext()` | `MoveScrollNext` | 스크롤 다음 |
| 24 | `MoveScrollPrev()` | `MoveScrollPrev` | 스크롤 이전 |
| 25 | `MoveViewBegin()` | `MoveViewBegin` | 뷰 시작으로 |
| 26 | `MoveViewEnd()` | `MoveViewEnd` | 뷰 끝으로 |
| 27 | `MoveViewDown()` | `MoveViewDown` | 뷰 아래로 |
| 28 | `MoveViewUp()` | `MoveViewUp` | 뷰 위로 |

---

### 6. 표 작업 (Table*) - 78개

#### 6.1 셀 이동

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `TableLeftCell()` | `TableLeftCell` | 왼쪽 셀로 |
| 2 | `TableRightCell()` | `TableRightCell` | 오른쪽 셀로 |
| 3 | `TableUpperCell()` | `TableUpperCell` | 위 셀로 |
| 4 | `TableLowerCell()` | `TableLowerCell` | 아래 셀로 |
| 5 | `TableColBegin()` | `TableColBegin` | 행 첫 셀로 |
| 6 | `TableColEnd()` | `TableColEnd` | 행 마지막 셀로 |
| 7 | `TableColPageUp()` | `TableColPageUp` | 열 맨 위로 |
| 8 | `TableColPageDown()` | `TableColPageDown` | 열 맨 아래로 |
| 9 | `TableRightCellAppend()` | `TableRightCellAppend` | 오른쪽 셀로 (행 추가) |

#### 6.2 셀 선택/블록

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `TableCellBlock()` | `TableCellBlock` | 셀 블록 상태로 |
| 2 | `TableCellBlockCol()` | `TableCellBlockCol` | 열 전체 선택 |
| 3 | `TableCellBlockRow()` | `TableCellBlockRow` | 행 전체 선택 |
| 4 | `TableCellBlockExtend()` | `TableCellBlockExtend` | 셀 블록 연장 |
| 5 | `TableCellBlockExtendAbs()` | `TableCellBlockExtendAbs` | 셀 블록 연장 (절대) |

#### 6.3 셀 정렬

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `TableCellAlignLeftTop()` | `TableCellAlignLeftTop` | 왼쪽 위 정렬 |
| 2 | `TableCellAlignCenterTop()` | `TableCellAlignCenterTop` | 가운데 위 정렬 |
| 3 | `TableCellAlignRightTop()` | `TableCellAlignRightTop` | 오른쪽 위 정렬 |
| 4 | `TableCellAlignLeftCenter()` | `TableCellAlignLeftCenter` | 왼쪽 가운데 정렬 |
| 5 | `TableCellAlignCenterCenter()` | `TableCellAlignCenterCenter` | 가운데 정렬 |
| 6 | `TableCellAlignRightCenter()` | `TableCellAlignRightCenter` | 오른쪽 가운데 정렬 |
| 7 | `TableCellAlignLeftBottom()` | `TableCellAlignLeftBottom` | 왼쪽 아래 정렬 |
| 8 | `TableCellAlignCenterBottom()` | `TableCellAlignCenterBottom` | 가운데 아래 정렬 |
| 9 | `TableCellAlignRightBottom()` | `TableCellAlignRightBottom` | 오른쪽 아래 정렬 |
| 10 | `TableVAlignTop()` | `TableVAlignTop` | 세로 위 정렬 |
| 11 | `TableVAlignCenter()` | `TableVAlignCenter` | 세로 가운데 정렬 |
| 12 | `TableVAlignBottom()` | `TableVAlignBottom` | 세로 아래 정렬 |

#### 6.4 셀 테두리

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `TableCellBorderAll()` | `TableCellBorderAll` | 모든 테두리 토글 |
| 2 | `TableCellBorderNo()` | `TableCellBorderNo` | 모든 테두리 제거 |
| 3 | `TableCellBorderOutside()` | `TableCellBorderOutside` | 바깥 테두리 토글 |
| 4 | `TableCellBorderInside()` | `TableCellBorderInside` | 안쪽 테두리 토글 |
| 5 | `TableCellBorderInsideHorz()` | `TableCellBorderInsideHorz` | 안쪽 가로 테두리 |
| 6 | `TableCellBorderInsideVert()` | `TableCellBorderInsideVert` | 안쪽 세로 테두리 |
| 7 | `TableCellBorderTop()` | `TableCellBorderTop` | 위 테두리 토글 |
| 8 | `TableCellBorderBottom()` | `TableCellBorderBottom` | 아래 테두리 토글 |
| 9 | `TableCellBorderLeft()` | `TableCellBorderLeft` | 왼쪽 테두리 토글 |
| 10 | `TableCellBorderRight()` | `TableCellBorderRight` | 오른쪽 테두리 토글 |
| 11 | `TableCellBorderDiagonalDown()` | `TableCellBorderDiagonalDown` | 대각선(⍂) 토글 |
| 12 | `TableCellBorderDiagonalUp()` | `TableCellBorderDiagonalUp` | 대각선(⍁) 토글 |

#### 6.5 셀/행/열 편집

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `TableAppendRow()` | `TableAppendRow` | 행 추가 |
| 2 | `TableSubtractRow()` | `TableSubtractRow` | 행 삭제 |
| 3 | `TableDeleteCell()` | `TableDeleteCell` | 셀 삭제 (파라미터 사용) |
| 4 | `TableMergeCell()` | `TableMergeCell` | 셀 합치기 |
| 5 | `TableSplitCell()` | `TableSplitCell` | 셀 나누기 (HParameterSet 사용) |
| 6 | `TableMergeTable()` | `TableMergeTable` | 표와 표 붙이기 |
| 7 | `TableSplitTable()` | `TableSplitTable` | 표 나누기 |
| 8 | `TableDistributeCellHeight()` | `TableDistributeCellHeight` | 셀 높이 같게 |
| 9 | `TableDistributeCellWidth()` | `TableDistributeCellWidth` | 셀 너비 같게 |

#### 6.6 셀 크기 조정

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `TableResizeDown()` | `TableResizeDown` | 아래로 크기 조정 |
| 2 | `TableResizeUp()` | `TableResizeUp` | 위로 크기 조정 |
| 3 | `TableResizeLeft()` | `TableResizeLeft` | 왼쪽으로 크기 조정 |
| 4 | `TableResizeRight()` | `TableResizeRight` | 오른쪽으로 크기 조정 |
| 5 | `TableResizeCellDown()` | `TableResizeCellDown` | 셀 아래로 |
| 6 | `TableResizeCellUp()` | `TableResizeCellUp` | 셀 위로 |
| 7 | `TableResizeCellLeft()` | `TableResizeCellLeft` | 셀 왼쪽으로 |
| 8 | `TableResizeCellRight()` | `TableResizeCellRight` | 셀 오른쪽으로 |
| 9 | `TableResizeExDown()` | `TableResizeExDown` | 아래로 (비블록) |
| 10 | `TableResizeExUp()` | `TableResizeExUp` | 위로 (비블록) |
| 11 | `TableResizeExLeft()` | `TableResizeExLeft` | 왼쪽으로 (비블록) |
| 12 | `TableResizeExRight()` | `TableResizeExRight` | 오른쪽으로 (비블록) |
| 13 | `TableResizeLineDown()` | `TableResizeLineDown` | 선 아래로 |
| 14 | `TableResizeLineUp()` | `TableResizeLineUp` | 선 위로 |
| 15 | `TableResizeLineLeft()` | `TableResizeLineLeft` | 선 왼쪽으로 |
| 16 | `TableResizeLineRight()` | `TableResizeLineRight` | 선 오른쪽으로 |

#### 6.7 표 수식

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `TableFormulaSumAuto()` | `TableFormulaSumAuto` | 블록 합계 |
| 2 | `TableFormulaSumHor()` | `TableFormulaSumHor` | 가로 합계 |
| 3 | `TableFormulaSumVer()` | `TableFormulaSumVer` | 세로 합계 |
| 4 | `TableFormulaAvgAuto()` | `TableFormulaAvgAuto` | 블록 평균 |
| 5 | `TableFormulaAvgHor()` | `TableFormulaAvgHor` | 가로 평균 |
| 6 | `TableFormulaAvgVer()` | `TableFormulaAvgVer` | 세로 평균 |
| 7 | `TableFormulaProAuto()` | `TableFormulaProAuto` | 블록 곱 |
| 8 | `TableFormulaProHor()` | `TableFormulaProHor` | 가로 곱 |
| 9 | `TableFormulaProVer()` | `TableFormulaProVer` | 세로 곱 |

#### 6.8 기타 표 기능

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `TableAutoFill()` | `TableAutoFill` | 자동 채우기 |
| 2 | `TableAutoFillDlg()` | `TableAutoFillDlg` | 자동 채우기 대화상자 |
| 3 | `TableDrawPen()` | `TableDrawPen` | 표 그리기 |
| 4 | `TableDrawPenStyle()` | `TableDrawPenStyle` | 표 그리기 선 모양 |
| 5 | `TableDrawPenWidth()` | `TableDrawPenWidth` | 표 그리기 선 굵기 |
| 6 | `TableEraser()` | `TableEraser` | 표 지우개 |
| 7 | `TableDeleteComma()` | `TableDeleteComma` | 천단위 콤마 제거 |
| 8 | `TableInsertComma()` | `TableInsertComma` | 천단위 콤마 삽입 |

---

### 7. 그리기 개체 (ShapeObj*) - 52개

#### 7.1 정렬

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `ShapeObjAlignLeft()` | `ShapeObjAlignLeft` | 왼쪽 정렬 |
| 2 | `ShapeObjAlignCenter()` | `ShapeObjAlignCenter` | 가운데 정렬 |
| 3 | `ShapeObjAlignRight()` | `ShapeObjAlignRight` | 오른쪽 정렬 |
| 4 | `ShapeObjAlignTop()` | `ShapeObjAlignTop` | 위 정렬 |
| 5 | `ShapeObjAlignMiddle()` | `ShapeObjAlignMiddle` | 중간 정렬 |
| 6 | `ShapeObjAlignBottom()` | `ShapeObjAlignBottom` | 아래 정렬 |
| 7 | `ShapeObjAlignWidth()` | `ShapeObjAlignWidth` | 폭 맞춤 |
| 8 | `ShapeObjAlignHeight()` | `ShapeObjAlignHeight` | 높이 맞춤 |
| 9 | `ShapeObjAlignSize()` | `ShapeObjAlignSize` | 크기 맞춤 |
| 10 | `ShapeObjAlignHorzSpacing()` | `ShapeObjAlignHorzSpacing` | 가로 간격 맞춤 |
| 11 | `ShapeObjAlignVertSpacing()` | `ShapeObjAlignVertSpacing` | 세로 간격 맞춤 |

#### 7.2 순서/레이어

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `ShapeObjBringToFront()` | `ShapeObjBringToFront` | 맨 앞으로 |
| 2 | `ShapeObjSendToBack()` | `ShapeObjSendToBack` | 맨 뒤로 |
| 3 | `ShapeObjBringForward()` | `ShapeObjBringForward` | 앞으로 |
| 4 | `ShapeObjSendBack()` | `ShapeObjSendBack` | 뒤로 |
| 5 | `ShapeObjBringInFrontOfText()` | `ShapeObjBringInFrontOfText` | 글 앞으로 |
| 6 | `ShapeObjCtrlSendBehindText()` | `ShapeObjCtrlSendBehindText` | 글 뒤로 |

#### 7.3 변형

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `ShapeObjHorzFlip()` | `ShapeObjHorzFlip` | 좌우 뒤집기 |
| 2 | `ShapeObjVertFlip()` | `ShapeObjVertFlip` | 상하 뒤집기 |
| 3 | `ShapeObjHorzFlipOrgState()` | `ShapeObjHorzFlipOrgState` | 좌우 뒤집기 복원 |
| 4 | `ShapeObjVertFlipOrgState()` | `ShapeObjVertFlipOrgState` | 상하 뒤집기 복원 |
| 5 | `ShapeObjRotater()` | `ShapeObjRotater` | 자유 회전 |
| 6 | `ShapeObjRightAngleRotater()` | `ShapeObjRightAngleRotater` | 90도 시계방향 회전 |
| 7 | `ShapeObjRightAngleRotaterAnticlockwise()` | `ShapeObjRightAngleRotaterAnticlockwise` | 90도 반시계 회전 |

#### 7.4 크기/이동

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `ShapeObjMoveUp()` | `ShapeObjMoveUp` | 위로 이동 |
| 2 | `ShapeObjMoveDown()` | `ShapeObjMoveDown` | 아래로 이동 |
| 3 | `ShapeObjMoveLeft()` | `ShapeObjMoveLeft` | 왼쪽으로 이동 |
| 4 | `ShapeObjMoveRight()` | `ShapeObjMoveRight` | 오른쪽으로 이동 |
| 5 | `ShapeObjResizeUp()` | `ShapeObjResizeUp` | 크기 위로 |
| 6 | `ShapeObjResizeDown()` | `ShapeObjResizeDown` | 크기 아래로 |
| 7 | `ShapeObjResizeLeft()` | `ShapeObjResizeLeft` | 크기 왼쪽으로 |
| 8 | `ShapeObjResizeRight()` | `ShapeObjResizeRight` | 크기 오른쪽으로 |

#### 7.5 그룹/선택

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `ShapeObjGroup()` | `ShapeObjGroup` | 그룹화 |
| 2 | `ShapeObjUngroup()` | `ShapeObjUngroup` | 그룹 해제 |
| 3 | `ShapeObjSelect()` | `ShapeObjSelect` | 개체 선택 도구 |
| 4 | `ShapeObjNextObject()` | `ShapeObjNextObject` | 다음 개체로 |
| 5 | `ShapeObjPrevObject()` | `ShapeObjPrevObject` | 이전 개체로 |
| 6 | `ShapeObjLock()` | `ShapeObjLock` | 개체 잠금 |
| 7 | `ShapeObjUnlockAll()` | `ShapeObjUnlockAll` | 모든 잠금 해제 |

#### 7.6 캡션/글상자

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `ShapeObjAttachCaption()` | `ShapeObjAttachCaption` | 캡션 추가 |
| 2 | `ShapeObjDetachCaption()` | `ShapeObjDetachCaption` | 캡션 제거 |
| 3 | `ShapeObjInsertCaptionNum()` | `ShapeObjInsertCaptionNum` | 캡션 번호 삽입 |
| 4 | `ShapeObjAttachTextBox()` | `ShapeObjAttachTextBox` | 글상자로 변경 |
| 5 | `ShapeObjDetachTextBox()` | `ShapeObjDetachTextBox` | 사각형으로 변경 |
| 6 | `ShapeObjToggleTextBox()` | `ShapeObjToggleTextBox` | 글상자 토글 |
| 7 | `ShapeObjTextBoxEdit()` | `ShapeObjTextBoxEdit` | 글상자 편집 모드 |
| 8 | `ShapeObjTableSelCell()` | `ShapeObjTableSelCell` | 표 첫 셀 선택 |

#### 7.7 속성

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `ShapeObjFillProperty()` | `ShapeObjFillProperty` | 채우기 속성 |
| 2 | `ShapeObjLineProperty()` | `ShapeObjLineProperty` | 선 속성 |
| 3 | `ShapeObjLineStyleOther()` | `ShapeObjLineStyleOhter` | 다른 선 종류 |
| 4 | `ShapeObjLineWidthOther()` | `ShapeObjLineWidthOhter` | 다른 선 굵기 |
| 5 | `ShapeObjNorm()` | `ShapeObjNorm` | 기본 도형 설정 |
| 6 | `ShapeObjGuideLine()` | `ShapeObjGuideLine` | 안내선 설정 |
| 7 | `ShapeObjShowGuideLine()` | `ShapeObjShowGuideLine` | 안내선 보기 토글 |
| 8 | `ShapeObjShowGuideLineBase()` | `ShapeObjShowGuideLineBase` | 그리기 안내선 (2024+) |
| 9 | `ShapeObjWrapSquare()` | `ShapeObjWrapSquare` | 직사각형 배치 |
| 10 | `ShapeObjWrapTopAndBottom()` | `ShapeObjWrapTopAndBottom` | 자리 차지 배치 |
| 11 | `ShapeObjSaveAsPicture()` | `ShapeObjSaveAsPicture` | 그림으로 저장 |

---

### 8. 파일 작업 (File*) - 18개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `FileNew()` | `FileNew` | 새 문서 |
| 2 | `FileNewTab()` | `FileNewTab` | 새 탭 |
| 3 | `FileOpen()` | `FileOpen` | 파일 열기 대화상자 |
| 4 | `FileOpenMRU()` | `FileOpenMRU` | 최근 파일 열기 |
| 5 | `FileSave()` | `FileSave` | 저장 |
| 6 | `FileSaveAs()` | `FileSaveAs` | 다른 이름으로 저장 |
| 7 | `FileSaveAsDRM()` | `FileSaveAsDRM` | DRM 저장 |
| 8 | `FileSaveOptionDlg()` | `FileSaveOptionDlg` | 저장 옵션 대화상자 |
| 9 | `FileClose()` | `FileClose` | 파일 닫기 |
| 10 | `FileFind()` | `FileFind` | 파일 찾기 |
| 11 | `FilePreview()` | `FilePreview` | 미리 보기 |
| 12 | `FileQuit()` | `FileQuit` | 프로그램 종료 |
| 13 | `FileNextVersionDiff()` | `FileNextVersionDiff` | 다음 버전 비교 |
| 14 | `FilePrevVersionDiff()` | `FilePrevVersionDiff` | 이전 버전 비교 |
| 15 | `FileVersionDiffChangeAlign()` | `FileVersionDiffChangeAlign` | 비교 정렬 변경 |
| 16 | `FileVersionDiffSameAlign()` | `FileVersionDiffSameAlign` | 비교 같은 정렬 |
| 17 | `FileVersionDiffSyncScroll()` | `FileVersionDiffSyncScroll` | 비교 동기 스크롤 |

---

### 9. 삽입 (Insert*) - 25개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `InsertAutoNum()` | `InsertAutoNum` | 자동 번호 삽입 |
| 2 | `InsertCpNo()` | `InsertCpNo` | 현재 쪽 번호 |
| 3 | `InsertCpTpNo()` | `InsertCpTpNo` | 현재 쪽/전체 쪽 번호 |
| 4 | `InsertTpNo()` | `InsertTpNo` | 전체 쪽 번호 |
| 5 | `InsertPageNum()` | `InsertPageNum` | 쪽 번호 매기기 |
| 6 | `InsertDateCode()` | `InsertDateCode` | 날짜 코드 |
| 7 | `InsertDocInfo()` | `InsertDocInfo` | 문서 정보 |
| 8 | `InsertEndnote()` | `InsertEndnote` | 미주 삽입 |
| 9 | `InsertFootnote()` | `InsertFootnote` | 각주 삽입 |
| 10 | `InsertFieldCitation()` | `InsertFieldCitation` | 인용 필드 |
| 11 | `InsertFieldDateTime()` | `InsertFieldDateTime` | 날짜/시간 필드 |
| 12 | `InsertFieldMemo()` | `InsertFieldMemo` | 메모 필드 |
| 13 | `InsertFieldRevisionChagne()` | `InsertFieldRevisionChagne` | 교정 필드 |
| 14 | `InsertFixedWidthSpace()` | `InsertFixedWidthSpace` | 고정폭 공백 |
| 15 | `InsertNonBreakingSpace()` | `InsertNonBreakingSpace` | 줄 바꿈 없는 공백 |
| 16 | `InsertSoftHyphen()` | `InsertSoftHyphen` | 소프트 하이픈 |
| 17 | `InsertSpace()` | `InsertSpace` | 공백 삽입 |
| 18 | `InsertTab()` | `InsertTab` | 탭 삽입 |
| 19 | `InsertLine()` | `InsertLine` | 줄 삽입 |
| 20 | `InsertStringDateTime()` | `InsertStringDateTime` | 날짜/시간 문자열 |
| 21 | `InsertLastPrintDate()` | `InsertLastPrintDate` | 마지막 인쇄 날짜 |
| 22 | `InsertLastSaveBy()` | `InsertLastSaveBy` | 마지막 저장자 |
| 23 | `InsertLastSaveDate()` | `InsertLastSaveDate` | 마지막 저장 날짜 |

---

### 10. 삭제 (Delete*) - 12개

| # | Python 메서드 | HAction 문자열 | 설명 | 파라미터 |
|---|--------------|----------------|------|----------|
| 1 | `Delete()` | `Delete` | 삭제 (Del 키) | delete_ctrl |
| 2 | `DeleteBack()` | `DeleteBack` | 뒤로 삭제 (Backspace) | delete_ctrl |
| 3 | `DeleteLine()` | `DeleteLine` | 줄 삭제 | delete_ctrl |
| 4 | `DeleteLineEnd()` | `DeleteLineEnd` | 줄 끝까지 삭제 | delete_ctrl |
| 5 | `DeleteWord()` | `DeleteWord` | 단어 삭제 | delete_ctrl |
| 6 | `DeleteWordBack()` | `DeleteWordBack` | 이전 단어 삭제 | delete_ctrl |
| 7 | `DeleteField()` | `DeleteField` | 필드 삭제 | |
| 8 | `DeleteFieldMemo()` | `DeleteFieldMemo` | 메모 필드 삭제 | |
| 9 | `DeletePage()` | `DeletePage` | 페이지 삭제 | |
| 10 | `DeletePrivateInfoMark()` | `DeletePrivateInfoMark` | 개인정보 마크 삭제 | |
| 11 | `DeletePrivateInfoMarkAtCurrentPos()` | `DeletePrivateInfoMarkAtCurrentPos` | 현재 위치 개인정보 마크 삭제 | |
| 12 | `DeleteDocumentMasterPage()` | `DeleteDocumentMasterPage` | 문서 바탕쪽 삭제 | |
| 13 | `DeleteSectionMasterPage()` | `DeleteSectionMasterPage` | 구역 바탕쪽 삭제 | |

---

### 11. 클립보드 (Copy/Cut/Paste) - 8개

| # | Python 메서드 | HAction 문자열 | 설명 | 파라미터 |
|---|--------------|----------------|------|----------|
| 1 | `Copy()` | `Copy` | 복사 | |
| 2 | `Cut()` | `Cut` | 잘라내기 | remove_cell |
| 3 | `Paste()` | `Paste` | 붙여넣기 | |
| 4 | `PasteSpecial()` | `PasteSpecial` | 선택하여 붙여넣기 | |
| 5 | `CopyPage()` | `CopyPage` | 페이지 복사 | |
| 6 | `PastePage()` | `PastePage` | 페이지 붙여넣기 | |
| 7 | `Erase()` | `Erase` | 지우기 | |

---

### 12. 변경 추적 (TrackChange*) - 15개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `TrackChangeApply()` | `TrackChangeApply` | 변경 적용 |
| 2 | `TrackChangeApplyAll()` | `TrackChangeApplyAll` | 모든 변경 적용 |
| 3 | `TrackChangeApplyNext()` | `TrackChangeApplyNext` | 적용 후 다음 |
| 4 | `TrackChangeApplyPrev()` | `TrackChangeApplyPrev` | 적용 후 이전 |
| 5 | `TrackChangeApplyViewAll()` | `TrackChangeApplyViewAll` | 표시된 변경 모두 적용 |
| 6 | `TrackChangeCancel()` | `TrackChangeCancel` | 변경 취소 |
| 7 | `TrackChangeCancelAll()` | `TrackChangeCancelAll` | 모든 변경 취소 |
| 8 | `TrackChangeCancelNext()` | `TrackChangeCancelNext` | 취소 후 다음 |
| 9 | `TrackChangeCancelPrev()` | `TrackChangeCancelPrev` | 취소 후 이전 |
| 10 | `TrackChangeCancelViewAll()` | `TrackZChangeCancelViewAll` | 표시된 변경 모두 취소 |
| 11 | `TrackChangeNext()` | `TrackChangeNext` | 다음 변경 |
| 12 | `TrackChangePrev()` | `TrackChangePrev` | 이전 변경 |
| 13 | `TrackChangeAuthor()` | `TrackChangeAuthor` | 작성자 변경 |

---

### 13. 보기 옵션 (ViewOption*) - 20개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `ViewOptionCtrlMark()` | `ViewOptionCtrlMark` | 조판 부호 보기 |
| 2 | `ViewOptionGuideLine()` | `ViewOptionGuideLine` | 안내선 보기 |
| 3 | `ViewOptionMemo()` | `ViewOptionMemo` | 메모 보기 |
| 4 | `ViewOptionMemoGuideline()` | `ViewOptionMemoGuideline` | 메모 안내선 |
| 5 | `ViewOptionPaper()` | `ViewOptionPaper` | 쪽 윤곽 보기 |
| 6 | `ViewOptionParaMark()` | `ViewOptionParaMark` | 문단 부호 보기 |
| 7 | `ViewOptionPicture()` | `ViewOptionPicture` | 그림 보기 |
| 8 | `ViewOptionRevision()` | `ViewOptionRevision` | 교정 부호 보기 |
| 9 | `ViewOptionTrackChange()` | `ViewOptionTrackChange` | 변경 추적 보기 |
| 10 | `ViewOptionTrackChangeFinal()` | `ViewOptionTrackChangeFinal` | 최종본 보기 |
| 11 | `ViewOptionTrackChangeFinalMemo()` | `ViewOptionTrackChangeFinalMemo` | 최종본+메모 |
| 12 | `ViewOptionTrackChangeInline()` | `ViewOptionTrackChangeInline` | 인라인 표시 |
| 13 | `ViewOptionTrackChangeInsertDelete()` | `ViewOptionTrackChangeInsertDelete` | 삽입/삭제 표시 |
| 14 | `ViewOptionTrackChangeOriginal()` | `ViewOptionTrackChangeOriginal` | 원본 보기 |
| 15 | `ViewOptionTrackChangeOriginalMemo()` | `ViewOptionTrackChangeOriginalMemo` | 원본+메모 |
| 16 | `ViewOptionTrackChangeShape()` | `ViewOptionTrackChangeShape` | 서식 변경 표시 |
| 17 | `ViewOptionTrackChnageInfo()` | `ViewOptionTrackChnageInfo` | 변경 정보 보기 |
| 18 | `ViewZoomFitPage()` | `ViewZoomFitPage` | 페이지에 맞춤 |
| 19 | `ViewZoomFitWidth()` | `ViewZoomFitWidth` | 폭에 맞춤 |
| 20 | `ViewZoomNormal()` | `ViewZoomFitNormal` | 기본 확대 |
| 21 | `ViewZoomRibon()` | `ViewZoomRibon` | 확대 리본 |
| 22 | `ViewTabButton()` | `ViewTabButton` | 문서 탭 보기 |
| 23 | `ViewIdiom()` | `ViewIdiom` | 상용구 보기 |

---

### 14. 매크로 (Macro*) - 24개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `MacroPause()` | `MacroPause` | 매크로 일시정지 |
| 2-12 | `MacroPlay1()` ~ `MacroPlay11()` | `MacroPlay1` ~ `MacroPlay11` | 매크로 1~11 실행 |
| 13 | `MacroRepeat()` | `MacroRepeat` | 매크로 반복 |
| 14 | `MacroStop()` | `MacroStop` | 매크로 중지 |
| 15 | `ScrMacroPause()` | `ScrMacroPause` | 스크립트 매크로 일시정지 |
| 16-26 | `ScrMacroPlay1()` ~ `ScrMacroPlay11()` | `ScrMacroPlay1` ~ `ScrMacroPlay11` | 스크립트 매크로 1~11 |
| 27 | `ScrMacroRepeat()` | `ScrMacroRepeat` | 스크립트 매크로 반복 |
| 28 | `ScrMacroStop()` | `ScrMacroStop` | 스크립트 매크로 중지 |

---

### 15. 빠른 교정/빠른 찾기 (Quick*) - 24개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `QuickCommandRun()` | `QuickCommandRun` | 빠른 명령 실행 |
| 2 | `QuickCorrect()` | `QuickCorrect` | 빠른 교정 |
| 3 | `QuickCorrectRun()` | `QuickCorrectRun` | 빠른 교정 실행 |
| 4 | `QuickCorrectSound()` | `QuickCorrectSound` | 빠른 교정 소리 |
| 5-14 | `QuickMarkInsert0()` ~ `QuickMarkInsert9()` | `QuickMarkInsert0` ~ `QuickMarkInsert9` | 빠른 책갈피 삽입 0~9 |
| 15-24 | `QuickMarkMove0()` ~ `QuickMarkMove9()` | `QuickMarkMove0` ~ `QuickMarkMove9` | 빠른 책갈피로 이동 0~9 |

---

### 16. 머리말/꼬리말 (HeaderFooter*) - 4개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `HeaderFooterDelete()` | `HeaderFooterDelete` | 머리말/꼬리말 삭제 |
| 2 | `HeaderFooterModify()` | `HeaderFooterModify` | 머리말/꼬리말 수정 |
| 3 | `HeaderFooterToNext()` | `HeaderFooterToNext` | 다음 머리말/꼬리말 |
| 4 | `HeaderFooterToPrev()` | `HeaderFooterToPrev` | 이전 머리말/꼬리말 |

---

### 17. 바탕쪽 (MasterPage*) - 10개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `MasterPage()` | `MasterPage` | 바탕쪽 |
| 2 | `MasterPageDuplicate()` | `MasterPageDuplicate` | 바탕쪽 복제 |
| 3 | `MasterPageExcept()` | `MasterPageExcept` | 바탕쪽 제외 |
| 4 | `MasterPageFront()` | `MasterPageFront` | 바탕쪽 앞면 |
| 5 | `MasterPagePrevSection()` | `MasterPagePrevSection` | 이전 구역 바탕쪽 |
| 6 | `MasterPageToNext()` | `MasterPageToNext` | 다음 바탕쪽 |
| 7 | `MasterPageToPrevious()` | `MasterPageToPrevious` | 이전 바탕쪽 |
| 8 | `MasterPageType()` | `MasterPageType` | 바탕쪽 유형 |
| 9 | `MPSectionToNext()` | `MPSectionToNext` | 다음 구역으로 |
| 10 | `MPSectionToPrevious()` | `MPSectionToPrevious` | 이전 구역으로 |
| 11 | `MPShowMarginBorder()` | `MPShowMarginBorder` | 여백 테두리 보기 |

---

### 18. 그림 (Picture*) - 9개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1-8 | `PictureEffect1()` ~ `PictureEffect8()` | `PictureEffect1` ~ `PictureEffect8` | 그림 효과 1~8 |
| 9 | `PictureInsertDialog()` | `PictureInsertDialog` | 그림 삽입 대화상자 |
| 10 | `PictureLinkedToEmbedded()` | `PictureLinkedToEmbedded` | 연결→포함 변환 |
| 11 | `PictureSave()` | `PictureSave` | 그림 저장 |
| 12 | `PictureScissor()` | `PictureScissor` | 그림 자르기 |
| 13 | `PictureToOriginal()` | `PictureToOriginal` | 원본 크기로 |

---

### 19. 주석/메모 (Note*, Memo*, Comment*) - 16개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `NoteDelete()` | `NoteDelete` | 각주/미주 삭제 |
| 2 | `NoteModify()` | `NoteModify` | 각주/미주 수정 |
| 3 | `NoteToNext()` | `NoteToNext` | 다음 각주/미주 |
| 4 | `NoteToPrev()` | `NoteToPrev` | 이전 각주/미주 |
| 5 | `NoteNumProperty()` | `NoteNumProperty` | 각주 번호 속성 |
| 6 | `NoteNumShape()` | `NoteNumShape` | 각주 번호 모양 |
| 7 | `NoteLineColor()` | `NoteLineColor` | 각주 선 색 |
| 8 | `NoteLineLength()` | `NoteLineLength` | 각주 선 길이 |
| 9 | `NoteLineShape()` | `NoteLineShape` | 각주 선 모양 |
| 10 | `NoteLineWeight()` | `NoteLineWeight` | 각주 선 굵기 |
| 11 | `NotePosition()` | `NotePosition` | 각주 위치 |
| 12 | `MemoToNext()` | `MemoToNext` | 다음 메모 |
| 13 | `MemoToPrev()` | `MemoToPrev` | 이전 메모 |
| 14 | `Comment()` | `Comment` | 덧글 삽입 |
| 15 | `CommentDelete()` | `CommentDelete` | 덧글 삭제 |
| 16 | `CommentModify()` | `CommentModify` | 덧글 수정 |
| 17 | `ReplyMemo()` | `ReplyMemo` | 메모 답장 |
| 18 | `EditFieldMemo()` | `EditFieldMemo` | 메모 필드 편집 |

---

### 20. 그리기 개체 작성 (DrawObj*) - 4개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `DrawObjCancelOneStep()` | `DrawObjCancelOneStep` | 그리기 한 단계 취소 |
| 2 | `DrawObjEditDetail()` | `DrawObjEditDetail` | 그리기 상세 편집 |
| 3 | `DrawObjOpenClosePolygon()` | `DrawObjOpenClosePolygon` | 다각형 열기/닫기 |
| 4 | `DrawObjTemplateSave()` | `DrawObjTemplateSave` | 그리기 템플릿 저장 |

---

### 21. 입력 (Input*) - 4개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `InputCodeChange()` | `InputCodeChange` | 입력 코드 변경 |
| 2 | `InputHanja()` | `InputHanja` | 한자 입력 |
| 3 | `InputHanjaBusu()` | `InputHanjaBusu` | 한자 부수 입력 |
| 4 | `InputHanjaMean()` | `InputHanjaMean` | 한자 뜻 입력 |

---

### 22. 폼 개체 (FormObj*) - 8개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `FormDesignMode()` | `FormDesignMode` | 폼 디자인 모드 |
| 2 | `FormObjCreatorCheckButton()` | `FormObjCreatorCheckButton` | 체크 버튼 생성 |
| 3 | `FormObjCreatorComboBox()` | `FormObjCreatorComboBox` | 콤보 박스 생성 |
| 4 | `FormObjCreatorEdit()` | `FormObjCreatorEdit` | 편집 상자 생성 |
| 5 | `FormObjCreatorListBox()` | `FormObjCreatorListBox` | 리스트 박스 생성 |
| 6 | `FormObjCreatorPushButton()` | `FormObjCreatorPushButton` | 푸시 버튼 생성 |
| 7 | `FormObjCreatorRadioButton()` | `FormObjCreatorRadioButton` | 라디오 버튼 생성 |
| 8 | `FormObjCreatorScrollBar()` | `FormObjCreatorScrollBar` | 스크롤바 생성 |
| 9 | `FormObjRadioGroup()` | `FormObjRadioGroup` | 라디오 그룹 |

---

### 23. 창/프레임 (Frame*, Window*, Split*) - 20개

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `FrameFullScreen()` | `FrameFullScreen` | 전체 화면 |
| 2 | `FrameFullScreenEnd()` | `FrameFullScreenEnd` | 전체 화면 종료 |
| 3 | `FrameHRuler()` | `FrameHRuler` | 가로 눈금자 |
| 4 | `FrameVRuler()` | `FrameVRuler` | 세로 눈금자 |
| 5 | `FrameStatusBar()` | `FrameStatusBar` | 상태 표시줄 |
| 6 | `FrameViewZoomRibon()` | `FrameViewZoomRibon` | 확대 리본 |
| 7 | `WindowAlignCascade()` | `WindowAlignCascade` | 창 계단식 배열 |
| 8 | `WindowAlignTileHorz()` | `WindowAlignTileHorz` | 창 가로 배열 |
| 9 | `WindowAlignTileVert()` | `WindowAlignTileVert` | 창 세로 배열 |
| 10 | `WindowList()` | `WindowList` | 창 목록 |
| 11 | `WindowMinimizeAll()` | `WindowMinimizeAll` | 모든 창 최소화 |
| 12 | `WindowNextPane()` | `WindowNextPane` | 다음 창 |
| 13 | `WindowNextTab()` | `WindowNextTab` | 다음 탭 |
| 14 | `WindowPrevTab()` | `WindowPrevTab` | 이전 탭 |
| 15 | `SplitAll()` | `SplitAll` | 창 분할 |
| 16 | `SplitHorz()` | `SplitHorz` | 가로 분할 |
| 17 | `SplitVert()` | `SplitVert` | 세로 분할 |
| 18 | `SplitMemo()` | `SplitMemo` | 메모창 분할 |
| 19 | `SplitMemoClose()` | `SplitMemoClose` | 메모창 닫기 |
| 20 | `SplitMemoOpen()` | `SplitMemoOpen` | 메모창 열기 |
| 21 | `SplitMainActive()` | `SplitMainActive` | 메인 활성화 |
| 22 | `NoSplit()` | `NoSplit` | 분할 해제 |

---

### 24. 기타 액션 (Misc)

| # | Python 메서드 | HAction 문자열 | 설명 |
|---|--------------|----------------|------|
| 1 | `Cancel()` | `Cancel` | 취소 (Esc) |
| 2 | `Close()` | `Close` | 닫기 |
| 3 | `CloseEx()` | `CloseEx` | 확장 닫기 (상위 리스트로) |
| 4 | `Undo()` | `Undo` | 실행 취소 |
| 5 | `Redo()` | `Redo` | 다시 실행 |
| 6 | `Select()` | `Select` | 선택 시작 |
| 7 | `SelectAll()` | `SelectAll` | 전체 선택 |
| 8 | `SelectColumn()` | `SelectColumn` | 열 선택 |
| 9 | `SelectCtrlFront()` | `SelectCtrlFront` | 앞 컨트롤 선택 |
| 10 | `SelectCtrlReverse()` | `SelectCtrlReverse` | 뒤 컨트롤 선택 |
| 11 | `UnSelectCtrl()` | `UnSelectCtrl` | 컨트롤 선택 해제 |
| 12 | `ToggleOverwrite()` | `ToggleOverwrite` | 삽입/수정 모드 전환 |
| 13 | `ReturnPrevPos()` | `ReturnPrevPos` | 이전 위치로 돌아가기 |
| 14 | `ReturnKeyInField()` | `ReturnKeyInField` | 필드 내 엔터 |
| 15 | `RecalcPageCount()` | `RecalcPageCount` | 페이지 수 재계산 |
| 16 | `SpellingCheck()` | `SpellingCheck` | 맞춤법 검사 |
| 17 | `HanThDIC()` | `HanThDIC` | 한글 사전 |
| 18 | `HwpDic()` | `HwpDic` | HWP 사전 |
| 19 | `HwpWSDic()` | `HwpWSDic` | HWP 웹 사전 |
| 20 | `EasyFind()` | `EasyFind` | 빠른 찾기 |
| 21 | `MailMergeField()` | `MailMergeField` | 메일 머지 필드 |
| 22 | `MakeIndex()` | `MakeIndex` | 색인 만들기 |
| 23 | `LabelAdd()` | `LabelAdd` | 라벨 추가 |
| 24 | `LabelTemplate()` | `LabelTemplate` | 라벨 템플릿 |
| 25 | `Jajun()` | `Jajun` | 자준 |
| 26 | `ChangeSkin()` | `ChangeSkin` | 스킨 변경 |
| 27 | `SoftKeyboard()` | `Soft Keyboard` | 소프트 키보드 |

---

## 특수 메서드 (HParameterSet 사용)

일부 메서드는 단순 `HAction.Run()` 호출이 아닌 `HParameterSet`을 사용합니다:

### TableSplitCell
```python
def TableSplitCell(self, Rows=2, Cols=0, DistributeHeight=0, Merge=0) -> bool:
    pset = self.hwp.HParameterSet.HTableSplitCell
    pset.Rows = Rows
    pset.Cols = Cols
    pset.DistributeHeight = DistributeHeight
    pset.Merge = Merge
    return self.hwp.HAction.Execute("TableSplitCell", pset.HSet)
```

### ShapeObjAttachCaption
```python
def ShapeObjAttachCaption(self, text="", add_num=True) -> bool:
    self.hwp.HAction.Run("ShapeObjAttachCaption")
    if not text:
        return True
    self.SelectAll()
    self.Delete()
    if add_num:
        self.ShapeObjInsertCaptionNum()
    try:
        return self.insert_text(text)
    finally:
        self.CloseEx()
```

### Delete/DeleteBack (delete_ctrl 파라미터)
```python
def Delete(self, delete_ctrl=True) -> bool:
    if delete_ctrl:
        return self.hwp.HAction.Run("Delete")
    else:
        # 컨트롤 삭제 방지 로직
        ...
```

---

## C++ 구현 가이드

### 기본 Run 액션 구현
```cpp
// RunMethods.h
class RunMethods {
public:
    bool ASendBrowserText();
    bool AutoChangeHangul();
    bool BreakPara();
    // ... 684개 메서드

private:
    IDispatch* m_pHwp;
    bool InvokeHActionRun(const wchar_t* actionName);
};

// RunMethods.cpp
bool RunMethods::InvokeHActionRun(const wchar_t* actionName) {
    DISPID dispid;
    OLECHAR* name = L"HAction";
    m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);

    // HAction 객체 획득
    VARIANT varHAction;
    VariantInit(&varHAction);
    DISPPARAMS dpNoArgs = {NULL, NULL, 0, 0};
    m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                   DISPATCH_PROPERTYGET, &dpNoArgs, &varHAction, NULL, NULL);

    IDispatch* pHAction = varHAction.pdispVal;

    // Run 메서드 호출
    OLECHAR* runName = L"Run";
    pHAction->GetIDsOfNames(IID_NULL, &runName, 1, LOCALE_USER_DEFAULT, &dispid);

    VARIANT varAction;
    varAction.vt = VT_BSTR;
    varAction.bstrVal = SysAllocString(actionName);

    DISPPARAMS dpArgs = {&varAction, NULL, 1, 0};
    VARIANT varResult;
    VariantInit(&varResult);

    HRESULT hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_METHOD, &dpArgs, &varResult, NULL, NULL);

    SysFreeString(varAction.bstrVal);
    pHAction->Release();

    return SUCCEEDED(hr) && varResult.boolVal;
}

bool RunMethods::BreakPara() {
    return InvokeHActionRun(L"BreakPara");
}
```

### pybind11 바인딩
```cpp
// bindings.cpp
#include <pybind11/pybind11.h>
namespace py = pybind11;

PYBIND11_MODULE(cpyhwpx, m) {
    py::class_<HwpWrapper>(m, "Hwp")
        // Run 액션들
        .def("ASendBrowserText", &HwpWrapper::ASendBrowserText)
        .def("AutoChangeHangul", &HwpWrapper::AutoChangeHangul)
        .def("BreakPara", &HwpWrapper::BreakPara)
        .def("BreakPage", &HwpWrapper::BreakPage)
        .def("BreakLine", &HwpWrapper::BreakLine)
        // ... 684개 메서드 바인딩
        ;
}
```

---

## 포팅 우선순위

### High (자주 사용)
1. Break* (줄바꿈/나누기)
2. Move* (이동)
3. Table* (표)
4. Copy/Cut/Paste (클립보드)
5. Delete* (삭제)
6. Select* (선택)
7. Undo/Redo

### Medium
1. CharShape* (글자 모양)
2. ParagraphShape* (문단 모양)
3. ShapeObj* (개체)
4. File* (파일)
5. Insert* (삽입)

### Low
1. Macro* (매크로)
2. ViewOption* (보기 옵션)
3. TrackChange* (변경 추적)
4. FormObj* (폼)
5. 기타

---

## 진행 상황

| 카테고리 | 총 개수 | 완료 | 진행률 |
|---------|--------|------|--------|
| 자동 기능 (Auto*) | 22 | 0 | 0% |
| 줄바꿈/나누기 (Break*) | 6 | 0 | 0% |
| 글자 모양 (CharShape*) | 33 | 0 | 0% |
| 문단 모양 (ParagraphShape*) | 20 | 0 | 0% |
| 이동 (Move*) | 85 | 0 | 0% |
| 표 (Table*) | 78 | 0 | 0% |
| 그리기 개체 (ShapeObj*) | 52 | 0 | 0% |
| 파일 (File*) | 18 | 0 | 0% |
| 삽입 (Insert*) | 25 | 0 | 0% |
| 삭제 (Delete*) | 13 | 0 | 0% |
| 클립보드 | 8 | 0 | 0% |
| 변경 추적 (TrackChange*) | 15 | 0 | 0% |
| 보기 옵션 (ViewOption*) | 23 | 0 | 0% |
| 매크로 (Macro*) | 28 | 0 | 0% |
| 빠른 교정 (Quick*) | 24 | 0 | 0% |
| 기타 | ~230 | 0 | 0% |
| **총계** | **684** | **0** | **0%** |
