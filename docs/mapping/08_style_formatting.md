# 스타일/서식 매핑

## 요약
- 총 함수: 약 70개 (core.py ~20개 + Run 액션 ~50개)
- 포팅 완료: 0개 (0%)
- 우선순위: High (Phase 3-4)

---

## 글자 모양 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 1 | `CharShape` (속성) | getter/setter | Any | 글자 모양 파라미터셋 | `GetCharShape()` / `SetCharShape()` | TODO |
| 2 | `get_charshape()` | () | HParameterSet | 글자 모양 조회 | `GetCharShape()` | TODO |
| 3 | `get_charshape_as_dict()` | () | dict | 글자 모양을 딕셔너리로 | `GetCharShapeAsDict()` | TODO |
| 4 | `set_charshape()` | (pset) | bool | 글자 모양 적용 | `SetCharShape()` | TODO |

## 문단 모양 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 5 | `ParaShape` (속성) | getter/setter | Any | 문단 모양 파라미터셋 | `GetParaShape()` / `SetParaShape()` | TODO |
| 6 | `get_parashape()` | () | HParameterSet | 문단 모양 조회 | `GetParaShape()` | TODO |
| 7 | `get_parashape_as_dict()` | () | dict | 문단 모양을 딕셔너리로 | `GetParaShapeAsDict()` | TODO |
| 8 | `set_parashape()` | (pset) | bool | 문단 모양 적용 | `SetParaShape()` | TODO |
| 9 | `set_para()` | (AlignType, LineSpacing, ...) | bool | 문단 모양 단축 설정 | `SetPara()` | TODO |

## 셀 모양 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 10 | `CellShape` (속성) | getter/setter | Any | 셀 모양 파라미터셋 | `GetCellShape()` / `SetCellShape()` | TODO |

## 스타일 관리 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 11 | `import_style()` | (sty_filepath) | bool | 스타일 파일 가져오기 | `ImportStyle()` | TODO |
| 12 | `ImportStyle()` | (sty_filepath) | bool | 저수준 스타일 가져오기 | `ImportStyle()` | TODO |
| 13 | `delete_style_by_name()` | (src, dst) | bool | 스타일 삭제/교체 | `DeleteStyleByName()` | TODO |

---

## 글자 모양 Run 액션 (HAction.Run)

### 기본 서식

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 1 | `CharShapeBold()` | 굵게 | `CharShapeBold()` | TODO |
| 2 | `CharShapeItalic()` | 기울임 | `CharShapeItalic()` | TODO |
| 3 | `CharShapeUnderline()` | 밑줄 | `CharShapeUnderline()` | TODO |
| 4 | `CharShapeNormal()` | 기본 글자 모양 | `CharShapeNormal()` | TODO |

### 특수 효과

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 5 | `CharShapeOutline()` | 외곽선 | `CharShapeOutline()` | TODO |
| 6 | `CharShapeShadow()` | 그림자 | `CharShapeShadow()` | TODO |
| 7 | `CharShapeEmboss()` | 양각 | `CharShapeEmboss()` | TODO |
| 8 | `CharShapeEngrave()` | 음각 | `CharShapeEngrave()` | TODO |
| 9 | `CharShapeCenterline()` | 취소선 | `CharShapeCenterline()` | TODO |

### 위치/첨자

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 10 | `CharShapeSuperscript()` | 위 첨자 | `CharShapeSuperscript()` | TODO |
| 11 | `CharShapeSubscript()` | 아래 첨자 | `CharShapeSubscript()` | TODO |
| 12 | `CharShapeSuperSubscript()` | 위/아래 첨자 전환 | `CharShapeSuperSubscript()` | TODO |

### 글자 크기

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 13 | `CharShapeHeight()` | 글자 크기 | `CharShapeHeight()` | TODO |
| 14 | `CharShapeHeightIncrease()` | 글자 크기 증가 | `CharShapeHeightIncrease()` | TODO |
| 15 | `CharShapeHeightDecrease()` | 글자 크기 감소 | `CharShapeHeightDecrease()` | TODO |

### 글자 간격/폭

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 16 | `CharShapeSpacing()` | 자간 | `CharShapeSpacing()` | TODO |
| 17 | `CharShapeSpacingIncrease()` | 자간 증가 | `CharShapeSpacingIncrease()` | TODO |
| 18 | `CharShapeSpacingDecrease()` | 자간 감소 | `CharShapeSpacingDecrease()` | TODO |
| 19 | `CharShapeWidth()` | 장평 | `CharShapeWidth()` | TODO |
| 20 | `CharShapeWidthIncrease()` | 장평 증가 | `CharShapeWidthIncrease()` | TODO |
| 21 | `CharShapeWidthDecrease()` | 장평 감소 | `CharShapeWidthDecrease()` | TODO |

### 글자색

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 22 | `CharShapeTextColorBlack()` | 검정색 | `CharShapeTextColorBlack()` | TODO |
| 23 | `CharShapeTextColorWhite()` | 흰색 | `CharShapeTextColorWhite()` | TODO |
| 24 | `CharShapeTextColorRed()` | 빨간색 | `CharShapeTextColorRed()` | TODO |
| 25 | `CharShapeTextColorBlue()` | 파란색 | `CharShapeTextColorBlue()` | TODO |
| 26 | `CharShapeTextColorGreen()` | 초록색 | `CharShapeTextColorGreen()` | TODO |
| 27 | `CharShapeTextColorYellow()` | 노란색 | `CharShapeTextColorYellow()` | TODO |
| 28 | `CharShapeTextColorViolet()` | 보라색 | `CharShapeTextColorViolet()` | TODO |
| 29 | `CharShapeTextColorBluish()` | 청록색 | `CharShapeTextColorBluish()` | TODO |

### 글꼴

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 30 | `CharShapeTypeface()` | 글꼴 선택 | `CharShapeTypeface()` | TODO |
| 31 | `CharShapeNextFaceName()` | 다음 글꼴 | `CharShapeNextFaceName()` | TODO |
| 32 | `CharShapePrevFaceName()` | 이전 글꼴 | `CharShapePrevFaceName()` | TODO |
| 33 | `CharShapeLang()` | 언어 설정 | `CharShapeLang()` | TODO |

---

## 문단 모양 Run 액션 (HAction.Run)

### 정렬

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 34 | `ParagraphShapeAlignLeft()` | 왼쪽 정렬 | `ParagraphShapeAlignLeft()` | TODO |
| 35 | `ParagraphShapeAlignCenter()` | 가운데 정렬 | `ParagraphShapeAlignCenter()` | TODO |
| 36 | `ParagraphShapeAlignRight()` | 오른쪽 정렬 | `ParagraphShapeAlignRight()` | TODO |
| 37 | `ParagraphShapeAlignJustify()` | 양쪽 정렬 | `ParagraphShapeAlignJustify()` | TODO |
| 38 | `ParagraphShapeAlignDistribute()` | 배분 정렬 | `ParagraphShapeAlignDistribute()` | TODO |
| 39 | `ParagraphShapeAlignDivision()` | 나눔 정렬 | `ParagraphShapeAlignDivision()` | TODO |

### 여백

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 40 | `ParagraphShapeIncreaseMargin()` | 여백 증가 | `ParagraphShapeIncreaseMargin()` | TODO |
| 41 | `ParagraphShapeDecreaseMargin()` | 여백 감소 | `ParagraphShapeDecreaseMargin()` | TODO |
| 42 | `ParagraphShapeIncreaseLeftMargin()` | 왼쪽 여백 증가 | `ParagraphShapeIncreaseLeftMargin()` | TODO |
| 43 | `ParagraphShapeDecreaseLeftMargin()` | 왼쪽 여백 감소 | `ParagraphShapeDecreaseLeftMargin()` | TODO |
| 44 | `ParagraphShapeIncreaseRightMargin()` | 오른쪽 여백 증가 | `ParagraphShapeIncreaseRightMargin()` | TODO |
| 45 | `ParagraphShapeDecreaseRightMargin()` | 오른쪽 여백 감소 | `ParagraphShapeDecreaseRightMargin()` | TODO |

### 줄 간격

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 46 | `ParagraphShapeIncreaseLineSpacing()` | 줄 간격 증가 | `ParagraphShapeIncreaseLineSpacing()` | TODO |
| 47 | `ParagraphShapeDecreaseLineSpacing()` | 줄 간격 감소 | `ParagraphShapeDecreaseLineSpacing()` | TODO |

### 들여쓰기

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 48 | `ParagraphShapeIndentAtCaret()` | 캐럿 위치에서 들여쓰기 | `ParagraphShapeIndentAtCaret()` | TODO |
| 49 | `ParagraphShapeIndentPositive()` | 들여쓰기 | `ParagraphShapeIndentPositive()` | TODO |
| 50 | `ParagraphShapeIndentNegative()` | 내어쓰기 | `ParagraphShapeIndentNegative()` | TODO |

---

## 상세 설명

### CharShape (속성)
글자 모양 파라미터셋에 접근한다.

```python
@property
def CharShape(self) -> "Hwp.CharShape":
    return self.hwp.CharShape

@CharShape.setter
def CharShape(self, prop: Any) -> None:
    self.hwp.CharShape = prop
```

**사용 예:**
```python
# 글자 크기 조회
height = hwp.CharShape.Item("Height")
print(hwp.HwpUnitToPoint(height))  # 포인트 단위로 출력

# 글자 모양 변경
prop = hwp.CharShape
prop.SetItem("Height", hwp.PointToHwpUnit(20))
hwp.CharShape = prop
```

### get_charshape() / set_charshape()
글자 모양을 조회하고 적용한다.

```python
def get_charshape(self):
    """현재 캐럿의 글자 모양 파라미터셋 반환"""
    pset = self.hwp.HParameterSet.HCharShape
    self.hwp.HAction.GetDefault("CharShape", pset.HSet)
    return pset

def set_charshape(self, pset):
    """
    Args:
        pset: HCharShape 파라미터셋 또는 dict
    """
    if isinstance(pset, dict):
        new_pset = self.hwp.HParameterSet.HCharShape
        for key in pset.keys():
            new_pset.__setattr__(key, pset[key])
    else:
        new_pset = pset
    return self.hwp.HAction.Execute("CharShape", new_pset.HSet)
```

**사용 패턴:**
```python
# 패턴 1: 파라미터셋 사용
pset = hwp.get_charshape()
pset.Height = hwp.PointToHwpUnit(14)
pset.Bold = True
hwp.SelectAll()
hwp.set_charshape(pset)

# 패턴 2: 딕셔너리 사용
char_dict = hwp.get_charshape_as_dict()
char_dict["Height"] = hwp.PointToHwpUnit(12)
hwp.set_charshape(char_dict)
```

### ParaShape (속성)
문단 모양 파라미터셋에 접근한다.

```python
@property
def ParaShape(self):
    return self.hwp.ParaShape

@ParaShape.setter
def ParaShape(self, prop):
    self.hwp.ParaShape = prop
```

**사용 예:**
```python
# 줄간격 조회
linespacing = hwp.ParaShape.Item("LineSpacing")
print(f"줄간격: {linespacing}%")

# 줄간격 200%로 변경
hwp.SelectAll()
prop = hwp.ParaShape
prop.SetItem("LineSpacing", 200)
hwp.ParaShape = prop
```

### set_para()
문단 모양을 한 번에 설정하는 단축 메서드

```python
def set_para(self,
    AlignType: Optional[Literal["Justify", "Left", "Center", "Right", "Distribute", "DistributeSpace"]] = None,
    BreakNonLatinWord: Optional[Literal[0, 1]] = None,
    LineSpacing: Optional[int] = None,
    Condense: Optional[int] = None,
    SnapToGrid: Optional[Literal[0, 1]] = None,
    NextSpacing: Optional[float] = None,
    PrevSpacing: Optional[float] = None,
    Indentation: Optional[float] = None,
    RightMargin: Optional[float] = None,
    LeftMargin: Optional[float] = None,
    PagebreakBefore: Optional[Literal[0, 1]] = None,
    KeepLinesTogether: Optional[Literal[0, 1]] = None,
    KeepWithNext: Optional[Literal[0, 1]] = None,
    WidowOrphan: Optional[Literal[0, 1]] = None,
    AutoSpaceEAsianNum: Optional[Literal[0, 1]] = None,
    AutoSpaceEAsianEng: Optional[Literal[0, 1]] = None,
    LineWrap: Optional[Literal[0, 1]] = None,
    FontLineHeight: Optional[Literal[0, 1]] = None,
    TextAlignment: Optional[Literal[0, 1, 2, 3]] = None
) -> bool:
    """
    Args:
        AlignType: 정렬 방식
        LineSpacing: 줄 간격 (%)
        NextSpacing: 문단 아래 간격
        PrevSpacing: 문단 위 간격
        Indentation: 들여쓰기
        LeftMargin: 왼쪽 여백
        RightMargin: 오른쪽 여백
        # ... 기타 옵션
    """
```

### import_style()
스타일 파일(.sty)을 가져온다.

```python
def import_style(self, sty_filepath: str) -> bool:
    """
    Args:
        sty_filepath: .sty 파일 경로

    Returns:
        성공시 True

    Examples:
        >>> hwp.import_style("C:/styles/custom.sty")
    """
    style_set = self.hwp.HParameterSet.HStyleTemplate
    style_set.filename = sty_filepath
    return self.hwp.ImportStyle(style_set.HSet)
```

---

## C++ 클래스 구조

```cpp
// Hwp.h - 스타일/서식 메서드
class Hwp : public ParamHelpers, public RunMethods {
public:
    // 글자 모양
    IDispatch* GetCharShape();
    void SetCharShape(IDispatch* prop);
    std::map<std::wstring, VARIANT> GetCharShapeAsDict();
    bool SetCharShape(IDispatch* pset);
    bool SetCharShape(const std::map<std::wstring, VARIANT>& dict);

    // 문단 모양
    IDispatch* GetParaShape();
    void SetParaShape(IDispatch* prop);
    std::map<std::wstring, VARIANT> GetParaShapeAsDict();
    bool SetParaShape(IDispatch* pset);
    bool SetPara(const std::wstring& alignType = L"",
                 int lineSpacing = -1,
                 double leftMargin = -1,
                 double rightMargin = -1,
                 double indentation = -1);

    // 셀 모양
    IDispatch* GetCellShape();
    void SetCellShape(IDispatch* prop);

    // 스타일 관리
    bool ImportStyle(const std::wstring& styFilepath);
    bool DeleteStyleByName(const std::wstring& src, const std::wstring& dst);

    // Run 액션 - 글자 모양
    bool CharShapeBold();
    bool CharShapeItalic();
    bool CharShapeUnderline();
    bool CharShapeNormal();
    bool CharShapeOutline();
    bool CharShapeShadow();
    bool CharShapeEmboss();
    bool CharShapeEngrave();
    bool CharShapeSuperscript();
    bool CharShapeSubscript();
    bool CharShapeHeightIncrease();
    bool CharShapeHeightDecrease();
    bool CharShapeSpacingIncrease();
    bool CharShapeSpacingDecrease();
    bool CharShapeTextColorBlack();
    bool CharShapeTextColorRed();
    bool CharShapeTextColorBlue();
    // ... 기타 색상

    // Run 액션 - 문단 모양
    bool ParagraphShapeAlignLeft();
    bool ParagraphShapeAlignCenter();
    bool ParagraphShapeAlignRight();
    bool ParagraphShapeAlignJustify();
    bool ParagraphShapeAlignDistribute();
    bool ParagraphShapeIncreaseMargin();
    bool ParagraphShapeDecreaseMargin();
    bool ParagraphShapeIncreaseLineSpacing();
    bool ParagraphShapeDecreaseLineSpacing();
    bool ParagraphShapeIndentPositive();
    bool ParagraphShapeIndentNegative();
};
```

---

## 주요 HCharShape 속성

| 속성 | 타입 | 설명 |
|------|------|------|
| `Height` | int | 글자 크기 (HwpUnit) |
| `Bold` | bool | 굵게 |
| `Italic` | bool | 기울임 |
| `UnderlineType` | int | 밑줄 종류 |
| `OutlineType` | int | 외곽선 |
| `Shadow` | int | 그림자 |
| `TextColor` | int | 글자색 (RGBColor) |
| `FaceNameUser` | str | 사용자 정의 글꼴 |
| `FaceNameSymbol` | str | 기호 글꼴 |
| `FaceNameOther` | str | 기타 글꼴 |
| `FaceNameJapanese` | str | 일본어 글꼴 |
| `FaceNameHanja` | str | 한자 글꼴 |
| `FaceNameLatin` | str | 영어 글꼴 |
| `FaceNameHangul` | str | 한글 글꼴 |
| `Spacing` | int | 자간 (%) |
| `Ratio` | int | 장평 (%) |

## 주요 HParaShape 속성

| 속성 | 타입 | 설명 |
|------|------|------|
| `Align` | int | 정렬 방식 |
| `LineSpacing` | int | 줄 간격 (%) |
| `LeftMargin` | int | 왼쪽 여백 (HwpUnit) |
| `RightMargin` | int | 오른쪽 여백 (HwpUnit) |
| `Indent` | int | 들여쓰기 (HwpUnit) |
| `PrevSpacing` | int | 문단 위 간격 |
| `NextSpacing` | int | 문단 아래 간격 |
| `TabDef` | int | 탭 정의 |
| `HeadingType` | int | 개요 수준 |

---

## 구현 우선순위

1. **Critical**: CharShape, ParaShape 속성 (getter/setter)
2. **High**: get_charshape(), set_charshape(), get_parashape(), set_parashape()
3. **High**: CharShapeBold(), CharShapeItalic(), CharShapeUnderline()
4. **High**: ParagraphShapeAlign* (정렬 관련)
5. **Medium**: set_para() - 단축 메서드
6. **Medium**: CharShapeHeight*, CharShapeSpacing* (크기/간격)
7. **Low**: CharShapeTextColor* (색상)
8. **Low**: ImportStyle(), DeleteStyleByName()
