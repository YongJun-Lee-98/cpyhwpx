# 파라미터 헬퍼 매핑

## 요약
- 총 함수: 90개
- 포팅 완료: 0개 (0%)
- 우선순위: Medium
- 소스 파일: param_helpers.py (510줄)

---

## ParamHelpers 클래스 (param_helpers.py:4~510)

파라미터 변환 헬퍼 메서드 모음. 문자열을 정수 상수로 변환하는 역할.

### 메서드 매핑 테이블

| # | Python 함수 | 파라미터 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 1 | `BorderShape()` | (border_type) | int | 테두리 모양 | `BorderShape()` | TODO |
| 2 | `ArcType()` | (arc_type) | int | 호 타입 | `ArcType()` | TODO |
| 3 | `AutoNumType()` | (autonum) | int | 자동 번호 타입 | `AutoNumType()` | TODO |
| 4 | `BreakWordLatin()` | (break_latin_word) | int | 영문 단어 나누기 | `BreakWordLatin()` | TODO |
| 5 | `BrushType()` | (brush_type) | int | 브러시 타입 | `BrushType()` | TODO |
| 6 | `Canonical()` | (canonical) | int | 정규화 | `Canonical()` | TODO |
| 7 | `CellApply()` | (cell_apply) | int | 셀 적용 | `CellApply()` | TODO |
| 8 | `CharShadowType()` | (shadow_type) | int | 글자 그림자 타입 | `CharShadowType()` | TODO |
| 9 | `ColDefType()` | (col_def_type) | int | 단 정의 타입 | `ColDefType()` | TODO |
| 10 | `ColLayoutType()` | (col_layout_type) | int | 단 레이아웃 타입 | `ColLayoutType()` | TODO |
| 11 | `ConvertPUAHangulToUnicode()` | (reverse) | int | PUA 한글→유니코드 | `ConvertPUAHangulToUnicode()` | TODO |
| 12 | `CrookedSlash()` | (crooked_slash) | int | 꺾인 슬래시 | `CrookedSlash()` | TODO |
| 13 | `DSMark()` | (diac_sym_mark) | int | 발음 구별 기호 | `DSMark()` | TODO |
| 14 | `DbfCodeType()` | (dbf_code) | int | DBF 코드 타입 | `DbfCodeType()` | TODO |
| 15 | `Delimiter()` | (delimiter) | int | 구분자 | `Delimiter()` | TODO |
| 16 | `DrawAspect()` | (draw_aspect) | int | 그리기 화면비 | `DrawAspect()` | TODO |
| 17 | `DrawFillImage()` | (fillimage) | int | 그리기 채우기 이미지 | `DrawFillImage()` | TODO |
| 18 | `DrawShadowType()` | (shadow_type) | int | 그리기 그림자 타입 | `DrawShadowType()` | TODO |
| 19 | `Encrypt()` | (encrypt) | int | 암호화 | `Encrypt()` | TODO |
| 20 | `EndSize()` | (end_size) | int | 끝 크기 | `EndSize()` | TODO |
| 21 | `EndStyle()` | (end_style) | int | 끝 스타일 | `EndStyle()` | TODO |
| 22 | `FillAreaType()` | (fill_area) | int | 채우기 영역 타입 | `FillAreaType()` | TODO |
| 23 | `FindDir()` | (find_dir) | int | 찾기 방향 | `FindDir()` | TODO |
| 24 | `FontType()` | (font_type) | int | 폰트 타입 | `FontType()` | TODO |
| 25 | `GetTranslateLangList()` | (cur_lang) | list | 번역 언어 목록 | `GetTranslateLangList()` | TODO |
| 26 | `GetUserInfo()` | (user_info_id) | str | 사용자 정보 조회 | `GetUserInfo()` | TODO |
| 27 | `Gradation()` | (gradation) | int | 그라데이션 | `Gradation()` | TODO |
| 28 | `GridMethod()` | (grid_method) | int | 격자 방식 | `GridMethod()` | TODO |
| 29 | `GridViewLine()` | (grid_view_line) | int | 격자선 보기 | `GridViewLine()` | TODO |
| 30 | `GutterMethod()` | (gutter_type) | int | 제본 여백 방식 | `GutterMethod()` | TODO |
| 31 | `HAlign()` | (h_align) | int | 가로 정렬 | `HAlign()` | TODO |
| 32 | `Handler()` | (handler) | int | 핸들러 | `Handler()` | TODO |
| 33 | `Hash()` | (hash) | int | 해시 | `Hash()` | TODO |
| 34 | `HatchStyle()` | (hatch_style) | int | 빗금 스타일 | `HatchStyle()` | TODO |
| 35 | `HeadType()` | (heading_type) | int | 문단 종류 | `HeadType()` | TODO |
| 36 | `HeightRel()` | (height_rel) | int | 상대 높이 | `HeightRel()` | TODO |
| 37 | `Hiding()` | (hiding) | int | 숨기기 | `Hiding()` | TODO |
| 38 | `HorzRel()` | (horz_rel) | int | 가로 기준 | `HorzRel()` | TODO |
| 39 | `HwpLineType()` | (line_type) | int | 선 타입 | `HwpLineType()` | TODO |
| 40 | `HwpLineWidth()` | (line_width) | int | 선 굵기 | `HwpLineWidth()` | TODO |
| 41 | `HwpOutlineStyle()` | (hwp_outline_style) | int | 개요 스타일 | `HwpOutlineStyle()` | TODO |
| 42 | `HwpOutlineType()` | (hwp_outline_type) | int | 개요 타입 | `HwpOutlineType()` | TODO |
| 43 | `HwpUnderlineShape()` | (hwp_underline_shape) | int | 밑줄 모양 | `HwpUnderlineShape()` | TODO |
| 44 | `HwpUnderlineType()` | (hwp_underline_type) | int | 밑줄 타입 | `HwpUnderlineType()` | TODO |
| 45 | `HwpZoomType()` | (zoom_type) | int | 확대/축소 타입 | `HwpZoomType()` | TODO |
| 46 | `ImageFormat()` | (image_format) | int | 이미지 형식 | `ImageFormat()` | TODO |
| 47 | `LineSpacingMethod()` | (line_spacing) | int | 줄간격 방식 | `LineSpacingMethod()` | TODO |
| 48 | `LineWrapType()` | (line_wrap) | int | 줄바꿈 타입 | `LineWrapType()` | TODO |
| 49 | `LunarToSolar()` | (l_year, l_month, l_day, l_leap, s_year, s_month, s_day) | int | 음력→양력 | `LunarToSolar()` | TODO |
| 50 | `LunarToSolarBySet()` | (l_year, l_month, l_day, l_leap) | ParameterSet | 음력→양력(셋) | `LunarToSolarBySet()` | TODO |
| 51 | `MacroState()` | (macro_state) | int | 매크로 상태 | `MacroState()` | TODO |
| 52 | `MailType()` | (mail_type) | int | 메일 타입 | `MailType()` | TODO |
| 53 | `mili_to_hwp_unit()` | (mili: float) | int | mm → HwpUnit | `MiliToHwpUnit()` | TODO |
| 54 | `MiliToHwpUnit()` | (mili: float) | int | mm → HwpUnit | `MiliToHwpUnit()` | TODO |
| 55 | `hwp_unit_to_mili()` | (hwp_unit: int) | float | HwpUnit → mm (정적) | `HwpUnitToMili()` | TODO |
| 56 | `HwpUnitToMili()` | (hwp_unit: int) | float | HwpUnit → mm | `HwpUnitToMili()` | TODO |
| 57 | `NumberFormat()` | (num_format) | int | 번호 형식 | `NumberFormat()` | TODO |
| 58 | `Numbering()` | (numbering) | int | 번호 매기기 | `Numbering()` | TODO |
| 59 | `PageNumPosition()` | (pagenumpos) | int | 쪽번호 위치 | `PageNumPosition()` | TODO |
| 60 | `PageType()` | (page_type) | int | 페이지 타입 | `PageType()` | TODO |
| 61 | `ParaHeadAlign()` | (para_head_align) | int | 문단 머리 정렬 | `ParaHeadAlign()` | TODO |
| 62 | `PicEffect()` | (pic_effect) | int | 그림 효과 | `PicEffect()` | TODO |
| 63 | `PlacementType()` | (restart) | int | 배치 타입 | `PlacementType()` | TODO |
| 64 | `PresentEffect()` | (prsnteffect) | int | 프레젠테이션 효과 | `PresentEffect()` | TODO |
| 65 | `PrintDevice()` | (print_device) | int | 프린터 장치 | `PrintDevice()` | TODO |
| 66 | `PrintPaper()` | (print_paper) | int | 인쇄 용지 | `PrintPaper()` | TODO |
| 67 | `PrintRange()` | (print_range) | int | 인쇄 범위 | `PrintRange()` | TODO |
| 68 | `PrintType()` | (print_method) | int | 인쇄 타입 | `PrintType()` | TODO |
| 69 | `SetCurMetatagName()` | (tag) | bool | 메타태그 설정 | `SetCurMetatagName()` | TODO |
| 70 | `SetDRMAuthority()` | (authority) | bool | DRM 권한 설정 | `SetDRMAuthority()` | TODO |
| 71 | `SetUserInfo()` | (user_info_id, value) | bool | 사용자 정보 설정 | `SetUserInfo()` | TODO |
| 72 | `SideType()` | (side_type) | int | 면 타입 | `SideType()` | TODO |
| 73 | `Signature()` | (signature) | int | 서명 | `Signature()` | TODO |
| 74 | `Slash()` | (slash) | int | 슬래시 | `Slash()` | TODO |
| 75 | `SolarToLunar()` | (s_year, s_month, s_day, l_year, l_month, l_day, l_leap) | int | 양력→음력 | `SolarToLunar()` | TODO |
| 76 | `SolarToLunarBySet()` | (s_year, s_month, s_day) | ParameterSet | 양력→음력(셋) | `SolarToLunarBySet()` | TODO |
| 77 | `SortDelimiter()` | (sort_delimiter) | int | 정렬 구분자 | `SortDelimiter()` | TODO |
| 78 | `StrikeOut()` | (strike_out_type) | int | 취소선 | `StrikeOut()` | TODO |
| 79 | `StyleType()` | (style_type) | int | 스타일 타입 | `StyleType()` | TODO |
| 80 | `SubtPos()` | (subt_pos) | int | 자막 위치 | `SubtPos()` | TODO |
| 81 | `TableBreak()` | (page_break) | int | 표 나누기 | `TableBreak()` | TODO |
| 82 | `TableFormat()` | (table_format) | int | 표 형식 | `TableFormat()` | TODO |
| 83 | `TableSwapType()` | (tableswap) | int | 표 스왑 타입 | `TableSwapType()` | TODO |
| 84 | `TableTarget()` | (table_target) | int | 표 대상 | `TableTarget()` | TODO |
| 85 | `TextAlign()` | (text_align) | int | 텍스트 정렬 | `TextAlign()` | TODO |
| 86 | `TextArtAlign()` | (text_art_align) | int | 글맵시 정렬 | `TextArtAlign()` | TODO |
| 87 | `TextDir()` | (text_direction) | int | 텍스트 방향 | `TextDir()` | TODO |
| 88 | `TextFlowType()` | (text_flow) | int | 텍스트 흐름 | `TextFlowType()` | TODO |
| 89 | `TextWrapType()` | (text_wrap) | int | 텍스트 감싸기 | `TextWrapType()` | TODO |
| 90 | `VAlign()` | (v_align) | int | 세로 정렬 | `VAlign()` | TODO |
| 91 | `VertRel()` | (vert_rel) | int | 세로 기준 | `VertRel()` | TODO |
| 92 | `ViewFlag()` | (view_flag) | int | 보기 플래그 | `ViewFlag()` | TODO |
| 93 | `WatermarkBrush()` | (watermark_brush) | int | 워터마크 브러시 | `WatermarkBrush()` | TODO |
| 94 | `WidthRel()` | (width_rel) | int | 상대 너비 | `WidthRel()` | TODO |

---

## 상세 설명 (자주 사용되는 헬퍼)

### HwpLineType()
선 타입을 문자열에서 정수로 변환

```python
def HwpLineType(self, line_type: Literal[
    "None",        # 0: 없음
    "Solid",       # 1: 실선
    "Dash",        # 2: 파선
    "Dot",         # 3: 점선
    "DashDot",     # 4: 일점쇄선
    "DashDotDot",  # 5: 이점쇄선
    "LongDash",    # 6: 긴 파선
    "Circle",      # 7: 원형 점선
    "DoubleSlim",  # 8: 이중 실선
    "SlimThick",   # 9: 얇고 굵은 이중선
    "ThickSlim",   # 10: 굵고 얇은 이중선
    "SlimThickSlim",  # 11: 얇고 굵고 얇은 삼중선
] = "Solid") -> int
```

**C++ 구현:**
```cpp
int ParamHelpers::HwpLineType(const std::wstring& lineType) {
    static const std::map<std::wstring, int> lineTypes = {
        {L"None", 0}, {L"Solid", 1}, {L"Dash", 2}, {L"Dot", 3},
        {L"DashDot", 4}, {L"DashDotDot", 5}, {L"LongDash", 6},
        {L"Circle", 7}, {L"DoubleSlim", 8}, {L"SlimThick", 9},
        {L"ThickSlim", 10}, {L"SlimThickSlim", 11}
    };
    auto it = lineTypes.find(lineType);
    return (it != lineTypes.end()) ? it->second : 1;  // 기본값: Solid
}
```

### HwpLineWidth()
선 굵기를 문자열에서 정수로 변환

```python
def HwpLineWidth(self, line_width: Literal[
    "0.1mm",   # 0
    "0.12mm",  # 1
    "0.15mm",  # 2
    "0.2mm",   # 3
    "0.25mm",  # 4
    "0.3mm",   # 5
    "0.4mm",   # 6
    "0.5mm",   # 7
    "0.6mm",   # 8
    "0.7mm",   # 9
    "1.0mm",   # 10
    "1.5mm",   # 11
    "2.0mm",   # 12
    "3.0mm",   # 13
    "4.0mm",   # 14
    "5.0mm",   # 15
] = "0.1mm") -> int
```

### HeadType()
문단 종류 (개요, 번호, 글머리표)

```python
def HeadType(self, heading_type: Literal[
    "None",     # 없음 (보통 문단)
    "Outline",  # 개요 문단
    "Number",   # 번호 문단
    "Bullet",   # 글머리표 문단
]) -> int
```

### NumberFormat()
개요 번호 형식

```python
def NumberFormat(self, num_format: Literal[
    "Digit",              # 123
    "CircledDigit",       # ①
    "RomanCapital",       # I
    "RomanSmall",         # i
    "LatinCapital",       # A
    "LatinSmall",         # a
    "CircledLatinCapital", # Ⓐ
    "CircledLatinSmall",  # ⓐ
    "HangulSyllable",     # 가나다
    "CircledHangulSyllable", # ㉯
    "HangulJamo",         # ㄱㄴㄷ
    "CircledHangulJamo",  # ㉠
    "HangulPhonetic",     # 일이삼
    "Ideograph",          # 一
    "CircledIdeograph",   # ㊀
    "DecagonCircle",      # 갑을병
    "DecagonCircleHanja", # 甲
]) -> int
```

### PicEffect()
그림 효과

```python
def PicEffect(self, pic_effect: Literal[
    "RealPic",    # 효과 없음
    "GrayScale",  # 회색조
    "BlackWhite", # 흑백
] = "RealPic") -> int
```

### FindDir()
찾기 방향

```python
def FindDir(self, find_dir: Literal[
    "Forward",   # 앞으로
    "Backward",  # 뒤로
    "AllDoc",    # 문서 전체
] = "Forward") -> int
```

### 단위 변환 함수

```python
# mm → HwpUnit (1 HwpUnit = 1/7200 인치)
def MiliToHwpUnit(self, mili: float) -> int:
    return self.hwp.MiliToHwpUnit(mili=mili)

# HwpUnit → mm (정적 메서드)
@staticmethod
def hwp_unit_to_mili(hwp_unit: int) -> float:
    if hwp_unit == 0:
        return 0
    return round(hwp_unit / 7200 * 25.4, 2)
```

**C++ 구현:**
```cpp
int ParamHelpers::MiliToHwpUnit(double mili) {
    // COM 호출 또는 직접 계산
    return static_cast<int>(mili / 25.4 * 7200);
}

double ParamHelpers::HwpUnitToMili(int hwpUnit) {
    if (hwpUnit == 0) return 0.0;
    return std::round(hwpUnit / 7200.0 * 25.4 * 100) / 100;  // 소수점 2자리
}
```

---

## C++ 클래스 구조

```cpp
// ParamHelpers.h
class ParamHelpers {
public:
    // 선 관련
    int HwpLineType(const std::wstring& lineType);
    int HwpLineWidth(const std::wstring& lineWidth);

    // 정렬 관련
    int HAlign(const std::wstring& hAlign);
    int VAlign(const std::wstring& vAlign);
    int TextAlign(const std::wstring& textAlign);

    // 테이블 관련
    int TableBreak(const std::wstring& pageBreak);
    int TableFormat(const std::wstring& tableFormat);
    int TableTarget(const std::wstring& tableTarget);

    // 문단 관련
    int HeadType(const std::wstring& headingType);
    int NumberFormat(const std::wstring& numFormat);
    int LineSpacingMethod(const std::wstring& lineSpacing);

    // 단위 변환
    int MiliToHwpUnit(double mili);
    static double HwpUnitToMili(int hwpUnit);

    // 날짜 변환
    int LunarToSolar(int lYear, int lMonth, int lDay, int lLeap,
                     int& sYear, int& sMonth, int& sDay);
    int SolarToLunar(int sYear, int sMonth, int sDay,
                     int& lYear, int& lMonth, int& lDay, int& lLeap);

    // 찾기 관련
    int FindDir(const std::wstring& findDir);

    // 그림 관련
    int PicEffect(const std::wstring& picEffect);
    int ImageFormat(const std::wstring& imageFormat);

    // 기타
    int CharShadowType(const std::wstring& shadowType);
    int StrikeOut(const std::wstring& strikeOutType);
    // ... 나머지 90개 메서드

protected:
    IDispatch* m_pHwp;  // Hwp 클래스에서 접근
};
```

---

## 구현 우선순위

1. **High**: 단위 변환 (MiliToHwpUnit, HwpUnitToMili)
2. **High**: 선 관련 (HwpLineType, HwpLineWidth)
3. **High**: 정렬 관련 (HAlign, VAlign, TextAlign)
4. **Medium**: 테이블 관련 (TableBreak, TableFormat, TableTarget)
5. **Medium**: 문단 관련 (HeadType, NumberFormat, LineSpacingMethod)
6. **Low**: 기타 모든 헬퍼
