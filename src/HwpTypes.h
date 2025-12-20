/**
 * @file HwpTypes.h
 * @brief HWP 자동화를 위한 타입 및 열거형 정의
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 */

#pragma once

// Windows 헤더 설정
#ifndef WIN32_LEAN_AND_MEAN
#define WIN32_LEAN_AND_MEAN
#endif

#ifndef NOMINMAX
#define NOMINMAX
#endif

#include <Windows.h>
#include <string>
#include <vector>
#include <map>
#include <tuple>
#include <optional>

namespace cpyhwpx {

//=============================================================================
// 기본 타입 정의
//=============================================================================

using HwpUnit = int;  // HWP 내부 단위 (1/7200 inch)

//=============================================================================
// 위치 관련
//=============================================================================

struct HwpPos {
    int list;   // 리스트 ID
    int para;   // 문단 번호
    int pos;    // 문단 내 위치
};

//=============================================================================
// 뷰 상태 (ViewState)
//=============================================================================

enum class ViewState : int {
    Normal = 0,         // 일반 보기
    MultiColumn = 1,    // 다단 보기
    MultiPage = 2,      // 다중 페이지
    MasterPage = 3,     // 바탕쪽 보기
    Page = 4,           // 쪽 보기
    Text = 5,           // 텍스트 보기
    Outline = 6         // 개요 보기
};

//=============================================================================
// 이동 ID (MovePos)
//=============================================================================

enum class MoveID : int {
    None = 0,           // 이동 없음
    Top = 1,            // 문서 처음
    Bottom = 2,         // 문서 끝
    ParaBegin = 3,      // 문단 처음
    ParaEnd = 4,        // 문단 끝
    LineBegin = 5,      // 줄 처음
    LineEnd = 6,        // 줄 끝
    WordBegin = 7,      // 단어 처음
    WordEnd = 8         // 단어 끝
};

//=============================================================================
// 컨트롤 타입 (CtrlID)
//=============================================================================

enum class CtrlType {
    Unknown,
    Table,          // tbl - 표
    Equation,       // eqed - 수식
    Picture,        // pic - 그림
    OLE,            // ole - OLE 개체
    Line,           // lin - 선
    Rectangle,      // rec - 사각형
    Ellipse,        // ell - 타원
    Arc,            // arc - 호
    Polygon,        // pol - 다각형
    Curve,          // cur - 곡선
    TextBox,        // $txt - 글상자
    Container,      // $con - 컨테이너
    GenShapeObj,    // gso - 일반 그리기 개체
    Header,         // head - 머리말
    Footer,         // foot - 꼬리말
    Footnote,       // fn - 각주
    Endnote,        // en - 미주
    AutoNum,        // atno - 자동 번호
    PageNum,        // pgno - 쪽 번호
    Field,          // %fld - 필드
    Bookmark,       // bokm - 책갈피
    Button,         // btn - 버튼
    RadioButton,    // rdo - 라디오 버튼
    CheckButton,    // chk - 체크 버튼
    ComboBox,       // cbo - 콤보박스
    Edit,           // edt - 편집상자
    ListBox,        // lst - 리스트박스
    ScrollBar,      // scr - 스크롤바
    Video,          // vid - 비디오
    ColumnDef       // cold - 단 정의
};

//=============================================================================
// 파일 형식
//=============================================================================

enum class FileFormat {
    HWP,            // 한글 문서 (.hwp)
    HWPX,           // 한글 문서 OOXML (.hwpx)
    HWT,            // 한글 서식 (.hwt)
    HTML,           // HTML
    MHTML,          // MHTML
    TXT,            // 텍스트
    UnicodeText,    // 유니코드 텍스트 (UNICODE는 Windows 매크로와 충돌)
    RTF,            // RTF
    DOC,            // MS Word 문서
    DOCX,           // MS Word 2007+
    ODT,            // OpenDocument Text
    PDF             // PDF
};

//=============================================================================
// 선 스타일
//=============================================================================

enum class LineStyle : int {
    None = 0,           // 없음
    Solid = 1,          // 실선
    Dash = 2,           // 파선
    Dot = 3,            // 점선
    DashDot = 4,        // 일점쇄선
    DashDotDot = 5,     // 이점쇄선
    LongDash = 6,       // 긴 파선
    Circle = 7,         // 원형 점선
    DoubleSlim = 8,     // 이중선 (가는선)
    SlimThick = 9,      // 이중선 (가는-굵은)
    ThickSlim = 10,     // 이중선 (굵은-가는)
    SlimThickSlim = 11  // 삼중선
};

//=============================================================================
// 정렬 방식
//=============================================================================

enum class HAlign : int {
    Left = 0,           // 왼쪽 정렬
    Center = 1,         // 가운데 정렬
    Right = 2,          // 오른쪽 정렬
    Justify = 3,        // 양쪽 정렬
    Distribute = 4,     // 배분 정렬
    Division = 5        // 나눔 정렬
};

enum class VAlign : int {
    Top = 0,            // 위
    Center = 1,         // 가운데
    Bottom = 2          // 아래
};

//=============================================================================
// 글자 모양 속성
//=============================================================================

struct CharShape {
    std::wstring FaceNameHangul;
    std::wstring FaceNameLatin;
    std::wstring FaceNameHanja;
    std::wstring FaceNameJapanese;
    std::wstring FaceNameOther;
    std::wstring FaceNameSymbol;
    std::wstring FaceNameUser;

    int FontTypeHangul = 2;     // TTF
    int FontTypeLatin = 2;
    int FontTypeHanja = 2;
    int FontTypeJapanese = 2;
    int FontTypeOther = 2;
    int FontTypeSymbol = 2;
    int FontTypeUser = 2;

    int Height = 1000;          // 크기 (1/100 pt)
    int TextColor = 0;          // 글자색 (RGB)
    int ShadeColor = -1;        // 음영색
    int Spacing = 0;            // 자간 (%)
    int Ratio = 100;            // 장평 (%)
    int Offset = 0;             // 위치 조정

    bool Bold = false;          // 진하게
    bool Italic = false;        // 이탤릭
    bool Underline = false;     // 밑줄
    bool Strikeout = false;     // 취소선
    bool Outline = false;       // 외곽선
    bool Shadow = false;        // 그림자
    bool Emboss = false;        // 양각
    bool Engrave = false;       // 음각
    bool Superscript = false;   // 위첨자
    bool Subscript = false;     // 아래첨자
};

//=============================================================================
// 문단 모양 속성
//=============================================================================

struct ParaShape {
    HAlign Align = HAlign::Justify;     // 정렬
    int LeftMargin = 0;                 // 왼쪽 여백 (HwpUnit)
    int RightMargin = 0;                // 오른쪽 여백
    int Indent = 0;                     // 들여쓰기
    int LineSpacing = 160;              // 줄 간격 (%)
    int SpaceBefore = 0;                // 문단 위 간격
    int SpaceAfter = 0;                 // 문단 아래 간격
    int LineSpacingType = 0;            // 줄 간격 타입
    bool KeepWithNext = false;          // 다음 문단과 함께
    bool KeepLines = false;             // 문단 보호
    bool PageBreakBefore = false;       // 문단 앞에서 나눔
    bool WidowOrphan = false;           // 외톨이줄 방지
};

//=============================================================================
// 폰트 프리셋
//=============================================================================

struct FontPreset {
    std::wstring FaceNameHangul;
    std::wstring FaceNameLatin;
    std::wstring FaceNameHanja;
    std::wstring FaceNameJapanese;
    std::wstring FaceNameOther;
    std::wstring FaceNameSymbol;
    std::wstring FaceNameUser;
    int FontTypeHangul = 2;
    int FontTypeHanja = 2;
    int FontTypeJapanese = 2;
    int FontTypeLatin = 2;
    int FontTypeOther = 2;
    int FontTypeSymbol = 2;
    int FontTypeUser = 2;
};

//=============================================================================
// 색상 유틸리티
//=============================================================================

inline COLORREF RGB_TO_COLORREF(int r, int g, int b) {
    return RGB(r, g, b);
}

inline std::tuple<int, int, int> COLORREF_TO_RGB(COLORREF color) {
    return std::make_tuple(GetRValue(color), GetGValue(color), GetBValue(color));
}

//=============================================================================
// 단위 변환
//=============================================================================

namespace Units {
    // 1 inch = 25.4 mm = 72 pt = 7200 HwpUnit

    constexpr double HWPUNIT_PER_INCH = 7200.0;
    constexpr double HWPUNIT_PER_MM = HWPUNIT_PER_INCH / 25.4;
    constexpr double HWPUNIT_PER_PT = 100.0;
    constexpr double HWPUNIT_PER_CM = HWPUNIT_PER_MM * 10.0;

    inline HwpUnit FromMM(double mm) {
        return static_cast<HwpUnit>(mm * HWPUNIT_PER_MM);
    }

    inline HwpUnit FromInch(double inch) {
        return static_cast<HwpUnit>(inch * HWPUNIT_PER_INCH);
    }

    inline HwpUnit FromPoint(double pt) {
        return static_cast<HwpUnit>(pt * HWPUNIT_PER_PT);
    }

    inline HwpUnit FromCM(double cm) {
        return static_cast<HwpUnit>(cm * HWPUNIT_PER_CM);
    }

    inline double ToMM(HwpUnit unit) {
        return unit / HWPUNIT_PER_MM;
    }

    inline double ToInch(HwpUnit unit) {
        return unit / HWPUNIT_PER_INCH;
    }

    inline double ToPoint(HwpUnit unit) {
        return unit / HWPUNIT_PER_PT;
    }

    inline double ToCM(HwpUnit unit) {
        return unit / HWPUNIT_PER_CM;
    }
}

//=============================================================================
// 메시지 박스 모드
//=============================================================================

namespace MsgBoxMode {
    constexpr int ShowAll = 0xF0000;        // 모든 다이얼로그 표시 (기본)
    constexpr int AutoOK = 0x10000;         // 확인 자동 클릭
    constexpr int AutoCancel = 0x20000;     // 취소 자동 클릭
    constexpr int AutoYes = 0x1000;         // 예 자동 클릭
    constexpr int AutoNo = 0x2000;          // 아니오 자동 클릭
}

//=============================================================================
// 붙여넣기 옵션
//=============================================================================

enum class PasteOption : int {
    Default = 0,        // 기본
    Special = 1,        // 특수
    NoFormat = 2,       // 서식 없이
    HTML = 3,           // HTML
    TextOnly = 4,       // 텍스트만
    Picture = 5,        // 그림으로
    OLE = 6             // OLE
};

} // namespace cpyhwpx
