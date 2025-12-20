/**
 * @file bindings.cpp
 * @brief pybind11 Python 바인딩 정의
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 * C++ 클래스를 Python 모듈로 노출
 */

#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <pybind11/functional.h>

#include "HwpTypes.h"
#include "HwpWrapper.h"
#include "HwpCtrl.h"
#include "HwpAction.h"
#include "HwpParameter.h"
#include "FontDefs.h"
#include "Utils.h"

namespace py = pybind11;

PYBIND11_MODULE(cpyhwpx, m) {
    m.doc() = "cpyhwpx - C++ HWP Automation Library (pyhwpx C++ port)";

    //=========================================================================
    // 열거형 바인딩
    //=========================================================================

    py::enum_<cpyhwpx::ViewState>(m, "ViewState")
        .value("Normal", cpyhwpx::ViewState::Normal)
        .value("MultiColumn", cpyhwpx::ViewState::MultiColumn)
        .value("MultiPage", cpyhwpx::ViewState::MultiPage)
        .value("MasterPage", cpyhwpx::ViewState::MasterPage)
        .value("Page", cpyhwpx::ViewState::Page)
        .value("Text", cpyhwpx::ViewState::Text)
        .value("Outline", cpyhwpx::ViewState::Outline)
        .export_values();

    py::enum_<cpyhwpx::MoveID>(m, "MoveID")
        .value("None_", cpyhwpx::MoveID::None)
        .value("Top", cpyhwpx::MoveID::Top)
        .value("Bottom", cpyhwpx::MoveID::Bottom)
        .value("ParaBegin", cpyhwpx::MoveID::ParaBegin)
        .value("ParaEnd", cpyhwpx::MoveID::ParaEnd)
        .value("LineBegin", cpyhwpx::MoveID::LineBegin)
        .value("LineEnd", cpyhwpx::MoveID::LineEnd)
        .value("WordBegin", cpyhwpx::MoveID::WordBegin)
        .value("WordEnd", cpyhwpx::MoveID::WordEnd)
        .export_values();

    py::enum_<cpyhwpx::CtrlType>(m, "CtrlType")
        .value("Unknown", cpyhwpx::CtrlType::Unknown)
        .value("Table", cpyhwpx::CtrlType::Table)
        .value("Equation", cpyhwpx::CtrlType::Equation)
        .value("Picture", cpyhwpx::CtrlType::Picture)
        .value("OLE", cpyhwpx::CtrlType::OLE)
        .value("Line", cpyhwpx::CtrlType::Line)
        .value("Rectangle", cpyhwpx::CtrlType::Rectangle)
        .value("Ellipse", cpyhwpx::CtrlType::Ellipse)
        .value("Arc", cpyhwpx::CtrlType::Arc)
        .value("Polygon", cpyhwpx::CtrlType::Polygon)
        .value("Curve", cpyhwpx::CtrlType::Curve)
        .value("TextBox", cpyhwpx::CtrlType::TextBox)
        .value("Container", cpyhwpx::CtrlType::Container)
        .value("GenShapeObj", cpyhwpx::CtrlType::GenShapeObj)
        .value("Header", cpyhwpx::CtrlType::Header)
        .value("Footer", cpyhwpx::CtrlType::Footer)
        .value("Footnote", cpyhwpx::CtrlType::Footnote)
        .value("Endnote", cpyhwpx::CtrlType::Endnote)
        .value("AutoNum", cpyhwpx::CtrlType::AutoNum)
        .value("PageNum", cpyhwpx::CtrlType::PageNum)
        .value("Field", cpyhwpx::CtrlType::Field)
        .value("Bookmark", cpyhwpx::CtrlType::Bookmark)
        .export_values();

    py::enum_<cpyhwpx::FileFormat>(m, "FileFormat")
        .value("HWP", cpyhwpx::FileFormat::HWP)
        .value("HWPX", cpyhwpx::FileFormat::HWPX)
        .value("HWT", cpyhwpx::FileFormat::HWT)
        .value("HTML", cpyhwpx::FileFormat::HTML)
        .value("TXT", cpyhwpx::FileFormat::TXT)
        .value("RTF", cpyhwpx::FileFormat::RTF)
        .value("DOC", cpyhwpx::FileFormat::DOC)
        .value("DOCX", cpyhwpx::FileFormat::DOCX)
        .value("PDF", cpyhwpx::FileFormat::PDF)
        .export_values();

    py::enum_<cpyhwpx::LineStyle>(m, "LineStyle")
        .value("None_", cpyhwpx::LineStyle::None)
        .value("Solid", cpyhwpx::LineStyle::Solid)
        .value("Dash", cpyhwpx::LineStyle::Dash)
        .value("Dot", cpyhwpx::LineStyle::Dot)
        .value("DashDot", cpyhwpx::LineStyle::DashDot)
        .value("DashDotDot", cpyhwpx::LineStyle::DashDotDot)
        .export_values();

    py::enum_<cpyhwpx::HAlign>(m, "HAlign")
        .value("Left", cpyhwpx::HAlign::Left)
        .value("Center", cpyhwpx::HAlign::Center)
        .value("Right", cpyhwpx::HAlign::Right)
        .value("Justify", cpyhwpx::HAlign::Justify)
        .value("Distribute", cpyhwpx::HAlign::Distribute)
        .export_values();

    py::enum_<cpyhwpx::VAlign>(m, "VAlign")
        .value("Top", cpyhwpx::VAlign::Top)
        .value("Center", cpyhwpx::VAlign::Center)
        .value("Bottom", cpyhwpx::VAlign::Bottom)
        .export_values();

    //=========================================================================
    // 구조체 바인딩
    //=========================================================================

    py::class_<cpyhwpx::HwpPos>(m, "HwpPos")
        .def(py::init<>())
        .def_readwrite("list", &cpyhwpx::HwpPos::list)
        .def_readwrite("para", &cpyhwpx::HwpPos::para)
        .def_readwrite("pos", &cpyhwpx::HwpPos::pos)
        .def("__repr__", [](const cpyhwpx::HwpPos& p) {
            return "HwpPos(list=" + std::to_string(p.list) +
                   ", para=" + std::to_string(p.para) +
                   ", pos=" + std::to_string(p.pos) + ")";
        });

    py::class_<cpyhwpx::CharShape>(m, "CharShape")
        .def(py::init<>())
        .def_readwrite("FaceNameHangul", &cpyhwpx::CharShape::FaceNameHangul)
        .def_readwrite("FaceNameLatin", &cpyhwpx::CharShape::FaceNameLatin)
        .def_readwrite("Height", &cpyhwpx::CharShape::Height)
        .def_readwrite("TextColor", &cpyhwpx::CharShape::TextColor)
        .def_readwrite("Bold", &cpyhwpx::CharShape::Bold)
        .def_readwrite("Italic", &cpyhwpx::CharShape::Italic)
        .def_readwrite("Underline", &cpyhwpx::CharShape::Underline)
        .def_readwrite("Strikeout", &cpyhwpx::CharShape::Strikeout);

    py::class_<cpyhwpx::ParaShape>(m, "ParaShape")
        .def(py::init<>())
        .def_readwrite("Align", &cpyhwpx::ParaShape::Align)
        .def_readwrite("LeftMargin", &cpyhwpx::ParaShape::LeftMargin)
        .def_readwrite("RightMargin", &cpyhwpx::ParaShape::RightMargin)
        .def_readwrite("Indent", &cpyhwpx::ParaShape::Indent)
        .def_readwrite("LineSpacing", &cpyhwpx::ParaShape::LineSpacing);

    py::class_<cpyhwpx::FontPreset>(m, "FontPreset")
        .def(py::init<>())
        .def_readwrite("FaceNameHangul", &cpyhwpx::FontPreset::FaceNameHangul)
        .def_readwrite("FaceNameLatin", &cpyhwpx::FontPreset::FaceNameLatin)
        .def_readwrite("FaceNameHanja", &cpyhwpx::FontPreset::FaceNameHanja)
        .def_readwrite("FontTypeHangul", &cpyhwpx::FontPreset::FontTypeHangul)
        .def_readwrite("FontTypeLatin", &cpyhwpx::FontPreset::FontTypeLatin);

    //=========================================================================
    // HwpWrapper 클래스 바인딩 (메인 클래스 - Hwp)
    //=========================================================================

    py::class_<cpyhwpx::HwpWrapper>(m, "Hwp")
        .def(py::init<bool, bool>(),
             py::arg("visible") = true,
             py::arg("new_instance") = true,
             "HWP 자동화 객체 생성")

        // 초기화/종료
        .def("initialize", &cpyhwpx::HwpWrapper::Initialize,
             "HWP COM 객체 초기화")
        .def("register_module", &cpyhwpx::HwpWrapper::RegisterModule,
             py::arg("module_type"), py::arg("module_data"),
             "보안 모듈 등록")
        .def("quit", &cpyhwpx::HwpWrapper::Quit,
             py::arg("save") = false,
             "HWP 종료")
        .def("is_initialized", &cpyhwpx::HwpWrapper::IsInitialized,
             "초기화 완료 여부")

        // 파일 I/O
        .def("open", &cpyhwpx::HwpWrapper::Open,
             py::arg("filename"),
             py::arg("format") = L"",
             py::arg("arg") = L"",
             "파일 열기")
        .def("save", &cpyhwpx::HwpWrapper::Save,
             py::arg("save_if_dirty") = true,
             "파일 저장")
        .def("save_as", &cpyhwpx::HwpWrapper::SaveAs,
             py::arg("filename"),
             py::arg("format") = L"HWP",
             py::arg("arg") = L"",
             "다른 이름으로 저장")
        .def("clear", &cpyhwpx::HwpWrapper::Clear,
             py::arg("option") = 1,
             "문서 초기화")
        .def("close", &cpyhwpx::HwpWrapper::Close,
             py::arg("is_dirty") = false,
             "문서 닫기")

        // 텍스트 편집
        .def("insert_text", &cpyhwpx::HwpWrapper::InsertText,
             py::arg("text"),
             "텍스트 삽입")
        .def("get_text", &cpyhwpx::HwpWrapper::GetText,
             "현재 위치의 텍스트 가져오기")
        .def("get_selected_text", &cpyhwpx::HwpWrapper::GetSelectedText,
             py::arg("keep_select") = false,
             "선택된 텍스트 가져오기")

        // 위치 관리
        .def("get_pos", &cpyhwpx::HwpWrapper::GetPos,
             "현재 위치 가져오기")
        .def("set_pos", &cpyhwpx::HwpWrapper::SetPos,
             py::arg("list"), py::arg("para"), py::arg("pos"),
             "위치 설정")
        .def("move_pos", &cpyhwpx::HwpWrapper::MovePos,
             py::arg("move_id"), py::arg("para") = 0, py::arg("pos") = 0,
             "위치 이동")

        // 창/UI 관리
        .def("set_visible", &cpyhwpx::HwpWrapper::SetVisible,
             py::arg("visible"),
             "창 표시/숨김")
        .def("maximize_window", &cpyhwpx::HwpWrapper::MaximizeWindow,
             "창 최대화")
        .def("minimize_window", &cpyhwpx::HwpWrapper::MinimizeWindow,
             "창 최소화")

        // 문서 상태
        .def("is_empty", &cpyhwpx::HwpWrapper::IsEmpty,
             "빈 문서 여부")
        .def("is_modified", &cpyhwpx::HwpWrapper::IsModified,
             "수정 여부")
        .def("is_cell", &cpyhwpx::HwpWrapper::IsCell,
             "셀 안 여부")

        // 찾기/바꾸기
        .def("find", &cpyhwpx::HwpWrapper::Find,
             py::arg("text"),
             py::arg("forward") = true,
             py::arg("match_case") = false,
             py::arg("regex") = false,
             py::arg("replace_mode") = false,
             "텍스트 찾기")
        .def("replace", &cpyhwpx::HwpWrapper::Replace,
             py::arg("find_text"),
             py::arg("replace_text"),
             py::arg("forward") = true,
             py::arg("match_case") = false,
             py::arg("regex") = false,
             "텍스트 바꾸기")
        .def("replace_all", &cpyhwpx::HwpWrapper::ReplaceAll,
             py::arg("find_text"),
             py::arg("replace_text"),
             py::arg("match_case") = false,
             py::arg("regex") = false,
             "모두 바꾸기")

        // HAction 관련
        .def("run", &cpyhwpx::HwpWrapper::RunAction,
             py::arg("act_id"),
             "HAction.Run() 실행")
        .def("Run", &cpyhwpx::HwpWrapper::RunAction,
             py::arg("act_id"),
             "HAction.Run() 실행 (pyhwpx 호환)")
        .def("find_ctrl", &cpyhwpx::HwpWrapper::FindCtrl,
             "현재 위치의 컨트롤을 찾아 선택")

        // 속성
        .def_property_readonly("version", &cpyhwpx::HwpWrapper::GetVersion,
                               "HWP 버전")
        .def_property_readonly("build_number", &cpyhwpx::HwpWrapper::GetBuildNumber,
                               "빌드 번호")
        .def_property_readonly("current_page", &cpyhwpx::HwpWrapper::GetCurrentPage,
                               "현재 페이지")
        .def_property_readonly("page_count", &cpyhwpx::HwpWrapper::GetPageCount,
                               "총 페이지 수")
        .def_property("edit_mode",
                      &cpyhwpx::HwpWrapper::GetEditMode,
                      &cpyhwpx::HwpWrapper::SetEditMode,
                      "편집 모드")

        //=========================================================================
        // 필드 작업 (Field Operations)
        //=========================================================================
        .def("create_field", &cpyhwpx::HwpWrapper::CreateField,
             py::arg("name"),
             py::arg("direction") = L"",
             py::arg("memo") = L"",
             "누름틀 필드 생성")
        .def("get_field_list", &cpyhwpx::HwpWrapper::GetFieldList,
             py::arg("number") = 1,
             py::arg("option") = 0,
             "필드 목록 조회 (\\x02로 구분)")
        .def("get_field_text", &cpyhwpx::HwpWrapper::GetFieldText,
             py::arg("field"),
             py::arg("idx") = 0,
             "필드 텍스트 조회")
        .def("put_field_text", &cpyhwpx::HwpWrapper::PutFieldText,
             py::arg("field"),
             py::arg("text"),
             "필드 텍스트 설정")
        .def("field_exist", &cpyhwpx::HwpWrapper::FieldExist,
             py::arg("field"),
             "필드 존재 확인")
        .def("move_to_field", &cpyhwpx::HwpWrapper::MoveToField,
             py::arg("field"),
             py::arg("idx") = 0,
             py::arg("text") = true,
             py::arg("start") = true,
             py::arg("select") = false,
             "필드로 캐럿 이동")
        .def("rename_field", &cpyhwpx::HwpWrapper::RenameField,
             py::arg("oldname"),
             py::arg("newname"),
             "필드 이름 변경")
        .def("get_cur_field_name", &cpyhwpx::HwpWrapper::GetCurFieldName,
             py::arg("option") = 0,
             "현재 위치 필드 이름 조회")
        .def("set_cur_field_name", &cpyhwpx::HwpWrapper::SetCurFieldName,
             py::arg("field"),
             py::arg("direction") = L"",
             py::arg("memo") = L"",
             py::arg("option") = 0,
             "현재 셀 필드 이름 설정")
        .def("set_field_view_option", &cpyhwpx::HwpWrapper::SetFieldViewOption,
             py::arg("option"),
             "필드 뷰 옵션 설정")
        .def("delete_all_fields", &cpyhwpx::HwpWrapper::DeleteAllFields,
             "모든 필드 삭제")
        .def("delete_field_by_name", &cpyhwpx::HwpWrapper::DeleteFieldByName,
             py::arg("field_name"),
             py::arg("idx") = -1,
             "이름으로 필드 삭제 (idx=-1이면 모든 동일 이름 필드)")
        .def("fields_to_map", &cpyhwpx::HwpWrapper::FieldsToMap,
             "필드를 딕셔너리로 변환")

        //=========================================================================
        // 테이블 작업 (Table Operations)
        //=========================================================================
        .def("create_table", &cpyhwpx::HwpWrapper::CreateTable,
             py::arg("rows") = 2,
             py::arg("cols") = 2,
             py::arg("treat_as_char") = true,
             py::arg("width_type") = 0,
             py::arg("height_type") = 0,
             py::arg("header") = false,
             "테이블 생성 (rows, cols, treat_as_char, width_type, height_type, header)")
        .def("get_into_nth_table", &cpyhwpx::HwpWrapper::GetIntoNthTable,
             py::arg("n") = 0,
             py::arg("select_cell") = false,
             "n번째 테이블로 이동 (음수: 뒤에서부터)")
        .def("get_table_row_count", &cpyhwpx::HwpWrapper::GetTableRowCount,
             "테이블 행 개수 조회 (-1: 테이블 아님)")
        .def("get_table_col_count", &cpyhwpx::HwpWrapper::GetTableColCount,
             "테이블 열 개수 조회 (-1: 테이블 아님)")
        .def("table_left_cell", &cpyhwpx::HwpWrapper::TableLeftCell,
             "왼쪽 셀로 이동")
        .def("table_right_cell", &cpyhwpx::HwpWrapper::TableRightCell,
             "오른쪽 셀로 이동")
        .def("table_upper_cell", &cpyhwpx::HwpWrapper::TableUpperCell,
             "위쪽 셀로 이동")
        .def("table_lower_cell", &cpyhwpx::HwpWrapper::TableLowerCell,
             "아래쪽 셀로 이동")
        .def("table_right_cell_append", &cpyhwpx::HwpWrapper::TableRightCellAppend,
             "오른쪽 셀로 이동 (행 끝이면 다음 행)");

    //=========================================================================
    // HwpCtrl 클래스 바인딩
    //=========================================================================

    py::class_<cpyhwpx::HwpCtrl>(m, "Ctrl")
        .def("get_ctrl_id", &cpyhwpx::HwpCtrl::GetCtrlID,
             "컨트롤 ID")
        .def("get_ctrl_type", &cpyhwpx::HwpCtrl::GetCtrlType,
             "컨트롤 타입")
        .def("is_valid", &cpyhwpx::HwpCtrl::IsValid,
             "유효한 컨트롤인지")
        .def("is_table", &cpyhwpx::HwpCtrl::IsTable,
             "표인지 확인")
        .def("is_picture", &cpyhwpx::HwpCtrl::IsPicture,
             "그림인지 확인")
        .def("get_row_count", &cpyhwpx::HwpCtrl::GetRowCount,
             "표 행 수")
        .def("get_col_count", &cpyhwpx::HwpCtrl::GetColCount,
             "표 열 수")
        .def("get_user_data", &cpyhwpx::HwpCtrl::GetUserData,
             "사용자 데이터")
        .def("set_user_data", &cpyhwpx::HwpCtrl::SetUserData,
             py::arg("data"),
             "사용자 데이터 설정")
        .def("get_size", &cpyhwpx::HwpCtrl::GetSize,
             "크기 가져오기 (width, height)")
        .def("set_size", &cpyhwpx::HwpCtrl::SetSize,
             py::arg("width"), py::arg("height"),
             "크기 설정");

    //=========================================================================
    // FontDefs 클래스 바인딩
    //=========================================================================

    py::class_<cpyhwpx::FontDefs>(m, "FontDefs")
        .def_static("get_preset", &cpyhwpx::FontDefs::GetPreset,
                    py::arg("preset_name"),
                    "프리셋 이름으로 폰트 정보 가져오기")
        .def_static("has_preset", &cpyhwpx::FontDefs::HasPreset,
                    py::arg("preset_name"),
                    "프리셋 존재 여부")
        .def_static("get_preset_names", &cpyhwpx::FontDefs::GetPresetNames,
                    "모든 프리셋 이름 목록")
        .def_static("malgun_gothic", &cpyhwpx::FontDefs::MalgunGothic,
                    "맑은 고딕 프리셋")
        .def_static("nanum_gothic", &cpyhwpx::FontDefs::NanumGothic,
                    "나눔고딕 프리셋")
        .def_static("nanum_myeongjo", &cpyhwpx::FontDefs::NanumMyeongjo,
                    "나눔명조 프리셋");

    //=========================================================================
    // Utils 서브모듈
    //=========================================================================

    py::module_ utils = m.def_submodule("utils", "유틸리티 함수");

    utils.def("addr_to_tuple", &cpyhwpx::Utils::AddrToTuple,
              py::arg("address"),
              "엑셀 주소를 튜플로 변환 (1-based)");
    utils.def("tuple_to_addr", &cpyhwpx::Utils::TupleToAddr,
              py::arg("row"), py::arg("col"),
              "튜플을 엑셀 주소로 변환 (1-based)");
    utils.def("parse_range", &cpyhwpx::Utils::ParseRange,
              py::arg("range"),
              "셀 범위 파싱");
    utils.def("expand_range", &cpyhwpx::Utils::ExpandRange,
              py::arg("range"),
              "범위 내 셀 주소 목록");
    utils.def("trim", &cpyhwpx::Utils::Trim,
              py::arg("str"),
              "문자열 공백 제거");
    utils.def("split", &cpyhwpx::Utils::Split,
              py::arg("str"), py::arg("delimiter"),
              "문자열 분할");
    utils.def("join", &cpyhwpx::Utils::Join,
              py::arg("parts"), py::arg("delimiter"),
              "문자열 결합");
    utils.def("file_exists", &cpyhwpx::Utils::FileExists,
              py::arg("path"),
              "파일 존재 여부");
    utils.def("get_extension", &cpyhwpx::Utils::GetExtension,
              py::arg("path"),
              "파일 확장자");
    utils.def("hex_to_colorref", &cpyhwpx::Utils::HexToColorRef,
              py::arg("hex"),
              "16진수를 COLORREF로 변환");
    utils.def("colorref_to_hex", &cpyhwpx::Utils::ColorRefToHex,
              py::arg("color"),
              "COLORREF를 16진수로 변환");

    //=========================================================================
    // Units 서브모듈
    //=========================================================================

    py::module_ units = m.def_submodule("units", "단위 변환 함수");

    units.def("from_mm", &cpyhwpx::Units::FromMM,
              py::arg("mm"),
              "mm를 HwpUnit으로 변환");
    units.def("from_cm", &cpyhwpx::Units::FromCM,
              py::arg("cm"),
              "cm를 HwpUnit으로 변환");
    units.def("from_inch", &cpyhwpx::Units::FromInch,
              py::arg("inch"),
              "인치를 HwpUnit으로 변환");
    units.def("from_point", &cpyhwpx::Units::FromPoint,
              py::arg("pt"),
              "포인트를 HwpUnit으로 변환");
    units.def("to_mm", &cpyhwpx::Units::ToMM,
              py::arg("unit"),
              "HwpUnit을 mm로 변환");
    units.def("to_cm", &cpyhwpx::Units::ToCM,
              py::arg("unit"),
              "HwpUnit을 cm로 변환");
    units.def("to_inch", &cpyhwpx::Units::ToInch,
              py::arg("unit"),
              "HwpUnit을 인치로 변환");
    units.def("to_point", &cpyhwpx::Units::ToPoint,
              py::arg("unit"),
              "HwpUnit을 포인트로 변환");

    // 상수
    units.attr("HWPUNIT_PER_MM") = cpyhwpx::Units::HWPUNIT_PER_MM;
    units.attr("HWPUNIT_PER_CM") = cpyhwpx::Units::HWPUNIT_PER_CM;
    units.attr("HWPUNIT_PER_INCH") = cpyhwpx::Units::HWPUNIT_PER_INCH;
    units.attr("HWPUNIT_PER_PT") = cpyhwpx::Units::HWPUNIT_PER_PT;
}
