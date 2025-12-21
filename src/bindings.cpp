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
        .def(py::init<bool, bool, bool>(),
             py::arg("visible") = true,
             py::arg("new_instance") = true,
             py::arg("register_module") = true,
             "HWP 자동화 객체 생성 (register_module=True: 보안 모듈 자동 등록)")

        // 초기화/종료
        .def("initialize", &cpyhwpx::HwpWrapper::Initialize,
             "HWP COM 객체 초기화")
        .def("register_module", &cpyhwpx::HwpWrapper::RegisterModule,
             py::arg("module_type"), py::arg("module_data"),
             "보안 모듈 등록")

        // 보안 모듈 자동 등록
        .def_static("check_registry_key", &cpyhwpx::HwpWrapper::CheckRegistryKey,
             py::arg("key_name") = L"FilePathCheckerModule",
             "보안 모듈 레지스트리 등록 여부 확인")
        .def_static("register_to_registry", &cpyhwpx::HwpWrapper::RegisterToRegistry,
             py::arg("dll_path") = L"",
             py::arg("key_name") = L"FilePathCheckerModule",
             "보안 모듈 DLL을 레지스트리에 등록")
        .def_static("find_dll_path", &cpyhwpx::HwpWrapper::FindDllPath,
             "DLL 파일 경로 자동 감지")
        .def("auto_register_module", &cpyhwpx::HwpWrapper::AutoRegisterModule,
             py::arg("module_type") = L"FilePathCheckDLL",
             py::arg("module_data") = L"FilePathCheckerModule",
             "보안 모듈 자동 등록 (레지스트리 확인/등록 + COM API)")

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
        .def("insert_file", &cpyhwpx::HwpWrapper::InsertFile,
             py::arg("filename"),
             py::arg("keep_section") = 1,
             py::arg("keep_charshape") = 1,
             py::arg("keep_parashape") = 1,
             py::arg("keep_style") = 1,
             py::arg("move_doc_end") = false,
             "현재 위치에 파일 삽입 (keep_*: 1=유지, 0=무시)")
        .def("get_text_file", &cpyhwpx::HwpWrapper::GetTextFile,
             py::arg("format") = L"UNICODE",
             py::arg("option") = L"",
             "문서 텍스트 추출 (format: HWP/HWPML2X/HTML/UNICODE/TEXT, option: saveblock:true=선택블록)")
        .def("set_text_file", &cpyhwpx::HwpWrapper::SetTextFile,
             py::arg("data"),
             py::arg("format") = L"HWPML2X",
             py::arg("option") = L"insertfile",
             "텍스트 데이터 삽입 (format: HWP/HWPML2X/HTML/UNICODE/TEXT)")
        .def("open_pdf", &cpyhwpx::HwpWrapper::OpenPdf,
             py::arg("pdf_path"),
             py::arg("this_window") = 1,
             "PDF 파일 열기 (this_window: 1=현재창, 0=새창)")
        .def("save_block_as", &cpyhwpx::HwpWrapper::SaveBlockAs,
             py::arg("path"),
             py::arg("format") = L"HWP",
             py::arg("attributes") = 1,
             "선택 블록을 파일로 저장")
        .def("get_file_info", &cpyhwpx::HwpWrapper::GetFileInfo,
             py::arg("filename"),
             "파일 정보 조회 (Format, VersionStr, VersionNum, Encrypted)")

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

        // 텍스트 스캔/선택
        .def("init_scan", &cpyhwpx::HwpWrapper::InitScan,
             py::arg("option") = 0x07, py::arg("range") = 0x77,
             py::arg("spara") = 0, py::arg("spos") = 0,
             py::arg("epara") = -1, py::arg("epos") = -1,
             "텍스트 스캔 초기화")
        .def("release_scan", &cpyhwpx::HwpWrapper::ReleaseScan,
             "텍스트 스캔 해제")
        .def("select_text", &cpyhwpx::HwpWrapper::SelectText,
             py::arg("spara"), py::arg("spos"),
             py::arg("epara"), py::arg("epos"),
             py::arg("slist") = 0,
             "텍스트 선택")
        .def("select_text_by_get_pos", &cpyhwpx::HwpWrapper::SelectTextByGetPos,
             py::arg("s_pos"), py::arg("e_pos"),
             "GetPos 튜플로 텍스트 선택")
        .def("get_pos_by_set", &cpyhwpx::HwpWrapper::GetPosBySetPy,
             "위치 저장 후 인덱스 반환 (-1: 실패)")
        .def("set_pos_by_set", &cpyhwpx::HwpWrapper::SetPosBySetPy,
             py::arg("idx"),
             "인덱스로 위치 복원")
        .def("clear_pos_cache", &cpyhwpx::HwpWrapper::ClearPosCache,
             "위치 캐시 정리")

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
        .def("insert_ctrl", [](cpyhwpx::HwpWrapper& self, const std::wstring& ctrl_id) {
                 return self.InsertCtrl(ctrl_id, nullptr);
             },
             py::arg("ctrl_id"),
             py::return_value_policy::take_ownership,
             "컨트롤 삽입 (ctrl_id: tbl/pic/gso/eqed 등)")
        .def("delete_ctrl", py::overload_cast<cpyhwpx::HwpCtrl*>(&cpyhwpx::HwpWrapper::DeleteCtrl),
             py::arg("ctrl"),
             "컨트롤 삭제")

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
        // COM 속성 접근 (Low-level API)
        //=========================================================================
        // Note: IDispatch* 반환 속성들 (Application, EngineProperties, ViewProperties,
        //       XHwpMessageBox, XHwpODBC, XHwpWindows, HAction, HParameterSet)은
        //       pybind11에서 직접 지원되지 않아 제외됨
        .def_property_readonly("CLSID", &cpyhwpx::HwpWrapper::GetCLSID,
                               "클래스 ID")
        .def_property_readonly("CurFieldState", &cpyhwpx::HwpWrapper::GetCurFieldState,
                               "현재 필드 상태 (0=본문, 1=셀, 4=글상자)")
        .def_property_readonly("CurMetatagState", &cpyhwpx::HwpWrapper::GetCurMetatagState,
                               "현재 메타태그 상태 (한글2024+)")
        .def_property_readonly("IsPrivateInfoProtected", &cpyhwpx::HwpWrapper::GetIsPrivateInfoProtected,
                               "개인정보 보호 여부")
        .def_property_readonly("IsTrackChange", &cpyhwpx::HwpWrapper::GetIsTrackChange,
                               "변경 추적 여부")
        .def_property_readonly("Path", &cpyhwpx::HwpWrapper::GetDocPath,
                               "문서 경로")
        .def_property_readonly("SelectionMode", &cpyhwpx::HwpWrapper::GetSelectionMode,
                               "선택 모드 (0=일반, 1=블록)")
        .def_property_readonly("Title", &cpyhwpx::HwpWrapper::GetTitle,
                               "창 제목")
        .def_property_readonly("current_printpage", &cpyhwpx::HwpWrapper::GetCurrentPrintPage,
                               "현재 인쇄 페이지 번호")
        .def_property_readonly("current_font", &cpyhwpx::HwpWrapper::GetCurrentFont,
                               "현재 폰트명")

        //=========================================================================
        // 컨트롤 속성 (Control Properties)
        //=========================================================================
        .def_property_readonly("cur_selected_ctrl", &cpyhwpx::HwpWrapper::GetCurSelectedCtrl,
                               "현재 선택된 컨트롤")
        .def_property_readonly("head_ctrl", &cpyhwpx::HwpWrapper::GetHeadCtrl,
                               "문서의 첫 번째 컨트롤 (secd)")
        .def_property_readonly("last_ctrl", &cpyhwpx::HwpWrapper::GetLastCtrl,
                               "문서의 마지막 컨트롤")
        .def_property_readonly("parent_ctrl", &cpyhwpx::HwpWrapper::GetParentCtrl,
                               "현재 위치를 포함하는 상위 컨트롤")
        .def_property_readonly("ctrl_list", &cpyhwpx::HwpWrapper::GetCtrlList,
                               "문서 내 모든 사용자 컨트롤 목록 (secd, cold 제외)")

        //=========================================================================
        // 문서 컬렉션 (Document Collection)
        //=========================================================================
        .def_property_readonly("XHwpDocuments", &cpyhwpx::HwpWrapper::GetXHwpDocuments,
                               "문서 컬렉션")
        .def("switch_to", &cpyhwpx::HwpWrapper::SwitchTo,
             py::arg("num"),
             "문서 전환 (인덱스)")
        .def("add_tab", &cpyhwpx::HwpWrapper::AddTab,
             "새 탭으로 문서 추가")
        .def("add_doc", &cpyhwpx::HwpWrapper::AddDoc,
             "새 창으로 문서 추가")

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
             "오른쪽 셀로 이동 (행 끝이면 다음 행)")
        .def("table_col_begin", &cpyhwpx::HwpWrapper::TableColBegin,
             "첫 번째 열로 이동")
        .def("table_col_end", &cpyhwpx::HwpWrapper::TableColEnd,
             "마지막 열로 이동")
        .def("table_col_page_up", &cpyhwpx::HwpWrapper::TableColPageUp,
             "열의 맨 위로 이동")
        .def("table_cell_block_extend_abs", &cpyhwpx::HwpWrapper::TableCellBlockExtendAbs,
             "셀 블록 선택 확장 (절대)")
        .def("cancel", &cpyhwpx::HwpWrapper::Cancel,
             "선택 취소 (ESC)")
        .def("cell_fill", &cpyhwpx::HwpWrapper::CellFill,
             py::arg("r") = 217,
             py::arg("g") = 217,
             py::arg("b") = 217,
             "셀 배경색 채우기 (기본: 연한 회색)")
        .def("table_from_data", &cpyhwpx::HwpWrapper::TableFromData,
             py::arg("data"),
             py::arg("treat_as_char") = false,
             py::arg("header") = true,
             py::arg("header_bold") = true,
             py::arg("cell_fill_r") = -1,
             py::arg("cell_fill_g") = -1,
             py::arg("cell_fill_b") = -1,
             "2D 리스트 데이터로 테이블 생성")

        .def("get_table_xml", &cpyhwpx::HwpWrapper::GetTableXml,
             "현재 커서 위치의 테이블 XML 추출 (HWPML2X 형식)")

        //=========================================================================
        // 스타일 관리 (CharShape/ParaShape)
        //=========================================================================
        .def("get_charshape", &cpyhwpx::HwpWrapper::GetCharShape,
             "현재 글자모양 속성 가져오기 (dict 반환)")
        .def("set_charshape", &cpyhwpx::HwpWrapper::SetCharShape,
             py::arg("props"),
             "글자모양 속성 설정 (dict)")
        .def("set_font", &cpyhwpx::HwpWrapper::SetFont,
             py::arg("face_name") = L"",
             py::arg("height") = -1,
             py::arg("bold") = -1,
             py::arg("italic") = -1,
             py::arg("text_color") = -1,
             "글자모양 간편 설정 (face_name: 글꼴, height: pt*100, bold/italic: 0=해제,1=설정,-1=미변경)")
        .def("get_parashape", &cpyhwpx::HwpWrapper::GetParaShape,
             "현재 문단모양 속성 가져오기 (dict 반환)")
        .def("set_parashape", &cpyhwpx::HwpWrapper::SetParaShape,
             py::arg("props"),
             "문단모양 속성 설정 (dict)")
        .def("set_para", &cpyhwpx::HwpWrapper::SetPara,
             py::arg("align_type") = -1,
             py::arg("line_spacing") = -1,
             py::arg("left_margin") = -1,
             py::arg("indentation") = -1,
             "문단모양 간편 설정 (align_type: 0=양쪽,1=왼쪽,2=가운데,3=오른쪽)")

        //=========================================================================
        // 이미지 삽입 (Image Insertion)
        //=========================================================================
        .def("insert_picture", &cpyhwpx::HwpWrapper::InsertPicture,
             py::arg("path"),
             py::arg("embedded") = true,
             py::arg("sizeoption") = 0,
             py::arg("reverse") = false,
             py::arg("watermark") = false,
             py::arg("effect") = 0,
             py::arg("width") = 0,
             py::arg("height") = 0,
             "이미지 삽입 (sizeoption: 0=원본, 1=지정크기, 2=셀맞춤, 3=셀맞춤+종횡비)");

    //=========================================================================
    // XHwpDocument 클래스 바인딩
    //=========================================================================

    py::class_<cpyhwpx::XHwpDocument>(m, "XHwpDocument")
        // 속성
        .def_property_readonly("DocumentID", &cpyhwpx::XHwpDocument::GetDocumentID,
                               "문서 고유 ID")
        .def_property_readonly("FullName", &cpyhwpx::XHwpDocument::GetFullName,
                               "문서 전체 경로")
        .def_property_readonly("Path", &cpyhwpx::XHwpDocument::GetPath,
                               "문서 디렉토리 경로")
        .def_property_readonly("Format", &cpyhwpx::XHwpDocument::GetFormat,
                               "문서 포맷")
        .def_property_readonly("Modified", &cpyhwpx::XHwpDocument::IsModified,
                               "수정 여부")
        .def_property_readonly("EditMode", &cpyhwpx::XHwpDocument::GetEditMode,
                               "편집 모드")
        .def("is_valid", &cpyhwpx::XHwpDocument::IsValid,
             "유효한 문서인지")
        // 메서드
        .def("SetActive_XHwpDocument", &cpyhwpx::XHwpDocument::SetActiveDocument,
             "이 문서를 활성화")
        .def("Open", &cpyhwpx::XHwpDocument::Open,
             py::arg("filename"),
             py::arg("format") = L"",
             py::arg("arg") = L"",
             "파일 열기")
        .def("Save", &cpyhwpx::XHwpDocument::Save,
             py::arg("save_if_dirty") = false,
             "저장")
        .def("SaveAs", &cpyhwpx::XHwpDocument::SaveAs,
             py::arg("path"),
             py::arg("format") = L"",
             py::arg("arg") = L"",
             "다른 이름으로 저장")
        .def("Close", &cpyhwpx::XHwpDocument::Close,
             py::arg("is_dirty") = false,
             "닫기")
        .def("Clear", &cpyhwpx::XHwpDocument::Clear,
             py::arg("option") = false,
             "내용 지우기")
        .def("Undo", &cpyhwpx::XHwpDocument::Undo,
             py::arg("count") = 1,
             "실행 취소")
        .def("Redo", &cpyhwpx::XHwpDocument::Redo,
             py::arg("count") = 1,
             "다시 실행");

    //=========================================================================
    // XHwpDocuments 클래스 바인딩
    //=========================================================================

    py::class_<cpyhwpx::XHwpDocuments>(m, "XHwpDocuments")
        // 컬렉션 프로토콜
        .def("__len__", &cpyhwpx::XHwpDocuments::Len)
        .def("__getitem__", &cpyhwpx::XHwpDocuments::GetItem,
             py::arg("index"))
        // 속성
        .def_property_readonly("Count", &cpyhwpx::XHwpDocuments::GetCount,
                               "문서 개수")
        .def_property_readonly("Active_XHwpDocument", &cpyhwpx::XHwpDocuments::GetActiveDocument,
                               "활성 문서")
        .def("is_valid", &cpyhwpx::XHwpDocuments::IsValid,
             "유효한 컬렉션인지")
        // 메서드
        .def("Item", &cpyhwpx::XHwpDocuments::Item,
             py::arg("index"),
             "인덱스로 문서 접근")
        .def("FindItem", &cpyhwpx::XHwpDocuments::FindItem,
             py::arg("docID"),
             "ID로 문서 검색")
        .def("Add", &cpyhwpx::XHwpDocuments::Add,
             py::arg("isTab") = false,
             "새 문서 추가 (isTab=True: 탭, False: 창)")
        .def("Close", &cpyhwpx::XHwpDocuments::Close,
             py::arg("isDirty") = false,
             "활성 문서 닫기");

    //=========================================================================
    // HwpCtrl 클래스 바인딩
    //=========================================================================

    py::class_<cpyhwpx::HwpCtrl>(m, "Ctrl")
        // 속성
        .def_property_readonly("ctrlid", &cpyhwpx::HwpCtrl::GetCtrlID,
             "컨트롤 ID")
        .def_property_readonly("ctrl_type", &cpyhwpx::HwpCtrl::GetCtrlType,
             "컨트롤 타입")
        // 메서드 (하위 호환성)
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
             "크기 설정")
        // 컨트롤 순회
        .def("next", &cpyhwpx::HwpCtrl::Next,
             "다음 컨트롤")
        .def("prev", &cpyhwpx::HwpCtrl::Prev,
             "이전 컨트롤");

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

    //=========================================================================
    // HwpActionHelper 서브모듈 (HAction.Run 헬퍼)
    //=========================================================================

    py::module_ actions = m.def_submodule("actions", "HAction.Run 헬퍼 함수");

    // 기본 Run
    actions.def("Run", &cpyhwpx::HwpActionHelper::Run,
                py::arg("hwp"), py::arg("action_name"),
                "지정된 액션 실행");

    // 문단
    actions.def("BreakPara", &cpyhwpx::HwpActionHelper::BreakPara, py::arg("hwp"), "문단 나누기");
    actions.def("BreakPage", &cpyhwpx::HwpActionHelper::BreakPage, py::arg("hwp"), "페이지 나누기");
    actions.def("BreakSection", &cpyhwpx::HwpActionHelper::BreakSection, py::arg("hwp"), "구역 나누기");
    actions.def("BreakColumn", &cpyhwpx::HwpActionHelper::BreakColumn, py::arg("hwp"), "다단 나누기");
    actions.def("BreakLine", &cpyhwpx::HwpActionHelper::BreakLine, py::arg("hwp"), "줄 나누기");
    actions.def("BreakColDef", &cpyhwpx::HwpActionHelper::BreakColDef, py::arg("hwp"), "단 정의 나누기");

    // 선택/취소
    actions.def("SelectAll", &cpyhwpx::HwpActionHelper::SelectAll, py::arg("hwp"), "모두 선택");
    actions.def("Cancel", &cpyhwpx::HwpActionHelper::Cancel, py::arg("hwp"), "선택 취소");
    actions.def("Select", &cpyhwpx::HwpActionHelper::Select, py::arg("hwp"), "선택");
    actions.def("SelectColumn", &cpyhwpx::HwpActionHelper::SelectColumn, py::arg("hwp"), "열 선택");
    actions.def("SelectCtrlFront", &cpyhwpx::HwpActionHelper::SelectCtrlFront, py::arg("hwp"), "컨트롤 앞으로 선택");
    actions.def("SelectCtrlReverse", &cpyhwpx::HwpActionHelper::SelectCtrlReverse, py::arg("hwp"), "컨트롤 뒤로 선택");
    actions.def("UnSelectCtrl", &cpyhwpx::HwpActionHelper::UnSelectCtrl, py::arg("hwp"), "컨트롤 선택 해제");

    // 커서 이동 (기본)
    actions.def("MoveLeft", &cpyhwpx::HwpActionHelper::MoveLeft, py::arg("hwp"), "왼쪽으로 이동");
    actions.def("MoveRight", &cpyhwpx::HwpActionHelper::MoveRight, py::arg("hwp"), "오른쪽으로 이동");
    actions.def("MoveUp", &cpyhwpx::HwpActionHelper::MoveUp, py::arg("hwp"), "위로 이동");
    actions.def("MoveDown", &cpyhwpx::HwpActionHelper::MoveDown, py::arg("hwp"), "아래로 이동");
    actions.def("MoveLineBegin", &cpyhwpx::HwpActionHelper::MoveLineBegin, py::arg("hwp"), "줄 시작으로 이동");
    actions.def("MoveLineEnd", &cpyhwpx::HwpActionHelper::MoveLineEnd, py::arg("hwp"), "줄 끝으로 이동");
    actions.def("MoveDocBegin", &cpyhwpx::HwpActionHelper::MoveDocBegin, py::arg("hwp"), "문서 시작으로 이동");
    actions.def("MoveDocEnd", &cpyhwpx::HwpActionHelper::MoveDocEnd, py::arg("hwp"), "문서 끝으로 이동");
    actions.def("MoveParaBegin", &cpyhwpx::HwpActionHelper::MoveParaBegin, py::arg("hwp"), "문단 시작으로 이동");
    actions.def("MoveParaEnd", &cpyhwpx::HwpActionHelper::MoveParaEnd, py::arg("hwp"), "문단 끝으로 이동");
    actions.def("MoveWordBegin", &cpyhwpx::HwpActionHelper::MoveWordBegin, py::arg("hwp"), "단어 시작으로 이동");
    actions.def("MoveWordEnd", &cpyhwpx::HwpActionHelper::MoveWordEnd, py::arg("hwp"), "단어 끝으로 이동");
    actions.def("MovePageUp", &cpyhwpx::HwpActionHelper::MovePageUp, py::arg("hwp"), "페이지 위로");
    actions.def("MovePageDown", &cpyhwpx::HwpActionHelper::MovePageDown, py::arg("hwp"), "페이지 아래로");

    // 커서 이동 (추가)
    actions.def("MoveLineDown", &cpyhwpx::HwpActionHelper::MoveLineDown, py::arg("hwp"), "줄 아래로");
    actions.def("MoveLineUp", &cpyhwpx::HwpActionHelper::MoveLineUp, py::arg("hwp"), "줄 위로");
    actions.def("MoveNextWord", &cpyhwpx::HwpActionHelper::MoveNextWord, py::arg("hwp"), "다음 단어로");
    actions.def("MovePrevWord", &cpyhwpx::HwpActionHelper::MovePrevWord, py::arg("hwp"), "이전 단어로");
    actions.def("MoveNextChar", &cpyhwpx::HwpActionHelper::MoveNextChar, py::arg("hwp"), "다음 문자로");
    actions.def("MovePrevChar", &cpyhwpx::HwpActionHelper::MovePrevChar, py::arg("hwp"), "이전 문자로");
    actions.def("MovePageBegin", &cpyhwpx::HwpActionHelper::MovePageBegin, py::arg("hwp"), "페이지 시작으로");
    actions.def("MovePageEnd", &cpyhwpx::HwpActionHelper::MovePageEnd, py::arg("hwp"), "페이지 끝으로");
    actions.def("MoveColumnBegin", &cpyhwpx::HwpActionHelper::MoveColumnBegin, py::arg("hwp"), "단 시작으로");
    actions.def("MoveColumnEnd", &cpyhwpx::HwpActionHelper::MoveColumnEnd, py::arg("hwp"), "단 끝으로");
    actions.def("MoveListBegin", &cpyhwpx::HwpActionHelper::MoveListBegin, py::arg("hwp"), "리스트 시작으로");
    actions.def("MoveListEnd", &cpyhwpx::HwpActionHelper::MoveListEnd, py::arg("hwp"), "리스트 끝으로");
    actions.def("MoveParentList", &cpyhwpx::HwpActionHelper::MoveParentList, py::arg("hwp"), "상위 리스트로");
    actions.def("MoveRootList", &cpyhwpx::HwpActionHelper::MoveRootList, py::arg("hwp"), "루트 리스트로");
    actions.def("MoveTopLevelBegin", &cpyhwpx::HwpActionHelper::MoveTopLevelBegin, py::arg("hwp"), "최상위 시작으로");
    actions.def("MoveTopLevelEnd", &cpyhwpx::HwpActionHelper::MoveTopLevelEnd, py::arg("hwp"), "최상위 끝으로");
    actions.def("MoveTopLevelList", &cpyhwpx::HwpActionHelper::MoveTopLevelList, py::arg("hwp"), "최상위 리스트로");
    actions.def("MoveNextColumn", &cpyhwpx::HwpActionHelper::MoveNextColumn, py::arg("hwp"), "다음 단으로");
    actions.def("MovePrevColumn", &cpyhwpx::HwpActionHelper::MovePrevColumn, py::arg("hwp"), "이전 단으로");
    actions.def("MoveNextPos", &cpyhwpx::HwpActionHelper::MoveNextPos, py::arg("hwp"), "다음 위치로");
    actions.def("MovePrevPos", &cpyhwpx::HwpActionHelper::MovePrevPos, py::arg("hwp"), "이전 위치로");
    actions.def("MoveNextPosEx", &cpyhwpx::HwpActionHelper::MoveNextPosEx, py::arg("hwp"), "다음 위치로 (확장)");
    actions.def("MovePrevPosEx", &cpyhwpx::HwpActionHelper::MovePrevPosEx, py::arg("hwp"), "이전 위치로 (확장)");
    actions.def("MoveNextParaBegin", &cpyhwpx::HwpActionHelper::MoveNextParaBegin, py::arg("hwp"), "다음 문단 시작으로");
    actions.def("MovePrevParaBegin", &cpyhwpx::HwpActionHelper::MovePrevParaBegin, py::arg("hwp"), "이전 문단 시작으로");
    actions.def("MovePrevParaEnd", &cpyhwpx::HwpActionHelper::MovePrevParaEnd, py::arg("hwp"), "이전 문단 끝으로");
    actions.def("MoveSectionDown", &cpyhwpx::HwpActionHelper::MoveSectionDown, py::arg("hwp"), "섹션 아래로");
    actions.def("MoveSectionUp", &cpyhwpx::HwpActionHelper::MoveSectionUp, py::arg("hwp"), "섹션 위로");
    actions.def("MoveScrollDown", &cpyhwpx::HwpActionHelper::MoveScrollDown, py::arg("hwp"), "스크롤 아래로");
    actions.def("MoveScrollUp", &cpyhwpx::HwpActionHelper::MoveScrollUp, py::arg("hwp"), "스크롤 위로");
    actions.def("MoveScrollNext", &cpyhwpx::HwpActionHelper::MoveScrollNext, py::arg("hwp"), "스크롤 다음");
    actions.def("MoveScrollPrev", &cpyhwpx::HwpActionHelper::MoveScrollPrev, py::arg("hwp"), "스크롤 이전");
    actions.def("MoveViewBegin", &cpyhwpx::HwpActionHelper::MoveViewBegin, py::arg("hwp"), "뷰 시작으로");
    actions.def("MoveViewEnd", &cpyhwpx::HwpActionHelper::MoveViewEnd, py::arg("hwp"), "뷰 끝으로");
    actions.def("MoveViewDown", &cpyhwpx::HwpActionHelper::MoveViewDown, py::arg("hwp"), "뷰 아래로");
    actions.def("MoveViewUp", &cpyhwpx::HwpActionHelper::MoveViewUp, py::arg("hwp"), "뷰 위로");

    // 선택 이동
    actions.def("MoveSelDown", &cpyhwpx::HwpActionHelper::MoveSelDown, py::arg("hwp"), "선택 아래로");
    actions.def("MoveSelLeft", &cpyhwpx::HwpActionHelper::MoveSelLeft, py::arg("hwp"), "선택 왼쪽으로");
    actions.def("MoveSelRight", &cpyhwpx::HwpActionHelper::MoveSelRight, py::arg("hwp"), "선택 오른쪽으로");
    actions.def("MoveSelUp", &cpyhwpx::HwpActionHelper::MoveSelUp, py::arg("hwp"), "선택 위로");
    actions.def("MoveSelDocBegin", &cpyhwpx::HwpActionHelper::MoveSelDocBegin, py::arg("hwp"), "선택 문서 시작까지");
    actions.def("MoveSelDocEnd", &cpyhwpx::HwpActionHelper::MoveSelDocEnd, py::arg("hwp"), "선택 문서 끝까지");
    actions.def("MoveSelLineBegin", &cpyhwpx::HwpActionHelper::MoveSelLineBegin, py::arg("hwp"), "선택 줄 시작까지");
    actions.def("MoveSelLineEnd", &cpyhwpx::HwpActionHelper::MoveSelLineEnd, py::arg("hwp"), "선택 줄 끝까지");
    actions.def("MoveSelLineDown", &cpyhwpx::HwpActionHelper::MoveSelLineDown, py::arg("hwp"), "선택 줄 아래로");
    actions.def("MoveSelLineUp", &cpyhwpx::HwpActionHelper::MoveSelLineUp, py::arg("hwp"), "선택 줄 위로");
    actions.def("MoveSelNextChar", &cpyhwpx::HwpActionHelper::MoveSelNextChar, py::arg("hwp"), "선택 다음 문자");
    actions.def("MoveSelPrevChar", &cpyhwpx::HwpActionHelper::MoveSelPrevChar, py::arg("hwp"), "선택 이전 문자");
    actions.def("MoveSelNextWord", &cpyhwpx::HwpActionHelper::MoveSelNextWord, py::arg("hwp"), "선택 다음 단어");
    actions.def("MoveSelPrevWord", &cpyhwpx::HwpActionHelper::MoveSelPrevWord, py::arg("hwp"), "선택 이전 단어");
    actions.def("MoveSelPageDown", &cpyhwpx::HwpActionHelper::MoveSelPageDown, py::arg("hwp"), "선택 페이지 아래로");
    actions.def("MoveSelPageUp", &cpyhwpx::HwpActionHelper::MoveSelPageUp, py::arg("hwp"), "선택 페이지 위로");
    actions.def("MoveSelParaBegin", &cpyhwpx::HwpActionHelper::MoveSelParaBegin, py::arg("hwp"), "선택 문단 시작까지");
    actions.def("MoveSelParaEnd", &cpyhwpx::HwpActionHelper::MoveSelParaEnd, py::arg("hwp"), "선택 문단 끝까지");
    actions.def("MoveSelWordBegin", &cpyhwpx::HwpActionHelper::MoveSelWordBegin, py::arg("hwp"), "선택 단어 시작까지");
    actions.def("MoveSelWordEnd", &cpyhwpx::HwpActionHelper::MoveSelWordEnd, py::arg("hwp"), "선택 단어 끝까지");

    // 편집
    actions.def("Delete", &cpyhwpx::HwpActionHelper::Delete, py::arg("hwp"), "삭제");
    actions.def("DeleteBack", &cpyhwpx::HwpActionHelper::DeleteBack, py::arg("hwp"), "앞 글자 삭제");
    actions.def("DeleteLine", &cpyhwpx::HwpActionHelper::DeleteLine, py::arg("hwp"), "줄 삭제");
    actions.def("DeleteLineEnd", &cpyhwpx::HwpActionHelper::DeleteLineEnd, py::arg("hwp"), "줄 끝까지 삭제");
    actions.def("DeleteWord", &cpyhwpx::HwpActionHelper::DeleteWord, py::arg("hwp"), "단어 삭제");
    actions.def("DeleteWordBack", &cpyhwpx::HwpActionHelper::DeleteWordBack, py::arg("hwp"), "이전 단어 삭제");
    actions.def("DeleteField", &cpyhwpx::HwpActionHelper::DeleteField, py::arg("hwp"), "필드 삭제");
    actions.def("DeleteFieldMemo", &cpyhwpx::HwpActionHelper::DeleteFieldMemo, py::arg("hwp"), "메모 필드 삭제");
    actions.def("DeletePage", &cpyhwpx::HwpActionHelper::DeletePage, py::arg("hwp"), "페이지 삭제");
    actions.def("DeletePrivateInfoMark", &cpyhwpx::HwpActionHelper::DeletePrivateInfoMark, py::arg("hwp"), "개인정보 표시 삭제");
    actions.def("DeletePrivateInfoMarkAtCurrentPos", &cpyhwpx::HwpActionHelper::DeletePrivateInfoMarkAtCurrentPos, py::arg("hwp"), "현재 위치 개인정보 표시 삭제");
    actions.def("DeleteDocumentMasterPage", &cpyhwpx::HwpActionHelper::DeleteDocumentMasterPage, py::arg("hwp"), "문서 마스터 페이지 삭제");
    actions.def("DeleteSectionMasterPage", &cpyhwpx::HwpActionHelper::DeleteSectionMasterPage, py::arg("hwp"), "섹션 마스터 페이지 삭제");

    // 클립보드
    actions.def("Cut", &cpyhwpx::HwpActionHelper::Cut, py::arg("hwp"), "잘라내기");
    actions.def("Copy", &cpyhwpx::HwpActionHelper::Copy, py::arg("hwp"), "복사");
    actions.def("Paste", &cpyhwpx::HwpActionHelper::Paste, py::arg("hwp"), "붙여넣기");
    actions.def("PasteSpecial", &cpyhwpx::HwpActionHelper::PasteSpecial, py::arg("hwp"), "선택하여 붙여넣기");
    actions.def("CopyPage", &cpyhwpx::HwpActionHelper::CopyPage, py::arg("hwp"), "페이지 복사");
    actions.def("PastePage", &cpyhwpx::HwpActionHelper::PastePage, py::arg("hwp"), "페이지 붙여넣기");
    actions.def("Erase", &cpyhwpx::HwpActionHelper::Erase, py::arg("hwp"), "지우기");
    actions.def("Undo", &cpyhwpx::HwpActionHelper::Undo, py::arg("hwp"), "실행 취소");
    actions.def("Redo", &cpyhwpx::HwpActionHelper::Redo, py::arg("hwp"), "다시 실행");

    // 글자 모양 (기본)
    actions.def("CharShapeBold", &cpyhwpx::HwpActionHelper::CharShapeBold, py::arg("hwp"), "굵게");
    actions.def("CharShapeItalic", &cpyhwpx::HwpActionHelper::CharShapeItalic, py::arg("hwp"), "기울임");
    actions.def("CharShapeUnderline", &cpyhwpx::HwpActionHelper::CharShapeUnderline, py::arg("hwp"), "밑줄");
    actions.def("CharShapeStrikeout", &cpyhwpx::HwpActionHelper::CharShapeStrikeout, py::arg("hwp"), "취소선");
    actions.def("CharShapeSuperscript", &cpyhwpx::HwpActionHelper::CharShapeSuperscript, py::arg("hwp"), "위 첨자");
    actions.def("CharShapeSubscript", &cpyhwpx::HwpActionHelper::CharShapeSubscript, py::arg("hwp"), "아래 첨자");

    // 글자 모양 (추가)
    actions.def("CharShapeCenterline", &cpyhwpx::HwpActionHelper::CharShapeCenterline, py::arg("hwp"), "중앙선");
    actions.def("CharShapeEmboss", &cpyhwpx::HwpActionHelper::CharShapeEmboss, py::arg("hwp"), "양각");
    actions.def("CharShapeEngrave", &cpyhwpx::HwpActionHelper::CharShapeEngrave, py::arg("hwp"), "음각");
    actions.def("CharShapeHeight", &cpyhwpx::HwpActionHelper::CharShapeHeight, py::arg("hwp"), "글자 크기");
    actions.def("CharShapeHeightDecrease", &cpyhwpx::HwpActionHelper::CharShapeHeightDecrease, py::arg("hwp"), "글자 크기 감소");
    actions.def("CharShapeHeightIncrease", &cpyhwpx::HwpActionHelper::CharShapeHeightIncrease, py::arg("hwp"), "글자 크기 증가");
    actions.def("CharShapeLang", &cpyhwpx::HwpActionHelper::CharShapeLang, py::arg("hwp"), "언어");
    actions.def("CharShapeNextFaceName", &cpyhwpx::HwpActionHelper::CharShapeNextFaceName, py::arg("hwp"), "다음 글꼴");
    actions.def("CharShapeNormal", &cpyhwpx::HwpActionHelper::CharShapeNormal, py::arg("hwp"), "보통 모양");
    actions.def("CharShapeOutline", &cpyhwpx::HwpActionHelper::CharShapeOutline, py::arg("hwp"), "외곽선");
    actions.def("CharShapePrevFaceName", &cpyhwpx::HwpActionHelper::CharShapePrevFaceName, py::arg("hwp"), "이전 글꼴");
    actions.def("CharShapeShadow", &cpyhwpx::HwpActionHelper::CharShapeShadow, py::arg("hwp"), "그림자");
    actions.def("CharShapeSpacing", &cpyhwpx::HwpActionHelper::CharShapeSpacing, py::arg("hwp"), "자간");
    actions.def("CharShapeSpacingDecrease", &cpyhwpx::HwpActionHelper::CharShapeSpacingDecrease, py::arg("hwp"), "자간 감소");
    actions.def("CharShapeSpacingIncrease", &cpyhwpx::HwpActionHelper::CharShapeSpacingIncrease, py::arg("hwp"), "자간 증가");
    actions.def("CharShapeSuperSubscript", &cpyhwpx::HwpActionHelper::CharShapeSuperSubscript, py::arg("hwp"), "위/아래 첨자 토글");
    actions.def("CharShapeTextColorBlack", &cpyhwpx::HwpActionHelper::CharShapeTextColorBlack, py::arg("hwp"), "글자색 검정");
    actions.def("CharShapeTextColorBlue", &cpyhwpx::HwpActionHelper::CharShapeTextColorBlue, py::arg("hwp"), "글자색 파랑");
    actions.def("CharShapeTextColorBluish", &cpyhwpx::HwpActionHelper::CharShapeTextColorBluish, py::arg("hwp"), "글자색 청색");
    actions.def("CharShapeTextColorGreen", &cpyhwpx::HwpActionHelper::CharShapeTextColorGreen, py::arg("hwp"), "글자색 녹색");
    actions.def("CharShapeTextColorRed", &cpyhwpx::HwpActionHelper::CharShapeTextColorRed, py::arg("hwp"), "글자색 빨강");
    actions.def("CharShapeTextColorViolet", &cpyhwpx::HwpActionHelper::CharShapeTextColorViolet, py::arg("hwp"), "글자색 보라");
    actions.def("CharShapeTextColorWhite", &cpyhwpx::HwpActionHelper::CharShapeTextColorWhite, py::arg("hwp"), "글자색 흰색");
    actions.def("CharShapeTextColorYellow", &cpyhwpx::HwpActionHelper::CharShapeTextColorYellow, py::arg("hwp"), "글자색 노랑");
    actions.def("CharShapeTypeface", &cpyhwpx::HwpActionHelper::CharShapeTypeface, py::arg("hwp"), "글꼴");
    actions.def("CharShapeWidth", &cpyhwpx::HwpActionHelper::CharShapeWidth, py::arg("hwp"), "장평");
    actions.def("CharShapeWidthDecrease", &cpyhwpx::HwpActionHelper::CharShapeWidthDecrease, py::arg("hwp"), "장평 감소");
    actions.def("CharShapeWidthIncrease", &cpyhwpx::HwpActionHelper::CharShapeWidthIncrease, py::arg("hwp"), "장평 증가");

    // 문단 모양 (기본)
    actions.def("ParagraphShapeAlignLeft", &cpyhwpx::HwpActionHelper::ParagraphShapeAlignLeft, py::arg("hwp"), "왼쪽 정렬");
    actions.def("ParagraphShapeAlignCenter", &cpyhwpx::HwpActionHelper::ParagraphShapeAlignCenter, py::arg("hwp"), "가운데 정렬");
    actions.def("ParagraphShapeAlignRight", &cpyhwpx::HwpActionHelper::ParagraphShapeAlignRight, py::arg("hwp"), "오른쪽 정렬");
    actions.def("ParagraphShapeAlignJustify", &cpyhwpx::HwpActionHelper::ParagraphShapeAlignJustify, py::arg("hwp"), "양쪽 정렬");

    // 문단 모양 (추가)
    actions.def("ParagraphShapeAlignDistribute", &cpyhwpx::HwpActionHelper::ParagraphShapeAlignDistribute, py::arg("hwp"), "배분 정렬");
    actions.def("ParagraphShapeAlignDivision", &cpyhwpx::HwpActionHelper::ParagraphShapeAlignDivision, py::arg("hwp"), "나눔 정렬");
    actions.def("ParagraphShapeDecreaseLeftMargin", &cpyhwpx::HwpActionHelper::ParagraphShapeDecreaseLeftMargin, py::arg("hwp"), "왼쪽 여백 감소");
    actions.def("ParagraphShapeDecreaseLineSpacing", &cpyhwpx::HwpActionHelper::ParagraphShapeDecreaseLineSpacing, py::arg("hwp"), "줄간격 감소");
    actions.def("ParagraphShapeDecreaseMargin", &cpyhwpx::HwpActionHelper::ParagraphShapeDecreaseMargin, py::arg("hwp"), "여백 감소");
    actions.def("ParagraphShapeDecreaseRightMargin", &cpyhwpx::HwpActionHelper::ParagraphShapeDecreaseRightMargin, py::arg("hwp"), "오른쪽 여백 감소");
    actions.def("ParagraphShapeIncreaseLeftMargin", &cpyhwpx::HwpActionHelper::ParagraphShapeIncreaseLeftMargin, py::arg("hwp"), "왼쪽 여백 증가");
    actions.def("ParagraphShapeIncreaseLineSpacing", &cpyhwpx::HwpActionHelper::ParagraphShapeIncreaseLineSpacing, py::arg("hwp"), "줄간격 증가");
    actions.def("ParagraphShapeIncreaseMargin", &cpyhwpx::HwpActionHelper::ParagraphShapeIncreaseMargin, py::arg("hwp"), "여백 증가");
    actions.def("ParagraphShapeIncreaseRightMargin", &cpyhwpx::HwpActionHelper::ParagraphShapeIncreaseRightMargin, py::arg("hwp"), "오른쪽 여백 증가");
    actions.def("ParagraphShapeIndentAtCaret", &cpyhwpx::HwpActionHelper::ParagraphShapeIndentAtCaret, py::arg("hwp"), "캐럿 위치에서 들여쓰기");
    actions.def("ParagraphShapeIndentNegative", &cpyhwpx::HwpActionHelper::ParagraphShapeIndentNegative, py::arg("hwp"), "내어쓰기");
    actions.def("ParagraphShapeIndentPositive", &cpyhwpx::HwpActionHelper::ParagraphShapeIndentPositive, py::arg("hwp"), "들여쓰기");
    actions.def("ParagraphShapeProtect", &cpyhwpx::HwpActionHelper::ParagraphShapeProtect, py::arg("hwp"), "문단 보호");
    actions.def("ParagraphShapeSingleRow", &cpyhwpx::HwpActionHelper::ParagraphShapeSingleRow, py::arg("hwp"), "한 줄로");
    actions.def("ParagraphShapeWithNext", &cpyhwpx::HwpActionHelper::ParagraphShapeWithNext, py::arg("hwp"), "다음 문단과 함께");

    // 테이블 (기본)
    actions.def("TableCellBlock", &cpyhwpx::HwpActionHelper::TableCellBlock, py::arg("hwp"), "셀 블록 선택");
    actions.def("TableColBegin", &cpyhwpx::HwpActionHelper::TableColBegin, py::arg("hwp"), "첫 열로");
    actions.def("TableColEnd", &cpyhwpx::HwpActionHelper::TableColEnd, py::arg("hwp"), "마지막 열로");
    actions.def("TableRowBegin", &cpyhwpx::HwpActionHelper::TableRowBegin, py::arg("hwp"), "첫 행으로");
    actions.def("TableRowEnd", &cpyhwpx::HwpActionHelper::TableRowEnd, py::arg("hwp"), "마지막 행으로");
    actions.def("TableCellAddr", &cpyhwpx::HwpActionHelper::TableCellAddr, py::arg("hwp"), "셀 주소");
    actions.def("TableAppendRow", &cpyhwpx::HwpActionHelper::TableAppendRow, py::arg("hwp"), "행 추가");
    actions.def("TableAppendColumn", &cpyhwpx::HwpActionHelper::TableAppendColumn, py::arg("hwp"), "열 추가");
    actions.def("TableDeleteRow", &cpyhwpx::HwpActionHelper::TableDeleteRow, py::arg("hwp"), "행 삭제");
    actions.def("TableDeleteColumn", &cpyhwpx::HwpActionHelper::TableDeleteColumn, py::arg("hwp"), "열 삭제");
    actions.def("TableMergeCell", &cpyhwpx::HwpActionHelper::TableMergeCell, py::arg("hwp"), "셀 병합");
    actions.def("TableSplitCell", &cpyhwpx::HwpActionHelper::TableSplitCell, py::arg("hwp"), "셀 분할");
    actions.def("TableColPageUp", &cpyhwpx::HwpActionHelper::TableColPageUp, py::arg("hwp"), "열 맨 위로");
    actions.def("TableCellBlockExtendAbs", &cpyhwpx::HwpActionHelper::TableCellBlockExtendAbs, py::arg("hwp"), "셀 블록 확장 (절대)");

    // 테이블 (추가)
    actions.def("TableLeftCell", &cpyhwpx::HwpActionHelper::TableLeftCell, py::arg("hwp"), "왼쪽 셀로");
    actions.def("TableRightCell", &cpyhwpx::HwpActionHelper::TableRightCell, py::arg("hwp"), "오른쪽 셀로");
    actions.def("TableUpperCell", &cpyhwpx::HwpActionHelper::TableUpperCell, py::arg("hwp"), "위 셀로");
    actions.def("TableLowerCell", &cpyhwpx::HwpActionHelper::TableLowerCell, py::arg("hwp"), "아래 셀로");
    actions.def("TableColPageDown", &cpyhwpx::HwpActionHelper::TableColPageDown, py::arg("hwp"), "열 맨 아래로");
    actions.def("TableRightCellAppend", &cpyhwpx::HwpActionHelper::TableRightCellAppend, py::arg("hwp"), "오른쪽 셀로 (다음 행)");
    actions.def("TableCellBlockCol", &cpyhwpx::HwpActionHelper::TableCellBlockCol, py::arg("hwp"), "열 블록 선택");
    actions.def("TableCellBlockRow", &cpyhwpx::HwpActionHelper::TableCellBlockRow, py::arg("hwp"), "행 블록 선택");
    actions.def("TableCellBlockExtend", &cpyhwpx::HwpActionHelper::TableCellBlockExtend, py::arg("hwp"), "셀 블록 확장");

    // 테이블 셀 정렬
    actions.def("TableCellAlignLeftTop", &cpyhwpx::HwpActionHelper::TableCellAlignLeftTop, py::arg("hwp"), "왼쪽 위 정렬");
    actions.def("TableCellAlignCenterTop", &cpyhwpx::HwpActionHelper::TableCellAlignCenterTop, py::arg("hwp"), "가운데 위 정렬");
    actions.def("TableCellAlignRightTop", &cpyhwpx::HwpActionHelper::TableCellAlignRightTop, py::arg("hwp"), "오른쪽 위 정렬");
    actions.def("TableCellAlignLeftCenter", &cpyhwpx::HwpActionHelper::TableCellAlignLeftCenter, py::arg("hwp"), "왼쪽 가운데 정렬");
    actions.def("TableCellAlignCenterCenter", &cpyhwpx::HwpActionHelper::TableCellAlignCenterCenter, py::arg("hwp"), "가운데 가운데 정렬");
    actions.def("TableCellAlignRightCenter", &cpyhwpx::HwpActionHelper::TableCellAlignRightCenter, py::arg("hwp"), "오른쪽 가운데 정렬");
    actions.def("TableCellAlignLeftBottom", &cpyhwpx::HwpActionHelper::TableCellAlignLeftBottom, py::arg("hwp"), "왼쪽 아래 정렬");
    actions.def("TableCellAlignCenterBottom", &cpyhwpx::HwpActionHelper::TableCellAlignCenterBottom, py::arg("hwp"), "가운데 아래 정렬");
    actions.def("TableCellAlignRightBottom", &cpyhwpx::HwpActionHelper::TableCellAlignRightBottom, py::arg("hwp"), "오른쪽 아래 정렬");
    actions.def("TableVAlignTop", &cpyhwpx::HwpActionHelper::TableVAlignTop, py::arg("hwp"), "세로 위 정렬");
    actions.def("TableVAlignCenter", &cpyhwpx::HwpActionHelper::TableVAlignCenter, py::arg("hwp"), "세로 가운데 정렬");
    actions.def("TableVAlignBottom", &cpyhwpx::HwpActionHelper::TableVAlignBottom, py::arg("hwp"), "세로 아래 정렬");

    // 테이블 테두리
    actions.def("TableCellBorderAll", &cpyhwpx::HwpActionHelper::TableCellBorderAll, py::arg("hwp"), "모든 테두리");
    actions.def("TableCellBorderNo", &cpyhwpx::HwpActionHelper::TableCellBorderNo, py::arg("hwp"), "테두리 없음");
    actions.def("TableCellBorderOutside", &cpyhwpx::HwpActionHelper::TableCellBorderOutside, py::arg("hwp"), "바깥 테두리");
    actions.def("TableCellBorderInside", &cpyhwpx::HwpActionHelper::TableCellBorderInside, py::arg("hwp"), "안쪽 테두리");
    actions.def("TableCellBorderInsideHorz", &cpyhwpx::HwpActionHelper::TableCellBorderInsideHorz, py::arg("hwp"), "안쪽 가로 테두리");
    actions.def("TableCellBorderInsideVert", &cpyhwpx::HwpActionHelper::TableCellBorderInsideVert, py::arg("hwp"), "안쪽 세로 테두리");
    actions.def("TableCellBorderTop", &cpyhwpx::HwpActionHelper::TableCellBorderTop, py::arg("hwp"), "위 테두리");
    actions.def("TableCellBorderBottom", &cpyhwpx::HwpActionHelper::TableCellBorderBottom, py::arg("hwp"), "아래 테두리");
    actions.def("TableCellBorderLeft", &cpyhwpx::HwpActionHelper::TableCellBorderLeft, py::arg("hwp"), "왼쪽 테두리");
    actions.def("TableCellBorderRight", &cpyhwpx::HwpActionHelper::TableCellBorderRight, py::arg("hwp"), "오른쪽 테두리");
    actions.def("TableCellBorderDiagonalDown", &cpyhwpx::HwpActionHelper::TableCellBorderDiagonalDown, py::arg("hwp"), "대각선 테두리 (↘)");
    actions.def("TableCellBorderDiagonalUp", &cpyhwpx::HwpActionHelper::TableCellBorderDiagonalUp, py::arg("hwp"), "대각선 테두리 (↗)");

    // 테이블 행/열 관리
    actions.def("TableSubtractRow", &cpyhwpx::HwpActionHelper::TableSubtractRow, py::arg("hwp"), "행 빼기");
    actions.def("TableDeleteCell", &cpyhwpx::HwpActionHelper::TableDeleteCell, py::arg("hwp"), "셀 삭제");
    actions.def("TableMergeTable", &cpyhwpx::HwpActionHelper::TableMergeTable, py::arg("hwp"), "표 병합");
    actions.def("TableSplitTable", &cpyhwpx::HwpActionHelper::TableSplitTable, py::arg("hwp"), "표 분할");
    actions.def("TableDistributeCellHeight", &cpyhwpx::HwpActionHelper::TableDistributeCellHeight, py::arg("hwp"), "셀 높이 균등");
    actions.def("TableDistributeCellWidth", &cpyhwpx::HwpActionHelper::TableDistributeCellWidth, py::arg("hwp"), "셀 너비 균등");

    // 테이블 크기 조절
    actions.def("TableResizeDown", &cpyhwpx::HwpActionHelper::TableResizeDown, py::arg("hwp"), "표 아래로 크기 조절");
    actions.def("TableResizeUp", &cpyhwpx::HwpActionHelper::TableResizeUp, py::arg("hwp"), "표 위로 크기 조절");
    actions.def("TableResizeLeft", &cpyhwpx::HwpActionHelper::TableResizeLeft, py::arg("hwp"), "표 왼쪽으로 크기 조절");
    actions.def("TableResizeRight", &cpyhwpx::HwpActionHelper::TableResizeRight, py::arg("hwp"), "표 오른쪽으로 크기 조절");
    actions.def("TableResizeCellDown", &cpyhwpx::HwpActionHelper::TableResizeCellDown, py::arg("hwp"), "셀 아래로 크기 조절");
    actions.def("TableResizeCellUp", &cpyhwpx::HwpActionHelper::TableResizeCellUp, py::arg("hwp"), "셀 위로 크기 조절");
    actions.def("TableResizeCellLeft", &cpyhwpx::HwpActionHelper::TableResizeCellLeft, py::arg("hwp"), "셀 왼쪽으로 크기 조절");
    actions.def("TableResizeCellRight", &cpyhwpx::HwpActionHelper::TableResizeCellRight, py::arg("hwp"), "셀 오른쪽으로 크기 조절");
    actions.def("TableResizeExDown", &cpyhwpx::HwpActionHelper::TableResizeExDown, py::arg("hwp"), "표 확장 아래로");
    actions.def("TableResizeExUp", &cpyhwpx::HwpActionHelper::TableResizeExUp, py::arg("hwp"), "표 확장 위로");
    actions.def("TableResizeExLeft", &cpyhwpx::HwpActionHelper::TableResizeExLeft, py::arg("hwp"), "표 확장 왼쪽으로");
    actions.def("TableResizeExRight", &cpyhwpx::HwpActionHelper::TableResizeExRight, py::arg("hwp"), "표 확장 오른쪽으로");
    actions.def("TableResizeLineDown", &cpyhwpx::HwpActionHelper::TableResizeLineDown, py::arg("hwp"), "줄 아래로 크기 조절");
    actions.def("TableResizeLineUp", &cpyhwpx::HwpActionHelper::TableResizeLineUp, py::arg("hwp"), "줄 위로 크기 조절");
    actions.def("TableResizeLineLeft", &cpyhwpx::HwpActionHelper::TableResizeLineLeft, py::arg("hwp"), "줄 왼쪽으로 크기 조절");
    actions.def("TableResizeLineRight", &cpyhwpx::HwpActionHelper::TableResizeLineRight, py::arg("hwp"), "줄 오른쪽으로 크기 조절");

    // 테이블 수식
    actions.def("TableFormulaSumAuto", &cpyhwpx::HwpActionHelper::TableFormulaSumAuto, py::arg("hwp"), "합계 (자동)");
    actions.def("TableFormulaSumHor", &cpyhwpx::HwpActionHelper::TableFormulaSumHor, py::arg("hwp"), "합계 (가로)");
    actions.def("TableFormulaSumVer", &cpyhwpx::HwpActionHelper::TableFormulaSumVer, py::arg("hwp"), "합계 (세로)");
    actions.def("TableFormulaAvgAuto", &cpyhwpx::HwpActionHelper::TableFormulaAvgAuto, py::arg("hwp"), "평균 (자동)");
    actions.def("TableFormulaAvgHor", &cpyhwpx::HwpActionHelper::TableFormulaAvgHor, py::arg("hwp"), "평균 (가로)");
    actions.def("TableFormulaAvgVer", &cpyhwpx::HwpActionHelper::TableFormulaAvgVer, py::arg("hwp"), "평균 (세로)");
    actions.def("TableFormulaProAuto", &cpyhwpx::HwpActionHelper::TableFormulaProAuto, py::arg("hwp"), "곱 (자동)");
    actions.def("TableFormulaProHor", &cpyhwpx::HwpActionHelper::TableFormulaProHor, py::arg("hwp"), "곱 (가로)");
    actions.def("TableFormulaProVer", &cpyhwpx::HwpActionHelper::TableFormulaProVer, py::arg("hwp"), "곱 (세로)");

    // 테이블 도구
    actions.def("TableAutoFill", &cpyhwpx::HwpActionHelper::TableAutoFill, py::arg("hwp"), "자동 채우기");
    actions.def("TableAutoFillDlg", &cpyhwpx::HwpActionHelper::TableAutoFillDlg, py::arg("hwp"), "자동 채우기 대화상자");
    actions.def("TableDrawPen", &cpyhwpx::HwpActionHelper::TableDrawPen, py::arg("hwp"), "표 그리기 펜");
    actions.def("TableDrawPenStyle", &cpyhwpx::HwpActionHelper::TableDrawPenStyle, py::arg("hwp"), "펜 스타일");
    actions.def("TableDrawPenWidth", &cpyhwpx::HwpActionHelper::TableDrawPenWidth, py::arg("hwp"), "펜 너비");
    actions.def("TableEraser", &cpyhwpx::HwpActionHelper::TableEraser, py::arg("hwp"), "표 지우개");
    actions.def("TableDeleteComma", &cpyhwpx::HwpActionHelper::TableDeleteComma, py::arg("hwp"), "쉼표 삭제");
    actions.def("TableInsertComma", &cpyhwpx::HwpActionHelper::TableInsertComma, py::arg("hwp"), "쉼표 삽입");

    // 개체 (기본)
    actions.def("ShapeObjSelect", &cpyhwpx::HwpActionHelper::ShapeObjSelect, py::arg("hwp"), "개체 선택");
    actions.def("ShapeObjDelete", &cpyhwpx::HwpActionHelper::ShapeObjDelete, py::arg("hwp"), "개체 삭제");
    actions.def("ShapeObjCopy", &cpyhwpx::HwpActionHelper::ShapeObjCopy, py::arg("hwp"), "개체 복사");
    actions.def("ShapeObjCut", &cpyhwpx::HwpActionHelper::ShapeObjCut, py::arg("hwp"), "개체 잘라내기");
    actions.def("ShapeObjBringToFront", &cpyhwpx::HwpActionHelper::ShapeObjBringToFront, py::arg("hwp"), "맨 앞으로");
    actions.def("ShapeObjSendToBack", &cpyhwpx::HwpActionHelper::ShapeObjSendToBack, py::arg("hwp"), "맨 뒤로");

    // 개체 정렬
    actions.def("ShapeObjAlignLeft", &cpyhwpx::HwpActionHelper::ShapeObjAlignLeft, py::arg("hwp"), "개체 왼쪽 정렬");
    actions.def("ShapeObjAlignCenter", &cpyhwpx::HwpActionHelper::ShapeObjAlignCenter, py::arg("hwp"), "개체 가로 가운데 정렬");
    actions.def("ShapeObjAlignRight", &cpyhwpx::HwpActionHelper::ShapeObjAlignRight, py::arg("hwp"), "개체 오른쪽 정렬");
    actions.def("ShapeObjAlignTop", &cpyhwpx::HwpActionHelper::ShapeObjAlignTop, py::arg("hwp"), "개체 위쪽 정렬");
    actions.def("ShapeObjAlignMiddle", &cpyhwpx::HwpActionHelper::ShapeObjAlignMiddle, py::arg("hwp"), "개체 세로 가운데 정렬");
    actions.def("ShapeObjAlignBottom", &cpyhwpx::HwpActionHelper::ShapeObjAlignBottom, py::arg("hwp"), "개체 아래쪽 정렬");
    actions.def("ShapeObjAlignWidth", &cpyhwpx::HwpActionHelper::ShapeObjAlignWidth, py::arg("hwp"), "개체 너비 맞춤");
    actions.def("ShapeObjAlignHeight", &cpyhwpx::HwpActionHelper::ShapeObjAlignHeight, py::arg("hwp"), "개체 높이 맞춤");
    actions.def("ShapeObjAlignSize", &cpyhwpx::HwpActionHelper::ShapeObjAlignSize, py::arg("hwp"), "개체 크기 맞춤");
    actions.def("ShapeObjAlignHorzSpacing", &cpyhwpx::HwpActionHelper::ShapeObjAlignHorzSpacing, py::arg("hwp"), "가로 간격 균등");
    actions.def("ShapeObjAlignVertSpacing", &cpyhwpx::HwpActionHelper::ShapeObjAlignVertSpacing, py::arg("hwp"), "세로 간격 균등");

    // 개체 순서
    actions.def("ShapeObjBringForward", &cpyhwpx::HwpActionHelper::ShapeObjBringForward, py::arg("hwp"), "앞으로");
    actions.def("ShapeObjSendBack", &cpyhwpx::HwpActionHelper::ShapeObjSendBack, py::arg("hwp"), "뒤로");
    actions.def("ShapeObjBringInFrontOfText", &cpyhwpx::HwpActionHelper::ShapeObjBringInFrontOfText, py::arg("hwp"), "글 앞으로");
    actions.def("ShapeObjCtrlSendBehindText", &cpyhwpx::HwpActionHelper::ShapeObjCtrlSendBehindText, py::arg("hwp"), "글 뒤로");

    // 개체 뒤집기/회전
    actions.def("ShapeObjHorzFlip", &cpyhwpx::HwpActionHelper::ShapeObjHorzFlip, py::arg("hwp"), "좌우 뒤집기");
    actions.def("ShapeObjVertFlip", &cpyhwpx::HwpActionHelper::ShapeObjVertFlip, py::arg("hwp"), "상하 뒤집기");
    actions.def("ShapeObjHorzFlipOrgState", &cpyhwpx::HwpActionHelper::ShapeObjHorzFlipOrgState, py::arg("hwp"), "좌우 원래 상태");
    actions.def("ShapeObjVertFlipOrgState", &cpyhwpx::HwpActionHelper::ShapeObjVertFlipOrgState, py::arg("hwp"), "상하 원래 상태");
    actions.def("ShapeObjRotater", &cpyhwpx::HwpActionHelper::ShapeObjRotater, py::arg("hwp"), "회전");
    actions.def("ShapeObjRightAngleRotater", &cpyhwpx::HwpActionHelper::ShapeObjRightAngleRotater, py::arg("hwp"), "90도 회전");
    actions.def("ShapeObjRightAngleRotaterAnticlockwise", &cpyhwpx::HwpActionHelper::ShapeObjRightAngleRotaterAnticlockwise, py::arg("hwp"), "90도 반시계 회전");

    // 개체 이동/크기
    actions.def("ShapeObjMoveUp", &cpyhwpx::HwpActionHelper::ShapeObjMoveUp, py::arg("hwp"), "개체 위로 이동");
    actions.def("ShapeObjMoveDown", &cpyhwpx::HwpActionHelper::ShapeObjMoveDown, py::arg("hwp"), "개체 아래로 이동");
    actions.def("ShapeObjMoveLeft", &cpyhwpx::HwpActionHelper::ShapeObjMoveLeft, py::arg("hwp"), "개체 왼쪽으로 이동");
    actions.def("ShapeObjMoveRight", &cpyhwpx::HwpActionHelper::ShapeObjMoveRight, py::arg("hwp"), "개체 오른쪽으로 이동");
    actions.def("ShapeObjResizeUp", &cpyhwpx::HwpActionHelper::ShapeObjResizeUp, py::arg("hwp"), "개체 위로 크기 조절");
    actions.def("ShapeObjResizeDown", &cpyhwpx::HwpActionHelper::ShapeObjResizeDown, py::arg("hwp"), "개체 아래로 크기 조절");
    actions.def("ShapeObjResizeLeft", &cpyhwpx::HwpActionHelper::ShapeObjResizeLeft, py::arg("hwp"), "개체 왼쪽으로 크기 조절");
    actions.def("ShapeObjResizeRight", &cpyhwpx::HwpActionHelper::ShapeObjResizeRight, py::arg("hwp"), "개체 오른쪽으로 크기 조절");

    // 개체 그룹화
    actions.def("ShapeObjGroup", &cpyhwpx::HwpActionHelper::ShapeObjGroup, py::arg("hwp"), "개체 그룹화");
    actions.def("ShapeObjUngroup", &cpyhwpx::HwpActionHelper::ShapeObjUngroup, py::arg("hwp"), "개체 그룹 해제");
    actions.def("ShapeObjNextObject", &cpyhwpx::HwpActionHelper::ShapeObjNextObject, py::arg("hwp"), "다음 개체");
    actions.def("ShapeObjPrevObject", &cpyhwpx::HwpActionHelper::ShapeObjPrevObject, py::arg("hwp"), "이전 개체");
    actions.def("ShapeObjLock", &cpyhwpx::HwpActionHelper::ShapeObjLock, py::arg("hwp"), "개체 잠금");
    actions.def("ShapeObjUnlockAll", &cpyhwpx::HwpActionHelper::ShapeObjUnlockAll, py::arg("hwp"), "개체 잠금 해제");

    // 개체 캡션/글상자
    actions.def("ShapeObjAttachCaption", &cpyhwpx::HwpActionHelper::ShapeObjAttachCaption, py::arg("hwp"), "캡션 달기");
    actions.def("ShapeObjDetachCaption", &cpyhwpx::HwpActionHelper::ShapeObjDetachCaption, py::arg("hwp"), "캡션 떼기");
    actions.def("ShapeObjInsertCaptionNum", &cpyhwpx::HwpActionHelper::ShapeObjInsertCaptionNum, py::arg("hwp"), "캡션 번호 삽입");
    actions.def("ShapeObjAttachTextBox", &cpyhwpx::HwpActionHelper::ShapeObjAttachTextBox, py::arg("hwp"), "글상자 달기");
    actions.def("ShapeObjDetachTextBox", &cpyhwpx::HwpActionHelper::ShapeObjDetachTextBox, py::arg("hwp"), "글상자 떼기");
    actions.def("ShapeObjToggleTextBox", &cpyhwpx::HwpActionHelper::ShapeObjToggleTextBox, py::arg("hwp"), "글상자 토글");
    actions.def("ShapeObjTextBoxEdit", &cpyhwpx::HwpActionHelper::ShapeObjTextBoxEdit, py::arg("hwp"), "글상자 편집");
    actions.def("ShapeObjTableSelCell", &cpyhwpx::HwpActionHelper::ShapeObjTableSelCell, py::arg("hwp"), "표 셀 선택");

    // 개체 속성
    actions.def("ShapeObjFillProperty", &cpyhwpx::HwpActionHelper::ShapeObjFillProperty, py::arg("hwp"), "채우기 속성");
    actions.def("ShapeObjLineProperty", &cpyhwpx::HwpActionHelper::ShapeObjLineProperty, py::arg("hwp"), "선 속성");
    actions.def("ShapeObjLineStyleOther", &cpyhwpx::HwpActionHelper::ShapeObjLineStyleOther, py::arg("hwp"), "기타 선 스타일");
    actions.def("ShapeObjLineWidthOther", &cpyhwpx::HwpActionHelper::ShapeObjLineWidthOther, py::arg("hwp"), "기타 선 너비");
    actions.def("ShapeObjNorm", &cpyhwpx::HwpActionHelper::ShapeObjNorm, py::arg("hwp"), "개체 표준");
    actions.def("ShapeObjGuideLine", &cpyhwpx::HwpActionHelper::ShapeObjGuideLine, py::arg("hwp"), "안내선");
    actions.def("ShapeObjShowGuideLine", &cpyhwpx::HwpActionHelper::ShapeObjShowGuideLine, py::arg("hwp"), "안내선 표시");
    actions.def("ShapeObjShowGuideLineBase", &cpyhwpx::HwpActionHelper::ShapeObjShowGuideLineBase, py::arg("hwp"), "기준 안내선 표시");
    actions.def("ShapeObjWrapSquare", &cpyhwpx::HwpActionHelper::ShapeObjWrapSquare, py::arg("hwp"), "글 배치 사각형");
    actions.def("ShapeObjWrapTopAndBottom", &cpyhwpx::HwpActionHelper::ShapeObjWrapTopAndBottom, py::arg("hwp"), "글 배치 위아래");
    actions.def("ShapeObjSaveAsPicture", &cpyhwpx::HwpActionHelper::ShapeObjSaveAsPicture, py::arg("hwp"), "그림으로 저장");

    // 파일
    actions.def("FileNew", &cpyhwpx::HwpActionHelper::FileNew, py::arg("hwp"), "새 문서");
    actions.def("FileNewTab", &cpyhwpx::HwpActionHelper::FileNewTab, py::arg("hwp"), "새 탭");
    actions.def("FileOpen", &cpyhwpx::HwpActionHelper::FileOpen, py::arg("hwp"), "파일 열기");
    actions.def("FileOpenMRU", &cpyhwpx::HwpActionHelper::FileOpenMRU, py::arg("hwp"), "최근 문서 열기");
    actions.def("FileSave", &cpyhwpx::HwpActionHelper::FileSave, py::arg("hwp"), "저장");
    actions.def("FileSaveAs", &cpyhwpx::HwpActionHelper::FileSaveAs, py::arg("hwp"), "다른 이름으로 저장");
    actions.def("FileSaveAsDRM", &cpyhwpx::HwpActionHelper::FileSaveAsDRM, py::arg("hwp"), "DRM으로 저장");
    actions.def("FileSaveOptionDlg", &cpyhwpx::HwpActionHelper::FileSaveOptionDlg, py::arg("hwp"), "저장 옵션 대화상자");
    actions.def("FileFind", &cpyhwpx::HwpActionHelper::FileFind, py::arg("hwp"), "파일 찾기");
    actions.def("FilePreview", &cpyhwpx::HwpActionHelper::FilePreview, py::arg("hwp"), "파일 미리보기");
    actions.def("FileClose", &cpyhwpx::HwpActionHelper::FileClose, py::arg("hwp"), "파일 닫기");
    actions.def("FileQuit", &cpyhwpx::HwpActionHelper::FileQuit, py::arg("hwp"), "종료");
    actions.def("FilePrint", &cpyhwpx::HwpActionHelper::FilePrint, py::arg("hwp"), "인쇄");
    actions.def("FilePrintPreview", &cpyhwpx::HwpActionHelper::FilePrintPreview, py::arg("hwp"), "인쇄 미리보기");
    actions.def("FileNextVersionDiff", &cpyhwpx::HwpActionHelper::FileNextVersionDiff, py::arg("hwp"), "다음 버전 비교");
    actions.def("FilePrevVersionDiff", &cpyhwpx::HwpActionHelper::FilePrevVersionDiff, py::arg("hwp"), "이전 버전 비교");
    actions.def("FileVersionDiffChangeAlign", &cpyhwpx::HwpActionHelper::FileVersionDiffChangeAlign, py::arg("hwp"), "버전 비교 변경 정렬");
    actions.def("FileVersionDiffSameAlign", &cpyhwpx::HwpActionHelper::FileVersionDiffSameAlign, py::arg("hwp"), "버전 비교 동일 정렬");
    actions.def("FileVersionDiffSyncScroll", &cpyhwpx::HwpActionHelper::FileVersionDiffSyncScroll, py::arg("hwp"), "버전 비교 동기 스크롤");

    // 삽입
    actions.def("InsertAutoNum", &cpyhwpx::HwpActionHelper::InsertAutoNum, py::arg("hwp"), "자동 번호 삽입");
    actions.def("InsertCpNo", &cpyhwpx::HwpActionHelper::InsertCpNo, py::arg("hwp"), "쪽 번호 삽입");
    actions.def("InsertCpTpNo", &cpyhwpx::HwpActionHelper::InsertCpTpNo, py::arg("hwp"), "쪽/전체쪽 번호 삽입");
    actions.def("InsertTpNo", &cpyhwpx::HwpActionHelper::InsertTpNo, py::arg("hwp"), "전체 쪽 번호 삽입");
    actions.def("InsertPageNum", &cpyhwpx::HwpActionHelper::InsertPageNum, py::arg("hwp"), "페이지 번호 삽입");
    actions.def("InsertDateCode", &cpyhwpx::HwpActionHelper::InsertDateCode, py::arg("hwp"), "날짜 코드 삽입");
    actions.def("InsertDocInfo", &cpyhwpx::HwpActionHelper::InsertDocInfo, py::arg("hwp"), "문서 정보 삽입");
    actions.def("InsertEndnote", &cpyhwpx::HwpActionHelper::InsertEndnote, py::arg("hwp"), "미주 삽입");
    actions.def("InsertFootnote", &cpyhwpx::HwpActionHelper::InsertFootnote, py::arg("hwp"), "각주 삽입");
    actions.def("InsertFieldCitation", &cpyhwpx::HwpActionHelper::InsertFieldCitation, py::arg("hwp"), "인용 필드 삽입");
    actions.def("InsertFieldDateTime", &cpyhwpx::HwpActionHelper::InsertFieldDateTime, py::arg("hwp"), "날짜/시간 필드 삽입");
    actions.def("InsertFieldMemo", &cpyhwpx::HwpActionHelper::InsertFieldMemo, py::arg("hwp"), "메모 필드 삽입");
    actions.def("InsertFieldRevisionChagne", &cpyhwpx::HwpActionHelper::InsertFieldRevisionChagne, py::arg("hwp"), "수정 필드 삽입");
    actions.def("InsertFixedWidthSpace", &cpyhwpx::HwpActionHelper::InsertFixedWidthSpace, py::arg("hwp"), "고정폭 빈칸 삽입");
    actions.def("InsertNonBreakingSpace", &cpyhwpx::HwpActionHelper::InsertNonBreakingSpace, py::arg("hwp"), "묶음 빈칸 삽입");
    actions.def("InsertSoftHyphen", &cpyhwpx::HwpActionHelper::InsertSoftHyphen, py::arg("hwp"), "소프트 하이픈 삽입");
    actions.def("InsertSpace", &cpyhwpx::HwpActionHelper::InsertSpace, py::arg("hwp"), "빈칸 삽입");
    actions.def("InsertTab", &cpyhwpx::HwpActionHelper::InsertTab, py::arg("hwp"), "탭 삽입");
    actions.def("InsertLine", &cpyhwpx::HwpActionHelper::InsertLine, py::arg("hwp"), "선 삽입");
    actions.def("InsertStringDateTime", &cpyhwpx::HwpActionHelper::InsertStringDateTime, py::arg("hwp"), "날짜/시간 문자열 삽입");
    actions.def("InsertLastPrintDate", &cpyhwpx::HwpActionHelper::InsertLastPrintDate, py::arg("hwp"), "마지막 인쇄일 삽입");
    actions.def("InsertLastSaveBy", &cpyhwpx::HwpActionHelper::InsertLastSaveBy, py::arg("hwp"), "마지막 저장자 삽입");
    actions.def("InsertLastSaveDate", &cpyhwpx::HwpActionHelper::InsertLastSaveDate, py::arg("hwp"), "마지막 저장일 삽입");

    // 기타 핵심
    actions.def("Close", &cpyhwpx::HwpActionHelper::Close, py::arg("hwp"), "닫기");
    actions.def("CloseEx", &cpyhwpx::HwpActionHelper::CloseEx, py::arg("hwp"), "닫기 (확장)");
    actions.def("ToggleOverwrite", &cpyhwpx::HwpActionHelper::ToggleOverwrite, py::arg("hwp"), "삽입/덮어쓰기 토글");
    actions.def("ReturnPrevPos", &cpyhwpx::HwpActionHelper::ReturnPrevPos, py::arg("hwp"), "이전 위치로 돌아가기");
    actions.def("RecalcPageCount", &cpyhwpx::HwpActionHelper::RecalcPageCount, py::arg("hwp"), "페이지 수 다시 계산");
    actions.def("SpellingCheck", &cpyhwpx::HwpActionHelper::SpellingCheck, py::arg("hwp"), "맞춤법 검사");
    actions.def("EasyFind", &cpyhwpx::HwpActionHelper::EasyFind, py::arg("hwp"), "빠른 찾기");

    // 변경 추적
    actions.def("TrackChangeApply", &cpyhwpx::HwpActionHelper::TrackChangeApply, py::arg("hwp"), "변경 적용");
    actions.def("TrackChangeApplyAll", &cpyhwpx::HwpActionHelper::TrackChangeApplyAll, py::arg("hwp"), "모든 변경 적용");
    actions.def("TrackChangeApplyNext", &cpyhwpx::HwpActionHelper::TrackChangeApplyNext, py::arg("hwp"), "다음 변경 적용");
    actions.def("TrackChangeApplyPrev", &cpyhwpx::HwpActionHelper::TrackChangeApplyPrev, py::arg("hwp"), "이전 변경 적용");
    actions.def("TrackChangeApplyViewAll", &cpyhwpx::HwpActionHelper::TrackChangeApplyViewAll, py::arg("hwp"), "모든 변경 보기 적용");
    actions.def("TrackChangeCancel", &cpyhwpx::HwpActionHelper::TrackChangeCancel, py::arg("hwp"), "변경 취소");
    actions.def("TrackChangeCancelAll", &cpyhwpx::HwpActionHelper::TrackChangeCancelAll, py::arg("hwp"), "모든 변경 취소");
    actions.def("TrackChangeCancelNext", &cpyhwpx::HwpActionHelper::TrackChangeCancelNext, py::arg("hwp"), "다음 변경 취소");
    actions.def("TrackChangeCancelPrev", &cpyhwpx::HwpActionHelper::TrackChangeCancelPrev, py::arg("hwp"), "이전 변경 취소");
    actions.def("TrackChangeCancelViewAll", &cpyhwpx::HwpActionHelper::TrackChangeCancelViewAll, py::arg("hwp"), "모든 변경 보기 취소");
    actions.def("TrackChangeNext", &cpyhwpx::HwpActionHelper::TrackChangeNext, py::arg("hwp"), "다음 변경");
    actions.def("TrackChangePrev", &cpyhwpx::HwpActionHelper::TrackChangePrev, py::arg("hwp"), "이전 변경");
    actions.def("TrackChangeAuthor", &cpyhwpx::HwpActionHelper::TrackChangeAuthor, py::arg("hwp"), "변경 작성자");

    // 보기 옵션
    actions.def("ViewOptionCtrlMark", &cpyhwpx::HwpActionHelper::ViewOptionCtrlMark, py::arg("hwp"), "조판 부호 보기");
    actions.def("ViewOptionGuideLine", &cpyhwpx::HwpActionHelper::ViewOptionGuideLine, py::arg("hwp"), "안내선 보기");
    actions.def("ViewOptionMemo", &cpyhwpx::HwpActionHelper::ViewOptionMemo, py::arg("hwp"), "메모 보기");
    actions.def("ViewOptionPaper", &cpyhwpx::HwpActionHelper::ViewOptionPaper, py::arg("hwp"), "용지 보기");
    actions.def("ViewOptionParaMark", &cpyhwpx::HwpActionHelper::ViewOptionParaMark, py::arg("hwp"), "문단 부호 보기");
    actions.def("ViewOptionPicture", &cpyhwpx::HwpActionHelper::ViewOptionPicture, py::arg("hwp"), "그림 보기");
    actions.def("ViewZoomIn", &cpyhwpx::HwpActionHelper::ViewZoomIn, py::arg("hwp"), "확대");
    actions.def("ViewZoomOut", &cpyhwpx::HwpActionHelper::ViewZoomOut, py::arg("hwp"), "축소");
    actions.def("ViewZoomFitPage", &cpyhwpx::HwpActionHelper::ViewZoomFitPage, py::arg("hwp"), "페이지 맞춤");
    actions.def("ViewZoomFitWidth", &cpyhwpx::HwpActionHelper::ViewZoomFitWidth, py::arg("hwp"), "폭 맞춤");
    actions.def("ViewTabButton", &cpyhwpx::HwpActionHelper::ViewTabButton, py::arg("hwp"), "탭 버튼 보기");

    // 머리글/꼬리글
    actions.def("HeaderFooterDelete", &cpyhwpx::HwpActionHelper::HeaderFooterDelete, py::arg("hwp"), "머리글/꼬리글 삭제");
    actions.def("HeaderFooterModify", &cpyhwpx::HwpActionHelper::HeaderFooterModify, py::arg("hwp"), "머리글/꼬리글 수정");
    actions.def("HeaderFooterToNext", &cpyhwpx::HwpActionHelper::HeaderFooterToNext, py::arg("hwp"), "다음 머리글/꼬리글");
    actions.def("HeaderFooterToPrev", &cpyhwpx::HwpActionHelper::HeaderFooterToPrev, py::arg("hwp"), "이전 머리글/꼬리글");

    // 각주/미주/메모
    actions.def("NoteDelete", &cpyhwpx::HwpActionHelper::NoteDelete, py::arg("hwp"), "각주/미주 삭제");
    actions.def("NoteModify", &cpyhwpx::HwpActionHelper::NoteModify, py::arg("hwp"), "각주/미주 수정");
    actions.def("NoteToNext", &cpyhwpx::HwpActionHelper::NoteToNext, py::arg("hwp"), "다음 각주/미주");
    actions.def("NoteToPrev", &cpyhwpx::HwpActionHelper::NoteToPrev, py::arg("hwp"), "이전 각주/미주");
    actions.def("MemoToNext", &cpyhwpx::HwpActionHelper::MemoToNext, py::arg("hwp"), "다음 메모");
    actions.def("MemoToPrev", &cpyhwpx::HwpActionHelper::MemoToPrev, py::arg("hwp"), "이전 메모");
    actions.def("Comment", &cpyhwpx::HwpActionHelper::Comment, py::arg("hwp"), "덧글");
    actions.def("CommentDelete", &cpyhwpx::HwpActionHelper::CommentDelete, py::arg("hwp"), "덧글 삭제");
    actions.def("CommentModify", &cpyhwpx::HwpActionHelper::CommentModify, py::arg("hwp"), "덧글 수정");
    actions.def("ReplyMemo", &cpyhwpx::HwpActionHelper::ReplyMemo, py::arg("hwp"), "메모 답글");

    // 마스터 페이지
    actions.def("MasterPage", &cpyhwpx::HwpActionHelper::MasterPage, py::arg("hwp"), "마스터 페이지");
    actions.def("MasterPageDuplicate", &cpyhwpx::HwpActionHelper::MasterPageDuplicate, py::arg("hwp"), "마스터 페이지 복제");
    actions.def("MasterPageExcept", &cpyhwpx::HwpActionHelper::MasterPageExcept, py::arg("hwp"), "마스터 페이지 제외");
    actions.def("MasterPageFront", &cpyhwpx::HwpActionHelper::MasterPageFront, py::arg("hwp"), "마스터 페이지 앞면");
    actions.def("MasterPageToNext", &cpyhwpx::HwpActionHelper::MasterPageToNext, py::arg("hwp"), "다음 마스터 페이지");
    actions.def("MasterPageToPrevious", &cpyhwpx::HwpActionHelper::MasterPageToPrevious, py::arg("hwp"), "이전 마스터 페이지");

    // 그림
    actions.def("PictureInsertDialog", &cpyhwpx::HwpActionHelper::PictureInsertDialog, py::arg("hwp"), "그림 삽입 대화상자");
    actions.def("PictureLinkedToEmbedded", &cpyhwpx::HwpActionHelper::PictureLinkedToEmbedded, py::arg("hwp"), "연결→포함 그림");
    actions.def("PictureSave", &cpyhwpx::HwpActionHelper::PictureSave, py::arg("hwp"), "그림 저장");
    actions.def("PictureScissor", &cpyhwpx::HwpActionHelper::PictureScissor, py::arg("hwp"), "그림 자르기");
    actions.def("PictureToOriginal", &cpyhwpx::HwpActionHelper::PictureToOriginal, py::arg("hwp"), "그림 원래 상태");

    // 창/프레임
    actions.def("FrameFullScreen", &cpyhwpx::HwpActionHelper::FrameFullScreen, py::arg("hwp"), "전체 화면");
    actions.def("FrameFullScreenEnd", &cpyhwpx::HwpActionHelper::FrameFullScreenEnd, py::arg("hwp"), "전체 화면 종료");
    actions.def("FrameHRuler", &cpyhwpx::HwpActionHelper::FrameHRuler, py::arg("hwp"), "가로 눈금자");
    actions.def("FrameVRuler", &cpyhwpx::HwpActionHelper::FrameVRuler, py::arg("hwp"), "세로 눈금자");
    actions.def("FrameStatusBar", &cpyhwpx::HwpActionHelper::FrameStatusBar, py::arg("hwp"), "상태 표시줄");
    actions.def("WindowAlignCascade", &cpyhwpx::HwpActionHelper::WindowAlignCascade, py::arg("hwp"), "창 계단식 배열");
    actions.def("WindowAlignTileHorz", &cpyhwpx::HwpActionHelper::WindowAlignTileHorz, py::arg("hwp"), "창 가로 배열");
    actions.def("WindowAlignTileVert", &cpyhwpx::HwpActionHelper::WindowAlignTileVert, py::arg("hwp"), "창 세로 배열");
    actions.def("WindowList", &cpyhwpx::HwpActionHelper::WindowList, py::arg("hwp"), "창 목록");
    actions.def("WindowMaximize", &cpyhwpx::HwpActionHelper::WindowMaximize, py::arg("hwp"), "창 최대화");
    actions.def("WindowMinimize", &cpyhwpx::HwpActionHelper::WindowMinimize, py::arg("hwp"), "창 최소화");
    actions.def("WindowMinimizeAll", &cpyhwpx::HwpActionHelper::WindowMinimizeAll, py::arg("hwp"), "모든 창 최소화");
    actions.def("WindowNextPane", &cpyhwpx::HwpActionHelper::WindowNextPane, py::arg("hwp"), "다음 창");
    actions.def("WindowNextTab", &cpyhwpx::HwpActionHelper::WindowNextTab, py::arg("hwp"), "다음 탭");
    actions.def("WindowPrevTab", &cpyhwpx::HwpActionHelper::WindowPrevTab, py::arg("hwp"), "이전 탭");
    actions.def("SplitAll", &cpyhwpx::HwpActionHelper::SplitAll, py::arg("hwp"), "전체 분할");
    actions.def("SplitHorz", &cpyhwpx::HwpActionHelper::SplitHorz, py::arg("hwp"), "가로 분할");
    actions.def("SplitVert", &cpyhwpx::HwpActionHelper::SplitVert, py::arg("hwp"), "세로 분할");
    actions.def("NoSplit", &cpyhwpx::HwpActionHelper::NoSplit, py::arg("hwp"), "분할 해제");

    // 입력
    actions.def("InputCodeChange", &cpyhwpx::HwpActionHelper::InputCodeChange, py::arg("hwp"), "코드 변환 입력");
    actions.def("InputHanja", &cpyhwpx::HwpActionHelper::InputHanja, py::arg("hwp"), "한자 입력");
    actions.def("InputHanjaBusu", &cpyhwpx::HwpActionHelper::InputHanjaBusu, py::arg("hwp"), "한자 부수 입력");
    actions.def("InputHanjaMean", &cpyhwpx::HwpActionHelper::InputHanjaMean, py::arg("hwp"), "한자 뜻 입력");

    // 양식 개체
    actions.def("FormDesignMode", &cpyhwpx::HwpActionHelper::FormDesignMode, py::arg("hwp"), "양식 디자인 모드");
    actions.def("FormObjCreatorCheckButton", &cpyhwpx::HwpActionHelper::FormObjCreatorCheckButton, py::arg("hwp"), "체크 버튼 생성");
    actions.def("FormObjCreatorComboBox", &cpyhwpx::HwpActionHelper::FormObjCreatorComboBox, py::arg("hwp"), "콤보 박스 생성");
    actions.def("FormObjCreatorEdit", &cpyhwpx::HwpActionHelper::FormObjCreatorEdit, py::arg("hwp"), "편집 상자 생성");
    actions.def("FormObjCreatorListBox", &cpyhwpx::HwpActionHelper::FormObjCreatorListBox, py::arg("hwp"), "목록 상자 생성");
    actions.def("FormObjCreatorPushButton", &cpyhwpx::HwpActionHelper::FormObjCreatorPushButton, py::arg("hwp"), "푸시 버튼 생성");
    actions.def("FormObjCreatorRadioButton", &cpyhwpx::HwpActionHelper::FormObjCreatorRadioButton, py::arg("hwp"), "라디오 버튼 생성");

    // 자동
    actions.def("ASendBrowserText", &cpyhwpx::HwpActionHelper::ASendBrowserText, py::arg("hwp"), "브라우저 텍스트 전송");
    actions.def("AutoChangeHangul", &cpyhwpx::HwpActionHelper::AutoChangeHangul, py::arg("hwp"), "한글 자동 변환");
    actions.def("AutoChangeRun", &cpyhwpx::HwpActionHelper::AutoChangeRun, py::arg("hwp"), "자동 변환 실행");
    actions.def("AutoSpellRun", &cpyhwpx::HwpActionHelper::AutoSpellRun, py::arg("hwp"), "맞춤법 자동 실행");

    // 그리기 개체
    actions.def("DrawObjCancelOneStep", &cpyhwpx::HwpActionHelper::DrawObjCancelOneStep, py::arg("hwp"), "한 단계 취소");
    actions.def("DrawObjEditDetail", &cpyhwpx::HwpActionHelper::DrawObjEditDetail, py::arg("hwp"), "상세 편집");
    actions.def("DrawObjOpenClosePolygon", &cpyhwpx::HwpActionHelper::DrawObjOpenClosePolygon, py::arg("hwp"), "다각형 열기/닫기");
    actions.def("DrawObjTemplateSave", &cpyhwpx::HwpActionHelper::DrawObjTemplateSave, py::arg("hwp"), "템플릿 저장");

    // 빠른 명령
    actions.def("QuickCommandRun", &cpyhwpx::HwpActionHelper::QuickCommandRun, py::arg("hwp"), "빠른 명령 실행");
    actions.def("QuickCorrect", &cpyhwpx::HwpActionHelper::QuickCorrect, py::arg("hwp"), "빠른 교정");
    actions.def("QuickCorrectRun", &cpyhwpx::HwpActionHelper::QuickCorrectRun, py::arg("hwp"), "빠른 교정 실행");

    // 매크로
    actions.def("MacroPause", &cpyhwpx::HwpActionHelper::MacroPause, py::arg("hwp"), "매크로 일시 정지");
    actions.def("MacroRepeat", &cpyhwpx::HwpActionHelper::MacroRepeat, py::arg("hwp"), "매크로 반복");
    actions.def("MacroStop", &cpyhwpx::HwpActionHelper::MacroStop, py::arg("hwp"), "매크로 중지");

    // 기타
    actions.def("MailMergeField", &cpyhwpx::HwpActionHelper::MailMergeField, py::arg("hwp"), "메일 병합 필드");
    actions.def("MakeIndex", &cpyhwpx::HwpActionHelper::MakeIndex, py::arg("hwp"), "찾아보기 만들기");
    actions.def("LabelAdd", &cpyhwpx::HwpActionHelper::LabelAdd, py::arg("hwp"), "라벨 추가");
    actions.def("LabelTemplate", &cpyhwpx::HwpActionHelper::LabelTemplate, py::arg("hwp"), "라벨 템플릿");
    actions.def("HanThDIC", &cpyhwpx::HwpActionHelper::HanThDIC, py::arg("hwp"), "한글 유의어 사전");
    actions.def("HwpDic", &cpyhwpx::HwpActionHelper::HwpDic, py::arg("hwp"), "한컴 사전");
    actions.def("ReturnKeyInField", &cpyhwpx::HwpActionHelper::ReturnKeyInField, py::arg("hwp"), "필드에서 엔터");
}
