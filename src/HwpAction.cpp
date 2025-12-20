/**
 * @file HwpAction.cpp
 * @brief HAction COM 객체 래퍼 클래스 구현
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 */

#include "HwpAction.h"
#include "HwpWrapper.h"

namespace cpyhwpx {

//=============================================================================
// HwpAction 구현
//=============================================================================

HwpAction::HwpAction(IDispatch* action, HwpWrapper* hwp)
    : m_pAction(action)
    , m_pHwp(hwp)
{
    if (m_pAction) {
        m_pAction->AddRef();
    }
}

HwpAction::~HwpAction()
{
    if (m_pAction) {
        m_pAction->Release();
        m_pAction = nullptr;
    }
}

bool HwpAction::Run()
{
    VARIANT result = InvokeMethod(L"Run");
    if (result.vt == VT_BOOL) {
        return result.boolVal != VARIANT_FALSE;
    }
    return false;
}

bool HwpAction::Execute(IDispatch* pset)
{
    if (!pset) return false;

    VARIANT args[1];
    VariantInit(&args[0]);
    args[0].vt = VT_DISPATCH;
    args[0].pdispVal = pset;

    VARIANT result = InvokeMethod(L"Execute", { args[0] });
    if (result.vt == VT_BOOL) {
        return result.boolVal != VARIANT_FALSE;
    }
    return false;
}

IDispatch* HwpAction::GetDefault()
{
    VARIANT result = InvokeMethod(L"GetDefault");
    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    return nullptr;
}

bool HwpAction::PopupDialog(IDispatch* pset)
{
    if (!pset) return false;

    VARIANT args[1];
    VariantInit(&args[0]);
    args[0].vt = VT_DISPATCH;
    args[0].pdispVal = pset;

    VARIANT result = InvokeMethod(L"PopupDialog", { args[0] });
    if (result.vt == VT_BOOL) {
        return result.boolVal != VARIANT_FALSE;
    }
    return false;
}

std::wstring HwpAction::GetActionID() const
{
    if (!m_pAction) return L"";

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"ActionID");
    HRESULT hr = m_pAction->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return L"";

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                           &params, &result, NULL, NULL);

    if (SUCCEEDED(hr) && result.vt == VT_BSTR && result.bstrVal) {
        std::wstring id(result.bstrVal);
        SysFreeString(result.bstrVal);
        return id;
    }
    return L"";
}

VARIANT HwpAction::InvokeMethod(const std::wstring& name,
                                 const std::vector<VARIANT>& args)
{
    VARIANT result;
    VariantInit(&result);

    if (!m_pAction) return result;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = m_pAction->GetIDsOfNames(IID_NULL, &methodName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return result;

    std::vector<VARIANT> reversedArgs(args.rbegin(), args.rend());

    DISPPARAMS params = {
        reversedArgs.empty() ? NULL : reversedArgs.data(),
        NULL,
        static_cast<UINT>(reversedArgs.size()),
        0
    };

    m_pAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                      &params, &result, NULL, NULL);

    return result;
}

//=============================================================================
// HwpActionHelper 구현
//=============================================================================

bool HwpActionHelper::Run(HwpWrapper* hwp, const std::wstring& action_name)
{
    if (!hwp) return false;
    return hwp->RunAction(action_name);
}

bool HwpActionHelper::RunWithSet(HwpWrapper* hwp,
                                  const std::wstring& action_name,
                                  const std::wstring& set_name,
                                  std::function<void(IDispatch*)> setter)
{
    if (!hwp) return false;

    // 액션 생성
    IDispatch* pAction = hwp->CreateAction(action_name);
    if (!pAction) return false;

    // 파라미터셋 생성
    IDispatch* pSet = hwp->CreateSet(set_name);
    if (!pSet) {
        pAction->Release();
        return false;
    }

    // GetDefault 호출
    HwpAction action(pAction, hwp);
    IDispatch* pDefault = action.GetDefault();

    // 커스텀 setter 호출
    if (setter && pSet) {
        setter(pSet);
    }

    // Execute 실행
    bool success = action.Execute(pSet);

    pSet->Release();
    pAction->Release();

    return success;
}

// 문단
bool HwpActionHelper::BreakPara(HwpWrapper* hwp) { return Run(hwp, L"BreakPara"); }
bool HwpActionHelper::BreakPage(HwpWrapper* hwp) { return Run(hwp, L"BreakPage"); }
bool HwpActionHelper::BreakSection(HwpWrapper* hwp) { return Run(hwp, L"BreakSection"); }
bool HwpActionHelper::BreakColumn(HwpWrapper* hwp) { return Run(hwp, L"BreakColumn"); }

// 선택
bool HwpActionHelper::SelectAll(HwpWrapper* hwp) { return Run(hwp, L"SelectAll"); }
bool HwpActionHelper::Cancel(HwpWrapper* hwp) { return Run(hwp, L"Cancel"); }

// 커서 이동
bool HwpActionHelper::MoveLeft(HwpWrapper* hwp) { return Run(hwp, L"MoveLeft"); }
bool HwpActionHelper::MoveRight(HwpWrapper* hwp) { return Run(hwp, L"MoveRight"); }
bool HwpActionHelper::MoveUp(HwpWrapper* hwp) { return Run(hwp, L"MoveUp"); }
bool HwpActionHelper::MoveDown(HwpWrapper* hwp) { return Run(hwp, L"MoveDown"); }
bool HwpActionHelper::MoveLineBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveLineBegin"); }
bool HwpActionHelper::MoveLineEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveLineEnd"); }
bool HwpActionHelper::MoveDocBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveDocBegin"); }
bool HwpActionHelper::MoveDocEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveDocEnd"); }
bool HwpActionHelper::MoveParaBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveParaBegin"); }
bool HwpActionHelper::MoveParaEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveParaEnd"); }
bool HwpActionHelper::MoveWordBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveWordBegin"); }
bool HwpActionHelper::MoveWordEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveWordEnd"); }
bool HwpActionHelper::MovePageUp(HwpWrapper* hwp) { return Run(hwp, L"MovePageUp"); }
bool HwpActionHelper::MovePageDown(HwpWrapper* hwp) { return Run(hwp, L"MovePageDown"); }

// 편집
bool HwpActionHelper::Delete(HwpWrapper* hwp) { return Run(hwp, L"Delete"); }
bool HwpActionHelper::DeleteBack(HwpWrapper* hwp) { return Run(hwp, L"DeleteBack"); }
bool HwpActionHelper::Cut(HwpWrapper* hwp) { return Run(hwp, L"Cut"); }
bool HwpActionHelper::Copy(HwpWrapper* hwp) { return Run(hwp, L"Copy"); }
bool HwpActionHelper::Paste(HwpWrapper* hwp) { return Run(hwp, L"Paste"); }
bool HwpActionHelper::Undo(HwpWrapper* hwp) { return Run(hwp, L"Undo"); }
bool HwpActionHelper::Redo(HwpWrapper* hwp) { return Run(hwp, L"Redo"); }

// 표 관련
bool HwpActionHelper::TableCellBlock(HwpWrapper* hwp) { return Run(hwp, L"TableCellBlock"); }
bool HwpActionHelper::TableColBegin(HwpWrapper* hwp) { return Run(hwp, L"TableColBegin"); }
bool HwpActionHelper::TableColEnd(HwpWrapper* hwp) { return Run(hwp, L"TableColEnd"); }
bool HwpActionHelper::TableRowBegin(HwpWrapper* hwp) { return Run(hwp, L"TableRowBegin"); }
bool HwpActionHelper::TableRowEnd(HwpWrapper* hwp) { return Run(hwp, L"TableRowEnd"); }
bool HwpActionHelper::TableCellAddr(HwpWrapper* hwp) { return Run(hwp, L"TableCellAddr"); }
bool HwpActionHelper::TableAppendRow(HwpWrapper* hwp) { return Run(hwp, L"TableAppendRow"); }
bool HwpActionHelper::TableAppendColumn(HwpWrapper* hwp) { return Run(hwp, L"TableAppendColumn"); }
bool HwpActionHelper::TableDeleteRow(HwpWrapper* hwp) { return Run(hwp, L"TableDeleteRow"); }
bool HwpActionHelper::TableDeleteColumn(HwpWrapper* hwp) { return Run(hwp, L"TableDeleteColumn"); }
bool HwpActionHelper::TableMergeCell(HwpWrapper* hwp) { return Run(hwp, L"TableMergeCell"); }
bool HwpActionHelper::TableSplitCell(HwpWrapper* hwp) { return Run(hwp, L"TableSplitCell"); }

// 스타일
bool HwpActionHelper::ParagraphShapeAlignLeft(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeAlignLeft"); }
bool HwpActionHelper::ParagraphShapeAlignCenter(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeAlignCenter"); }
bool HwpActionHelper::ParagraphShapeAlignRight(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeAlignRight"); }
bool HwpActionHelper::ParagraphShapeAlignJustify(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeAlignJustify"); }

// 글자 모양
bool HwpActionHelper::CharShapeBold(HwpWrapper* hwp) { return Run(hwp, L"CharShapeBold"); }
bool HwpActionHelper::CharShapeItalic(HwpWrapper* hwp) { return Run(hwp, L"CharShapeItalic"); }
bool HwpActionHelper::CharShapeUnderline(HwpWrapper* hwp) { return Run(hwp, L"CharShapeUnderline"); }
bool HwpActionHelper::CharShapeStrikeout(HwpWrapper* hwp) { return Run(hwp, L"CharShapeStrikeout"); }
bool HwpActionHelper::CharShapeSuperscript(HwpWrapper* hwp) { return Run(hwp, L"CharShapeSuperscript"); }
bool HwpActionHelper::CharShapeSubscript(HwpWrapper* hwp) { return Run(hwp, L"CharShapeSubscript"); }

// 창/보기
bool HwpActionHelper::FileClose(HwpWrapper* hwp) { return Run(hwp, L"FileClose"); }
bool HwpActionHelper::FileQuit(HwpWrapper* hwp) { return Run(hwp, L"FileQuit"); }
bool HwpActionHelper::WindowMaximize(HwpWrapper* hwp) { return Run(hwp, L"WindowMaximize"); }
bool HwpActionHelper::WindowMinimize(HwpWrapper* hwp) { return Run(hwp, L"WindowMinimize"); }
bool HwpActionHelper::ViewZoomIn(HwpWrapper* hwp) { return Run(hwp, L"ViewZoomIn"); }
bool HwpActionHelper::ViewZoomOut(HwpWrapper* hwp) { return Run(hwp, L"ViewZoomOut"); }

// 인쇄
bool HwpActionHelper::FilePrint(HwpWrapper* hwp) { return Run(hwp, L"FilePrint"); }
bool HwpActionHelper::FilePrintPreview(HwpWrapper* hwp) { return Run(hwp, L"FilePrintPreview"); }

// 개체
bool HwpActionHelper::ShapeObjSelect(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjSelect"); }
bool HwpActionHelper::ShapeObjDelete(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjDelete"); }
bool HwpActionHelper::ShapeObjCopy(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjCopy"); }
bool HwpActionHelper::ShapeObjCut(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjCut"); }
bool HwpActionHelper::ShapeObjBringToFront(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjBringToFront"); }
bool HwpActionHelper::ShapeObjSendToBack(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjSendToBack"); }

} // namespace cpyhwpx
