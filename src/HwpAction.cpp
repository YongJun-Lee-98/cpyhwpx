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
bool HwpActionHelper::TableColPageUp(HwpWrapper* hwp) { return Run(hwp, L"TableColPageUp"); }
bool HwpActionHelper::TableCellBlockExtendAbs(HwpWrapper* hwp) { return Run(hwp, L"TableCellBlockExtendAbs"); }

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

//=============================================================================
// 추가 Run 액션들
//=============================================================================

// Break (추가)
bool HwpActionHelper::BreakLine(HwpWrapper* hwp) { return Run(hwp, L"BreakLine"); }
bool HwpActionHelper::BreakColDef(HwpWrapper* hwp) { return Run(hwp, L"BreakColDef"); }

// CharShape (추가)
bool HwpActionHelper::CharShapeCenterline(HwpWrapper* hwp) { return Run(hwp, L"CharShapeCenterline"); }
bool HwpActionHelper::CharShapeEmboss(HwpWrapper* hwp) { return Run(hwp, L"CharShapeEmboss"); }
bool HwpActionHelper::CharShapeEngrave(HwpWrapper* hwp) { return Run(hwp, L"CharShapeEngrave"); }
bool HwpActionHelper::CharShapeHeight(HwpWrapper* hwp) { return Run(hwp, L"CharShapeHeight"); }
bool HwpActionHelper::CharShapeHeightDecrease(HwpWrapper* hwp) { return Run(hwp, L"CharShapeHeightDecrease"); }
bool HwpActionHelper::CharShapeHeightIncrease(HwpWrapper* hwp) { return Run(hwp, L"CharShapeHeightIncrease"); }
bool HwpActionHelper::CharShapeLang(HwpWrapper* hwp) { return Run(hwp, L"CharShapeLang"); }
bool HwpActionHelper::CharShapeNextFaceName(HwpWrapper* hwp) { return Run(hwp, L"CharShapeNextFaceName"); }
bool HwpActionHelper::CharShapeNormal(HwpWrapper* hwp) { return Run(hwp, L"CharShapeNormal"); }
bool HwpActionHelper::CharShapeOutline(HwpWrapper* hwp) { return Run(hwp, L"CharShapeOutline"); }
bool HwpActionHelper::CharShapePrevFaceName(HwpWrapper* hwp) { return Run(hwp, L"CharShapePrevFaceName"); }
bool HwpActionHelper::CharShapeShadow(HwpWrapper* hwp) { return Run(hwp, L"CharShapeShadow"); }
bool HwpActionHelper::CharShapeSpacing(HwpWrapper* hwp) { return Run(hwp, L"CharShapeSpacing"); }
bool HwpActionHelper::CharShapeSpacingDecrease(HwpWrapper* hwp) { return Run(hwp, L"CharShapeSpacingDecrease"); }
bool HwpActionHelper::CharShapeSpacingIncrease(HwpWrapper* hwp) { return Run(hwp, L"CharShapeSpacingIncrease"); }
bool HwpActionHelper::CharShapeSuperSubscript(HwpWrapper* hwp) { return Run(hwp, L"CharShapeSuperSubscript"); }
bool HwpActionHelper::CharShapeTextColorBlack(HwpWrapper* hwp) { return Run(hwp, L"CharShapeTextColorBlack"); }
bool HwpActionHelper::CharShapeTextColorBlue(HwpWrapper* hwp) { return Run(hwp, L"CharShapeTextColorBlue"); }
bool HwpActionHelper::CharShapeTextColorBluish(HwpWrapper* hwp) { return Run(hwp, L"CharShapeTextColorBluish"); }
bool HwpActionHelper::CharShapeTextColorGreen(HwpWrapper* hwp) { return Run(hwp, L"CharShapeTextColorGreen"); }
bool HwpActionHelper::CharShapeTextColorRed(HwpWrapper* hwp) { return Run(hwp, L"CharShapeTextColorRed"); }
bool HwpActionHelper::CharShapeTextColorViolet(HwpWrapper* hwp) { return Run(hwp, L"CharShapeTextColorViolet"); }
bool HwpActionHelper::CharShapeTextColorWhite(HwpWrapper* hwp) { return Run(hwp, L"CharShapeTextColorWhite"); }
bool HwpActionHelper::CharShapeTextColorYellow(HwpWrapper* hwp) { return Run(hwp, L"CharShapeTextColorYellow"); }
bool HwpActionHelper::CharShapeTypeface(HwpWrapper* hwp) { return Run(hwp, L"CharShapeTypeface"); }
bool HwpActionHelper::CharShapeWidth(HwpWrapper* hwp) { return Run(hwp, L"CharShapeWidth"); }
bool HwpActionHelper::CharShapeWidthDecrease(HwpWrapper* hwp) { return Run(hwp, L"CharShapeWidthDecrease"); }
bool HwpActionHelper::CharShapeWidthIncrease(HwpWrapper* hwp) { return Run(hwp, L"CharShapeWidthIncrease"); }

// ParagraphShape (추가)
bool HwpActionHelper::ParagraphShapeAlignDistribute(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeAlignDistribute"); }
bool HwpActionHelper::ParagraphShapeAlignDivision(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeAlignDivision"); }
bool HwpActionHelper::ParagraphShapeDecreaseLeftMargin(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeDecreaseLeftMargin"); }
bool HwpActionHelper::ParagraphShapeDecreaseLineSpacing(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeDecreaseLineSpacing"); }
bool HwpActionHelper::ParagraphShapeDecreaseMargin(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeDecreaseMargin"); }
bool HwpActionHelper::ParagraphShapeDecreaseRightMargin(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeDecreaseRightMargin"); }
bool HwpActionHelper::ParagraphShapeIncreaseLeftMargin(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeIncreaseLeftMargin"); }
bool HwpActionHelper::ParagraphShapeIncreaseLineSpacing(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeIncreaseLineSpacing"); }
bool HwpActionHelper::ParagraphShapeIncreaseMargin(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeIncreaseMargin"); }
bool HwpActionHelper::ParagraphShapeIncreaseRightMargin(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeIncreaseRightMargin"); }
bool HwpActionHelper::ParagraphShapeIndentAtCaret(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeIndentAtCaret"); }
bool HwpActionHelper::ParagraphShapeIndentNegative(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeIndentNegative"); }
bool HwpActionHelper::ParagraphShapeIndentPositive(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeIndentPositive"); }
bool HwpActionHelper::ParagraphShapeProtect(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeProtect"); }
bool HwpActionHelper::ParagraphShapeSingleRow(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeSingleRow"); }
bool HwpActionHelper::ParagraphShapeWithNext(HwpWrapper* hwp) { return Run(hwp, L"ParagraphShapeWithNext"); }

// Move (추가)
bool HwpActionHelper::MoveLineDown(HwpWrapper* hwp) { return Run(hwp, L"MoveLineDown"); }
bool HwpActionHelper::MoveLineUp(HwpWrapper* hwp) { return Run(hwp, L"MoveLineUp"); }
bool HwpActionHelper::MoveNextWord(HwpWrapper* hwp) { return Run(hwp, L"MoveNextWord"); }
bool HwpActionHelper::MovePrevWord(HwpWrapper* hwp) { return Run(hwp, L"MovePrevWord"); }
bool HwpActionHelper::MoveNextChar(HwpWrapper* hwp) { return Run(hwp, L"MoveNextChar"); }
bool HwpActionHelper::MovePrevChar(HwpWrapper* hwp) { return Run(hwp, L"MovePrevChar"); }
bool HwpActionHelper::MovePageBegin(HwpWrapper* hwp) { return Run(hwp, L"MovePageBegin"); }
bool HwpActionHelper::MovePageEnd(HwpWrapper* hwp) { return Run(hwp, L"MovePageEnd"); }
bool HwpActionHelper::MoveSelDown(HwpWrapper* hwp) { return Run(hwp, L"MoveSelDown"); }
bool HwpActionHelper::MoveSelLeft(HwpWrapper* hwp) { return Run(hwp, L"MoveSelLeft"); }
bool HwpActionHelper::MoveSelRight(HwpWrapper* hwp) { return Run(hwp, L"MoveSelRight"); }
bool HwpActionHelper::MoveSelUp(HwpWrapper* hwp) { return Run(hwp, L"MoveSelUp"); }
bool HwpActionHelper::MoveSelDocBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveSelDocBegin"); }
bool HwpActionHelper::MoveSelDocEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveSelDocEnd"); }
bool HwpActionHelper::MoveSelLineBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveSelLineBegin"); }
bool HwpActionHelper::MoveSelLineEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveSelLineEnd"); }
bool HwpActionHelper::MoveSelLineDown(HwpWrapper* hwp) { return Run(hwp, L"MoveSelLineDown"); }
bool HwpActionHelper::MoveSelLineUp(HwpWrapper* hwp) { return Run(hwp, L"MoveSelLineUp"); }
bool HwpActionHelper::MoveSelNextChar(HwpWrapper* hwp) { return Run(hwp, L"MoveSelNextChar"); }
bool HwpActionHelper::MoveSelPrevChar(HwpWrapper* hwp) { return Run(hwp, L"MoveSelPrevChar"); }
bool HwpActionHelper::MoveSelNextWord(HwpWrapper* hwp) { return Run(hwp, L"MoveSelNextWord"); }
bool HwpActionHelper::MoveSelPrevWord(HwpWrapper* hwp) { return Run(hwp, L"MoveSelPrevWord"); }
bool HwpActionHelper::MoveSelPageDown(HwpWrapper* hwp) { return Run(hwp, L"MoveSelPageDown"); }
bool HwpActionHelper::MoveSelPageUp(HwpWrapper* hwp) { return Run(hwp, L"MoveSelPageUp"); }
bool HwpActionHelper::MoveSelParaBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveSelParaBegin"); }
bool HwpActionHelper::MoveSelParaEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveSelParaEnd"); }
bool HwpActionHelper::MoveSelWordBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveSelWordBegin"); }
bool HwpActionHelper::MoveSelWordEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveSelWordEnd"); }
bool HwpActionHelper::MoveColumnBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveColumnBegin"); }
bool HwpActionHelper::MoveColumnEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveColumnEnd"); }
bool HwpActionHelper::MoveListBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveListBegin"); }
bool HwpActionHelper::MoveListEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveListEnd"); }
bool HwpActionHelper::MoveParentList(HwpWrapper* hwp) { return Run(hwp, L"MoveParentList"); }
bool HwpActionHelper::MoveRootList(HwpWrapper* hwp) { return Run(hwp, L"MoveRootList"); }
bool HwpActionHelper::MoveTopLevelBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveTopLevelBegin"); }
bool HwpActionHelper::MoveTopLevelEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveTopLevelEnd"); }
bool HwpActionHelper::MoveTopLevelList(HwpWrapper* hwp) { return Run(hwp, L"MoveTopLevelList"); }
bool HwpActionHelper::MoveNextColumn(HwpWrapper* hwp) { return Run(hwp, L"MoveNextColumn"); }
bool HwpActionHelper::MovePrevColumn(HwpWrapper* hwp) { return Run(hwp, L"MovePrevColumn"); }
bool HwpActionHelper::MoveNextPos(HwpWrapper* hwp) { return Run(hwp, L"MoveNextPos"); }
bool HwpActionHelper::MovePrevPos(HwpWrapper* hwp) { return Run(hwp, L"MovePrevPos"); }
bool HwpActionHelper::MoveNextPosEx(HwpWrapper* hwp) { return Run(hwp, L"MoveNextPosEx"); }
bool HwpActionHelper::MovePrevPosEx(HwpWrapper* hwp) { return Run(hwp, L"MovePrevPosEx"); }
bool HwpActionHelper::MoveNextParaBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveNextParaBegin"); }
bool HwpActionHelper::MovePrevParaBegin(HwpWrapper* hwp) { return Run(hwp, L"MovePrevParaBegin"); }
bool HwpActionHelper::MovePrevParaEnd(HwpWrapper* hwp) { return Run(hwp, L"MovePrevParaEnd"); }
bool HwpActionHelper::MoveSectionDown(HwpWrapper* hwp) { return Run(hwp, L"MoveSectionDown"); }
bool HwpActionHelper::MoveSectionUp(HwpWrapper* hwp) { return Run(hwp, L"MoveSectionUp"); }
bool HwpActionHelper::MoveScrollDown(HwpWrapper* hwp) { return Run(hwp, L"MoveScrollDown"); }
bool HwpActionHelper::MoveScrollUp(HwpWrapper* hwp) { return Run(hwp, L"MoveScrollUp"); }
bool HwpActionHelper::MoveScrollNext(HwpWrapper* hwp) { return Run(hwp, L"MoveScrollNext"); }
bool HwpActionHelper::MoveScrollPrev(HwpWrapper* hwp) { return Run(hwp, L"MoveScrollPrev"); }
bool HwpActionHelper::MoveViewBegin(HwpWrapper* hwp) { return Run(hwp, L"MoveViewBegin"); }
bool HwpActionHelper::MoveViewEnd(HwpWrapper* hwp) { return Run(hwp, L"MoveViewEnd"); }
bool HwpActionHelper::MoveViewDown(HwpWrapper* hwp) { return Run(hwp, L"MoveViewDown"); }
bool HwpActionHelper::MoveViewUp(HwpWrapper* hwp) { return Run(hwp, L"MoveViewUp"); }

// Table (추가)
bool HwpActionHelper::TableLeftCell(HwpWrapper* hwp) { return Run(hwp, L"TableLeftCell"); }
bool HwpActionHelper::TableRightCell(HwpWrapper* hwp) { return Run(hwp, L"TableRightCell"); }
bool HwpActionHelper::TableUpperCell(HwpWrapper* hwp) { return Run(hwp, L"TableUpperCell"); }
bool HwpActionHelper::TableLowerCell(HwpWrapper* hwp) { return Run(hwp, L"TableLowerCell"); }
bool HwpActionHelper::TableColPageDown(HwpWrapper* hwp) { return Run(hwp, L"TableColPageDown"); }
bool HwpActionHelper::TableRightCellAppend(HwpWrapper* hwp) { return Run(hwp, L"TableRightCellAppend"); }
bool HwpActionHelper::TableCellBlockCol(HwpWrapper* hwp) { return Run(hwp, L"TableCellBlockCol"); }
bool HwpActionHelper::TableCellBlockRow(HwpWrapper* hwp) { return Run(hwp, L"TableCellBlockRow"); }
bool HwpActionHelper::TableCellBlockExtend(HwpWrapper* hwp) { return Run(hwp, L"TableCellBlockExtend"); }
bool HwpActionHelper::TableCellAlignLeftTop(HwpWrapper* hwp) { return Run(hwp, L"TableCellAlignLeftTop"); }
bool HwpActionHelper::TableCellAlignCenterTop(HwpWrapper* hwp) { return Run(hwp, L"TableCellAlignCenterTop"); }
bool HwpActionHelper::TableCellAlignRightTop(HwpWrapper* hwp) { return Run(hwp, L"TableCellAlignRightTop"); }
bool HwpActionHelper::TableCellAlignLeftCenter(HwpWrapper* hwp) { return Run(hwp, L"TableCellAlignLeftCenter"); }
bool HwpActionHelper::TableCellAlignCenterCenter(HwpWrapper* hwp) { return Run(hwp, L"TableCellAlignCenterCenter"); }
bool HwpActionHelper::TableCellAlignRightCenter(HwpWrapper* hwp) { return Run(hwp, L"TableCellAlignRightCenter"); }
bool HwpActionHelper::TableCellAlignLeftBottom(HwpWrapper* hwp) { return Run(hwp, L"TableCellAlignLeftBottom"); }
bool HwpActionHelper::TableCellAlignCenterBottom(HwpWrapper* hwp) { return Run(hwp, L"TableCellAlignCenterBottom"); }
bool HwpActionHelper::TableCellAlignRightBottom(HwpWrapper* hwp) { return Run(hwp, L"TableCellAlignRightBottom"); }
bool HwpActionHelper::TableVAlignTop(HwpWrapper* hwp) { return Run(hwp, L"TableVAlignTop"); }
bool HwpActionHelper::TableVAlignCenter(HwpWrapper* hwp) { return Run(hwp, L"TableVAlignCenter"); }
bool HwpActionHelper::TableVAlignBottom(HwpWrapper* hwp) { return Run(hwp, L"TableVAlignBottom"); }
bool HwpActionHelper::TableCellBorderAll(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderAll"); }
bool HwpActionHelper::TableCellBorderNo(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderNo"); }
bool HwpActionHelper::TableCellBorderOutside(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderOutside"); }
bool HwpActionHelper::TableCellBorderInside(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderInside"); }
bool HwpActionHelper::TableCellBorderInsideHorz(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderInsideHorz"); }
bool HwpActionHelper::TableCellBorderInsideVert(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderInsideVert"); }
bool HwpActionHelper::TableCellBorderTop(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderTop"); }
bool HwpActionHelper::TableCellBorderBottom(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderBottom"); }
bool HwpActionHelper::TableCellBorderLeft(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderLeft"); }
bool HwpActionHelper::TableCellBorderRight(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderRight"); }
bool HwpActionHelper::TableCellBorderDiagonalDown(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderDiagonalDown"); }
bool HwpActionHelper::TableCellBorderDiagonalUp(HwpWrapper* hwp) { return Run(hwp, L"TableCellBorderDiagonalUp"); }
bool HwpActionHelper::TableSubtractRow(HwpWrapper* hwp) { return Run(hwp, L"TableSubtractRow"); }
bool HwpActionHelper::TableDeleteCell(HwpWrapper* hwp) { return Run(hwp, L"TableDeleteCell"); }
bool HwpActionHelper::TableMergeTable(HwpWrapper* hwp) { return Run(hwp, L"TableMergeTable"); }
bool HwpActionHelper::TableSplitTable(HwpWrapper* hwp) { return Run(hwp, L"TableSplitTable"); }
bool HwpActionHelper::TableDistributeCellHeight(HwpWrapper* hwp) { return Run(hwp, L"TableDistributeCellHeight"); }
bool HwpActionHelper::TableDistributeCellWidth(HwpWrapper* hwp) { return Run(hwp, L"TableDistributeCellWidth"); }
bool HwpActionHelper::TableResizeDown(HwpWrapper* hwp) { return Run(hwp, L"TableResizeDown"); }
bool HwpActionHelper::TableResizeUp(HwpWrapper* hwp) { return Run(hwp, L"TableResizeUp"); }
bool HwpActionHelper::TableResizeLeft(HwpWrapper* hwp) { return Run(hwp, L"TableResizeLeft"); }
bool HwpActionHelper::TableResizeRight(HwpWrapper* hwp) { return Run(hwp, L"TableResizeRight"); }
bool HwpActionHelper::TableResizeCellDown(HwpWrapper* hwp) { return Run(hwp, L"TableResizeCellDown"); }
bool HwpActionHelper::TableResizeCellUp(HwpWrapper* hwp) { return Run(hwp, L"TableResizeCellUp"); }
bool HwpActionHelper::TableResizeCellLeft(HwpWrapper* hwp) { return Run(hwp, L"TableResizeCellLeft"); }
bool HwpActionHelper::TableResizeCellRight(HwpWrapper* hwp) { return Run(hwp, L"TableResizeCellRight"); }
bool HwpActionHelper::TableResizeExDown(HwpWrapper* hwp) { return Run(hwp, L"TableResizeExDown"); }
bool HwpActionHelper::TableResizeExUp(HwpWrapper* hwp) { return Run(hwp, L"TableResizeExUp"); }
bool HwpActionHelper::TableResizeExLeft(HwpWrapper* hwp) { return Run(hwp, L"TableResizeExLeft"); }
bool HwpActionHelper::TableResizeExRight(HwpWrapper* hwp) { return Run(hwp, L"TableResizeExRight"); }
bool HwpActionHelper::TableResizeLineDown(HwpWrapper* hwp) { return Run(hwp, L"TableResizeLineDown"); }
bool HwpActionHelper::TableResizeLineUp(HwpWrapper* hwp) { return Run(hwp, L"TableResizeLineUp"); }
bool HwpActionHelper::TableResizeLineLeft(HwpWrapper* hwp) { return Run(hwp, L"TableResizeLineLeft"); }
bool HwpActionHelper::TableResizeLineRight(HwpWrapper* hwp) { return Run(hwp, L"TableResizeLineRight"); }
bool HwpActionHelper::TableFormulaSumAuto(HwpWrapper* hwp) { return Run(hwp, L"TableFormulaSumAuto"); }
bool HwpActionHelper::TableFormulaSumHor(HwpWrapper* hwp) { return Run(hwp, L"TableFormulaSumHor"); }
bool HwpActionHelper::TableFormulaSumVer(HwpWrapper* hwp) { return Run(hwp, L"TableFormulaSumVer"); }
bool HwpActionHelper::TableFormulaAvgAuto(HwpWrapper* hwp) { return Run(hwp, L"TableFormulaAvgAuto"); }
bool HwpActionHelper::TableFormulaAvgHor(HwpWrapper* hwp) { return Run(hwp, L"TableFormulaAvgHor"); }
bool HwpActionHelper::TableFormulaAvgVer(HwpWrapper* hwp) { return Run(hwp, L"TableFormulaAvgVer"); }
bool HwpActionHelper::TableFormulaProAuto(HwpWrapper* hwp) { return Run(hwp, L"TableFormulaProAuto"); }
bool HwpActionHelper::TableFormulaProHor(HwpWrapper* hwp) { return Run(hwp, L"TableFormulaProHor"); }
bool HwpActionHelper::TableFormulaProVer(HwpWrapper* hwp) { return Run(hwp, L"TableFormulaProVer"); }
bool HwpActionHelper::TableAutoFill(HwpWrapper* hwp) { return Run(hwp, L"TableAutoFill"); }
bool HwpActionHelper::TableAutoFillDlg(HwpWrapper* hwp) { return Run(hwp, L"TableAutoFillDlg"); }
bool HwpActionHelper::TableDrawPen(HwpWrapper* hwp) { return Run(hwp, L"TableDrawPen"); }
bool HwpActionHelper::TableDrawPenStyle(HwpWrapper* hwp) { return Run(hwp, L"TableDrawPenStyle"); }
bool HwpActionHelper::TableDrawPenWidth(HwpWrapper* hwp) { return Run(hwp, L"TableDrawPenWidth"); }
bool HwpActionHelper::TableEraser(HwpWrapper* hwp) { return Run(hwp, L"TableEraser"); }
bool HwpActionHelper::TableDeleteComma(HwpWrapper* hwp) { return Run(hwp, L"TableDeleteComma"); }
bool HwpActionHelper::TableInsertComma(HwpWrapper* hwp) { return Run(hwp, L"TableInsertComma"); }

// ShapeObj (추가)
bool HwpActionHelper::ShapeObjAlignLeft(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignLeft"); }
bool HwpActionHelper::ShapeObjAlignCenter(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignCenter"); }
bool HwpActionHelper::ShapeObjAlignRight(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignRight"); }
bool HwpActionHelper::ShapeObjAlignTop(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignTop"); }
bool HwpActionHelper::ShapeObjAlignMiddle(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignMiddle"); }
bool HwpActionHelper::ShapeObjAlignBottom(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignBottom"); }
bool HwpActionHelper::ShapeObjAlignWidth(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignWidth"); }
bool HwpActionHelper::ShapeObjAlignHeight(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignHeight"); }
bool HwpActionHelper::ShapeObjAlignSize(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignSize"); }
bool HwpActionHelper::ShapeObjAlignHorzSpacing(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignHorzSpacing"); }
bool HwpActionHelper::ShapeObjAlignVertSpacing(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAlignVertSpacing"); }
bool HwpActionHelper::ShapeObjBringForward(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjBringForward"); }
bool HwpActionHelper::ShapeObjSendBack(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjSendBack"); }
bool HwpActionHelper::ShapeObjBringInFrontOfText(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjBringInFrontOfText"); }
bool HwpActionHelper::ShapeObjCtrlSendBehindText(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjCtrlSendBehindText"); }
bool HwpActionHelper::ShapeObjHorzFlip(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjHorzFlip"); }
bool HwpActionHelper::ShapeObjVertFlip(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjVertFlip"); }
bool HwpActionHelper::ShapeObjHorzFlipOrgState(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjHorzFlipOrgState"); }
bool HwpActionHelper::ShapeObjVertFlipOrgState(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjVertFlipOrgState"); }
bool HwpActionHelper::ShapeObjRotater(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjRotater"); }
bool HwpActionHelper::ShapeObjRightAngleRotater(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjRightAngleRotater"); }
bool HwpActionHelper::ShapeObjRightAngleRotaterAnticlockwise(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjRightAngleRotaterAnticlockwise"); }
bool HwpActionHelper::ShapeObjMoveUp(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjMoveUp"); }
bool HwpActionHelper::ShapeObjMoveDown(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjMoveDown"); }
bool HwpActionHelper::ShapeObjMoveLeft(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjMoveLeft"); }
bool HwpActionHelper::ShapeObjMoveRight(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjMoveRight"); }
bool HwpActionHelper::ShapeObjResizeUp(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjResizeUp"); }
bool HwpActionHelper::ShapeObjResizeDown(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjResizeDown"); }
bool HwpActionHelper::ShapeObjResizeLeft(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjResizeLeft"); }
bool HwpActionHelper::ShapeObjResizeRight(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjResizeRight"); }
bool HwpActionHelper::ShapeObjGroup(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjGroup"); }
bool HwpActionHelper::ShapeObjUngroup(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjUngroup"); }
bool HwpActionHelper::ShapeObjNextObject(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjNextObject"); }
bool HwpActionHelper::ShapeObjPrevObject(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjPrevObject"); }
bool HwpActionHelper::ShapeObjLock(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjLock"); }
bool HwpActionHelper::ShapeObjUnlockAll(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjUnlockAll"); }
bool HwpActionHelper::ShapeObjAttachCaption(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAttachCaption"); }
bool HwpActionHelper::ShapeObjDetachCaption(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjDetachCaption"); }
bool HwpActionHelper::ShapeObjInsertCaptionNum(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjInsertCaptionNum"); }
bool HwpActionHelper::ShapeObjAttachTextBox(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjAttachTextBox"); }
bool HwpActionHelper::ShapeObjDetachTextBox(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjDetachTextBox"); }
bool HwpActionHelper::ShapeObjToggleTextBox(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjToggleTextBox"); }
bool HwpActionHelper::ShapeObjTextBoxEdit(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjTextBoxEdit"); }
bool HwpActionHelper::ShapeObjTableSelCell(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjTableSelCell"); }
bool HwpActionHelper::ShapeObjFillProperty(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjFillProperty"); }
bool HwpActionHelper::ShapeObjLineProperty(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjLineProperty"); }
bool HwpActionHelper::ShapeObjLineStyleOther(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjLineStyleOhter"); }
bool HwpActionHelper::ShapeObjLineWidthOther(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjLineWidthOhter"); }
bool HwpActionHelper::ShapeObjNorm(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjNorm"); }
bool HwpActionHelper::ShapeObjGuideLine(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjGuideLine"); }
bool HwpActionHelper::ShapeObjShowGuideLine(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjShowGuideLine"); }
bool HwpActionHelper::ShapeObjShowGuideLineBase(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjShowGuideLineBase"); }
bool HwpActionHelper::ShapeObjWrapSquare(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjWrapSquare"); }
bool HwpActionHelper::ShapeObjWrapTopAndBottom(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjWrapTopAndBottom"); }
bool HwpActionHelper::ShapeObjSaveAsPicture(HwpWrapper* hwp) { return Run(hwp, L"ShapeObjSaveAsPicture"); }

// File (추가)
bool HwpActionHelper::FileNew(HwpWrapper* hwp) { return Run(hwp, L"FileNew"); }
bool HwpActionHelper::FileNewTab(HwpWrapper* hwp) { return Run(hwp, L"FileNewTab"); }
bool HwpActionHelper::FileOpen(HwpWrapper* hwp) { return Run(hwp, L"FileOpen"); }
bool HwpActionHelper::FileOpenMRU(HwpWrapper* hwp) { return Run(hwp, L"FileOpenMRU"); }
bool HwpActionHelper::FileSave(HwpWrapper* hwp) { return Run(hwp, L"FileSave"); }
bool HwpActionHelper::FileSaveAs(HwpWrapper* hwp) { return Run(hwp, L"FileSaveAs"); }
bool HwpActionHelper::FileSaveAsDRM(HwpWrapper* hwp) { return Run(hwp, L"FileSaveAsDRM"); }
bool HwpActionHelper::FileSaveOptionDlg(HwpWrapper* hwp) { return Run(hwp, L"FileSaveOptionDlg"); }
bool HwpActionHelper::FileFind(HwpWrapper* hwp) { return Run(hwp, L"FileFind"); }
bool HwpActionHelper::FilePreview(HwpWrapper* hwp) { return Run(hwp, L"FilePreview"); }
bool HwpActionHelper::FileNextVersionDiff(HwpWrapper* hwp) { return Run(hwp, L"FileNextVersionDiff"); }
bool HwpActionHelper::FilePrevVersionDiff(HwpWrapper* hwp) { return Run(hwp, L"FilePrevVersionDiff"); }
bool HwpActionHelper::FileVersionDiffChangeAlign(HwpWrapper* hwp) { return Run(hwp, L"FileVersionDiffChangeAlign"); }
bool HwpActionHelper::FileVersionDiffSameAlign(HwpWrapper* hwp) { return Run(hwp, L"FileVersionDiffSameAlign"); }
bool HwpActionHelper::FileVersionDiffSyncScroll(HwpWrapper* hwp) { return Run(hwp, L"FileVersionDiffSyncScroll"); }

// Insert
bool HwpActionHelper::InsertAutoNum(HwpWrapper* hwp) { return Run(hwp, L"InsertAutoNum"); }
bool HwpActionHelper::InsertCpNo(HwpWrapper* hwp) { return Run(hwp, L"InsertCpNo"); }
bool HwpActionHelper::InsertCpTpNo(HwpWrapper* hwp) { return Run(hwp, L"InsertCpTpNo"); }
bool HwpActionHelper::InsertTpNo(HwpWrapper* hwp) { return Run(hwp, L"InsertTpNo"); }
bool HwpActionHelper::InsertPageNum(HwpWrapper* hwp) { return Run(hwp, L"InsertPageNum"); }
bool HwpActionHelper::InsertDateCode(HwpWrapper* hwp) { return Run(hwp, L"InsertDateCode"); }
bool HwpActionHelper::InsertDocInfo(HwpWrapper* hwp) { return Run(hwp, L"InsertDocInfo"); }
bool HwpActionHelper::InsertEndnote(HwpWrapper* hwp) { return Run(hwp, L"InsertEndnote"); }
bool HwpActionHelper::InsertFootnote(HwpWrapper* hwp) { return Run(hwp, L"InsertFootnote"); }
bool HwpActionHelper::InsertFieldCitation(HwpWrapper* hwp) { return Run(hwp, L"InsertFieldCitation"); }
bool HwpActionHelper::InsertFieldDateTime(HwpWrapper* hwp) { return Run(hwp, L"InsertFieldDateTime"); }
bool HwpActionHelper::InsertFieldMemo(HwpWrapper* hwp) { return Run(hwp, L"InsertFieldMemo"); }
bool HwpActionHelper::InsertFieldRevisionChagne(HwpWrapper* hwp) { return Run(hwp, L"InsertFieldRevisionChagne"); }
bool HwpActionHelper::InsertFixedWidthSpace(HwpWrapper* hwp) { return Run(hwp, L"InsertFixedWidthSpace"); }
bool HwpActionHelper::InsertNonBreakingSpace(HwpWrapper* hwp) { return Run(hwp, L"InsertNonBreakingSpace"); }
bool HwpActionHelper::InsertSoftHyphen(HwpWrapper* hwp) { return Run(hwp, L"InsertSoftHyphen"); }
bool HwpActionHelper::InsertSpace(HwpWrapper* hwp) { return Run(hwp, L"InsertSpace"); }
bool HwpActionHelper::InsertTab(HwpWrapper* hwp) { return Run(hwp, L"InsertTab"); }
bool HwpActionHelper::InsertLine(HwpWrapper* hwp) { return Run(hwp, L"InsertLine"); }
bool HwpActionHelper::InsertStringDateTime(HwpWrapper* hwp) { return Run(hwp, L"InsertStringDateTime"); }
bool HwpActionHelper::InsertLastPrintDate(HwpWrapper* hwp) { return Run(hwp, L"InsertLastPrintDate"); }
bool HwpActionHelper::InsertLastSaveBy(HwpWrapper* hwp) { return Run(hwp, L"InsertLastSaveBy"); }
bool HwpActionHelper::InsertLastSaveDate(HwpWrapper* hwp) { return Run(hwp, L"InsertLastSaveDate"); }

// Delete (추가)
bool HwpActionHelper::DeleteLine(HwpWrapper* hwp) { return Run(hwp, L"DeleteLine"); }
bool HwpActionHelper::DeleteLineEnd(HwpWrapper* hwp) { return Run(hwp, L"DeleteLineEnd"); }
bool HwpActionHelper::DeleteWord(HwpWrapper* hwp) { return Run(hwp, L"DeleteWord"); }
bool HwpActionHelper::DeleteWordBack(HwpWrapper* hwp) { return Run(hwp, L"DeleteWordBack"); }
bool HwpActionHelper::DeleteField(HwpWrapper* hwp) { return Run(hwp, L"DeleteField"); }
bool HwpActionHelper::DeleteFieldMemo(HwpWrapper* hwp) { return Run(hwp, L"DeleteFieldMemo"); }
bool HwpActionHelper::DeletePage(HwpWrapper* hwp) { return Run(hwp, L"DeletePage"); }
bool HwpActionHelper::DeletePrivateInfoMark(HwpWrapper* hwp) { return Run(hwp, L"DeletePrivateInfoMark"); }
bool HwpActionHelper::DeletePrivateInfoMarkAtCurrentPos(HwpWrapper* hwp) { return Run(hwp, L"DeletePrivateInfoMarkAtCurrentPos"); }
bool HwpActionHelper::DeleteDocumentMasterPage(HwpWrapper* hwp) { return Run(hwp, L"DeleteDocumentMasterPage"); }
bool HwpActionHelper::DeleteSectionMasterPage(HwpWrapper* hwp) { return Run(hwp, L"DeleteSectionMasterPage"); }

// Clipboard (추가)
bool HwpActionHelper::PasteSpecial(HwpWrapper* hwp) { return Run(hwp, L"PasteSpecial"); }
bool HwpActionHelper::CopyPage(HwpWrapper* hwp) { return Run(hwp, L"CopyPage"); }
bool HwpActionHelper::PastePage(HwpWrapper* hwp) { return Run(hwp, L"PastePage"); }
bool HwpActionHelper::Erase(HwpWrapper* hwp) { return Run(hwp, L"Erase"); }

// Select (추가)
bool HwpActionHelper::Select(HwpWrapper* hwp) { return Run(hwp, L"Select"); }
bool HwpActionHelper::SelectColumn(HwpWrapper* hwp) { return Run(hwp, L"SelectColumn"); }
bool HwpActionHelper::SelectCtrlFront(HwpWrapper* hwp) { return Run(hwp, L"SelectCtrlFront"); }
bool HwpActionHelper::SelectCtrlReverse(HwpWrapper* hwp) { return Run(hwp, L"SelectCtrlReverse"); }
bool HwpActionHelper::UnSelectCtrl(HwpWrapper* hwp) { return Run(hwp, L"UnSelectCtrl"); }

// 기타 핵심 액션
bool HwpActionHelper::Close(HwpWrapper* hwp) { return Run(hwp, L"Close"); }
bool HwpActionHelper::CloseEx(HwpWrapper* hwp) { return Run(hwp, L"CloseEx"); }
bool HwpActionHelper::ToggleOverwrite(HwpWrapper* hwp) { return Run(hwp, L"ToggleOverwrite"); }
bool HwpActionHelper::ReturnPrevPos(HwpWrapper* hwp) { return Run(hwp, L"ReturnPrevPos"); }
bool HwpActionHelper::RecalcPageCount(HwpWrapper* hwp) { return Run(hwp, L"RecalcPageCount"); }
bool HwpActionHelper::SpellingCheck(HwpWrapper* hwp) { return Run(hwp, L"SpellingCheck"); }
bool HwpActionHelper::EasyFind(HwpWrapper* hwp) { return Run(hwp, L"EasyFind"); }

// TrackChange
bool HwpActionHelper::TrackChangeApply(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeApply"); }
bool HwpActionHelper::TrackChangeApplyAll(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeApplyAll"); }
bool HwpActionHelper::TrackChangeApplyNext(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeApplyNext"); }
bool HwpActionHelper::TrackChangeApplyPrev(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeApplyPrev"); }
bool HwpActionHelper::TrackChangeApplyViewAll(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeApplyViewAll"); }
bool HwpActionHelper::TrackChangeCancel(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeCancel"); }
bool HwpActionHelper::TrackChangeCancelAll(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeCancelAll"); }
bool HwpActionHelper::TrackChangeCancelNext(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeCancelNext"); }
bool HwpActionHelper::TrackChangeCancelPrev(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeCancelPrev"); }
bool HwpActionHelper::TrackChangeCancelViewAll(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeCancelViewAll"); }
bool HwpActionHelper::TrackChangeNext(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeNext"); }
bool HwpActionHelper::TrackChangePrev(HwpWrapper* hwp) { return Run(hwp, L"TrackChangePrev"); }
bool HwpActionHelper::TrackChangeAuthor(HwpWrapper* hwp) { return Run(hwp, L"TrackChangeAuthor"); }

// ViewOption
bool HwpActionHelper::ViewOptionCtrlMark(HwpWrapper* hwp) { return Run(hwp, L"ViewOptionCtrlMark"); }
bool HwpActionHelper::ViewOptionGuideLine(HwpWrapper* hwp) { return Run(hwp, L"ViewOptionGuideLine"); }
bool HwpActionHelper::ViewOptionMemo(HwpWrapper* hwp) { return Run(hwp, L"ViewOptionMemo"); }
bool HwpActionHelper::ViewOptionPaper(HwpWrapper* hwp) { return Run(hwp, L"ViewOptionPaper"); }
bool HwpActionHelper::ViewOptionParaMark(HwpWrapper* hwp) { return Run(hwp, L"ViewOptionParaMark"); }
bool HwpActionHelper::ViewOptionPicture(HwpWrapper* hwp) { return Run(hwp, L"ViewOptionPicture"); }
bool HwpActionHelper::ViewZoomFitPage(HwpWrapper* hwp) { return Run(hwp, L"ViewZoomFitPage"); }
bool HwpActionHelper::ViewZoomFitWidth(HwpWrapper* hwp) { return Run(hwp, L"ViewZoomFitWidth"); }
bool HwpActionHelper::ViewTabButton(HwpWrapper* hwp) { return Run(hwp, L"ViewTabButton"); }

// HeaderFooter
bool HwpActionHelper::HeaderFooterDelete(HwpWrapper* hwp) { return Run(hwp, L"HeaderFooterDelete"); }
bool HwpActionHelper::HeaderFooterModify(HwpWrapper* hwp) { return Run(hwp, L"HeaderFooterModify"); }
bool HwpActionHelper::HeaderFooterToNext(HwpWrapper* hwp) { return Run(hwp, L"HeaderFooterToNext"); }
bool HwpActionHelper::HeaderFooterToPrev(HwpWrapper* hwp) { return Run(hwp, L"HeaderFooterToPrev"); }

// Note/Memo
bool HwpActionHelper::NoteDelete(HwpWrapper* hwp) { return Run(hwp, L"NoteDelete"); }
bool HwpActionHelper::NoteModify(HwpWrapper* hwp) { return Run(hwp, L"NoteModify"); }
bool HwpActionHelper::NoteToNext(HwpWrapper* hwp) { return Run(hwp, L"NoteToNext"); }
bool HwpActionHelper::NoteToPrev(HwpWrapper* hwp) { return Run(hwp, L"NoteToPrev"); }
bool HwpActionHelper::MemoToNext(HwpWrapper* hwp) { return Run(hwp, L"MemoToNext"); }
bool HwpActionHelper::MemoToPrev(HwpWrapper* hwp) { return Run(hwp, L"MemoToPrev"); }
bool HwpActionHelper::Comment(HwpWrapper* hwp) { return Run(hwp, L"Comment"); }
bool HwpActionHelper::CommentDelete(HwpWrapper* hwp) { return Run(hwp, L"CommentDelete"); }
bool HwpActionHelper::CommentModify(HwpWrapper* hwp) { return Run(hwp, L"CommentModify"); }
bool HwpActionHelper::ReplyMemo(HwpWrapper* hwp) { return Run(hwp, L"ReplyMemo"); }

// MasterPage
bool HwpActionHelper::MasterPage(HwpWrapper* hwp) { return Run(hwp, L"MasterPage"); }
bool HwpActionHelper::MasterPageDuplicate(HwpWrapper* hwp) { return Run(hwp, L"MasterPageDuplicate"); }
bool HwpActionHelper::MasterPageExcept(HwpWrapper* hwp) { return Run(hwp, L"MasterPageExcept"); }
bool HwpActionHelper::MasterPageFront(HwpWrapper* hwp) { return Run(hwp, L"MasterPageFront"); }
bool HwpActionHelper::MasterPageToNext(HwpWrapper* hwp) { return Run(hwp, L"MasterPageToNext"); }
bool HwpActionHelper::MasterPageToPrevious(HwpWrapper* hwp) { return Run(hwp, L"MasterPageToPrevious"); }

// Picture
bool HwpActionHelper::PictureInsertDialog(HwpWrapper* hwp) { return Run(hwp, L"PictureInsertDialog"); }
bool HwpActionHelper::PictureLinkedToEmbedded(HwpWrapper* hwp) { return Run(hwp, L"PictureLinkedToEmbedded"); }
bool HwpActionHelper::PictureSave(HwpWrapper* hwp) { return Run(hwp, L"PictureSave"); }
bool HwpActionHelper::PictureScissor(HwpWrapper* hwp) { return Run(hwp, L"PictureScissor"); }
bool HwpActionHelper::PictureToOriginal(HwpWrapper* hwp) { return Run(hwp, L"PictureToOriginal"); }

// Window/Frame
bool HwpActionHelper::FrameFullScreen(HwpWrapper* hwp) { return Run(hwp, L"FrameFullScreen"); }
bool HwpActionHelper::FrameFullScreenEnd(HwpWrapper* hwp) { return Run(hwp, L"FrameFullScreenEnd"); }
bool HwpActionHelper::FrameHRuler(HwpWrapper* hwp) { return Run(hwp, L"FrameHRuler"); }
bool HwpActionHelper::FrameVRuler(HwpWrapper* hwp) { return Run(hwp, L"FrameVRuler"); }
bool HwpActionHelper::FrameStatusBar(HwpWrapper* hwp) { return Run(hwp, L"FrameStatusBar"); }
bool HwpActionHelper::WindowAlignCascade(HwpWrapper* hwp) { return Run(hwp, L"WindowAlignCascade"); }
bool HwpActionHelper::WindowAlignTileHorz(HwpWrapper* hwp) { return Run(hwp, L"WindowAlignTileHorz"); }
bool HwpActionHelper::WindowAlignTileVert(HwpWrapper* hwp) { return Run(hwp, L"WindowAlignTileVert"); }
bool HwpActionHelper::WindowList(HwpWrapper* hwp) { return Run(hwp, L"WindowList"); }
bool HwpActionHelper::WindowMinimizeAll(HwpWrapper* hwp) { return Run(hwp, L"WindowMinimizeAll"); }
bool HwpActionHelper::WindowNextPane(HwpWrapper* hwp) { return Run(hwp, L"WindowNextPane"); }
bool HwpActionHelper::WindowNextTab(HwpWrapper* hwp) { return Run(hwp, L"WindowNextTab"); }
bool HwpActionHelper::WindowPrevTab(HwpWrapper* hwp) { return Run(hwp, L"WindowPrevTab"); }
bool HwpActionHelper::SplitAll(HwpWrapper* hwp) { return Run(hwp, L"SplitAll"); }
bool HwpActionHelper::SplitHorz(HwpWrapper* hwp) { return Run(hwp, L"SplitHorz"); }
bool HwpActionHelper::SplitVert(HwpWrapper* hwp) { return Run(hwp, L"SplitVert"); }
bool HwpActionHelper::NoSplit(HwpWrapper* hwp) { return Run(hwp, L"NoSplit"); }

// Input
bool HwpActionHelper::InputCodeChange(HwpWrapper* hwp) { return Run(hwp, L"InputCodeChange"); }
bool HwpActionHelper::InputHanja(HwpWrapper* hwp) { return Run(hwp, L"InputHanja"); }
bool HwpActionHelper::InputHanjaBusu(HwpWrapper* hwp) { return Run(hwp, L"InputHanjaBusu"); }
bool HwpActionHelper::InputHanjaMean(HwpWrapper* hwp) { return Run(hwp, L"InputHanjaMean"); }

// FormObj
bool HwpActionHelper::FormDesignMode(HwpWrapper* hwp) { return Run(hwp, L"FormDesignMode"); }
bool HwpActionHelper::FormObjCreatorCheckButton(HwpWrapper* hwp) { return Run(hwp, L"FormObjCreatorCheckButton"); }
bool HwpActionHelper::FormObjCreatorComboBox(HwpWrapper* hwp) { return Run(hwp, L"FormObjCreatorComboBox"); }
bool HwpActionHelper::FormObjCreatorEdit(HwpWrapper* hwp) { return Run(hwp, L"FormObjCreatorEdit"); }
bool HwpActionHelper::FormObjCreatorListBox(HwpWrapper* hwp) { return Run(hwp, L"FormObjCreatorListBox"); }
bool HwpActionHelper::FormObjCreatorPushButton(HwpWrapper* hwp) { return Run(hwp, L"FormObjCreatorPushButton"); }
bool HwpActionHelper::FormObjCreatorRadioButton(HwpWrapper* hwp) { return Run(hwp, L"FormObjCreatorRadioButton"); }

// Auto
bool HwpActionHelper::ASendBrowserText(HwpWrapper* hwp) { return Run(hwp, L"ASendBrowserText"); }
bool HwpActionHelper::AutoChangeHangul(HwpWrapper* hwp) { return Run(hwp, L"AutoChangeHangul"); }
bool HwpActionHelper::AutoChangeRun(HwpWrapper* hwp) { return Run(hwp, L"AutoChangeRun"); }
bool HwpActionHelper::AutoSpellRun(HwpWrapper* hwp) { return Run(hwp, L"AutoSpellRun"); }

// DrawObj
bool HwpActionHelper::DrawObjCancelOneStep(HwpWrapper* hwp) { return Run(hwp, L"DrawObjCancelOneStep"); }
bool HwpActionHelper::DrawObjEditDetail(HwpWrapper* hwp) { return Run(hwp, L"DrawObjEditDetail"); }
bool HwpActionHelper::DrawObjOpenClosePolygon(HwpWrapper* hwp) { return Run(hwp, L"DrawObjOpenClosePolygon"); }
bool HwpActionHelper::DrawObjTemplateSave(HwpWrapper* hwp) { return Run(hwp, L"DrawObjTemplateSave"); }

// Quick
bool HwpActionHelper::QuickCommandRun(HwpWrapper* hwp) { return Run(hwp, L"QuickCommandRun"); }
bool HwpActionHelper::QuickCorrect(HwpWrapper* hwp) { return Run(hwp, L"QuickCorrect"); }
bool HwpActionHelper::QuickCorrectRun(HwpWrapper* hwp) { return Run(hwp, L"QuickCorrectRun"); }

// Macro
bool HwpActionHelper::MacroPause(HwpWrapper* hwp) { return Run(hwp, L"MacroPause"); }
bool HwpActionHelper::MacroRepeat(HwpWrapper* hwp) { return Run(hwp, L"MacroRepeat"); }
bool HwpActionHelper::MacroStop(HwpWrapper* hwp) { return Run(hwp, L"MacroStop"); }

// Misc
bool HwpActionHelper::MailMergeField(HwpWrapper* hwp) { return Run(hwp, L"MailMergeField"); }
bool HwpActionHelper::MakeIndex(HwpWrapper* hwp) { return Run(hwp, L"MakeIndex"); }
bool HwpActionHelper::LabelAdd(HwpWrapper* hwp) { return Run(hwp, L"LabelAdd"); }
bool HwpActionHelper::LabelTemplate(HwpWrapper* hwp) { return Run(hwp, L"LabelTemplate"); }
bool HwpActionHelper::HanThDIC(HwpWrapper* hwp) { return Run(hwp, L"HanThDIC"); }
bool HwpActionHelper::HwpDic(HwpWrapper* hwp) { return Run(hwp, L"HwpDic"); }
bool HwpActionHelper::ReturnKeyInField(HwpWrapper* hwp) { return Run(hwp, L"ReturnKeyInField"); }

} // namespace cpyhwpx
