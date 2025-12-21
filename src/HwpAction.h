/**
 * @file HwpAction.h
 * @brief HAction COM 객체 래퍼 클래스
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 * HWP의 HAction 인터페이스 래핑
 */

#pragma once

#include "HwpTypes.h"
#include <Windows.h>
#include <comdef.h>
#include <string>
#include <vector>
#include <functional>

namespace cpyhwpx {

// 전방 선언
class HwpWrapper;
class HwpParameterSet;

/**
 * @class HwpAction
 * @brief HAction COM 객체 래퍼
 *
 * HWP의 액션(명령) 실행을 위한 클래스
 * CreateAction() 으로 생성된 액션 객체를 래핑
 */
class HwpAction {
public:
    //=========================================================================
    // 생성자/소멸자
    //=========================================================================

    /**
     * @brief HwpAction 생성자
     * @param action COM 액션 객체
     * @param hwp 부모 HwpWrapper 참조
     */
    HwpAction(IDispatch* action, HwpWrapper* hwp);

    /**
     * @brief 소멸자
     */
    ~HwpAction();

    // 복사 금지
    HwpAction(const HwpAction&) = delete;
    HwpAction& operator=(const HwpAction&) = delete;

    //=========================================================================
    // 액션 실행
    //=========================================================================

    /**
     * @brief 액션 실행 (Run)
     * @return 성공 여부
     */
    bool Run();

    /**
     * @brief 파라미터셋과 함께 액션 실행
     * @param pset 파라미터셋
     * @return 성공 여부
     */
    bool Execute(IDispatch* pset);

    /**
     * @brief 액션 기본 파라미터셋 가져오기
     * @return 파라미터셋 COM 객체
     */
    IDispatch* GetDefault();

    /**
     * @brief 다이얼로그 표시 후 실행
     * @param pset 파라미터셋
     * @return 성공 여부
     */
    bool PopupDialog(IDispatch* pset);

    //=========================================================================
    // 액션 정보
    //=========================================================================

    /**
     * @brief 액션 ID
     */
    std::wstring GetActionID() const;

    /**
     * @brief 유효한 액션인지
     */
    bool IsValid() const { return m_pAction != nullptr; }

    //=========================================================================
    // COM 인터페이스 직접 접근
    //=========================================================================

    /**
     * @brief 내부 COM 객체 접근
     */
    IDispatch* GetDispatch() const { return m_pAction; }

protected:
    /**
     * @brief 메서드 호출
     */
    VARIANT InvokeMethod(const std::wstring& name,
                         const std::vector<VARIANT>& args = {});

private:
    IDispatch* m_pAction;   // HAction COM 포인터
    HwpWrapper* m_pHwp;     // 부모 HwpWrapper
};

/**
 * @class HwpActionHelper
 * @brief HAction.Run() 정적 헬퍼
 *
 * 자주 사용되는 액션을 간편하게 실행
 */
class HwpActionHelper {
public:
    /**
     * @brief 단순 액션 실행
     * @param hwp HwpWrapper 포인터
     * @param action_name 액션 이름
     * @return 성공 여부
     */
    static bool Run(HwpWrapper* hwp, const std::wstring& action_name);

    /**
     * @brief 파라미터셋과 함께 액션 실행
     * @param hwp HwpWrapper 포인터
     * @param action_name 액션 이름
     * @param set_name 파라미터셋 이름
     * @param setter 파라미터 설정 콜백
     * @return 성공 여부
     */
    static bool RunWithSet(HwpWrapper* hwp,
                           const std::wstring& action_name,
                           const std::wstring& set_name,
                           std::function<void(IDispatch*)> setter = nullptr);

    //=========================================================================
    // 편의 메서드 - 자주 사용되는 액션
    //=========================================================================

    // 문단
    static bool BreakPara(HwpWrapper* hwp);
    static bool BreakPage(HwpWrapper* hwp);
    static bool BreakSection(HwpWrapper* hwp);
    static bool BreakColumn(HwpWrapper* hwp);

    // 선택
    static bool SelectAll(HwpWrapper* hwp);
    static bool Cancel(HwpWrapper* hwp);

    // 커서 이동
    static bool MoveLeft(HwpWrapper* hwp);
    static bool MoveRight(HwpWrapper* hwp);
    static bool MoveUp(HwpWrapper* hwp);
    static bool MoveDown(HwpWrapper* hwp);
    static bool MoveLineBegin(HwpWrapper* hwp);
    static bool MoveLineEnd(HwpWrapper* hwp);
    static bool MoveDocBegin(HwpWrapper* hwp);
    static bool MoveDocEnd(HwpWrapper* hwp);
    static bool MoveParaBegin(HwpWrapper* hwp);
    static bool MoveParaEnd(HwpWrapper* hwp);
    static bool MoveWordBegin(HwpWrapper* hwp);
    static bool MoveWordEnd(HwpWrapper* hwp);
    static bool MovePageUp(HwpWrapper* hwp);
    static bool MovePageDown(HwpWrapper* hwp);

    // 편집
    static bool Delete(HwpWrapper* hwp);
    static bool DeleteBack(HwpWrapper* hwp);
    static bool Cut(HwpWrapper* hwp);
    static bool Copy(HwpWrapper* hwp);
    static bool Paste(HwpWrapper* hwp);
    static bool Undo(HwpWrapper* hwp);
    static bool Redo(HwpWrapper* hwp);

    // 셀 이동
    static bool TableCellBlock(HwpWrapper* hwp);
    static bool TableColBegin(HwpWrapper* hwp);
    static bool TableColEnd(HwpWrapper* hwp);
    static bool TableRowBegin(HwpWrapper* hwp);
    static bool TableRowEnd(HwpWrapper* hwp);
    static bool TableCellAddr(HwpWrapper* hwp);
    static bool TableAppendRow(HwpWrapper* hwp);
    static bool TableAppendColumn(HwpWrapper* hwp);
    static bool TableDeleteRow(HwpWrapper* hwp);
    static bool TableDeleteColumn(HwpWrapper* hwp);
    static bool TableMergeCell(HwpWrapper* hwp);
    static bool TableSplitCell(HwpWrapper* hwp);
    static bool TableColPageUp(HwpWrapper* hwp);
    static bool TableCellBlockExtendAbs(HwpWrapper* hwp);

    // 스타일
    static bool ParagraphShapeAlignLeft(HwpWrapper* hwp);
    static bool ParagraphShapeAlignCenter(HwpWrapper* hwp);
    static bool ParagraphShapeAlignRight(HwpWrapper* hwp);
    static bool ParagraphShapeAlignJustify(HwpWrapper* hwp);

    // 글자 모양
    static bool CharShapeBold(HwpWrapper* hwp);
    static bool CharShapeItalic(HwpWrapper* hwp);
    static bool CharShapeUnderline(HwpWrapper* hwp);
    static bool CharShapeStrikeout(HwpWrapper* hwp);
    static bool CharShapeSuperscript(HwpWrapper* hwp);
    static bool CharShapeSubscript(HwpWrapper* hwp);

    // 창/보기
    static bool FileClose(HwpWrapper* hwp);
    static bool FileQuit(HwpWrapper* hwp);
    static bool WindowMaximize(HwpWrapper* hwp);
    static bool WindowMinimize(HwpWrapper* hwp);
    static bool ViewZoomIn(HwpWrapper* hwp);
    static bool ViewZoomOut(HwpWrapper* hwp);

    // 인쇄
    static bool FilePrint(HwpWrapper* hwp);
    static bool FilePrintPreview(HwpWrapper* hwp);

    // 개체 (기존)
    static bool ShapeObjSelect(HwpWrapper* hwp);
    static bool ShapeObjDelete(HwpWrapper* hwp);
    static bool ShapeObjCopy(HwpWrapper* hwp);
    static bool ShapeObjCut(HwpWrapper* hwp);
    static bool ShapeObjBringToFront(HwpWrapper* hwp);
    static bool ShapeObjSendToBack(HwpWrapper* hwp);

    //=========================================================================
    // 추가 Run 액션들
    //=========================================================================

    // Break (추가)
    static bool BreakLine(HwpWrapper* hwp);
    static bool BreakColDef(HwpWrapper* hwp);

    // CharShape (추가)
    static bool CharShapeCenterline(HwpWrapper* hwp);
    static bool CharShapeEmboss(HwpWrapper* hwp);
    static bool CharShapeEngrave(HwpWrapper* hwp);
    static bool CharShapeHeight(HwpWrapper* hwp);
    static bool CharShapeHeightDecrease(HwpWrapper* hwp);
    static bool CharShapeHeightIncrease(HwpWrapper* hwp);
    static bool CharShapeLang(HwpWrapper* hwp);
    static bool CharShapeNextFaceName(HwpWrapper* hwp);
    static bool CharShapeNormal(HwpWrapper* hwp);
    static bool CharShapeOutline(HwpWrapper* hwp);
    static bool CharShapePrevFaceName(HwpWrapper* hwp);
    static bool CharShapeShadow(HwpWrapper* hwp);
    static bool CharShapeSpacing(HwpWrapper* hwp);
    static bool CharShapeSpacingDecrease(HwpWrapper* hwp);
    static bool CharShapeSpacingIncrease(HwpWrapper* hwp);
    static bool CharShapeSuperSubscript(HwpWrapper* hwp);
    static bool CharShapeTextColorBlack(HwpWrapper* hwp);
    static bool CharShapeTextColorBlue(HwpWrapper* hwp);
    static bool CharShapeTextColorBluish(HwpWrapper* hwp);
    static bool CharShapeTextColorGreen(HwpWrapper* hwp);
    static bool CharShapeTextColorRed(HwpWrapper* hwp);
    static bool CharShapeTextColorViolet(HwpWrapper* hwp);
    static bool CharShapeTextColorWhite(HwpWrapper* hwp);
    static bool CharShapeTextColorYellow(HwpWrapper* hwp);
    static bool CharShapeTypeface(HwpWrapper* hwp);
    static bool CharShapeWidth(HwpWrapper* hwp);
    static bool CharShapeWidthDecrease(HwpWrapper* hwp);
    static bool CharShapeWidthIncrease(HwpWrapper* hwp);

    // ParagraphShape (추가)
    static bool ParagraphShapeAlignDistribute(HwpWrapper* hwp);
    static bool ParagraphShapeAlignDivision(HwpWrapper* hwp);
    static bool ParagraphShapeDecreaseLeftMargin(HwpWrapper* hwp);
    static bool ParagraphShapeDecreaseLineSpacing(HwpWrapper* hwp);
    static bool ParagraphShapeDecreaseMargin(HwpWrapper* hwp);
    static bool ParagraphShapeDecreaseRightMargin(HwpWrapper* hwp);
    static bool ParagraphShapeIncreaseLeftMargin(HwpWrapper* hwp);
    static bool ParagraphShapeIncreaseLineSpacing(HwpWrapper* hwp);
    static bool ParagraphShapeIncreaseMargin(HwpWrapper* hwp);
    static bool ParagraphShapeIncreaseRightMargin(HwpWrapper* hwp);
    static bool ParagraphShapeIndentAtCaret(HwpWrapper* hwp);
    static bool ParagraphShapeIndentNegative(HwpWrapper* hwp);
    static bool ParagraphShapeIndentPositive(HwpWrapper* hwp);
    static bool ParagraphShapeProtect(HwpWrapper* hwp);
    static bool ParagraphShapeSingleRow(HwpWrapper* hwp);
    static bool ParagraphShapeWithNext(HwpWrapper* hwp);

    // Move (추가)
    static bool MoveLineDown(HwpWrapper* hwp);
    static bool MoveLineUp(HwpWrapper* hwp);
    static bool MoveNextWord(HwpWrapper* hwp);
    static bool MovePrevWord(HwpWrapper* hwp);
    static bool MoveNextChar(HwpWrapper* hwp);
    static bool MovePrevChar(HwpWrapper* hwp);
    static bool MovePageBegin(HwpWrapper* hwp);
    static bool MovePageEnd(HwpWrapper* hwp);
    static bool MoveSelDown(HwpWrapper* hwp);
    static bool MoveSelLeft(HwpWrapper* hwp);
    static bool MoveSelRight(HwpWrapper* hwp);
    static bool MoveSelUp(HwpWrapper* hwp);
    static bool MoveSelDocBegin(HwpWrapper* hwp);
    static bool MoveSelDocEnd(HwpWrapper* hwp);
    static bool MoveSelLineBegin(HwpWrapper* hwp);
    static bool MoveSelLineEnd(HwpWrapper* hwp);
    static bool MoveSelLineDown(HwpWrapper* hwp);
    static bool MoveSelLineUp(HwpWrapper* hwp);
    static bool MoveSelNextChar(HwpWrapper* hwp);
    static bool MoveSelPrevChar(HwpWrapper* hwp);
    static bool MoveSelNextWord(HwpWrapper* hwp);
    static bool MoveSelPrevWord(HwpWrapper* hwp);
    static bool MoveSelPageDown(HwpWrapper* hwp);
    static bool MoveSelPageUp(HwpWrapper* hwp);
    static bool MoveSelParaBegin(HwpWrapper* hwp);
    static bool MoveSelParaEnd(HwpWrapper* hwp);
    static bool MoveSelWordBegin(HwpWrapper* hwp);
    static bool MoveSelWordEnd(HwpWrapper* hwp);
    static bool MoveColumnBegin(HwpWrapper* hwp);
    static bool MoveColumnEnd(HwpWrapper* hwp);
    static bool MoveListBegin(HwpWrapper* hwp);
    static bool MoveListEnd(HwpWrapper* hwp);
    static bool MoveParentList(HwpWrapper* hwp);
    static bool MoveRootList(HwpWrapper* hwp);
    static bool MoveTopLevelBegin(HwpWrapper* hwp);
    static bool MoveTopLevelEnd(HwpWrapper* hwp);
    static bool MoveTopLevelList(HwpWrapper* hwp);
    static bool MoveNextColumn(HwpWrapper* hwp);
    static bool MovePrevColumn(HwpWrapper* hwp);
    static bool MoveNextPos(HwpWrapper* hwp);
    static bool MovePrevPos(HwpWrapper* hwp);
    static bool MoveNextPosEx(HwpWrapper* hwp);
    static bool MovePrevPosEx(HwpWrapper* hwp);
    static bool MoveNextParaBegin(HwpWrapper* hwp);
    static bool MovePrevParaBegin(HwpWrapper* hwp);
    static bool MovePrevParaEnd(HwpWrapper* hwp);
    static bool MoveSectionDown(HwpWrapper* hwp);
    static bool MoveSectionUp(HwpWrapper* hwp);
    static bool MoveScrollDown(HwpWrapper* hwp);
    static bool MoveScrollUp(HwpWrapper* hwp);
    static bool MoveScrollNext(HwpWrapper* hwp);
    static bool MoveScrollPrev(HwpWrapper* hwp);
    static bool MoveViewBegin(HwpWrapper* hwp);
    static bool MoveViewEnd(HwpWrapper* hwp);
    static bool MoveViewDown(HwpWrapper* hwp);
    static bool MoveViewUp(HwpWrapper* hwp);

    // Table (추가)
    static bool TableLeftCell(HwpWrapper* hwp);
    static bool TableRightCell(HwpWrapper* hwp);
    static bool TableUpperCell(HwpWrapper* hwp);
    static bool TableLowerCell(HwpWrapper* hwp);
    static bool TableColPageDown(HwpWrapper* hwp);
    static bool TableRightCellAppend(HwpWrapper* hwp);
    static bool TableCellBlockCol(HwpWrapper* hwp);
    static bool TableCellBlockRow(HwpWrapper* hwp);
    static bool TableCellBlockExtend(HwpWrapper* hwp);
    static bool TableCellAlignLeftTop(HwpWrapper* hwp);
    static bool TableCellAlignCenterTop(HwpWrapper* hwp);
    static bool TableCellAlignRightTop(HwpWrapper* hwp);
    static bool TableCellAlignLeftCenter(HwpWrapper* hwp);
    static bool TableCellAlignCenterCenter(HwpWrapper* hwp);
    static bool TableCellAlignRightCenter(HwpWrapper* hwp);
    static bool TableCellAlignLeftBottom(HwpWrapper* hwp);
    static bool TableCellAlignCenterBottom(HwpWrapper* hwp);
    static bool TableCellAlignRightBottom(HwpWrapper* hwp);
    static bool TableVAlignTop(HwpWrapper* hwp);
    static bool TableVAlignCenter(HwpWrapper* hwp);
    static bool TableVAlignBottom(HwpWrapper* hwp);
    static bool TableCellBorderAll(HwpWrapper* hwp);
    static bool TableCellBorderNo(HwpWrapper* hwp);
    static bool TableCellBorderOutside(HwpWrapper* hwp);
    static bool TableCellBorderInside(HwpWrapper* hwp);
    static bool TableCellBorderInsideHorz(HwpWrapper* hwp);
    static bool TableCellBorderInsideVert(HwpWrapper* hwp);
    static bool TableCellBorderTop(HwpWrapper* hwp);
    static bool TableCellBorderBottom(HwpWrapper* hwp);
    static bool TableCellBorderLeft(HwpWrapper* hwp);
    static bool TableCellBorderRight(HwpWrapper* hwp);
    static bool TableCellBorderDiagonalDown(HwpWrapper* hwp);
    static bool TableCellBorderDiagonalUp(HwpWrapper* hwp);
    static bool TableSubtractRow(HwpWrapper* hwp);
    static bool TableDeleteCell(HwpWrapper* hwp);
    static bool TableMergeTable(HwpWrapper* hwp);
    static bool TableSplitTable(HwpWrapper* hwp);
    static bool TableDistributeCellHeight(HwpWrapper* hwp);
    static bool TableDistributeCellWidth(HwpWrapper* hwp);
    static bool TableResizeDown(HwpWrapper* hwp);
    static bool TableResizeUp(HwpWrapper* hwp);
    static bool TableResizeLeft(HwpWrapper* hwp);
    static bool TableResizeRight(HwpWrapper* hwp);
    static bool TableResizeCellDown(HwpWrapper* hwp);
    static bool TableResizeCellUp(HwpWrapper* hwp);
    static bool TableResizeCellLeft(HwpWrapper* hwp);
    static bool TableResizeCellRight(HwpWrapper* hwp);
    static bool TableResizeExDown(HwpWrapper* hwp);
    static bool TableResizeExUp(HwpWrapper* hwp);
    static bool TableResizeExLeft(HwpWrapper* hwp);
    static bool TableResizeExRight(HwpWrapper* hwp);
    static bool TableResizeLineDown(HwpWrapper* hwp);
    static bool TableResizeLineUp(HwpWrapper* hwp);
    static bool TableResizeLineLeft(HwpWrapper* hwp);
    static bool TableResizeLineRight(HwpWrapper* hwp);
    static bool TableFormulaSumAuto(HwpWrapper* hwp);
    static bool TableFormulaSumHor(HwpWrapper* hwp);
    static bool TableFormulaSumVer(HwpWrapper* hwp);
    static bool TableFormulaAvgAuto(HwpWrapper* hwp);
    static bool TableFormulaAvgHor(HwpWrapper* hwp);
    static bool TableFormulaAvgVer(HwpWrapper* hwp);
    static bool TableFormulaProAuto(HwpWrapper* hwp);
    static bool TableFormulaProHor(HwpWrapper* hwp);
    static bool TableFormulaProVer(HwpWrapper* hwp);
    static bool TableAutoFill(HwpWrapper* hwp);
    static bool TableAutoFillDlg(HwpWrapper* hwp);
    static bool TableDrawPen(HwpWrapper* hwp);
    static bool TableDrawPenStyle(HwpWrapper* hwp);
    static bool TableDrawPenWidth(HwpWrapper* hwp);
    static bool TableEraser(HwpWrapper* hwp);
    static bool TableDeleteComma(HwpWrapper* hwp);
    static bool TableInsertComma(HwpWrapper* hwp);

    // ShapeObj (추가)
    static bool ShapeObjAlignLeft(HwpWrapper* hwp);
    static bool ShapeObjAlignCenter(HwpWrapper* hwp);
    static bool ShapeObjAlignRight(HwpWrapper* hwp);
    static bool ShapeObjAlignTop(HwpWrapper* hwp);
    static bool ShapeObjAlignMiddle(HwpWrapper* hwp);
    static bool ShapeObjAlignBottom(HwpWrapper* hwp);
    static bool ShapeObjAlignWidth(HwpWrapper* hwp);
    static bool ShapeObjAlignHeight(HwpWrapper* hwp);
    static bool ShapeObjAlignSize(HwpWrapper* hwp);
    static bool ShapeObjAlignHorzSpacing(HwpWrapper* hwp);
    static bool ShapeObjAlignVertSpacing(HwpWrapper* hwp);
    static bool ShapeObjBringForward(HwpWrapper* hwp);
    static bool ShapeObjSendBack(HwpWrapper* hwp);
    static bool ShapeObjBringInFrontOfText(HwpWrapper* hwp);
    static bool ShapeObjCtrlSendBehindText(HwpWrapper* hwp);
    static bool ShapeObjHorzFlip(HwpWrapper* hwp);
    static bool ShapeObjVertFlip(HwpWrapper* hwp);
    static bool ShapeObjHorzFlipOrgState(HwpWrapper* hwp);
    static bool ShapeObjVertFlipOrgState(HwpWrapper* hwp);
    static bool ShapeObjRotater(HwpWrapper* hwp);
    static bool ShapeObjRightAngleRotater(HwpWrapper* hwp);
    static bool ShapeObjRightAngleRotaterAnticlockwise(HwpWrapper* hwp);
    static bool ShapeObjMoveUp(HwpWrapper* hwp);
    static bool ShapeObjMoveDown(HwpWrapper* hwp);
    static bool ShapeObjMoveLeft(HwpWrapper* hwp);
    static bool ShapeObjMoveRight(HwpWrapper* hwp);
    static bool ShapeObjResizeUp(HwpWrapper* hwp);
    static bool ShapeObjResizeDown(HwpWrapper* hwp);
    static bool ShapeObjResizeLeft(HwpWrapper* hwp);
    static bool ShapeObjResizeRight(HwpWrapper* hwp);
    static bool ShapeObjGroup(HwpWrapper* hwp);
    static bool ShapeObjUngroup(HwpWrapper* hwp);
    static bool ShapeObjNextObject(HwpWrapper* hwp);
    static bool ShapeObjPrevObject(HwpWrapper* hwp);
    static bool ShapeObjLock(HwpWrapper* hwp);
    static bool ShapeObjUnlockAll(HwpWrapper* hwp);
    static bool ShapeObjAttachCaption(HwpWrapper* hwp);
    static bool ShapeObjDetachCaption(HwpWrapper* hwp);
    static bool ShapeObjInsertCaptionNum(HwpWrapper* hwp);
    static bool ShapeObjAttachTextBox(HwpWrapper* hwp);
    static bool ShapeObjDetachTextBox(HwpWrapper* hwp);
    static bool ShapeObjToggleTextBox(HwpWrapper* hwp);
    static bool ShapeObjTextBoxEdit(HwpWrapper* hwp);
    static bool ShapeObjTableSelCell(HwpWrapper* hwp);
    static bool ShapeObjFillProperty(HwpWrapper* hwp);
    static bool ShapeObjLineProperty(HwpWrapper* hwp);
    static bool ShapeObjLineStyleOther(HwpWrapper* hwp);
    static bool ShapeObjLineWidthOther(HwpWrapper* hwp);
    static bool ShapeObjNorm(HwpWrapper* hwp);
    static bool ShapeObjGuideLine(HwpWrapper* hwp);
    static bool ShapeObjShowGuideLine(HwpWrapper* hwp);
    static bool ShapeObjShowGuideLineBase(HwpWrapper* hwp);
    static bool ShapeObjWrapSquare(HwpWrapper* hwp);
    static bool ShapeObjWrapTopAndBottom(HwpWrapper* hwp);
    static bool ShapeObjSaveAsPicture(HwpWrapper* hwp);

    // File (추가)
    static bool FileNew(HwpWrapper* hwp);
    static bool FileNewTab(HwpWrapper* hwp);
    static bool FileOpen(HwpWrapper* hwp);
    static bool FileOpenMRU(HwpWrapper* hwp);
    static bool FileSave(HwpWrapper* hwp);
    static bool FileSaveAs(HwpWrapper* hwp);
    static bool FileSaveAsDRM(HwpWrapper* hwp);
    static bool FileSaveOptionDlg(HwpWrapper* hwp);
    static bool FileFind(HwpWrapper* hwp);
    static bool FilePreview(HwpWrapper* hwp);
    static bool FileNextVersionDiff(HwpWrapper* hwp);
    static bool FilePrevVersionDiff(HwpWrapper* hwp);
    static bool FileVersionDiffChangeAlign(HwpWrapper* hwp);
    static bool FileVersionDiffSameAlign(HwpWrapper* hwp);
    static bool FileVersionDiffSyncScroll(HwpWrapper* hwp);

    // Insert
    static bool InsertAutoNum(HwpWrapper* hwp);
    static bool InsertCpNo(HwpWrapper* hwp);
    static bool InsertCpTpNo(HwpWrapper* hwp);
    static bool InsertTpNo(HwpWrapper* hwp);
    static bool InsertPageNum(HwpWrapper* hwp);
    static bool InsertDateCode(HwpWrapper* hwp);
    static bool InsertDocInfo(HwpWrapper* hwp);
    static bool InsertEndnote(HwpWrapper* hwp);
    static bool InsertFootnote(HwpWrapper* hwp);
    static bool InsertFieldCitation(HwpWrapper* hwp);
    static bool InsertFieldDateTime(HwpWrapper* hwp);
    static bool InsertFieldMemo(HwpWrapper* hwp);
    static bool InsertFieldRevisionChagne(HwpWrapper* hwp);
    static bool InsertFixedWidthSpace(HwpWrapper* hwp);
    static bool InsertNonBreakingSpace(HwpWrapper* hwp);
    static bool InsertSoftHyphen(HwpWrapper* hwp);
    static bool InsertSpace(HwpWrapper* hwp);
    static bool InsertTab(HwpWrapper* hwp);
    static bool InsertLine(HwpWrapper* hwp);
    static bool InsertStringDateTime(HwpWrapper* hwp);
    static bool InsertLastPrintDate(HwpWrapper* hwp);
    static bool InsertLastSaveBy(HwpWrapper* hwp);
    static bool InsertLastSaveDate(HwpWrapper* hwp);

    // Delete (추가)
    static bool DeleteLine(HwpWrapper* hwp);
    static bool DeleteLineEnd(HwpWrapper* hwp);
    static bool DeleteWord(HwpWrapper* hwp);
    static bool DeleteWordBack(HwpWrapper* hwp);
    static bool DeleteField(HwpWrapper* hwp);
    static bool DeleteFieldMemo(HwpWrapper* hwp);
    static bool DeletePage(HwpWrapper* hwp);
    static bool DeletePrivateInfoMark(HwpWrapper* hwp);
    static bool DeletePrivateInfoMarkAtCurrentPos(HwpWrapper* hwp);
    static bool DeleteDocumentMasterPage(HwpWrapper* hwp);
    static bool DeleteSectionMasterPage(HwpWrapper* hwp);

    // Clipboard (추가)
    static bool PasteSpecial(HwpWrapper* hwp);
    static bool CopyPage(HwpWrapper* hwp);
    static bool PastePage(HwpWrapper* hwp);
    static bool Erase(HwpWrapper* hwp);

    // Select (추가)
    static bool Select(HwpWrapper* hwp);
    static bool SelectColumn(HwpWrapper* hwp);
    static bool SelectCtrlFront(HwpWrapper* hwp);
    static bool SelectCtrlReverse(HwpWrapper* hwp);
    static bool UnSelectCtrl(HwpWrapper* hwp);

    // 기타 핵심 액션
    static bool Close(HwpWrapper* hwp);
    static bool CloseEx(HwpWrapper* hwp);
    static bool ToggleOverwrite(HwpWrapper* hwp);
    static bool ReturnPrevPos(HwpWrapper* hwp);
    static bool RecalcPageCount(HwpWrapper* hwp);
    static bool SpellingCheck(HwpWrapper* hwp);
    static bool EasyFind(HwpWrapper* hwp);

    // TrackChange
    static bool TrackChangeApply(HwpWrapper* hwp);
    static bool TrackChangeApplyAll(HwpWrapper* hwp);
    static bool TrackChangeApplyNext(HwpWrapper* hwp);
    static bool TrackChangeApplyPrev(HwpWrapper* hwp);
    static bool TrackChangeApplyViewAll(HwpWrapper* hwp);
    static bool TrackChangeCancel(HwpWrapper* hwp);
    static bool TrackChangeCancelAll(HwpWrapper* hwp);
    static bool TrackChangeCancelNext(HwpWrapper* hwp);
    static bool TrackChangeCancelPrev(HwpWrapper* hwp);
    static bool TrackChangeCancelViewAll(HwpWrapper* hwp);
    static bool TrackChangeNext(HwpWrapper* hwp);
    static bool TrackChangePrev(HwpWrapper* hwp);
    static bool TrackChangeAuthor(HwpWrapper* hwp);

    // ViewOption
    static bool ViewOptionCtrlMark(HwpWrapper* hwp);
    static bool ViewOptionGuideLine(HwpWrapper* hwp);
    static bool ViewOptionMemo(HwpWrapper* hwp);
    static bool ViewOptionPaper(HwpWrapper* hwp);
    static bool ViewOptionParaMark(HwpWrapper* hwp);
    static bool ViewOptionPicture(HwpWrapper* hwp);
    static bool ViewZoomFitPage(HwpWrapper* hwp);
    static bool ViewZoomFitWidth(HwpWrapper* hwp);
    static bool ViewTabButton(HwpWrapper* hwp);

    // HeaderFooter
    static bool HeaderFooterDelete(HwpWrapper* hwp);
    static bool HeaderFooterModify(HwpWrapper* hwp);
    static bool HeaderFooterToNext(HwpWrapper* hwp);
    static bool HeaderFooterToPrev(HwpWrapper* hwp);

    // Note/Memo
    static bool NoteDelete(HwpWrapper* hwp);
    static bool NoteModify(HwpWrapper* hwp);
    static bool NoteToNext(HwpWrapper* hwp);
    static bool NoteToPrev(HwpWrapper* hwp);
    static bool MemoToNext(HwpWrapper* hwp);
    static bool MemoToPrev(HwpWrapper* hwp);
    static bool Comment(HwpWrapper* hwp);
    static bool CommentDelete(HwpWrapper* hwp);
    static bool CommentModify(HwpWrapper* hwp);
    static bool ReplyMemo(HwpWrapper* hwp);

    // MasterPage
    static bool MasterPage(HwpWrapper* hwp);
    static bool MasterPageDuplicate(HwpWrapper* hwp);
    static bool MasterPageExcept(HwpWrapper* hwp);
    static bool MasterPageFront(HwpWrapper* hwp);
    static bool MasterPageToNext(HwpWrapper* hwp);
    static bool MasterPageToPrevious(HwpWrapper* hwp);

    // Picture
    static bool PictureInsertDialog(HwpWrapper* hwp);
    static bool PictureLinkedToEmbedded(HwpWrapper* hwp);
    static bool PictureSave(HwpWrapper* hwp);
    static bool PictureScissor(HwpWrapper* hwp);
    static bool PictureToOriginal(HwpWrapper* hwp);

    // Window/Frame
    static bool FrameFullScreen(HwpWrapper* hwp);
    static bool FrameFullScreenEnd(HwpWrapper* hwp);
    static bool FrameHRuler(HwpWrapper* hwp);
    static bool FrameVRuler(HwpWrapper* hwp);
    static bool FrameStatusBar(HwpWrapper* hwp);
    static bool WindowAlignCascade(HwpWrapper* hwp);
    static bool WindowAlignTileHorz(HwpWrapper* hwp);
    static bool WindowAlignTileVert(HwpWrapper* hwp);
    static bool WindowList(HwpWrapper* hwp);
    static bool WindowMinimizeAll(HwpWrapper* hwp);
    static bool WindowNextPane(HwpWrapper* hwp);
    static bool WindowNextTab(HwpWrapper* hwp);
    static bool WindowPrevTab(HwpWrapper* hwp);
    static bool SplitAll(HwpWrapper* hwp);
    static bool SplitHorz(HwpWrapper* hwp);
    static bool SplitVert(HwpWrapper* hwp);
    static bool NoSplit(HwpWrapper* hwp);

    // Input
    static bool InputCodeChange(HwpWrapper* hwp);
    static bool InputHanja(HwpWrapper* hwp);
    static bool InputHanjaBusu(HwpWrapper* hwp);
    static bool InputHanjaMean(HwpWrapper* hwp);

    // FormObj
    static bool FormDesignMode(HwpWrapper* hwp);
    static bool FormObjCreatorCheckButton(HwpWrapper* hwp);
    static bool FormObjCreatorComboBox(HwpWrapper* hwp);
    static bool FormObjCreatorEdit(HwpWrapper* hwp);
    static bool FormObjCreatorListBox(HwpWrapper* hwp);
    static bool FormObjCreatorPushButton(HwpWrapper* hwp);
    static bool FormObjCreatorRadioButton(HwpWrapper* hwp);

    // Auto
    static bool ASendBrowserText(HwpWrapper* hwp);
    static bool AutoChangeHangul(HwpWrapper* hwp);
    static bool AutoChangeRun(HwpWrapper* hwp);
    static bool AutoSpellRun(HwpWrapper* hwp);

    // DrawObj
    static bool DrawObjCancelOneStep(HwpWrapper* hwp);
    static bool DrawObjEditDetail(HwpWrapper* hwp);
    static bool DrawObjOpenClosePolygon(HwpWrapper* hwp);
    static bool DrawObjTemplateSave(HwpWrapper* hwp);

    // Quick
    static bool QuickCommandRun(HwpWrapper* hwp);
    static bool QuickCorrect(HwpWrapper* hwp);
    static bool QuickCorrectRun(HwpWrapper* hwp);

    // Macro
    static bool MacroPause(HwpWrapper* hwp);
    static bool MacroRepeat(HwpWrapper* hwp);
    static bool MacroStop(HwpWrapper* hwp);

    // Misc
    static bool MailMergeField(HwpWrapper* hwp);
    static bool MakeIndex(HwpWrapper* hwp);
    static bool LabelAdd(HwpWrapper* hwp);
    static bool LabelTemplate(HwpWrapper* hwp);
    static bool HanThDIC(HwpWrapper* hwp);
    static bool HwpDic(HwpWrapper* hwp);
    static bool ReturnKeyInField(HwpWrapper* hwp);

    //=========================================================================
    // 추가 Run 액션들 (228개 신규)
    //=========================================================================

    // Auto - AutoSpellSelect (17개)
    static bool AutoSpellSelect0(HwpWrapper* hwp);
    static bool AutoSpellSelect1(HwpWrapper* hwp);
    static bool AutoSpellSelect2(HwpWrapper* hwp);
    static bool AutoSpellSelect3(HwpWrapper* hwp);
    static bool AutoSpellSelect4(HwpWrapper* hwp);
    static bool AutoSpellSelect5(HwpWrapper* hwp);
    static bool AutoSpellSelect6(HwpWrapper* hwp);
    static bool AutoSpellSelect7(HwpWrapper* hwp);
    static bool AutoSpellSelect8(HwpWrapper* hwp);
    static bool AutoSpellSelect9(HwpWrapper* hwp);
    static bool AutoSpellSelect10(HwpWrapper* hwp);
    static bool AutoSpellSelect11(HwpWrapper* hwp);
    static bool AutoSpellSelect12(HwpWrapper* hwp);
    static bool AutoSpellSelect13(HwpWrapper* hwp);
    static bool AutoSpellSelect14(HwpWrapper* hwp);
    static bool AutoSpellSelect15(HwpWrapper* hwp);
    static bool AutoSpellSelect16(HwpWrapper* hwp);

    // ViewOption (14개)
    static bool ViewIdiom(HwpWrapper* hwp);
    static bool ViewOptionMemoGuideline(HwpWrapper* hwp);
    static bool ViewOptionRevision(HwpWrapper* hwp);
    static bool ViewOptionTrackChange(HwpWrapper* hwp);
    static bool ViewOptionTrackChangeFinal(HwpWrapper* hwp);
    static bool ViewOptionTrackChangeFinalMemo(HwpWrapper* hwp);
    static bool ViewOptionTrackChangeInline(HwpWrapper* hwp);
    static bool ViewOptionTrackChangeInsertDelete(HwpWrapper* hwp);
    static bool ViewOptionTrackChangeOriginal(HwpWrapper* hwp);
    static bool ViewOptionTrackChangeOriginalMemo(HwpWrapper* hwp);
    static bool ViewOptionTrackChangeShape(HwpWrapper* hwp);
    static bool ViewOptionTrackChnageInfo(HwpWrapper* hwp);
    static bool ViewZoomNormal(HwpWrapper* hwp);
    static bool ViewZoomRibon(HwpWrapper* hwp);

    // Macro (22개)
    static bool MacroPlay1(HwpWrapper* hwp);
    static bool MacroPlay2(HwpWrapper* hwp);
    static bool MacroPlay3(HwpWrapper* hwp);
    static bool MacroPlay4(HwpWrapper* hwp);
    static bool MacroPlay5(HwpWrapper* hwp);
    static bool MacroPlay6(HwpWrapper* hwp);
    static bool MacroPlay7(HwpWrapper* hwp);
    static bool MacroPlay8(HwpWrapper* hwp);
    static bool MacroPlay9(HwpWrapper* hwp);
    static bool MacroPlay10(HwpWrapper* hwp);
    static bool MacroPlay11(HwpWrapper* hwp);
    static bool ScrMacroPause(HwpWrapper* hwp);
    static bool ScrMacroRepeat(HwpWrapper* hwp);
    static bool ScrMacroStop(HwpWrapper* hwp);
    static bool ScrMacroPlay1(HwpWrapper* hwp);
    static bool ScrMacroPlay2(HwpWrapper* hwp);
    static bool ScrMacroPlay3(HwpWrapper* hwp);
    static bool ScrMacroPlay4(HwpWrapper* hwp);
    static bool ScrMacroPlay5(HwpWrapper* hwp);
    static bool ScrMacroPlay6(HwpWrapper* hwp);
    static bool ScrMacroPlay7(HwpWrapper* hwp);
    static bool ScrMacroPlay8(HwpWrapper* hwp);

    // Quick (21개)
    static bool QuickCorrectSound(HwpWrapper* hwp);
    static bool QuickMarkInsert0(HwpWrapper* hwp);
    static bool QuickMarkInsert1(HwpWrapper* hwp);
    static bool QuickMarkInsert2(HwpWrapper* hwp);
    static bool QuickMarkInsert3(HwpWrapper* hwp);
    static bool QuickMarkInsert4(HwpWrapper* hwp);
    static bool QuickMarkInsert5(HwpWrapper* hwp);
    static bool QuickMarkInsert6(HwpWrapper* hwp);
    static bool QuickMarkInsert7(HwpWrapper* hwp);
    static bool QuickMarkInsert8(HwpWrapper* hwp);
    static bool QuickMarkInsert9(HwpWrapper* hwp);
    static bool QuickMarkMove0(HwpWrapper* hwp);
    static bool QuickMarkMove1(HwpWrapper* hwp);
    static bool QuickMarkMove2(HwpWrapper* hwp);
    static bool QuickMarkMove3(HwpWrapper* hwp);
    static bool QuickMarkMove4(HwpWrapper* hwp);
    static bool QuickMarkMove5(HwpWrapper* hwp);
    static bool QuickMarkMove6(HwpWrapper* hwp);
    static bool QuickMarkMove7(HwpWrapper* hwp);
    static bool QuickMarkMove8(HwpWrapper* hwp);
    static bool QuickMarkMove9(HwpWrapper* hwp);

    // MasterPage (5개)
    static bool MasterPagePrevSection(HwpWrapper* hwp);
    static bool MasterPageType(HwpWrapper* hwp);
    static bool MPSectionToNext(HwpWrapper* hwp);
    static bool MPSectionToPrevious(HwpWrapper* hwp);
    static bool MPShowMarginBorder(HwpWrapper* hwp);

    // Picture (8개)
    static bool PictureEffect1(HwpWrapper* hwp);
    static bool PictureEffect2(HwpWrapper* hwp);
    static bool PictureEffect3(HwpWrapper* hwp);
    static bool PictureEffect4(HwpWrapper* hwp);
    static bool PictureEffect5(HwpWrapper* hwp);
    static bool PictureEffect6(HwpWrapper* hwp);
    static bool PictureEffect7(HwpWrapper* hwp);
    static bool PictureEffect8(HwpWrapper* hwp);

    // Note/Memo (8개)
    static bool NoteNumProperty(HwpWrapper* hwp);
    static bool NoteNumShape(HwpWrapper* hwp);
    static bool NoteLineColor(HwpWrapper* hwp);
    static bool NoteLineLength(HwpWrapper* hwp);
    static bool NoteLineShape(HwpWrapper* hwp);
    static bool NoteLineWeight(HwpWrapper* hwp);
    static bool NotePosition(HwpWrapper* hwp);
    static bool EditFieldMemo(HwpWrapper* hwp);

    // FormObj (2개)
    static bool FormObjCreatorScrollBar(HwpWrapper* hwp);
    static bool FormObjRadioGroup(HwpWrapper* hwp);

    // Window/Frame (5개)
    static bool FrameViewZoomRibon(HwpWrapper* hwp);
    static bool SplitMemo(HwpWrapper* hwp);
    static bool SplitMemoClose(HwpWrapper* hwp);
    static bool SplitMemoOpen(HwpWrapper* hwp);
    static bool SplitMainActive(HwpWrapper* hwp);

    // Misc 추가 (기타)
    static bool Jajun(HwpWrapper* hwp);
    static bool ChangeSkin(HwpWrapper* hwp);
    static bool SoftKeyboard(HwpWrapper* hwp);
};

} // namespace cpyhwpx
