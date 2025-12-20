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

    // 개체
    static bool ShapeObjSelect(HwpWrapper* hwp);
    static bool ShapeObjDelete(HwpWrapper* hwp);
    static bool ShapeObjCopy(HwpWrapper* hwp);
    static bool ShapeObjCut(HwpWrapper* hwp);
    static bool ShapeObjBringToFront(HwpWrapper* hwp);
    static bool ShapeObjSendToBack(HwpWrapper* hwp);
};

} // namespace cpyhwpx
