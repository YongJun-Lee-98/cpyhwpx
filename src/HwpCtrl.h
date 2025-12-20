/**
 * @file HwpCtrl.h
 * @brief HWP 컨트롤(표, 그림, 글상자 등) 래퍼 클래스
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 * Python의 Ctrl 클래스에 대응
 */

#pragma once

#include "HwpTypes.h"
#include <Windows.h>
#include <comdef.h>
#include <memory>
#include <string>

namespace cpyhwpx {

// 전방 선언
class HwpWrapper;

/**
 * @class HwpCtrl
 * @brief HWP 컨트롤 래퍼 클래스
 *
 * 표, 그림, 글상자, 수식 등 HWP 문서 내 컨트롤을 래핑
 */
class HwpCtrl {
public:
    //=========================================================================
    // 생성자/소멸자
    //=========================================================================

    /**
     * @brief HwpCtrl 생성자
     * @param ctrl COM 컨트롤 객체
     * @param hwp 부모 HwpWrapper 참조
     */
    HwpCtrl(IDispatch* ctrl, HwpWrapper* hwp);

    /**
     * @brief 소멸자
     */
    ~HwpCtrl();

    // 복사 금지, 이동 허용
    HwpCtrl(const HwpCtrl&) = delete;
    HwpCtrl& operator=(const HwpCtrl&) = delete;
    HwpCtrl(HwpCtrl&& other) noexcept;
    HwpCtrl& operator=(HwpCtrl&& other) noexcept;

    //=========================================================================
    // 컨트롤 정보
    //=========================================================================

    /**
     * @brief 컨트롤 ID 문자열 (tbl, gso, eqed 등)
     */
    std::wstring GetCtrlID() const;

    /**
     * @brief 컨트롤 타입 열거형
     */
    CtrlType GetCtrlType() const;

    /**
     * @brief 컨트롤 문자 (1~31)
     */
    int GetCtrlCh() const;

    /**
     * @brief 유효한 컨트롤인지 확인
     */
    bool IsValid() const { return m_pCtrl != nullptr; }

    //=========================================================================
    // 탐색
    //=========================================================================

    /**
     * @brief 다음 컨트롤
     * @return 다음 컨트롤 (없으면 nullptr)
     */
    std::unique_ptr<HwpCtrl> Next() const;

    /**
     * @brief 이전 컨트롤
     * @return 이전 컨트롤 (없으면 nullptr)
     */
    std::unique_ptr<HwpCtrl> Prev() const;

    /**
     * @brief 부모 컨트롤
     * @return 부모 컨트롤 (없으면 nullptr)
     */
    std::unique_ptr<HwpCtrl> Parent() const;

    /**
     * @brief 첫 번째 자식 컨트롤
     * @return 첫 번째 자식 (없으면 nullptr)
     */
    std::unique_ptr<HwpCtrl> FirstChild() const;

    //=========================================================================
    // 속성 접근
    //=========================================================================

    /**
     * @brief 컨트롤 속성 파라미터셋
     * @return HParameterSet COM 객체
     */
    IDispatch* GetProperties() const;

    /**
     * @brief 컨트롤 속성 설정
     * @param pset 파라미터셋
     * @return 성공 여부
     */
    bool SetProperties(IDispatch* pset);

    /**
     * @brief 컨트롤 사용자 데이터
     */
    std::wstring GetUserData() const;

    /**
     * @brief 컨트롤 사용자 데이터 설정
     */
    bool SetUserData(const std::wstring& data);

    /**
     * @brief 인스턴스 ID (고유 식별자)
     */
    int GetInstID() const;

    //=========================================================================
    // 표 전용 속성 (CtrlID == "tbl")
    //=========================================================================

    /**
     * @brief 표인지 확인
     */
    bool IsTable() const;

    /**
     * @brief 표 행 수
     */
    int GetRowCount() const;

    /**
     * @brief 표 열 수
     */
    int GetColCount() const;

    /**
     * @brief 표 셀 가져오기
     * @param row 행 (0-based)
     * @param col 열 (0-based)
     * @return 셀 컨트롤
     */
    std::unique_ptr<HwpCtrl> GetCell(int row, int col) const;

    /**
     * @brief 표 선택
     */
    bool SelectTable();

    /**
     * @brief 셀 주소로 이동
     * @param addr 셀 주소 ("A1", "B2" 등)
     */
    bool GoToCell(const std::wstring& addr);

    //=========================================================================
    // 그림 전용 속성 (CtrlID == "gso" 또는 "pic")
    //=========================================================================

    /**
     * @brief 그림인지 확인
     */
    bool IsPicture() const;

    /**
     * @brief 그림 경로 (BinData에서)
     */
    std::wstring GetPicturePath() const;

    /**
     * @brief 그림 크기 (너비, 높이 HwpUnit)
     */
    std::tuple<HwpUnit, HwpUnit> GetPictureSize() const;

    //=========================================================================
    // 위치/크기
    //=========================================================================

    /**
     * @brief 컨트롤 위치 (X, Y HwpUnit)
     */
    std::tuple<HwpUnit, HwpUnit> GetPosition() const;

    /**
     * @brief 컨트롤 크기 (너비, 높이 HwpUnit)
     */
    std::tuple<HwpUnit, HwpUnit> GetSize() const;

    /**
     * @brief 컨트롤 크기 설정
     * @param width 너비 (HwpUnit)
     * @param height 높이 (HwpUnit)
     * @return 성공 여부
     */
    bool SetSize(HwpUnit width, HwpUnit height);

    //=========================================================================
    // 컨트롤 조작
    //=========================================================================

    /**
     * @brief 컨트롤로 이동
     * @return 성공 여부
     */
    bool MoveTo();

    /**
     * @brief 컨트롤 선택
     * @return 성공 여부
     */
    bool Select();

    /**
     * @brief 컨트롤 삭제
     * @return 성공 여부
     */
    bool Delete();

    /**
     * @brief 컨트롤 복사
     * @return 성공 여부
     */
    bool Copy();

    //=========================================================================
    // COM 인터페이스 직접 접근
    //=========================================================================

    /**
     * @brief 내부 COM 객체 접근
     */
    IDispatch* GetDispatch() const { return m_pCtrl; }

protected:
    /**
     * @brief 속성 가져오기
     */
    VARIANT GetProperty(const std::wstring& name) const;

    /**
     * @brief 속성 설정
     */
    bool SetProperty(const std::wstring& name, const VARIANT& value);

    /**
     * @brief 메서드 호출
     */
    VARIANT InvokeMethod(const std::wstring& name,
                         const std::vector<VARIANT>& args = {}) const;

private:
    IDispatch* m_pCtrl;     // 컨트롤 COM 포인터
    HwpWrapper* m_pHwp;     // 부모 HwpWrapper

    /**
     * @brief CtrlID 문자열을 CtrlType으로 변환
     */
    static CtrlType CtrlIDToType(const std::wstring& id);
};

} // namespace cpyhwpx
