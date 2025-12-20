/**
 * @file HwpParameter.h
 * @brief HParameterSet COM 객체 래퍼 클래스
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 * HWP의 HParameterSet 인터페이스 래핑
 */

#pragma once

#include "HwpTypes.h"
#include <Windows.h>
#include <comdef.h>
#include <string>
#include <vector>
#include <memory>

namespace cpyhwpx {

// 전방 선언
class HwpWrapper;

/**
 * @class HwpParameterSet
 * @brief HParameterSet COM 객체 래퍼
 *
 * 액션 실행에 필요한 파라미터를 설정하는 클래스
 */
class HwpParameterSet {
public:
    //=========================================================================
    // 생성자/소멸자
    //=========================================================================

    /**
     * @brief HwpParameterSet 생성자
     * @param pset COM 파라미터셋 객체
     */
    HwpParameterSet(IDispatch* pset);

    /**
     * @brief 소멸자
     */
    ~HwpParameterSet();

    // 복사 금지
    HwpParameterSet(const HwpParameterSet&) = delete;
    HwpParameterSet& operator=(const HwpParameterSet&) = delete;

    //=========================================================================
    // 파라미터 설정 - 기본 타입
    //=========================================================================

    /**
     * @brief 정수 파라미터 설정
     */
    bool SetItem(const std::wstring& name, int value);

    /**
     * @brief 실수 파라미터 설정
     */
    bool SetItem(const std::wstring& name, double value);

    /**
     * @brief 문자열 파라미터 설정
     */
    bool SetItem(const std::wstring& name, const std::wstring& value);

    /**
     * @brief 불리언 파라미터 설정
     */
    bool SetItem(const std::wstring& name, bool value);

    /**
     * @brief VARIANT 파라미터 설정
     */
    bool SetItem(const std::wstring& name, const VARIANT& value);

    //=========================================================================
    // 파라미터 가져오기 - 기본 타입
    //=========================================================================

    /**
     * @brief 정수 파라미터 가져오기
     */
    int GetItemInt(const std::wstring& name, int default_value = 0);

    /**
     * @brief 실수 파라미터 가져오기
     */
    double GetItemDouble(const std::wstring& name, double default_value = 0.0);

    /**
     * @brief 문자열 파라미터 가져오기
     */
    std::wstring GetItemString(const std::wstring& name, const std::wstring& default_value = L"");

    /**
     * @brief 불리언 파라미터 가져오기
     */
    bool GetItemBool(const std::wstring& name, bool default_value = false);

    /**
     * @brief VARIANT 파라미터 가져오기
     */
    VARIANT GetItem(const std::wstring& name);

    //=========================================================================
    // 서브셋 관련
    //=========================================================================

    /**
     * @brief 서브셋 생성
     * @param set_id 서브셋 ID
     * @return 서브셋 COM 객체
     */
    IDispatch* CreateItemSet(const std::wstring& item_id, const std::wstring& set_id);

    /**
     * @brief 서브셋 가져오기
     */
    IDispatch* GetItemSet(const std::wstring& item_id);

    /**
     * @brief 배열 항목 가져오기
     */
    IDispatch* GetItemByIndex(int index);

    //=========================================================================
    // 유틸리티
    //=========================================================================

    /**
     * @brief 파라미터셋 ID
     */
    std::wstring GetSetID() const;

    /**
     * @brief 유효한 파라미터셋인지
     */
    bool IsValid() const { return m_pSet != nullptr; }

    /**
     * @brief 항목 수
     */
    int GetCount();

    /**
     * @brief 파라미터셋 초기화
     */
    void Clear();

    //=========================================================================
    // COM 인터페이스 직접 접근
    //=========================================================================

    /**
     * @brief 내부 COM 객체 접근
     */
    IDispatch* GetDispatch() const { return m_pSet; }

protected:
    /**
     * @brief SetItem 내부 구현
     */
    bool SetItemInternal(const std::wstring& name, VARIANT& value);

    /**
     * @brief GetItem 내부 구현
     */
    VARIANT GetItemInternal(const std::wstring& name);

private:
    IDispatch* m_pSet;  // HParameterSet COM 포인터
};

/**
 * @class HwpParameterSetHelper
 * @brief 자주 사용되는 파라미터셋 생성 헬퍼
 */
class HwpParameterSetHelper {
public:
    //=========================================================================
    // 일반 파라미터셋
    //=========================================================================

    /**
     * @brief 찾기/바꾸기 파라미터셋 생성
     */
    static std::unique_ptr<HwpParameterSet> CreateFindReplace(
        HwpWrapper* hwp,
        const std::wstring& find_text,
        const std::wstring& replace_text = L"",
        bool match_case = false,
        bool whole_word = false,
        bool regex = false);

    /**
     * @brief 표 생성 파라미터셋 생성
     */
    static std::unique_ptr<HwpParameterSet> CreateTable(
        HwpWrapper* hwp,
        int rows,
        int cols,
        HwpUnit width = 0,
        HwpUnit row_height = 0);

    /**
     * @brief 글자 모양 파라미터셋 생성
     */
    static std::unique_ptr<HwpParameterSet> CreateCharShape(
        HwpWrapper* hwp,
        const std::wstring& face_name = L"",
        int height = 0,
        COLORREF color = 0);

    /**
     * @brief 문단 모양 파라미터셋 생성
     */
    static std::unique_ptr<HwpParameterSet> CreateParaShape(
        HwpWrapper* hwp,
        HAlign align = HAlign::Justify,
        int line_spacing = 160);

    /**
     * @brief 그림 삽입 파라미터셋 생성
     */
    static std::unique_ptr<HwpParameterSet> CreateInsertPicture(
        HwpWrapper* hwp,
        const std::wstring& path,
        bool embedded = true,
        int size_option = 0);

    //=========================================================================
    // 셀/표 파라미터셋
    //=========================================================================

    /**
     * @brief 셀 모양 파라미터셋 생성
     */
    static std::unique_ptr<HwpParameterSet> CreateCellShape(
        HwpWrapper* hwp,
        COLORREF bg_color = 0xFFFFFF,
        VAlign valign = VAlign::Center);

    /**
     * @brief 테두리 파라미터셋 생성
     */
    static std::unique_ptr<HwpParameterSet> CreateBorderLine(
        HwpWrapper* hwp,
        LineStyle style = LineStyle::Solid,
        HwpUnit width = 1,
        COLORREF color = 0);
};

} // namespace cpyhwpx
