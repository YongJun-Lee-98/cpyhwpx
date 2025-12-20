/**
 * @file XHwpDocuments.h
 * @brief HWP 문서 컬렉션 래퍼 클래스
 *
 * pyhwpx -> cpyhwpx 포팅 프로젝트
 * XHwpDocuments COM 인터페이스를 래핑
 */

#pragma once

#include "XHwpDocument.h"
#include <Windows.h>
#include <comdef.h>
#include <memory>
#include <string>

namespace cpyhwpx {

/**
 * @class XHwpDocuments
 * @brief HWP 문서 컬렉션 래퍼 클래스
 *
 * 열린 모든 HWP 문서를 관리하는 컬렉션
 * Python의 pyhwpx.XHwpDocuments 클래스에 대응
 */
class XHwpDocuments {
public:
    //=========================================================================
    // 생성자/소멸자
    //=========================================================================

    /**
     * @brief XHwpDocuments 생성자
     * @param pDocuments COM 문서 컬렉션 객체
     */
    explicit XHwpDocuments(IDispatch* pDocuments);

    /**
     * @brief 소멸자
     */
    ~XHwpDocuments();

    // 복사 금지, 이동 허용
    XHwpDocuments(const XHwpDocuments&) = delete;
    XHwpDocuments& operator=(const XHwpDocuments&) = delete;
    XHwpDocuments(XHwpDocuments&& other) noexcept;
    XHwpDocuments& operator=(XHwpDocuments&& other) noexcept;

    //=========================================================================
    // 컬렉션 속성
    //=========================================================================

    /**
     * @brief 열린 문서 개수
     */
    int GetCount() const;

    /**
     * @brief 현재 활성 문서
     * @return 활성 문서 (없으면 nullptr)
     */
    std::unique_ptr<XHwpDocument> GetActiveDocument() const;

    /**
     * @brief 유효한 컬렉션인지 확인
     */
    bool IsValid() const { return m_pDocuments != nullptr; }

    //=========================================================================
    // 문서 접근
    //=========================================================================

    /**
     * @brief 인덱스로 문서 접근
     * @param index 0부터 시작하는 인덱스 (음수는 끝에서부터)
     * @return 해당 문서 (범위 초과 시 예외)
     */
    std::unique_ptr<XHwpDocument> Item(int index) const;

    /**
     * @brief 문서 ID로 검색
     * @param docID 문서 고유 ID
     * @return 해당 문서 (없으면 nullptr)
     */
    std::unique_ptr<XHwpDocument> FindItem(int docID) const;

    //=========================================================================
    // 문서 관리
    //=========================================================================

    /**
     * @brief 새 문서 추가
     * @param isTab true=새 탭, false=새 창
     * @return 생성된 문서
     */
    std::unique_ptr<XHwpDocument> Add(bool isTab = false);

    /**
     * @brief 활성 문서 닫기
     * @param isDirty true=저장 안함
     */
    void Close(bool isDirty = false);

    //=========================================================================
    // Python 컬렉션 프로토콜
    //=========================================================================

    /**
     * @brief __len__ - 문서 개수
     */
    int Len() const { return GetCount(); }

    /**
     * @brief __getitem__ - 인덱스 접근
     */
    std::unique_ptr<XHwpDocument> GetItem(int index) const { return Item(index); }

    //=========================================================================
    // COM 인터페이스 직접 접근
    //=========================================================================

    /**
     * @brief 내부 COM 객체 접근
     */
    IDispatch* GetDispatch() const { return m_pDocuments; }

private:
    IDispatch* m_pDocuments;  // 문서 컬렉션 COM 포인터
};

} // namespace cpyhwpx
