/**
 * @file XHwpDocument.h
 * @brief HWP 개별 문서 래퍼 클래스
 *
 * pyhwpx -> cpyhwpx 포팅 프로젝트
 * XHwpDocument COM 인터페이스를 래핑
 */

#pragma once

#include <Windows.h>
#include <comdef.h>
#include <string>

namespace cpyhwpx {

// 전방 선언
class HwpWrapper;

/**
 * @class XHwpDocument
 * @brief HWP 개별 문서 래퍼 클래스
 *
 * HWP 문서 하나를 래핑하여 문서별 조작 기능 제공
 * Python의 pyhwpx.XHwpDocument 클래스에 대응
 */
class XHwpDocument {
public:
    //=========================================================================
    // 생성자/소멸자
    //=========================================================================

    /**
     * @brief XHwpDocument 생성자
     * @param pDocument COM 문서 객체
     */
    explicit XHwpDocument(IDispatch* pDocument);

    /**
     * @brief 소멸자
     */
    ~XHwpDocument();

    // 복사 금지, 이동 허용
    XHwpDocument(const XHwpDocument&) = delete;
    XHwpDocument& operator=(const XHwpDocument&) = delete;
    XHwpDocument(XHwpDocument&& other) noexcept;
    XHwpDocument& operator=(XHwpDocument&& other) noexcept;

    //=========================================================================
    // 읽기 전용 속성
    //=========================================================================

    /**
     * @brief 문서 고유 ID
     */
    int GetDocumentID() const;

    /**
     * @brief 문서 전체 경로 (저장 안됐으면 빈 문자열)
     */
    std::wstring GetFullName() const;

    /**
     * @brief 문서 디렉토리 경로
     */
    std::wstring GetPath() const;

    /**
     * @brief 문서 포맷 ("HWP", "HWPX" 등)
     */
    std::wstring GetFormat() const;

    /**
     * @brief 수정 여부 (true=수정됨)
     */
    bool IsModified() const;

    /**
     * @brief 편집 모드 (1=일반, 2=개요 등)
     */
    int GetEditMode() const;

    /**
     * @brief 유효한 문서인지 확인
     */
    bool IsValid() const { return m_pDocument != nullptr; }

    //=========================================================================
    // 문서 활성화
    //=========================================================================

    /**
     * @brief 이 문서를 활성 문서로 설정
     * @return 성공 여부
     */
    bool SetActiveDocument();

    //=========================================================================
    // 파일 작업
    //=========================================================================

    /**
     * @brief 파일 열기
     * @param filename 파일 경로
     * @param format 포맷 ("HWP", "HWPX", "HTML" 등)
     * @param arg 추가 인자
     * @return 성공 여부
     */
    bool Open(const std::wstring& filename,
              const std::wstring& format = L"",
              const std::wstring& arg = L"");

    /**
     * @brief 문서 저장
     * @param saveIfDirty 수정된 경우에만 저장
     * @return 성공 여부
     */
    bool Save(bool saveIfDirty = false);

    /**
     * @brief 다른 이름으로 저장
     * @param path 저장 경로
     * @param format 포맷
     * @param arg 추가 인자
     * @return 성공 여부
     */
    bool SaveAs(const std::wstring& path,
                const std::wstring& format = L"",
                const std::wstring& arg = L"");

    /**
     * @brief 문서 닫기
     * @param isDirty true=저장 안함
     */
    void Close(bool isDirty = false);

    /**
     * @brief 문서 내용 지우기
     * @param option true=모든 내용 삭제
     */
    void Clear(bool option = false);

    //=========================================================================
    // 편집 작업
    //=========================================================================

    /**
     * @brief 실행 취소
     * @param count 취소할 횟수
     * @return 성공 여부
     */
    bool Undo(int count = 1);

    /**
     * @brief 다시 실행
     * @param count 다시 실행할 횟수
     * @return 성공 여부
     */
    bool Redo(int count = 1);

    //=========================================================================
    // COM 인터페이스 직접 접근
    //=========================================================================

    /**
     * @brief 내부 COM 객체 접근
     */
    IDispatch* GetDispatch() const { return m_pDocument; }

protected:
    /**
     * @brief 문자열 속성 가져오기
     */
    std::wstring GetStringProperty(const std::wstring& name) const;

    /**
     * @brief 정수 속성 가져오기
     */
    int GetIntProperty(const std::wstring& name) const;

    /**
     * @brief 불린 속성 가져오기
     */
    bool GetBoolProperty(const std::wstring& name) const;

    /**
     * @brief 메서드 호출 (인자 없음)
     */
    bool InvokeMethod(const std::wstring& name);

    /**
     * @brief 메서드 호출 (정수 인자 1개)
     */
    bool InvokeMethodInt(const std::wstring& name, int arg);

    /**
     * @brief 메서드 호출 (불린 인자 1개)
     */
    bool InvokeMethodBool(const std::wstring& name, bool arg);

private:
    IDispatch* m_pDocument;  // 문서 COM 포인터
};

} // namespace cpyhwpx
