/**
 * @file HwpWrapper.h
 * @brief HWP COM 객체 래퍼 클래스
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 * HWPFrame.HwpObject COM 인터페이스를 래핑
 */

#pragma once

#include "HwpTypes.h"
#include <Windows.h>
#include <comdef.h>
#include <oleidl.h>  // IOleWindow, IOleObject 지원
#include <memory>
#include <functional>
#include <unordered_map>
#include <map>
#include <string>

namespace cpyhwpx {

/**
 * @class DISPIDCache
 * @brief DISPID 캐싱을 위한 클래스 (성능 최적화)
 *
 * COM IDispatch::GetIDsOfNames 호출을 캐싱하여 반복 호출 비용 제거
 */
class DISPIDCache {
public:
    /**
     * @brief DISPID 조회 (캐시에 없으면 GetIDsOfNames 호출 후 캐시)
     * @param pObj IDispatch 객체
     * @param name 메서드/속성 이름
     * @return DISPID 값 (실패 시 DISPID_UNKNOWN)
     */
    DISPID GetOrLoad(IDispatch* pObj, const std::wstring& name);

    /**
     * @brief 캐시 초기화
     */
    void Clear() { m_cache.clear(); }

private:
    std::unordered_map<std::wstring, DISPID> m_cache;
};

// 전방 선언
class HwpCtrl;
class HwpAction;
class HwpParameterSet;
class XHwpDocument;
class XHwpDocuments;

/**
 * @class HwpWrapper
 * @brief HWP COM 객체를 래핑하는 메인 클래스
 *
 * Python의 pyhwpx.Hwp 클래스에 대응
 */
class HwpWrapper {
public:
    //=========================================================================
    // 생성자/소멸자
    //=========================================================================

    /**
     * @brief HwpWrapper 생성자
     * @param visible 창 표시 여부 (기본: true)
     * @param new_instance 새 인스턴스 생성 여부 (기본: true)
     */
    HwpWrapper(bool visible = true, bool new_instance = true);

    /**
     * @brief 소멸자 - COM 리소스 해제
     */
    ~HwpWrapper();

    // 복사/이동 금지
    HwpWrapper(const HwpWrapper&) = delete;
    HwpWrapper& operator=(const HwpWrapper&) = delete;
    HwpWrapper(HwpWrapper&&) = delete;
    HwpWrapper& operator=(HwpWrapper&&) = delete;

    //=========================================================================
    // 초기화 및 종료
    //=========================================================================

    /**
     * @brief HWP COM 객체 초기화
     * @return 성공 여부
     */
    bool Initialize();

    /**
     * @brief 보안 모듈 등록
     * @param module_type 모듈 유형 ("FilePathCheckDLL" 등)
     * @param module_data 모듈 경로
     * @return 성공 여부
     */
    bool RegisterModule(const std::wstring& module_type,
                        const std::wstring& module_data);

    /**
     * @brief HWP 종료
     * @param save 저장 여부
     */
    void Quit(bool save = false);

    /**
     * @brief 초기화 완료 여부
     */
    bool IsInitialized() const { return m_bInitialized; }

    //=========================================================================
    // 파일 I/O
    //=========================================================================

    /**
     * @brief 파일 열기
     * @param filename 파일 경로
     * @param format 파일 형식 (빈 문자열이면 자동 감지)
     * @param arg 추가 인자
     * @return 성공 여부
     */
    bool Open(const std::wstring& filename,
              const std::wstring& format = L"",
              const std::wstring& arg = L"");

    /**
     * @brief 파일 저장
     * @param save_if_dirty 수정된 경우만 저장
     * @return 성공 여부
     */
    bool Save(bool save_if_dirty = true);

    /**
     * @brief 다른 이름으로 저장
     * @param filename 파일 경로
     * @param format 파일 형식
     * @param arg 추가 인자
     * @return 성공 여부
     */
    bool SaveAs(const std::wstring& filename,
                const std::wstring& format = L"HWP",
                const std::wstring& arg = L"");

    /**
     * @brief 문서 초기화 (hwp.Clear)
     * @param option 초기화 옵션 (기본: 1)
     */
    void Clear(int option = 1);

    /**
     * @brief 문서 초기화 (pyhwpx 방식: XHwpDocuments.Active_XHwpDocument.Clear)
     * @param option 초기화 옵션 (1=hwpDiscard, 2=hwpSaveIfDirty, 3=hwpSave)
     * @return 성공 여부
     */
    bool ClearDocument(int option = 1);

    /**
     * @brief 문서 닫기
     * @param is_dirty dirty 상태 설정
     * @return 성공 여부
     */
    bool Close(bool is_dirty = false);

    /**
     * @brief 현재 위치에 파일 삽입
     * @param filename 삽입할 파일 경로
     * @param keep_section 구간 정보 유지 (1=유지, 0=무시)
     * @param keep_charshape 문자 모양 유지
     * @param keep_parashape 문단 모양 유지
     * @param keep_style 스타일 유지
     * @param move_doc_end 삽입 후 문서 끝으로 이동
     * @return 성공 여부
     */
    bool InsertFile(const std::wstring& filename,
                    int keep_section = 1,
                    int keep_charshape = 1,
                    int keep_parashape = 1,
                    int keep_style = 1,
                    bool move_doc_end = false);

    //=========================================================================
    // 텍스트 편집
    //=========================================================================

    /**
     * @brief 텍스트 삽입
     * @param text 삽입할 텍스트
     * @return 성공 여부
     */
    bool InsertText(const std::wstring& text);

    /**
     * @brief 현재 위치의 텍스트 가져오기
     * @return (상태코드, 텍스트) 튜플
     */
    std::tuple<int, std::wstring> GetText();

    /**
     * @brief 선택된 텍스트 가져오기
     * @param keep_select 선택 유지 여부
     * @return 선택된 텍스트
     */
    std::wstring GetSelectedText(bool keep_select = false);

    //=========================================================================
    // 위치 관리
    //=========================================================================

    /**
     * @brief 현재 위치 가져오기
     * @return (list, para, pos) 튜플
     */
    HwpPos GetPos();

    /**
     * @brief 위치 설정
     * @param list 리스트 ID
     * @param para 문단 번호
     * @param pos 문단 내 위치
     * @return 성공 여부
     */
    bool SetPos(int list, int para, int pos);

    /**
     * @brief 위치 이동
     * @param move_id 이동 ID (MoveID 열거형)
     * @param para 문단 번호 (기본: 0)
     * @param pos 위치 (기본: 0)
     * @return 성공 여부
     */
    bool MovePos(int move_id, int para = 0, int pos = 0);

    //=========================================================================
    // 창/UI 관리
    //=========================================================================

    /**
     * @brief 창 표시/숨김 설정
     */
    void SetVisible(bool visible);

    /**
     * @brief HWP 창 핸들(HWND) 획득
     * @return HWND (실패 시 NULL)
     */
    HWND GetHwnd();

    /**
     * @brief OLE 객체 활성화 (OLEIVERB_SHOW)
     * @return 성공 여부
     */
    bool ActivateOleObject();

    /**
     * @brief 창을 포그라운드로 표시
     * @param hwnd 창 핸들
     */
    void ShowWindowWithForeground(HWND hwnd);

    /**
     * @brief 창 최대화
     */
    void MaximizeWindow();

    /**
     * @brief 창 최소화
     */
    void MinimizeWindow();

    /**
     * @brief 뷰 상태 설정
     */
    bool SetViewState(int flag);

    /**
     * @brief 뷰 상태 가져오기
     */
    int GetViewState();

    //=========================================================================
    // 메시지 박스
    //=========================================================================

    /**
     * @brief 메시지 박스 표시
     */
    int MsgBox(const std::wstring& message, int flag = 0);

    /**
     * @brief 메시지 박스 모드 가져오기
     */
    int GetMessageBoxMode();

    /**
     * @brief 메시지 박스 모드 설정
     */
    int SetMessageBoxMode(int mode);

    //=========================================================================
    // 문서 상태
    //=========================================================================

    /**
     * @brief 빈 문서 여부
     */
    bool IsEmpty();

    /**
     * @brief 수정 여부
     */
    bool IsModified();

    /**
     * @brief 셀 안 여부
     */
    bool IsCell();

    //=========================================================================
    // 찾기/바꾸기
    //=========================================================================

    /**
     * @brief 텍스트 찾기
     * @param text 찾을 텍스트
     * @param forward 앞으로 찾기
     * @param match_case 대소문자 구분
     * @param regex 정규식 사용
     * @param replace_mode 바꾸기 모드
     * @return 성공 여부
     */
    bool Find(const std::wstring& text,
              bool forward = true,
              bool match_case = false,
              bool regex = false,
              bool replace_mode = false);

    /**
     * @brief 텍스트 바꾸기
     */
    bool Replace(const std::wstring& find_text,
                 const std::wstring& replace_text,
                 bool forward = true,
                 bool match_case = false,
                 bool regex = false);

    /**
     * @brief 모두 바꾸기
     */
    int ReplaceAll(const std::wstring& find_text,
                   const std::wstring& replace_text,
                   bool match_case = false,
                   bool regex = false);

    //=========================================================================
    // 필드 작업
    //=========================================================================

    /**
     * @brief 누름틀 필드 생성
     * @param name 필드 이름
     * @param direction 안내문/지시문
     * @param memo 설명/도움말
     * @return 성공 여부
     */
    bool CreateField(const std::wstring& name,
                     const std::wstring& direction = L"",
                     const std::wstring& memo = L"");

    /**
     * @brief 필드 목록 조회
     * @param number 0=plain, 1=numbered, 2=count
     * @param option 필드 유형 필터 (0=all, 1=cell, 2=clickhere, 4=selection)
     * @return 필드 목록 (0x02로 구분)
     */
    std::wstring GetFieldList(int number = 1, int option = 0);

    /**
     * @brief 필드 텍스트 조회
     * @param field 필드 이름 ({{n}} 인덱스 지원)
     * @param idx 동일 이름 필드의 인덱스
     * @return 필드 텍스트
     */
    std::wstring GetFieldText(const std::wstring& field, int idx = 0);

    /**
     * @brief 필드 텍스트 설정
     * @param field 필드 이름 (0x02로 구분된 복수 필드 지원)
     * @param text 설정할 텍스트 (0x02로 구분된 복수 텍스트 지원)
     * @return 성공 여부
     */
    bool PutFieldText(const std::wstring& field, const std::wstring& text);

    /**
     * @brief 필드 존재 확인
     * @param field 필드 이름
     * @return 존재 여부
     */
    bool FieldExist(const std::wstring& field);

    /**
     * @brief 필드로 캐럿 이동
     * @param field 필드 이름 ({{n}} 인덱스 지원)
     * @param idx 동일 이름 필드의 인덱스
     * @param text true=텍스트 영역으로 이동
     * @param start true=시작 위치, false=끝 위치
     * @param select true=필드 내용 선택
     * @return 성공 여부
     */
    bool MoveToField(const std::wstring& field, int idx = 0,
                     bool text = true, bool start = true, bool select = false);

    /**
     * @brief 필드 이름 변경
     * @param oldname 기존 이름 (0x02로 구분된 복수 필드 지원)
     * @param newname 새 이름 (0x02로 구분된 복수 이름 지원)
     * @return 성공 여부
     */
    bool RenameField(const std::wstring& oldname, const std::wstring& newname);

    /**
     * @brief 현재 위치 필드 이름 조회
     * @param option 0=all, 1=cell, 2=clickhere
     * @return 필드 이름
     */
    std::wstring GetCurFieldName(int option = 0);

    /**
     * @brief 현재 셀 필드 이름 설정 (표 안에서만)
     * @param field 필드 이름
     * @param direction 안내문 (clickhere 전용)
     * @param memo 메모 (clickhere 전용)
     * @param option 0=all, 1=cell, 2=clickhere
     * @return 성공 여부
     */
    bool SetCurFieldName(const std::wstring& field,
                         const std::wstring& direction = L"",
                         const std::wstring& memo = L"",
                         int option = 0);

    /**
     * @brief 필드 뷰 옵션 설정
     * @param option 옵션 값
     * @return 이전 옵션 값
     */
    int SetFieldViewOption(int option);

    /**
     * @brief 모든 필드 삭제
     * @return 성공 여부
     */
    bool DeleteAllFields();

    /**
     * @brief 이름으로 필드 삭제
     * @param field_name 필드 이름
     * @param idx 인덱스 (-1이면 모든 동일 이름 필드 삭제)
     * @return 성공 여부
     */
    bool DeleteFieldByName(const std::wstring& field_name, int idx = -1);

    /**
     * @brief 필드 목록을 맵으로 변환
     * @return 필드명:텍스트 맵
     */
    std::map<std::wstring, std::wstring> FieldsToMap();

    //=========================================================================
    // 테이블 작업
    //=========================================================================

    /**
     * @brief 테이블 생성
     * @param rows 행 수 (기본: 2)
     * @param cols 열 수 (기본: 2)
     * @param treat_as_char 글자처럼 취급 (기본: true)
     * @param width_type 너비 타입 (0=단에맞춤, 1=문단에맞춤, 2=임의값)
     * @param height_type 높이 타입 (0=자동, 1=임의값)
     * @param header 첫 행을 제목행으로 (기본: false)
     * @return 성공 여부
     */
    bool CreateTable(int rows = 2, int cols = 2,
                     bool treat_as_char = true,
                     int width_type = 0,
                     int height_type = 0,
                     bool header = false);

    /**
     * @brief n번째 테이블로 이동
     * @param n 테이블 인덱스 (0부터, 음수는 뒤에서부터)
     * @param select_cell 셀 선택 여부
     * @return 성공 여부
     */
    bool GetIntoNthTable(int n = 0, bool select_cell = false);

    /**
     * @brief 테이블 행 개수 조회
     * @return 행 개수 (-1: 테이블 아님)
     */
    int GetTableRowCount();

    /**
     * @brief 테이블 열 개수 조회
     * @return 열 개수 (-1: 테이블 아님)
     */
    int GetTableColCount();

    /**
     * @brief 왼쪽 셀로 이동
     */
    bool TableLeftCell();

    /**
     * @brief 오른쪽 셀로 이동
     */
    bool TableRightCell();

    /**
     * @brief 위쪽 셀로 이동
     */
    bool TableUpperCell();

    /**
     * @brief 아래쪽 셀로 이동
     */
    bool TableLowerCell();

    /**
     * @brief 오른쪽 셀로 이동 (행 끝이면 다음 행으로)
     */
    bool TableRightCellAppend();

    //=========================================================================
    // 이미지 삽입
    //=========================================================================

    /**
     * @brief 이미지 삽입
     * @param path 이미지 파일 경로 (절대 경로 권장)
     * @param embedded 문서 내 포함 여부 (기본: true)
     * @param sizeoption 크기 옵션 (0=원본, 1=지정, 2=셀맞춤, 3=셀맞춤+종횡비)
     * @param reverse 반전 여부 (기본: false)
     * @param watermark 워터마크 효과 (기본: false)
     * @param effect 이미지 효과 (0=원본, 1=그레이스케일, 2=흑백)
     * @param width 너비 (mm, sizeoption=1일 때 사용)
     * @param height 높이 (mm, sizeoption=1일 때 사용)
     * @return 성공 여부
     */
    bool InsertPicture(const std::wstring& path,
                       bool embedded = true,
                       int sizeoption = 0,
                       bool reverse = false,
                       bool watermark = false,
                       int effect = 0,
                       int width = 0,
                       int height = 0);

    //=========================================================================
    // HAction 관련
    //=========================================================================

    /**
     * @brief HAction.Run() 실행
     * @param action_name 액션 이름
     * @return 성공 여부
     */
    bool RunAction(const std::wstring& action_name);

    /**
     * @brief 액션 생성
     */
    IDispatch* CreateAction(const std::wstring& action_id);

    /**
     * @brief 파라미터셋 생성
     */
    IDispatch* CreateSet(const std::wstring& set_id);

    /**
     * @brief 현재 위치의 컨트롤을 찾아 선택
     * @return 성공 여부 (컨트롤을 찾으면 true)
     */
    bool FindCtrl();

    /**
     * @brief ParameterSet으로 위치 설정 (SetPosBySet API)
     * @param pDispVal GetAnchorPos 등에서 반환된 ParameterSet
     * @return 성공 여부
     */
    bool SetPosBySet(IDispatch* pDispVal);

    //=========================================================================
    // 속성 접근
    //=========================================================================

    /**
     * @brief HWP 버전 문자열
     */
    std::wstring GetVersion();

    /**
     * @brief 빌드 번호
     */
    std::wstring GetBuildNumber();

    /**
     * @brief 현재 페이지 번호 (0-based)
     */
    int GetCurrentPage();

    /**
     * @brief 현재 인쇄 페이지 번호 (1-based)
     */
    int GetCurrentPrintPage();

    /**
     * @brief 총 페이지 수
     */
    int GetPageCount();

    /**
     * @brief 수정 여부 속성
     */
    bool GetEditMode();
    void SetEditMode(bool mode);

    //=========================================================================
    // COM 인터페이스 직접 접근 (고급)
    //=========================================================================

    /**
     * @brief 내부 COM 객체 접근
     */
    IDispatch* GetHwpObject() const { return m_pHwp; }

    /**
     * @brief HParameterSet 접근
     */
    IDispatch* GetHParameterSet();

    /**
     * @brief HAction 접근
     */
    IDispatch* GetHAction();

protected:
    //=========================================================================
    // COM 헬퍼 메서드
    //=========================================================================

    /**
     * @brief Active_XHwpWindow에서 WindowHandle 획득
     * @param pActiveWindow Active_XHwpWindow IDispatch 포인터
     * @return HWND (실패 시 NULL)
     */
    HWND GetActiveWindowHandle(IDispatch* pActiveWindow);

    /**
     * @brief 속성 가져오기 (VARIANT)
     */
    VARIANT GetProperty(const std::wstring& name);

    /**
     * @brief 속성 설정
     */
    bool SetProperty(const std::wstring& name, const VARIANT& value);

    /**
     * @brief 메서드 호출 (반환값 없음)
     */
    bool InvokeMethod(const std::wstring& name);

    /**
     * @brief 메서드 호출 (VARIANT 반환)
     */
    VARIANT InvokeMethodWithResult(const std::wstring& name,
                                   const std::vector<VARIANT>& args = {});

private:
    IDispatch* m_pHwp;              // HWPFrame.HwpObject COM 포인터
    IDispatch* m_pHParameterSet;    // HParameterSet 캐시
    IDispatch* m_pHAction;          // HAction 캐시

    bool m_bInitialized;            // 초기화 완료 여부
    bool m_bVisible;                // 창 표시 여부
    bool m_bNewInstance;            // 새 인스턴스 여부

    DISPIDCache m_dispidCache;      // DISPID 캐시 (성능 최적화)

    /**
     * @brief COM 초기화
     */
    bool InitializeCOM();

    /**
     * @brief HWP COM 객체 생성
     */
    bool CreateHwpObject();

    /**
     * @brief 리소스 해제
     */
    void Release();
};

} // namespace cpyhwpx
