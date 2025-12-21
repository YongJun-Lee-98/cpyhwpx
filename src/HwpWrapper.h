/**
 * @file HwpWrapper.h
 * @brief HWP COM 객체 래퍼 클래스
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 * HWPFrame.HwpObject COM 인터페이스를 래핑
 */

#pragma once

#include "HwpTypes.h"
#include "XHwpDocument.h"
#include "XHwpDocuments.h"
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
     * @param register_module 보안 모듈 자동 등록 여부 (기본: true, pyhwpx 호환)
     */
    HwpWrapper(bool visible = true,
               bool new_instance = true,
               bool register_module = true);

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
     * @brief 보안 모듈 등록 (COM API 직접 호출)
     * @param module_type 모듈 유형 ("FilePathCheckDLL" 등)
     * @param module_data 모듈 경로
     * @return 성공 여부
     */
    bool RegisterModule(const std::wstring& module_type,
                        const std::wstring& module_data);

    /**
     * @brief 보안 모듈 레지스트리 등록 여부 확인
     * @param key_name 모듈 이름 (기본: "FilePathCheckerModule")
     * @return 등록되어 있으면 true
     */
    static bool CheckRegistryKey(const std::wstring& key_name = L"FilePathCheckerModule");

    /**
     * @brief 보안 모듈 DLL을 레지스트리에 등록
     * @param dll_path DLL 파일 경로 (빈 문자열이면 자동 감지)
     * @param key_name 등록할 키 이름
     * @return 성공 여부
     */
    static bool RegisterToRegistry(const std::wstring& dll_path = L"",
                                    const std::wstring& key_name = L"FilePathCheckerModule");

    /**
     * @brief DLL 파일 경로 자동 감지
     * @return DLL 파일 전체 경로 (없으면 빈 문자열)
     */
    static std::wstring FindDllPath();

    /**
     * @brief 보안 모듈 자동 등록 (check + regedit + RegisterModule)
     * @param module_type 모듈 타입 (기본: "FilePathCheckDLL")
     * @param module_data 모듈 데이터 (기본: "FilePathCheckerModule")
     * @return 성공 여부
     */
    bool AutoRegisterModule(const std::wstring& module_type = L"FilePathCheckDLL",
                            const std::wstring& module_data = L"FilePathCheckerModule");

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

    /**
     * @brief 문서 텍스트 추출 (GetTextFile API)
     * @param format 형식 ("HWP", "HWPML2X", "HTML", "UNICODE", "TEXT")
     * @param option 옵션 ("saveblock:true"=선택 블록만, ""=전체)
     * @return 추출된 텍스트 (형식에 따라 다름)
     *
     * format 상세:
     * - "HWP": BASE64 인코딩, 모든 정보 유지 (가장 효율적)
     * - "HWPML2X": XML 형식, 모든 정보 유지
     * - "HTML": HTML 형식, 표 번호 유지
     * - "UNICODE": 유니코드 텍스트 (서식정보 없음)
     * - "TEXT": 일반 텍스트 (특수문자 손실)
     */
    std::wstring GetTextFile(const std::wstring& format = L"UNICODE",
                             const std::wstring& option = L"");

    /**
     * @brief 텍스트 데이터 삽입 (SetTextFile API)
     * @param data GetTextFile로 추출한 텍스트 데이터
     * @param format 형식 ("HWP", "HWPML2X", "HTML", "UNICODE", "TEXT")
     * @param option 옵션 ("insertfile"=현재 커서에 삽입)
     * @return 성공 1, 실패 0
     */
    int SetTextFile(const std::wstring& data,
                    const std::wstring& format = L"HWPML2X",
                    const std::wstring& option = L"insertfile");

    /**
     * @brief PDF 파일 열기
     * pyhwpx의 open_pdf()에 대응
     * @param pdfPath PDF 파일 경로
     * @param thisWindow 현재 창에 열기 (1=현재창, 0=새창)
     * @return 성공 여부
     */
    bool OpenPdf(const std::wstring& pdfPath, int thisWindow = 1);

    /**
     * @brief 선택 블록을 파일로 저장
     * pyhwpx의 save_block_as()에 대응
     * @param path 저장 경로
     * @param format 파일 형식 (기본: "HWP")
     * @param attributes 저장 속성 (기본: 1)
     * @return 성공 여부
     */
    bool SaveBlockAs(const std::wstring& path,
                     const std::wstring& format = L"HWP",
                     int attributes = 1);

    /**
     * @brief 파일 정보 조회
     * pyhwpx의 get_file_info()에 대응
     * @param filename 파일 경로
     * @return 파일 정보 맵 (Format, VersionStr, VersionNum, Encrypted)
     */
    std::map<std::wstring, std::wstring> GetFileInfo(const std::wstring& filename);

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

    /**
     * @brief 텍스트 스캔 초기화
     * pyhwpx의 init_scan()에 대응
     * @param option 검색 대상 (0x07: maskNormal|maskChar|maskInline|maskCtrl)
     * @param range 검색 범위 (0x77: 문서 전체, 0xff: 선택 블록)
     * @param spara 시작 문단 번호
     * @param spos 시작 위치
     * @param epara 끝 문단 번호 (-1: 끝까지)
     * @param epos 끝 위치 (-1: 끝까지)
     * @return 성공 여부
     */
    bool InitScan(int option = 0x07, int range = 0x77,
                  int spara = 0, int spos = 0, int epara = -1, int epos = -1);

    /**
     * @brief 텍스트 스캔 해제
     * pyhwpx의 release_scan()에 대응
     */
    void ReleaseScan();

    /**
     * @brief 텍스트 선택
     * pyhwpx의 select_text()에 대응
     * @param spara 시작 문단 번호
     * @param spos 시작 위치
     * @param epara 끝 문단 번호
     * @param epos 끝 위치 (-1: 문단 끝)
     * @param slist 시작 리스트 ID (기본: 0)
     * @return 성공 여부
     */
    bool SelectText(int spara, int spos, int epara, int epos, int slist = 0);

    /**
     * @brief GetPos 튜플로 텍스트 선택
     * pyhwpx의 select_text_by_get_pos()에 대응
     * @param s_pos 시작 위치 (list, para, pos)
     * @param e_pos 끝 위치 (list, para, pos)
     * @return 성공 여부
     */
    bool SelectTextByGetPos(const HwpPos& s_pos, const HwpPos& e_pos);

    /**
     * @brief ParameterSet으로 위치 정보 조회 (내부용)
     * pyhwpx의 get_pos_by_set()에 대응
     * @return ParameterSet (List, Para, Pos 포함)
     */
    IDispatch* GetPosBySet();

    /**
     * @brief ParameterSet으로 위치 설정 (내부용)
     * pyhwpx의 set_pos_by_set()에 대응
     * @param pDispVal GetPosBySet()에서 반환된 ParameterSet
     * @return 성공 여부
     */
    bool SetPosBySet(IDispatch* pDispVal);

    /**
     * @brief Python용 위치 정보 조회
     * IDispatch*를 캐시에 저장하고 인덱스 반환
     * @return 캐시 인덱스 (-1: 실패)
     */
    int GetPosBySetPy();

    /**
     * @brief Python용 위치 설정
     * 캐시에서 인덱스로 IDispatch*를 꺼내서 설정
     * @param idx GetPosBySetPy()에서 반환된 인덱스
     * @return 성공 여부
     */
    bool SetPosBySetPy(int idx);

    /**
     * @brief 위치 캐시 정리
     * COM 객체 Release 및 캐시 초기화
     */
    void ClearPosCache();

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
    // 유틸리티
    //=========================================================================

    /**
     * @brief 키 인디케이터 (상태 표시줄 정보)
     * @return (성공여부, 구역수, 구역번호, 인쇄페이지, 단번호, 줄번호, 위치, 삽입/수정, 컨트롤명)
     */
    std::tuple<int, int, int, int, int, int, int, int, std::wstring> KeyIndicator();

    /**
     * @brief 페이지로 이동
     * @param pageIndex 페이지 번호 (1-based)
     * @return (현재 인쇄 페이지, 현재 페이지)
     */
    std::pair<int, int> GotoPage(int pageIndex);

    /**
     * @brief 밀리미터를 HWP 단위로 변환
     * @param mili 밀리미터 값
     * @return HWP 단위 값
     */
    int MiliToHwpUnit(double mili);

    /**
     * @brief HWP 단위를 밀리미터로 변환
     * @param hwpUnit HWP 단위 값
     * @return 밀리미터 값
     */
    static double HwpUnitToMili(int hwpUnit);

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

    /**
     * @brief 아래 방향으로 텍스트 찾기
     * @param src 찾을 텍스트
     * @param regex 정규식 사용 여부
     * @return 찾으면 true, 없으면 false
     */
    bool FindForward(const std::wstring& src, bool regex = false);

    /**
     * @brief 위 방향으로 텍스트 찾기
     * @param src 찾을 텍스트
     * @param regex 정규식 사용 여부
     * @return 찾으면 true, 없으면 false
     */
    bool FindBackward(const std::wstring& src, bool regex = false);

    /**
     * @brief 찾아 바꾸기 (방향 지정)
     * @param src 찾을 텍스트
     * @param dst 바꿀 텍스트
     * @param regex 정규식 사용 여부
     * @param direction 방향 (0=Forward, 1=Backward, 2=AllDoc)
     * @return 바꾼 개수
     */
    int FindReplace(const std::wstring& src,
                    const std::wstring& dst,
                    bool regex = false,
                    int direction = 0);

    /**
     * @brief 붙여넣기 (확장)
     * @param option 옵션 (0=왼쪽, 1=오른쪽, 2=위, 3=아래, 4=덮어쓰기, 5=내용만, 6=셀안에표)
     * @return 성공 여부
     */
    bool Paste(int option = 4);

    //=========================================================================
    // 파일 I/O 확장
    //=========================================================================

    /**
     * @brief 스타일 내보내기
     * @param styFilepath STY 파일 경로
     * @return 성공 여부
     */
    bool ExportStyle(const std::wstring& styFilepath);

    /**
     * @brief 스타일 가져오기
     * @param styFilepath STY 파일 경로
     * @return 성공 여부
     */
    bool ImportStyle(const std::wstring& styFilepath);

    /**
     * @brief 명령 잠금/해제
     * @param actId 액션 ID (예: "Undo", "Redo")
     * @param isLock true=잠금, false=해제
     */
    void LockCommand(const std::wstring& actId, bool isLock);

    /**
     * @brief 페이지 이미지 생성
     * @param path 이미지 파일 경로
     * @param pgno 페이지 번호 (1부터 시작, 0=현재 페이지)
     * @param resolution DPI (기본 300)
     * @param depth Color Depth (1, 4, 8, 24)
     * @param format 포맷 ("bmp", "gif")
     * @return 성공 여부
     */
    bool CreatePageImage(const std::wstring& path,
                         int pgno = 0,
                         int resolution = 300,
                         int depth = 24,
                         const std::wstring& format = L"bmp");

    /**
     * @brief 문서 인쇄
     * @return 성공 여부
     */
    bool PrintDocument();

    /**
     * @brief 메일 머지 실행
     * @return 성공 여부
     */
    bool MailMerge();

    //=========================================================================
    // 텍스트 편집 확장
    //=========================================================================

    /**
     * @brief 현재 캐럿 위치에 문서 삽입
     * @param path 삽입할 파일 경로
     * @param format 파일 형식 (빈 문자열=자동 감지)
     * @param arg 추가 인자
     * @param moveDocEnd 삽입 후 문서 끝으로 이동
     * @return 성공 여부
     */
    bool Insert(const std::wstring& path,
                const std::wstring& format = L"",
                const std::wstring& arg = L"",
                bool moveDocEnd = false);

    /**
     * @brief 표 셀에 배경 이미지 삽입
     * @param path 이미지 파일 경로
     * @param borderType "SelectedCell" 또는 "SelectedCellDelete"
     * @param embedded 문서에 포함 여부
     * @param fillOption 채우기 옵션 (0=바둑판식, 1=가로확대, 2=세로확대, 3=확대, 4=축소, 5=중앙)
     * @param effect 효과 (0=없음, 1=그레이스케일, 2=흑백)
     * @param watermark 워터마크 여부
     * @param brightness 밝기 (-100~100)
     * @param contrast 대비 (-100~100)
     * @return 성공 여부
     */
    bool InsertBackgroundPicture(const std::wstring& path,
                                  const std::wstring& borderType = L"SelectedCell",
                                  bool embedded = true,
                                  int fillOption = 5,
                                  int effect = 0,
                                  bool watermark = false,
                                  int brightness = 0,
                                  int contrast = 0);

    /**
     * @brief 메타태그로 캐럿 이동
     * @param tag 메타태그 이름
     * @param text 찾을 텍스트 (빈 문자열=첫 번째)
     * @param start 시작 위치로 이동
     * @param select 선택 여부
     * @return 성공 여부
     */
    bool MoveToMetatag(const std::wstring& tag,
                       const std::wstring& text = L"",
                       bool start = true,
                       bool select = false);

    /**
     * @brief 모든 필드의 텍스트 지우기
     */
    void ClearFieldText();

    /**
     * @brief 하이퍼링크 삽입
     * @param hypertext 링크 대상 (URL 또는 북마크)
     * @param description 링크 설명
     * @return 성공 여부
     */
    bool InsertHyperlink(const std::wstring& hypertext,
                         const std::wstring& description = L"");

    /**
     * @brief 메모 삽입
     * @param text 메모 내용
     * @param memoType 메모 유형 ("memo" 또는 "revision")
     */
    void InsertMemo(const std::wstring& text = L"",
                    const std::wstring& memoType = L"memo");

    /**
     * @brief 원문자/글자 겹치기 삽입
     * @param chars 겹칠 문자열
     * @param charSize 글자 크기 (-3~3, 음수=작게, 양수=크게)
     * @param checkCompose 합성 방식 (0=원문자, 1=겹치기)
     * @param circleType 테두리 유형 (0=원, 1=사각형 등)
     * @return 성공 여부
     */
    bool ComposeChars(const std::wstring& chars = L"",
                      int charSize = -3,
                      int checkCompose = 0,
                      int circleType = 0);

    /**
     * @brief 컨트롤로 캐럿 이동
     * @param pCtrl 컨트롤 객체
     * @param option 위치 옵션 (0=시작, 1=끝, 2=중간)
     * @return 성공 여부
     */
    bool MoveToCtrl(IDispatch* pCtrl, int option = 0);

    /**
     * @brief 컨트롤 선택
     * @param pCtrl 컨트롤 객체
     * @param anchorType 앵커 타입 (0=시작, 1=끝, 2=전체)
     * @param option 선택 옵션 (1=기본)
     * @return 성공 여부
     */
    bool SelectCtrl(IDispatch* pCtrl, int anchorType = 0, int option = 1);

    /**
     * @brief 모든 캡션 위치 일괄 변경
     * @param location 위치 ("Top", "Bottom", "Left", "Right")
     * @param align 정렬 ("Left", "Center", "Right", "Justify")
     * @return 성공 여부
     */
    bool MoveAllCaption(const std::wstring& location = L"Bottom",
                        const std::wstring& align = L"Justify");

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
    // 필드/메타태그 확장
    //=========================================================================

    /**
     * @brief 필드 속성 수정
     * @param field 필드 이름
     * @param remove 속성 제거 여부
     * @param add 속성 추가 여부
     * @return 성공 여부
     */
    bool ModifyFieldProperties(const std::wstring& field, bool remove, bool add);

    /**
     * @brief 개인정보 찾기
     * @param privateType 개인정보 유형
     * @param privateString 검색 문자열
     * @return 결과 (-1=끝, 0=없음, 비트마스크=유형)
     */
    int FindPrivateInfo(int privateType, const std::wstring& privateString);

    /**
     * @brief 현재 메타태그명 조회
     * @return 메타태그 이름 (없으면 빈 문자열)
     */
    std::wstring GetCurMetatagName();

    /**
     * @brief 메타태그 목록 조회
     * @param number 0=plain, 1=numbered
     * @param option 옵션
     * @return 메타태그 목록 (0x02로 구분)
     */
    std::wstring GetMetatagList(int number, int option);

    /**
     * @brief 메타태그 텍스트 조회
     * @param tag 메타태그 이름
     * @return 메타태그 텍스트
     */
    std::wstring GetMetatagNameText(const std::wstring& tag);

    /**
     * @brief 메타태그 텍스트 설정
     * @param tag 메타태그 이름
     * @param text 설정할 텍스트
     * @return 성공 여부
     */
    bool PutMetatagNameText(const std::wstring& tag, const std::wstring& text);

    /**
     * @brief 메타태그 이름 변경
     * @param oldtag 기존 이름
     * @param newtag 새 이름
     * @return 성공 여부
     */
    bool RenameMetatag(const std::wstring& oldtag, const std::wstring& newtag);

    /**
     * @brief 메타태그 속성 수정
     * @param tag 메타태그 이름
     * @param remove 속성 제거 여부
     * @param add 속성 추가 여부
     * @return 성공 여부
     */
    bool ModifyMetatagProperties(const std::wstring& tag, bool remove, bool add);

    /**
     * @brief 필드 정보 리스트 (HWPML2X 파싱)
     * @return 필드 정보 목록 [{name, direction, memo}, ...]
     */
    std::vector<std::map<std::wstring, std::wstring>> GetFieldInfo();

    /**
     * @brief 중괄호 구문을 필드로 변환
     * {{name:direction:memo}} → 누름틀
     * [[name:direction:memo]] → 셀 필드
     */
    void SetFieldByBracket();

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

    /**
     * @brief 첫 번째 열로 이동
     */
    bool TableColBegin();

    /**
     * @brief 마지막 열로 이동
     */
    bool TableColEnd();

    /**
     * @brief 열의 맨 위로 이동
     */
    bool TableColPageUp();

    /**
     * @brief 셀 블록 선택 확장 (절대)
     */
    bool TableCellBlockExtendAbs();

    /**
     * @brief 선택 취소 (ESC)
     */
    bool Cancel();

    /**
     * @brief 셀 배경색 채우기
     * @param r Red (0-255, 기본: 217)
     * @param g Green (0-255, 기본: 217)
     * @param b Blue (0-255, 기본: 217)
     * @return 성공 여부
     */
    bool CellFill(int r = 217, int g = 217, int b = 217);

    /**
     * @brief 2D 벡터 데이터로 테이블 생성
     * @param data 2D 문자열 벡터 (첫 행은 헤더)
     * @param treat_as_char 글자처럼 취급 (기본: false)
     * @param header 제목행 설정 (기본: true)
     * @param header_bold 헤더에 볼드 적용 (기본: true)
     * @param cell_fill_r 헤더 배경색 R (-1이면 미적용)
     * @param cell_fill_g 헤더 배경색 G
     * @param cell_fill_b 헤더 배경색 B
     * @return 성공 여부
     */
    bool TableFromData(
        const std::vector<std::vector<std::wstring>>& data,
        bool treat_as_char = false,
        bool header = true,
        bool header_bold = true,
        int cell_fill_r = -1,
        int cell_fill_g = -1,
        int cell_fill_b = -1);

    /**
     * @brief 현재 테이블의 XML 추출 (HWPML2X 형식)
     * @return XML 문자열 (실패 시 빈 문자열)
     *
     * Python의 table_to_df()에서 파싱하여 DataFrame으로 변환.
     * 병합 셀 포함 모든 셀 데이터 정확히 추출 가능.
     */
    std::wstring GetTableXml();

    //=========================================================================
    // 스타일 관리 (CharShape/ParaShape)
    //=========================================================================

    /**
     * @brief 현재 글자모양 가져오기
     * @return 글자모양 속성 맵 (Height, Bold, Italic, FaceNameHangul 등)
     *
     * pyhwpx의 get_charshape()에 대응.
     * HAction.GetDefault("CharShape")로 현재 설정을 가져옴.
     */
    std::map<std::wstring, int> GetCharShape();

    /**
     * @brief 글자모양 설정
     * @param props 설정할 속성 맵
     * @return 성공 여부
     *
     * pyhwpx의 set_charshape()에 대응.
     * 선택 영역에 글자모양 적용.
     */
    bool SetCharShape(const std::map<std::wstring, int>& props);

    /**
     * @brief 글자모양 간편 설정
     * @param face_name 글꼴 이름 (빈 문자열이면 미변경)
     * @param height 글자 크기 (pt * 100, 예: 10pt = 1000, -1이면 미변경)
     * @param bold 볼드 (0=해제, 1=설정, -1=미변경)
     * @param italic 이탤릭 (0=해제, 1=설정, -1=미변경)
     * @param text_color 글자색 RGB (COLORREF, -1이면 미변경)
     * @return 성공 여부
     *
     * 자주 사용하는 글자모양 속성을 간편하게 설정.
     */
    bool SetFont(const std::wstring& face_name = L"",
                 int height = -1,
                 int bold = -1,
                 int italic = -1,
                 int text_color = -1);

    /**
     * @brief 현재 문단모양 가져오기
     * @return 문단모양 속성 맵 (AlignType, LineSpacing, LeftMargin 등)
     *
     * pyhwpx의 get_parashape()에 대응.
     */
    std::map<std::wstring, int> GetParaShape();

    /**
     * @brief 문단모양 설정
     * @param props 설정할 속성 맵
     * @return 성공 여부
     *
     * pyhwpx의 set_parashape()에 대응.
     */
    bool SetParaShape(const std::map<std::wstring, int>& props);

    /**
     * @brief 문단모양 간편 설정
     * @param align_type 정렬 (0=양쪽, 1=왼쪽, 2=가운데, 3=오른쪽, -1=미변경)
     * @param line_spacing 줄간격 (%, -1이면 미변경)
     * @param left_margin 왼쪽 여백 (HwpUnit, -1이면 미변경)
     * @param indentation 첫줄 들여쓰기 (HwpUnit, -1이면 미변경)
     * @return 성공 여부
     */
    bool SetPara(int align_type = -1,
                 int line_spacing = -1,
                 int left_margin = -1,
                 int indentation = -1);

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

    //=========================================================================
    // 컨트롤 관리
    //=========================================================================

    /**
     * @brief 컨트롤 삽입
     * @param ctrl_id 컨트롤 ID ("tbl"=표, "pic"=그림, "gso"=도형, "eqed"=수식 등)
     * @param initparam 초기 파라미터셋 (nullptr이면 기본값)
     * @return 생성된 컨트롤 (실패 시 nullptr)
     *
     * 사용 예시:
     * ```cpp
     * // 표 삽입
     * auto pset = hwp.CreateSet(L"TableCreation");
     * // ... 파라미터 설정 ...
     * auto ctrl = hwp.InsertCtrl(L"tbl", pset);
     * ```
     */
    std::unique_ptr<HwpCtrl> InsertCtrl(const std::wstring& ctrl_id,
                                         IDispatch* initparam = nullptr);

    /**
     * @brief 컨트롤 삭제
     * @param ctrl 삭제할 컨트롤
     * @return 성공 여부
     */
    bool DeleteCtrl(HwpCtrl* ctrl);

    /**
     * @brief COM 객체로 컨트롤 삭제
     * @param pCtrl 삭제할 컨트롤 COM 객체
     * @return 성공 여부
     */
    bool DeleteCtrl(IDispatch* pCtrl);

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
    // 컨트롤 속성 접근
    //=========================================================================

    /**
     * @brief 현재 선택된 컨트롤
     * pyhwpx의 CurSelectedCtrl에 대응
     * @return 선택된 컨트롤 (없으면 nullptr)
     */
    std::unique_ptr<HwpCtrl> GetCurSelectedCtrl();

    /**
     * @brief 문서의 첫 번째 컨트롤 (항상 secd)
     * pyhwpx의 HeadCtrl에 대응
     * @return 첫 번째 컨트롤
     */
    std::unique_ptr<HwpCtrl> GetHeadCtrl();

    /**
     * @brief 문서의 마지막 컨트롤
     * pyhwpx의 LastCtrl에 대응
     * @return 마지막 컨트롤
     */
    std::unique_ptr<HwpCtrl> GetLastCtrl();

    /**
     * @brief 현재 위치를 포함하는 상위 컨트롤
     * pyhwpx의 ParentCtrl에 대응
     * @return 부모 컨트롤 (없으면 nullptr)
     */
    std::unique_ptr<HwpCtrl> GetParentCtrl();

    /**
     * @brief 문서 내 모든 사용자 컨트롤 목록
     * pyhwpx의 ctrl_list에 대응
     * secd(섹션정의)와 cold(단정의)는 제외
     * @return 컨트롤 목록
     */
    std::vector<std::unique_ptr<HwpCtrl>> GetCtrlList();

    //=========================================================================
    // 문서 컬렉션 접근
    //=========================================================================

    /**
     * @brief XHwpDocuments 컬렉션 접근
     * @return 문서 컬렉션
     */
    std::unique_ptr<XHwpDocuments> GetXHwpDocuments();

    /**
     * @brief 문서 전환
     * @param num 문서 인덱스 (0-based)
     * @return 활성화된 문서 (nullptr이면 실패)
     */
    std::unique_ptr<XHwpDocument> SwitchTo(int num);

    /**
     * @brief 새 탭으로 문서 추가
     * @return 추가된 문서
     */
    std::unique_ptr<XHwpDocument> AddTab();

    /**
     * @brief 새 창으로 문서 추가
     * @return 추가된 문서
     */
    std::unique_ptr<XHwpDocument> AddDoc();

    //=========================================================================
    // COM 속성 접근 (Low-level)
    //=========================================================================

    /**
     * @brief Application 객체 접근 (Low-level API)
     */
    IDispatch* GetApplication();

    /**
     * @brief 클래스 ID
     */
    std::wstring GetCLSID();

    /**
     * @brief 현재 필드 상태 (0=본문, 1=셀, 4=글상자)
     */
    int GetCurFieldState();

    /**
     * @brief 현재 메타태그 상태 (한글2024+)
     */
    int GetCurMetatagState();

    /**
     * @brief 엔진 속성 객체
     */
    IDispatch* GetEngineProperties();

    /**
     * @brief 개인정보 보호 여부
     */
    bool GetIsPrivateInfoProtected();

    /**
     * @brief 변경 추적 여부
     */
    bool GetIsTrackChange();

    /**
     * @brief 문서 경로
     */
    std::wstring GetDocPath();

    /**
     * @brief 선택 모드 (0=일반, 1=블록)
     */
    int GetSelectionMode();

    /**
     * @brief 창 제목
     */
    std::wstring GetTitle();

    /**
     * @brief ViewProperties 객체 (getter)
     */
    IDispatch* GetViewProperties();

    /**
     * @brief ViewProperties 객체 (setter)
     */
    void SetViewProperties(IDispatch* props);

    /**
     * @brief XHwpMessageBox 객체
     */
    IDispatch* GetXHwpMessageBox();

    /**
     * @brief XHwpODBC 객체
     */
    IDispatch* GetXHwpODBC();

    /**
     * @brief XHwpWindows 객체
     */
    IDispatch* GetXHwpWindows();

    /**
     * @brief 현재 폰트명
     */
    std::wstring GetCurrentFont();

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

    //=========================================================================
    // 파라미터 헬퍼 (Parameter Helpers)
    // pyhwpx param_helpers.py 포팅
    //=========================================================================

    // === 정렬 관련 ===
    int HAlign(const std::wstring& h_align);
    int VAlign(const std::wstring& v_align);
    int TextAlign(const std::wstring& text_align);
    int ParaHeadAlign(const std::wstring& para_head_align);
    int TextArtAlign(const std::wstring& text_art_align);

    // === 선/테두리 관련 ===
    int HwpLineType(const std::wstring& line_type);
    int HwpLineWidth(const std::wstring& line_width);
    int BorderShape(const std::wstring& border_type);
    int EndStyle(const std::wstring& end_style);
    int EndSize(const std::wstring& end_size);

    // === 서식 관련 ===
    int NumberFormat(const std::wstring& num_format);
    int HeadType(const std::wstring& heading_type);
    int FontType(const std::wstring& font_type);
    int StrikeOut(const std::wstring& strike_out_type);
    int HwpUnderlineType(const std::wstring& underline_type);
    int HwpUnderlineShape(const std::wstring& underline_shape);
    int StyleType(const std::wstring& style_type);

    // === 검색/효과 ===
    int FindDir(const std::wstring& find_dir);
    int PicEffect(const std::wstring& pic_effect);
    int HwpZoomType(const std::wstring& zoom_type);

    // === 페이지/인쇄 ===
    int PageNumPosition(const std::wstring& pagenum_pos);
    int PageType(const std::wstring& page_type);
    int PrintRange(const std::wstring& print_range);
    int PrintType(const std::wstring& print_method);
    int PrintDevice(const std::wstring& print_device);
    int PrintPaper(const std::wstring& print_paper);
    int SideType(const std::wstring& side_type);

    // === 채우기/그라데이션 ===
    int BrushType(const std::wstring& brush_type);
    int FillAreaType(const std::wstring& fill_area);
    int Gradation(const std::wstring& gradation);
    int HatchStyle(const std::wstring& hatch_style);
    int WatermarkBrush(const std::wstring& watermark_brush);

    // === 표 관련 ===
    int TableFormat(const std::wstring& table_format);
    int TableBreak(const std::wstring& page_break);
    int TableTarget(const std::wstring& table_target);
    int TableSwapType(const std::wstring& tableswap);
    int CellApply(const std::wstring& cell_apply);
    int GridMethod(const std::wstring& grid_method);
    int GridViewLine(const std::wstring& grid_view_line);

    // === 텍스트 흐름/배치 ===
    int TextDir(const std::wstring& text_direction);
    int TextWrapType(const std::wstring& text_wrap);
    int TextFlowType(const std::wstring& text_flow);
    int LineWrapType(const std::wstring& line_wrap);
    int LineSpacingMethod(const std::wstring& line_spacing);

    // === 도형/이미지 ===
    int ArcType(const std::wstring& arc_type);
    int DrawAspect(const std::wstring& draw_aspect);
    int DrawFillImage(const std::wstring& fillimage);
    int DrawShadowType(const std::wstring& shadow_type);
    int CharShadowType(const std::wstring& shadow_type);
    int ImageFormat(const std::wstring& image_format);
    int PlacementType(const std::wstring& restart);

    // === 위치/크기 관련 ===
    int HorzRel(const std::wstring& horz_rel);
    int VertRel(const std::wstring& vert_rel);
    int HeightRel(const std::wstring& height_rel);
    int WidthRel(const std::wstring& width_rel);

    // === 개요/번호 ===
    int AutoNumType(const std::wstring& autonum);
    int Numbering(const std::wstring& numbering);
    int HwpOutlineStyle(const std::wstring& hwp_outline_style);
    int HwpOutlineType(const std::wstring& hwp_outline_type);

    // === 열/단 정의 ===
    int ColDefType(const std::wstring& col_def_type);
    int ColLayoutType(const std::wstring& col_layout_type);
    int GutterMethod(const std::wstring& gutter_type);

    // === 기타 옵션 ===
    int BreakWordLatin(const std::wstring& break_latin_word);
    int Canonical(const std::wstring& canonical);
    int ConvertPUAHangulToUnicode(bool reverse);
    int CrookedSlash(const std::wstring& crooked_slash);
    int DbfCodeType(const std::wstring& dbf_code);
    int Delimiter(const std::wstring& delimiter);
    int DSMark(const std::wstring& diac_sym_mark);
    int Encrypt(const std::wstring& encrypt);
    int Handler(const std::wstring& handler);
    int Hash(const std::wstring& hash);
    int Hiding(const std::wstring& hiding);
    int MacroState(const std::wstring& macro_state);
    int MailType(const std::wstring& mail_type);
    int PresentEffect(const std::wstring& prsnteffect);
    int Signature(const std::wstring& signature);
    int Slash(const std::wstring& slash);
    int SortDelimiter(const std::wstring& sort_delimiter);
    int SubtPos(const std::wstring& subt_pos);
    int ViewFlag(const std::wstring& view_flag);

    // === 사용자 정보 ===
    std::wstring GetUserInfo(const std::wstring& user_info_id);
    bool SetUserInfo(const std::wstring& user_info_id, const std::wstring& value);

    // === 메타태그/DRM ===
    bool SetCurMetatagName(const std::wstring& tag);
    bool SetDRMAuthority(const std::wstring& authority);

    // === 번역 ===
    std::vector<std::wstring> GetTranslateLangList(const std::wstring& cur_lang);

    // === 음력/양력 변환 ===
    std::tuple<int, int, int> LunarToSolarBySet(int l_year, int l_month, int l_day, bool l_leap);
    std::tuple<int, int, int, bool> SolarToLunarBySet(int s_year, int s_month, int s_day);

    // === 단위 변환 확장 ===
    double HwpUnitToInch(int hwp_unit);
    double HwpUnitToPoint(int hwp_unit);
    int PointToHwpUnit(double point);

protected:
    //=========================================================================
    // COM 헬퍼 메서드
    //=========================================================================

    /**
     * @brief 파라미터 헬퍼 공통 호출 (문자열 파라미터 1개)
     * @param method 메서드 이름
     * @param param 파라미터 값
     * @return 결과 정수값
     */
    int InvokeParamHelper(const std::wstring& method, const std::wstring& param);

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
    bool m_bRegisterModule;         // 보안 모듈 자동 등록 여부 (pyhwpx 호환)

    DISPIDCache m_dispidCache;      // DISPID 캐시 (성능 최적화)
    std::vector<IDispatch*> m_posCache;  // GetPosBySet 결과 캐시 (Python용)

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
