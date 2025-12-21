/**
 * @file HwpWrapper.cpp
 * @brief HWP COM 객체 래퍼 클래스 구현
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 */

#include "HwpWrapper.h"
#include "HwpCtrl.h"
#include <stdexcept>
#include <atlbase.h>
#include <atlcom.h>
#include <shlwapi.h>

#pragma comment(lib, "shlwapi.lib")

namespace cpyhwpx {

//=============================================================================
// DISPIDCache 구현
//=============================================================================

DISPID DISPIDCache::GetOrLoad(IDispatch* pObj, const std::wstring& name)
{
    if (!pObj) return DISPID_UNKNOWN;

    // 캐시에서 먼저 검색
    auto it = m_cache.find(name);
    if (it != m_cache.end()) {
        return it->second;
    }

    // 캐시에 없으면 GetIDsOfNames 호출
    DISPID dispid;
    OLECHAR* pName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = pObj->GetIDsOfNames(IID_NULL, &pName, 1, LOCALE_USER_DEFAULT, &dispid);

    if (SUCCEEDED(hr)) {
        m_cache[name] = dispid;
        return dispid;
    }

    return DISPID_UNKNOWN;
}

//=============================================================================
// 생성자/소멸자
//=============================================================================

HwpWrapper::HwpWrapper(bool visible, bool new_instance, bool register_module)
    : m_pHwp(nullptr)
    , m_pHParameterSet(nullptr)
    , m_pHAction(nullptr)
    , m_bInitialized(false)
    , m_bVisible(visible)
    , m_bNewInstance(new_instance)
    , m_bRegisterModule(register_module)
{
}

HwpWrapper::~HwpWrapper()
{
    Release();
}

//=============================================================================
// 초기화 및 종료
//=============================================================================

bool HwpWrapper::InitializeCOM()
{
    HRESULT hr = CoInitialize(NULL);
    if (FAILED(hr) && hr != RPC_E_CHANGED_MODE) {
        return false;
    }
    return true;
}

bool HwpWrapper::CreateHwpObject()
{
    if (m_pHwp) {
        return true;  // 이미 생성됨
    }

    CLSID clsid;
    HRESULT hr = CLSIDFromProgID(L"HWPFrame.HwpObject", &clsid);
    if (FAILED(hr)) {
        return false;
    }

    DWORD clsContext = m_bNewInstance ? CLSCTX_LOCAL_SERVER : CLSCTX_ALL;
    hr = CoCreateInstance(clsid, NULL, clsContext, IID_IDispatch, (void**)&m_pHwp);
    if (FAILED(hr)) {
        return false;
    }

    return true;
}

bool HwpWrapper::Initialize()
{
    if (m_bInitialized) {
        return true;
    }

    if (!InitializeCOM()) {
        return false;
    }

    if (!CreateHwpObject()) {
        return false;
    }

    // COM 객체 안정화 대기 (제거됨 - 성능 최적화)

    // 편집 모드 활성화 (텍스트 삽입을 위해 필요)
    SetEditMode(true);

    // 창 표시 설정
    SetVisible(m_bVisible);

    m_bInitialized = true;

    // 보안 모듈 자동 등록 (pyhwpx 방식)
    if (m_bRegisterModule) {
        try {
            AutoRegisterModule();
        } catch (...) {
            // 등록 실패해도 계속 진행 (pyhwpx와 동일)
            // 사용자가 수동으로 등록할 수 있음
        }
    }

    return true;
}

bool HwpWrapper::RegisterModule(const std::wstring& module_type,
                                 const std::wstring& module_data)
{
    if (!m_pHwp) return false;

    // DISPID 캐시 사용 (성능 최적화)
    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"RegisterModule");
    if (dispid == DISPID_UNKNOWN) return false;

    VARIANT args[2];
    VariantInit(&args[0]);
    VariantInit(&args[1]);

    // 인자는 역순으로
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(module_data.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(module_type.c_str());

    DISPPARAMS params = { args, NULL, 2, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    SysFreeString(args[0].bstrVal);
    SysFreeString(args[1].bstrVal);

    if (FAILED(hr)) return false;
    return (result.vt == VT_BOOL) ? (result.boolVal != VARIANT_FALSE) : true;
}

//=============================================================================
// 보안 모듈 관리
//=============================================================================

bool HwpWrapper::CheckRegistryKey(const std::wstring& key_name)
{
    HKEY hKey;
    std::wstring subKey = L"Software\\HNC\\HwpAutomation\\Modules";

    // 첫 번째 경로 시도
    if (RegOpenKeyExW(HKEY_CURRENT_USER, subKey.c_str(), 0, KEY_READ, &hKey) != ERROR_SUCCESS) {
        // 대체 경로 시도
        subKey = L"Software\\Hnc\\HwpUserAction\\Modules";
        if (RegOpenKeyExW(HKEY_CURRENT_USER, subKey.c_str(), 0, KEY_READ, &hKey) != ERROR_SUCCESS) {
            return false;
        }
    }

    wchar_t buffer[MAX_PATH];
    DWORD bufferSize = sizeof(buffer);
    DWORD type;

    LONG result = RegQueryValueExW(hKey, key_name.c_str(), NULL, &type,
                                    (LPBYTE)buffer, &bufferSize);
    RegCloseKey(hKey);

    if (result != ERROR_SUCCESS || type != REG_SZ) {
        return false;
    }

    // DLL 파일 존재 확인
    return PathFileExistsW(buffer) == TRUE;
}

std::wstring HwpWrapper::FindDllPath()
{
    wchar_t modulePath[MAX_PATH];

    // 1. 현재 모듈(cpyhwpx.pyd) 경로에서 검색
    HMODULE hModule = NULL;
    GetModuleHandleExW(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS,
                       (LPCWSTR)&HwpWrapper::FindDllPath, &hModule);
    GetModuleFileNameW(hModule, modulePath, MAX_PATH);

    std::wstring basePath(modulePath);
    size_t pos = basePath.find_last_of(L"\\/");
    if (pos != std::wstring::npos) {
        basePath = basePath.substr(0, pos + 1);
    }

    // 같은 폴더에서 검색 (_native 폴더)
    std::wstring dllPath = basePath + L"FilePathCheckerModule.dll";
    if (PathFileExistsW(dllPath.c_str())) {
        return dllPath;
    }

    // 2. 상위 폴더/cpyhwpx/_native 검색
    pos = basePath.find_last_of(L"\\/", basePath.length() - 2);
    if (pos != std::wstring::npos) {
        std::wstring parentPath = basePath.substr(0, pos + 1);
        dllPath = parentPath + L"cpyhwpx\\_native\\FilePathCheckerModule.dll";
        if (PathFileExistsW(dllPath.c_str())) {
            return dllPath;
        }
    }

    return L"";
}

bool HwpWrapper::RegisterToRegistry(const std::wstring& dll_path,
                                     const std::wstring& key_name)
{
    std::wstring actualPath = dll_path.empty() ? FindDllPath() : dll_path;
    if (actualPath.empty()) {
        return false;
    }

    HKEY hKey;
    std::wstring subKey = L"Software\\HNC\\HwpAutomation\\Modules";

    // 키 열기/생성
    if (RegCreateKeyExW(HKEY_CURRENT_USER, subKey.c_str(), 0, NULL,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) != ERROR_SUCCESS) {
        // 대체 경로 시도
        subKey = L"Software\\Hnc\\HwpUserAction\\Modules";
        if (RegCreateKeyExW(HKEY_CURRENT_USER, subKey.c_str(), 0, NULL,
                            REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) != ERROR_SUCCESS) {
            return false;
        }
    }

    LONG result = RegSetValueExW(hKey, key_name.c_str(), 0, REG_SZ,
                                  (const BYTE*)actualPath.c_str(),
                                  (DWORD)((actualPath.length() + 1) * sizeof(wchar_t)));
    RegCloseKey(hKey);

    return result == ERROR_SUCCESS;
}

bool HwpWrapper::AutoRegisterModule(const std::wstring& module_type,
                                     const std::wstring& module_data)
{
    // 1. 레지스트리 확인
    if (!CheckRegistryKey(module_data)) {
        // 2. 미등록이면 등록
        if (!RegisterToRegistry(L"", module_data)) {
            return false;
        }
    }

    // 3. COM API 호출
    return RegisterModule(module_type, module_data);
}

void HwpWrapper::Quit(bool save)
{
    if (!m_pHwp) return;

    // pyhwpx 방식: Clear() 후 Quit()
    if (save) {
        // save=true → Save() 호출
        Save(true);
    } else {
        // save=false → ClearDocument(1) 호출 (문서 내용 버림)
        // XHwpDocuments.Active_XHwpDocument.Clear(1) = hwpDiscard
        ClearDocument(1);
    }

    // hwp.Quit() 직접 호출
    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"Quit");
    if (dispid != DISPID_UNKNOWN) {
        DISPPARAMS params = { NULL, NULL, 0, 0 };
        m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                       DISPATCH_METHOD, &params, NULL, NULL, NULL);
    }

    Release();
}

void HwpWrapper::Release()
{
    if (m_pHAction) {
        m_pHAction->Release();
        m_pHAction = nullptr;
    }
    if (m_pHParameterSet) {
        m_pHParameterSet->Release();
        m_pHParameterSet = nullptr;
    }
    if (m_pHwp) {
        m_pHwp->Release();
        m_pHwp = nullptr;
    }
    m_bInitialized = false;
}

//=============================================================================
// 파일 I/O
//=============================================================================

bool HwpWrapper::Open(const std::wstring& filename,
                       const std::wstring& format,
                       const std::wstring& arg)
{
    if (!m_pHwp) return false;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"Open");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    VARIANT args[3];
    VariantInit(&args[0]);
    VariantInit(&args[1]);
    VariantInit(&args[2]);

    // 인자는 역순
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(arg.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(format.c_str());
    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(filename.c_str());

    DISPPARAMS params = { args, NULL, 3, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    SysFreeString(args[0].bstrVal);
    SysFreeString(args[1].bstrVal);
    SysFreeString(args[2].bstrVal);

    if (FAILED(hr)) return false;
    return (result.vt == VT_BOOL) ? (result.boolVal != VARIANT_FALSE) : true;
}

bool HwpWrapper::Save(bool save_if_dirty)
{
    if (!m_pHwp) return false;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"Save");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    VARIANT args[1];
    VariantInit(&args[0]);
    args[0].vt = VT_BOOL;
    args[0].boolVal = save_if_dirty ? VARIANT_TRUE : VARIANT_FALSE;

    DISPPARAMS params = { args, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    if (FAILED(hr)) return false;
    return (result.vt == VT_BOOL) ? (result.boolVal != VARIANT_FALSE) : true;
}

bool HwpWrapper::SaveAs(const std::wstring& filename,
                         const std::wstring& format,
                         const std::wstring& arg)
{
    if (!m_pHwp) return false;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"SaveAs");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    VARIANT args[3];
    VariantInit(&args[0]);
    VariantInit(&args[1]);
    VariantInit(&args[2]);

    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(arg.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(format.c_str());
    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(filename.c_str());

    DISPPARAMS params = { args, NULL, 3, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    SysFreeString(args[0].bstrVal);
    SysFreeString(args[1].bstrVal);
    SysFreeString(args[2].bstrVal);

    if (FAILED(hr)) return false;
    return (result.vt == VT_BOOL) ? (result.boolVal != VARIANT_FALSE) : true;
}

void HwpWrapper::Clear(int option)
{
    if (!m_pHwp) return;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"Clear");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return;

    VARIANT args[1];
    VariantInit(&args[0]);
    args[0].vt = VT_I4;
    args[0].lVal = option;

    DISPPARAMS params = { args, NULL, 1, 0 };
    m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                   &params, NULL, NULL, NULL);
}

bool HwpWrapper::ClearDocument(int option)
{
    if (!m_pHwp) return false;

    // pyhwpx 방식: XHwpDocuments.Active_XHwpDocument.Clear(option)

    // 1. XHwpDocuments 속성 획득
    VARIANT xhwpDocsVar = GetProperty(L"XHwpDocuments");
    if (xhwpDocsVar.vt != VT_DISPATCH || !xhwpDocsVar.pdispVal) {
        VariantClear(&xhwpDocsVar);
        return false;
    }

    IDispatch* pXHwpDocuments = xhwpDocsVar.pdispVal;

    // 2. Active_XHwpDocument 속성 획득
    DISPID dispidActive;
    OLECHAR* activeName = const_cast<OLECHAR*>(L"Active_XHwpDocument");
    HRESULT hr = pXHwpDocuments->GetIDsOfNames(IID_NULL, &activeName, 1,
                                                LOCALE_USER_DEFAULT, &dispidActive);
    if (FAILED(hr)) {
        pXHwpDocuments->Release();
        return false;
    }

    DISPPARAMS paramsEmpty = { NULL, NULL, 0, 0 };
    VARIANT activeDocVar;
    VariantInit(&activeDocVar);
    hr = pXHwpDocuments->Invoke(dispidActive, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &paramsEmpty, &activeDocVar, NULL, NULL);

    if (FAILED(hr) || activeDocVar.vt != VT_DISPATCH || !activeDocVar.pdispVal) {
        VariantClear(&activeDocVar);
        pXHwpDocuments->Release();
        return false;
    }

    IDispatch* pActiveDoc = activeDocVar.pdispVal;

    // 3. Clear(option) 메서드 호출
    DISPID dispidClear;
    OLECHAR* clearName = const_cast<OLECHAR*>(L"Clear");
    hr = pActiveDoc->GetIDsOfNames(IID_NULL, &clearName, 1,
                                    LOCALE_USER_DEFAULT, &dispidClear);
    bool result = false;
    if (SUCCEEDED(hr)) {
        VARIANT arg;
        VariantInit(&arg);
        arg.vt = VT_I4;
        arg.lVal = option;  // 1 = hwpDiscard

        DISPPARAMS params = { &arg, NULL, 1, 0 };
        VARIANT clearResult;
        VariantInit(&clearResult);
        hr = pActiveDoc->Invoke(dispidClear, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_METHOD, &params, &clearResult, NULL, NULL);
        result = SUCCEEDED(hr);
        VariantClear(&clearResult);
    }

    pActiveDoc->Release();
    pXHwpDocuments->Release();

    return result;
}

bool HwpWrapper::Close(bool is_dirty)
{
    if (!m_pHwp) return false;

    // dirty 상태 설정
    VARIANT val;
    VariantInit(&val);
    val.vt = VT_BOOL;
    val.boolVal = is_dirty ? VARIANT_TRUE : VARIANT_FALSE;
    SetProperty(L"IsModified", val);

    return RunAction(L"FileClose");
}

bool HwpWrapper::InsertFile(const std::wstring& filename,
                             int keep_section,
                             int keep_charshape,
                             int keep_parashape,
                             int keep_style,
                             bool move_doc_end)
{
    if (!m_pHwp) return false;

    HRESULT hr;
    DISPID dispid;
    VARIANT result;
    VariantInit(&result);

    // 1. HParameterSet 속성 가져오기
    OLECHAR* psetName = const_cast<OLECHAR*>(L"HParameterSet");
    hr = m_pHwp->GetIDsOfNames(IID_NULL, &psetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                        &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pHParameterSet = result.pdispVal;

    // 2. HInsertFile 속성 가져오기
    OLECHAR* insertFileName = const_cast<OLECHAR*>(L"HInsertFile");
    hr = pHParameterSet->GetIDsOfNames(IID_NULL, &insertFileName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHParameterSet->Release();
        return false;
    }

    VariantInit(&result);
    hr = pHParameterSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                 &noParams, &result, NULL, NULL);
    pHParameterSet->Release();
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pHInsertFile = result.pdispVal;

    // 3. HSet 속성 가져오기
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pHInsertFile->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHInsertFile->Release();
        return false;
    }

    VariantInit(&result);
    hr = pHInsertFile->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                               &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        pHInsertFile->Release();
        return false;
    }

    IDispatch* pHSet = result.pdispVal;

    // 4. HAction 가져오기
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pHSet->Release();
        pHInsertFile->Release();
        return false;
    }

    // 5. HAction.GetDefault("InsertFile", HSet) 호출
    OLECHAR* getDefaultName = const_cast<OLECHAR*>(L"GetDefault");
    hr = pHAction->GetIDsOfNames(IID_NULL, &getDefaultName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHSet->Release();
        pHInsertFile->Release();
        return false;
    }

    VARIANT getDefaultArgs[2];
    VariantInit(&getDefaultArgs[0]);
    VariantInit(&getDefaultArgs[1]);
    getDefaultArgs[0].vt = VT_DISPATCH;
    getDefaultArgs[0].pdispVal = pHSet;
    getDefaultArgs[1].vt = VT_BSTR;
    getDefaultArgs[1].bstrVal = SysAllocString(L"InsertFile");

    DISPPARAMS getDefaultParams = { getDefaultArgs, NULL, 2, 0 };
    VariantInit(&result);
    hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &getDefaultParams, &result, NULL, NULL);
    SysFreeString(getDefaultArgs[1].bstrVal);

    // 6. 파라미터 설정: filename
    OLECHAR* filenamePropName = const_cast<OLECHAR*>(L"filename");
    hr = pHInsertFile->GetIDsOfNames(IID_NULL, &filenamePropName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT filenameVal;
        VariantInit(&filenameVal);
        filenameVal.vt = VT_BSTR;
        filenameVal.bstrVal = SysAllocString(filename.c_str());

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS filenameParams = { &filenameVal, &putid, 1, 1 };
        pHInsertFile->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                              &filenameParams, NULL, NULL, NULL);
        SysFreeString(filenameVal.bstrVal);
    }

    // 7. 파라미터 설정: KeepSection
    OLECHAR* keepSectionName = const_cast<OLECHAR*>(L"KeepSection");
    hr = pHInsertFile->GetIDsOfNames(IID_NULL, &keepSectionName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT keepVal;
        VariantInit(&keepVal);
        keepVal.vt = VT_I4;
        keepVal.lVal = keep_section;

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS keepParams = { &keepVal, &putid, 1, 1 };
        pHInsertFile->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                              &keepParams, NULL, NULL, NULL);
    }

    // 8. 파라미터 설정: KeepCharshape
    OLECHAR* keepCharshapeName = const_cast<OLECHAR*>(L"KeepCharshape");
    hr = pHInsertFile->GetIDsOfNames(IID_NULL, &keepCharshapeName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT keepVal;
        VariantInit(&keepVal);
        keepVal.vt = VT_I4;
        keepVal.lVal = keep_charshape;

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS keepParams = { &keepVal, &putid, 1, 1 };
        pHInsertFile->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                              &keepParams, NULL, NULL, NULL);
    }

    // 9. 파라미터 설정: KeepParashape
    OLECHAR* keepParashapeName = const_cast<OLECHAR*>(L"KeepParashape");
    hr = pHInsertFile->GetIDsOfNames(IID_NULL, &keepParashapeName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT keepVal;
        VariantInit(&keepVal);
        keepVal.vt = VT_I4;
        keepVal.lVal = keep_parashape;

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS keepParams = { &keepVal, &putid, 1, 1 };
        pHInsertFile->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                              &keepParams, NULL, NULL, NULL);
    }

    // 10. 파라미터 설정: KeepStyle
    OLECHAR* keepStyleName = const_cast<OLECHAR*>(L"KeepStyle");
    hr = pHInsertFile->GetIDsOfNames(IID_NULL, &keepStyleName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT keepVal;
        VariantInit(&keepVal);
        keepVal.vt = VT_I4;
        keepVal.lVal = keep_style;

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS keepParams = { &keepVal, &putid, 1, 1 };
        pHInsertFile->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                              &keepParams, NULL, NULL, NULL);
    }

    // 11. HAction.Execute("InsertFile", HSet) 호출
    OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
    hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHSet->Release();
        pHInsertFile->Release();
        return false;
    }

    VARIANT executeArgs[2];
    VariantInit(&executeArgs[0]);
    VariantInit(&executeArgs[1]);
    executeArgs[0].vt = VT_DISPATCH;
    executeArgs[0].pdispVal = pHSet;
    executeArgs[1].vt = VT_BSTR;
    executeArgs[1].bstrVal = SysAllocString(L"InsertFile");

    DISPPARAMS executeParams = { executeArgs, NULL, 2, 0 };
    VariantInit(&result);
    hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &executeParams, &result, NULL, NULL);
    SysFreeString(executeArgs[1].bstrVal);

    pHSet->Release();
    pHInsertFile->Release();

    bool success = SUCCEEDED(hr) && (result.vt == VT_BOOL ? result.boolVal != VARIANT_FALSE : true);

    // 12. move_doc_end 처리
    if (success && move_doc_end) {
        MovePos(3, 0, 0);  // moveDocEnd = 3
    }

    return success;
}

std::wstring HwpWrapper::GetTextFile(const std::wstring& format,
                                      const std::wstring& option)
{
    if (!m_pHwp) return L"";

    // COM 이름 지정 파라미터 호출:
    // hwp.GetTextFile(Format=format, option=option)

    HRESULT hr;
    DISPID dispid;

    // 1. GetTextFile 메서드의 DISPID 획득
    OLECHAR* methodName = const_cast<OLECHAR*>(L"Gettextfile");
    hr = m_pHwp->GetIDsOfNames(IID_NULL, &methodName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return L"";

    // 2. 이름 지정 파라미터의 DISPID 획득
    OLECHAR* paramNames[2] = {
        const_cast<OLECHAR*>(L"Format"),
        const_cast<OLECHAR*>(L"option")
    };
    DISPID namedDispids[2];
    hr = m_pHwp->GetIDsOfNames(IID_NULL, paramNames, 2, LOCALE_USER_DEFAULT, namedDispids);
    if (FAILED(hr)) {
        // 이름 지정 파라미터 실패 시 위치 기반으로 시도
        VARIANT args[2];
        VariantInit(&args[0]);
        VariantInit(&args[1]);

        // 인자는 역순 (option, format)
        args[0].vt = VT_BSTR;
        args[0].bstrVal = SysAllocString(option.c_str());
        args[1].vt = VT_BSTR;
        args[1].bstrVal = SysAllocString(format.c_str());

        DISPPARAMS params = { args, NULL, 2, 0 };
        VARIANT result;
        VariantInit(&result);

        hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                            &params, &result, NULL, NULL);

        SysFreeString(args[0].bstrVal);
        SysFreeString(args[1].bstrVal);

        if (FAILED(hr)) return L"";

        std::wstring text;
        if (result.vt == VT_BSTR && result.bstrVal) {
            text = result.bstrVal;
            SysFreeString(result.bstrVal);
        }
        return text;
    }

    // 3. 이름 지정 파라미터로 호출
    VARIANT args[2];
    VariantInit(&args[0]);
    VariantInit(&args[1]);

    // 이름 지정 시 순서는 namedDispids 순서와 일치해야 함
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(format.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(option.c_str());

    DISPPARAMS params = { args, namedDispids, 2, 2 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    SysFreeString(args[0].bstrVal);
    SysFreeString(args[1].bstrVal);

    if (FAILED(hr)) return L"";

    std::wstring text;
    if (result.vt == VT_BSTR && result.bstrVal) {
        text = result.bstrVal;
        SysFreeString(result.bstrVal);
    }
    return text;
}

int HwpWrapper::SetTextFile(const std::wstring& data,
                             const std::wstring& format,
                             const std::wstring& option)
{
    if (!m_pHwp) return 0;

    // COM 이름 지정 파라미터 호출:
    // hwp.SetTextFile(data=data, Format=format, option=option)

    HRESULT hr;
    DISPID dispid;

    // 1. SetTextFile 메서드의 DISPID 획득
    OLECHAR* methodName = const_cast<OLECHAR*>(L"SetTextFile");
    hr = m_pHwp->GetIDsOfNames(IID_NULL, &methodName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return 0;

    // 2. 위치 기반 파라미터로 호출 (data, Format, option 순서)
    VARIANT args[3];
    VariantInit(&args[0]);
    VariantInit(&args[1]);
    VariantInit(&args[2]);

    // 인자는 역순 (option, Format, data)
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(option.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(format.c_str());
    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(data.c_str());

    DISPPARAMS params = { args, NULL, 3, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    SysFreeString(args[0].bstrVal);
    SysFreeString(args[1].bstrVal);
    SysFreeString(args[2].bstrVal);

    if (FAILED(hr)) return 0;

    // 반환값 처리 (성공=1, 실패=0)
    if (result.vt == VT_I4) return result.lVal;
    if (result.vt == VT_BOOL) return result.boolVal != VARIANT_FALSE ? 1 : 0;
    return 1;  // 성공으로 간주
}

bool HwpWrapper::OpenPdf(const std::wstring& pdfPath, int thisWindow)
{
    if (!m_pHwp) return false;

    HRESULT hr;
    DISPID dispid;
    VARIANT result;
    VariantInit(&result);

    // 1. HParameterSet 속성 가져오기
    OLECHAR* psetName = const_cast<OLECHAR*>(L"HParameterSet");
    hr = m_pHwp->GetIDsOfNames(IID_NULL, &psetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                        &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pHParameterSet = result.pdispVal;

    // 2. HFileOpenSave 속성 가져오기
    OLECHAR* fileOpenSaveName = const_cast<OLECHAR*>(L"HFileOpenSave");
    hr = pHParameterSet->GetIDsOfNames(IID_NULL, &fileOpenSaveName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHParameterSet->Release();
        return false;
    }

    VariantInit(&result);
    hr = pHParameterSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                 &noParams, &result, NULL, NULL);
    pHParameterSet->Release();
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pHFileOpenSave = result.pdispVal;

    // 3. HSet 속성 가져오기
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pHFileOpenSave->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHFileOpenSave->Release();
        return false;
    }

    VariantInit(&result);
    hr = pHFileOpenSave->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                 &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        pHFileOpenSave->Release();
        return false;
    }

    IDispatch* pHSet = result.pdispVal;

    // 4. HAction 가져오기
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pHSet->Release();
        pHFileOpenSave->Release();
        return false;
    }

    // 5. CallPDFConverter 먼저 실행
    OLECHAR* runName = const_cast<OLECHAR*>(L"Run");
    hr = pHAction->GetIDsOfNames(IID_NULL, &runName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT runArgs[2];
        VariantInit(&runArgs[0]);
        VariantInit(&runArgs[1]);
        runArgs[0].vt = VT_BSTR;
        runArgs[0].bstrVal = SysAllocString(L"");
        runArgs[1].vt = VT_BSTR;
        runArgs[1].bstrVal = SysAllocString(L"CallPDFConverter");

        DISPPARAMS runParams = { runArgs, NULL, 2, 0 };
        VariantInit(&result);
        pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &runParams, &result, NULL, NULL);
        SysFreeString(runArgs[0].bstrVal);
        SysFreeString(runArgs[1].bstrVal);
        VariantClear(&result);
    }

    // 6. HAction.GetDefault("FileOpenPDF", HSet) 호출
    OLECHAR* getDefaultName = const_cast<OLECHAR*>(L"GetDefault");
    hr = pHAction->GetIDsOfNames(IID_NULL, &getDefaultName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT getDefaultArgs[2];
        VariantInit(&getDefaultArgs[0]);
        VariantInit(&getDefaultArgs[1]);
        getDefaultArgs[0].vt = VT_DISPATCH;
        getDefaultArgs[0].pdispVal = pHSet;
        getDefaultArgs[1].vt = VT_BSTR;
        getDefaultArgs[1].bstrVal = SysAllocString(L"FileOpenPDF");

        DISPPARAMS getDefaultParams = { getDefaultArgs, NULL, 2, 0 };
        VariantInit(&result);
        pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &getDefaultParams, &result, NULL, NULL);
        SysFreeString(getDefaultArgs[1].bstrVal);
        VariantClear(&result);
    }

    // 7. filename 속성 설정
    OLECHAR* filenamePropName = const_cast<OLECHAR*>(L"filename");
    hr = pHFileOpenSave->GetIDsOfNames(IID_NULL, &filenamePropName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT filenameVal;
        VariantInit(&filenameVal);
        filenameVal.vt = VT_BSTR;
        filenameVal.bstrVal = SysAllocString(pdfPath.c_str());

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS filenameParams = { &filenameVal, &putid, 1, 1 };
        pHFileOpenSave->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                                &filenameParams, NULL, NULL, NULL);
        SysFreeString(filenameVal.bstrVal);
    }

    // 8. OpenFlag 속성 설정
    OLECHAR* openFlagName = const_cast<OLECHAR*>(L"OpenFlag");
    hr = pHFileOpenSave->GetIDsOfNames(IID_NULL, &openFlagName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT flagVal;
        VariantInit(&flagVal);
        flagVal.vt = VT_I4;
        flagVal.lVal = thisWindow;

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS flagParams = { &flagVal, &putid, 1, 1 };
        pHFileOpenSave->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                                &flagParams, NULL, NULL, NULL);
    }

    // 9. HAction.Execute("FileOpenPDF", HSet) 호출
    OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
    hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1, LOCALE_USER_DEFAULT, &dispid);
    bool success = false;
    if (SUCCEEDED(hr)) {
        VARIANT executeArgs[2];
        VariantInit(&executeArgs[0]);
        VariantInit(&executeArgs[1]);
        executeArgs[0].vt = VT_DISPATCH;
        executeArgs[0].pdispVal = pHSet;
        executeArgs[1].vt = VT_BSTR;
        executeArgs[1].bstrVal = SysAllocString(L"FileOpenPDF");

        DISPPARAMS executeParams = { executeArgs, NULL, 2, 0 };
        VariantInit(&result);
        hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                              &executeParams, &result, NULL, NULL);
        SysFreeString(executeArgs[1].bstrVal);

        if (SUCCEEDED(hr) && result.vt == VT_BOOL) {
            success = (result.boolVal != VARIANT_FALSE);
        }
        VariantClear(&result);
    }

    pHSet->Release();
    pHFileOpenSave->Release();
    pHAction->Release();
    return success;
}

bool HwpWrapper::SaveBlockAs(const std::wstring& path,
                              const std::wstring& format,
                              int attributes)
{
    if (!m_pHwp) return false;

    // 선택 모드 확인
    VARIANT selModeVar = GetProperty(L"SelectionMode");
    if (selModeVar.vt == VT_I4 && selModeVar.lVal == 0) {
        return false;  // 선택된 블록 없음
    }

    HRESULT hr;
    DISPID dispid;
    VARIANT result;
    VariantInit(&result);

    // 1. HParameterSet 속성 가져오기
    OLECHAR* psetName = const_cast<OLECHAR*>(L"HParameterSet");
    hr = m_pHwp->GetIDsOfNames(IID_NULL, &psetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                        &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pHParameterSet = result.pdispVal;

    // 2. HFileOpenSave 속성 가져오기
    OLECHAR* fileOpenSaveName = const_cast<OLECHAR*>(L"HFileOpenSave");
    hr = pHParameterSet->GetIDsOfNames(IID_NULL, &fileOpenSaveName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHParameterSet->Release();
        return false;
    }

    VariantInit(&result);
    hr = pHParameterSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                 &noParams, &result, NULL, NULL);
    pHParameterSet->Release();
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pHFileOpenSave = result.pdispVal;

    // 3. HSet 속성 가져오기
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pHFileOpenSave->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHFileOpenSave->Release();
        return false;
    }

    VariantInit(&result);
    hr = pHFileOpenSave->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                 &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        pHFileOpenSave->Release();
        return false;
    }

    IDispatch* pHSet = result.pdispVal;

    // 4. HAction 가져오기
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pHSet->Release();
        pHFileOpenSave->Release();
        return false;
    }

    // 5. HAction.GetDefault("FileSaveBlock_S", HSet) 호출
    OLECHAR* getDefaultName = const_cast<OLECHAR*>(L"GetDefault");
    hr = pHAction->GetIDsOfNames(IID_NULL, &getDefaultName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT getDefaultArgs[2];
        VariantInit(&getDefaultArgs[0]);
        VariantInit(&getDefaultArgs[1]);
        getDefaultArgs[0].vt = VT_DISPATCH;
        getDefaultArgs[0].pdispVal = pHSet;
        getDefaultArgs[1].vt = VT_BSTR;
        getDefaultArgs[1].bstrVal = SysAllocString(L"FileSaveBlock_S");

        DISPPARAMS getDefaultParams = { getDefaultArgs, NULL, 2, 0 };
        VariantInit(&result);
        pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &getDefaultParams, &result, NULL, NULL);
        SysFreeString(getDefaultArgs[1].bstrVal);
        VariantClear(&result);
    }

    // 6. filename 속성 설정
    OLECHAR* filenamePropName = const_cast<OLECHAR*>(L"filename");
    hr = pHFileOpenSave->GetIDsOfNames(IID_NULL, &filenamePropName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT filenameVal;
        VariantInit(&filenameVal);
        filenameVal.vt = VT_BSTR;
        filenameVal.bstrVal = SysAllocString(path.c_str());

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS filenameParams = { &filenameVal, &putid, 1, 1 };
        pHFileOpenSave->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                                &filenameParams, NULL, NULL, NULL);
        SysFreeString(filenameVal.bstrVal);
    }

    // 7. Format 속성 설정
    OLECHAR* formatPropName = const_cast<OLECHAR*>(L"Format");
    hr = pHFileOpenSave->GetIDsOfNames(IID_NULL, &formatPropName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT formatVal;
        VariantInit(&formatVal);
        formatVal.vt = VT_BSTR;
        formatVal.bstrVal = SysAllocString(format.c_str());

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS formatParams = { &formatVal, &putid, 1, 1 };
        pHFileOpenSave->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                                &formatParams, NULL, NULL, NULL);
        SysFreeString(formatVal.bstrVal);
    }

    // 8. Attributes 속성 설정
    OLECHAR* attrPropName = const_cast<OLECHAR*>(L"Attributes");
    hr = pHFileOpenSave->GetIDsOfNames(IID_NULL, &attrPropName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT attrVal;
        VariantInit(&attrVal);
        attrVal.vt = VT_I4;
        attrVal.lVal = attributes;

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS attrParams = { &attrVal, &putid, 1, 1 };
        pHFileOpenSave->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                                &attrParams, NULL, NULL, NULL);
    }

    // 9. HAction.Execute("FileSaveBlock_S", HSet) 호출
    OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
    hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1, LOCALE_USER_DEFAULT, &dispid);
    bool success = false;
    if (SUCCEEDED(hr)) {
        VARIANT executeArgs[2];
        VariantInit(&executeArgs[0]);
        VariantInit(&executeArgs[1]);
        executeArgs[0].vt = VT_DISPATCH;
        executeArgs[0].pdispVal = pHSet;
        executeArgs[1].vt = VT_BSTR;
        executeArgs[1].bstrVal = SysAllocString(L"FileSaveBlock_S");

        DISPPARAMS executeParams = { executeArgs, NULL, 2, 0 };
        VariantInit(&result);
        hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                              &executeParams, &result, NULL, NULL);
        SysFreeString(executeArgs[1].bstrVal);

        if (SUCCEEDED(hr) && result.vt == VT_BOOL) {
            success = (result.boolVal != VARIANT_FALSE);
        }
        VariantClear(&result);
    }

    pHSet->Release();
    pHFileOpenSave->Release();
    pHAction->Release();
    return success;
}

std::map<std::wstring, std::wstring> HwpWrapper::GetFileInfo(const std::wstring& filename)
{
    std::map<std::wstring, std::wstring> result;
    if (!m_pHwp) return result;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"GetFileInfo");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return result;

    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_BSTR;
    arg.bstrVal = SysAllocString(filename.c_str());

    DISPPARAMS params = { &arg, NULL, 1, 0 };
    VARIANT retVal;
    VariantInit(&retVal);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &retVal, NULL, NULL);
    SysFreeString(arg.bstrVal);

    if (FAILED(hr) || retVal.vt != VT_DISPATCH || !retVal.pdispVal) {
        VariantClear(&retVal);
        return result;
    }

    IDispatch* pFileInfo = retVal.pdispVal;

    // Item 메서드로 각 속성 추출
    auto getItem = [&](const wchar_t* itemName) -> std::wstring {
        DISPID dispidItem;
        OLECHAR* itemMethod = const_cast<OLECHAR*>(L"Item");
        hr = pFileInfo->GetIDsOfNames(IID_NULL, &itemMethod, 1, LOCALE_USER_DEFAULT, &dispidItem);
        if (FAILED(hr)) return L"";

        VARIANT vItemName;
        VariantInit(&vItemName);
        vItemName.vt = VT_BSTR;
        vItemName.bstrVal = SysAllocString(itemName);

        DISPPARAMS itemParams = { &vItemName, NULL, 1, 0 };
        VARIANT itemResult;
        VariantInit(&itemResult);

        hr = pFileInfo->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_METHOD | DISPATCH_PROPERTYGET, &itemParams, &itemResult, NULL, NULL);
        SysFreeString(vItemName.bstrVal);

        std::wstring value;
        if (SUCCEEDED(hr)) {
            if (itemResult.vt == VT_BSTR && itemResult.bstrVal) {
                value = itemResult.bstrVal;
            } else if (itemResult.vt == VT_I4) {
                value = std::to_wstring(itemResult.lVal);
            } else if (itemResult.vt == VT_UI4) {
                wchar_t buf[32];
                swprintf_s(buf, L"0x%08X", itemResult.ulVal);
                value = buf;
            }
        }
        VariantClear(&itemResult);
        return value;
    };

    result[L"Format"] = getItem(L"Format");
    result[L"VersionStr"] = getItem(L"VersionStr");
    result[L"VersionNum"] = getItem(L"VersionNum");
    result[L"Encrypted"] = getItem(L"Encrypted");

    pFileInfo->Release();
    return result;
}

//=============================================================================
// 텍스트 편집
//=============================================================================

bool HwpWrapper::InsertText(const std::wstring& text)
{
    if (!m_pHwp) return false;

    // pyhwpx 방식: HParameterSet.HInsertText + HAction.Execute 사용
    // 이 방식이 직접 InsertText 호출보다 더 안정적임

    HRESULT hr;
    DISPID dispid;
    VARIANT result;
    VariantInit(&result);

    // 1. HParameterSet 속성 가져오기
    OLECHAR* psetName = const_cast<OLECHAR*>(L"HParameterSet");
    hr = m_pHwp->GetIDsOfNames(IID_NULL, &psetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                        &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pHParameterSet = result.pdispVal;

    // 2. HInsertText 속성 가져오기
    OLECHAR* insertTextName = const_cast<OLECHAR*>(L"HInsertText");
    hr = pHParameterSet->GetIDsOfNames(IID_NULL, &insertTextName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHParameterSet->Release();
        return false;
    }

    VariantInit(&result);
    hr = pHParameterSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                 &noParams, &result, NULL, NULL);
    pHParameterSet->Release();
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pHInsertText = result.pdispVal;

    // 3. HSet 속성 가져오기
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pHInsertText->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHInsertText->Release();
        return false;
    }

    VariantInit(&result);
    hr = pHInsertText->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                               &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        pHInsertText->Release();
        return false;
    }

    IDispatch* pHSet = result.pdispVal;

    // 4. HAction 가져오기
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pHSet->Release();
        pHInsertText->Release();
        return false;
    }

    // 5. HAction.GetDefault("InsertText", HSet) 호출
    OLECHAR* getDefaultName = const_cast<OLECHAR*>(L"GetDefault");
    hr = pHAction->GetIDsOfNames(IID_NULL, &getDefaultName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHSet->Release();
        pHInsertText->Release();
        return false;
    }

    VARIANT getDefaultArgs[2];
    VariantInit(&getDefaultArgs[0]);
    VariantInit(&getDefaultArgs[1]);
    getDefaultArgs[0].vt = VT_DISPATCH;
    getDefaultArgs[0].pdispVal = pHSet;
    getDefaultArgs[1].vt = VT_BSTR;
    getDefaultArgs[1].bstrVal = SysAllocString(L"InsertText");

    DISPPARAMS getDefaultParams = { getDefaultArgs, NULL, 2, 0 };
    VariantInit(&result);
    hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &getDefaultParams, &result, NULL, NULL);
    SysFreeString(getDefaultArgs[1].bstrVal);

    // 6. HInsertText.Text = text 설정
    OLECHAR* textPropName = const_cast<OLECHAR*>(L"Text");
    hr = pHInsertText->GetIDsOfNames(IID_NULL, &textPropName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHSet->Release();
        pHInsertText->Release();
        return false;
    }

    VARIANT textVal;
    VariantInit(&textVal);
    textVal.vt = VT_BSTR;
    textVal.bstrVal = SysAllocString(text.c_str());

    DISPID putid = DISPID_PROPERTYPUT;
    DISPPARAMS textParams = { &textVal, &putid, 1, 1 };
    hr = pHInsertText->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &textParams, NULL, NULL, NULL);
    SysFreeString(textVal.bstrVal);

    if (FAILED(hr)) {
        pHSet->Release();
        pHInsertText->Release();
        return false;
    }

    // 7. HAction.Execute("InsertText", HSet) 호출
    OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
    hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHSet->Release();
        pHInsertText->Release();
        return false;
    }

    VARIANT executeArgs[2];
    VariantInit(&executeArgs[0]);
    VariantInit(&executeArgs[1]);
    executeArgs[0].vt = VT_DISPATCH;
    executeArgs[0].pdispVal = pHSet;
    executeArgs[1].vt = VT_BSTR;
    executeArgs[1].bstrVal = SysAllocString(L"InsertText");

    DISPPARAMS executeParams = { executeArgs, NULL, 2, 0 };
    VariantInit(&result);
    hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &executeParams, &result, NULL, NULL);
    SysFreeString(executeArgs[1].bstrVal);

    pHSet->Release();
    pHInsertText->Release();

    if (FAILED(hr)) return false;
    return (result.vt == VT_BOOL) ? (result.boolVal != VARIANT_FALSE) : true;
}

std::tuple<int, std::wstring> HwpWrapper::GetText()
{
    if (!m_pHwp) return std::make_tuple(-1, L"");

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"GetText");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return std::make_tuple(-1, L"");

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    if (FAILED(hr)) return std::make_tuple(-1, L"");

    // GetText는 (상태, 텍스트) 튜플 반환
    // COM에서는 SAFEARRAY로 반환될 수 있음
    if (result.vt == VT_BSTR) {
        std::wstring text(result.bstrVal ? result.bstrVal : L"");
        SysFreeString(result.bstrVal);
        return std::make_tuple(0, text);
    }

    return std::make_tuple(-1, L"");
}

std::wstring HwpWrapper::GetSelectedText(bool keep_select)
{
    // 현재 선택 영역의 텍스트 반환
    // HAction.GetSelectedText() 호출
    if (!m_pHwp) return L"";

    // TODO: 구현
    return L"";
}

//=============================================================================
// 위치 관리
//=============================================================================

HwpPos HwpWrapper::GetPos()
{
    HwpPos pos = { 0, 0, 0 };
    if (!m_pHwp) return pos;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"GetPos");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return pos;

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    if (SUCCEEDED(hr) && result.vt == (VT_ARRAY | VT_VARIANT)) {
        SAFEARRAY* psa = result.parray;
        VARIANT* pData;
        SafeArrayAccessData(psa, (void**)&pData);

        if (pData[0].vt == VT_I4) pos.list = pData[0].lVal;
        if (pData[1].vt == VT_I4) pos.para = pData[1].lVal;
        if (pData[2].vt == VT_I4) pos.pos = pData[2].lVal;

        SafeArrayUnaccessData(psa);
    }

    VariantClear(&result);
    return pos;
}

bool HwpWrapper::SetPos(int list, int para, int pos)
{
    if (!m_pHwp) return false;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"SetPos");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    VARIANT args[3];
    VariantInit(&args[0]);
    VariantInit(&args[1]);
    VariantInit(&args[2]);

    args[0].vt = VT_I4;
    args[0].lVal = pos;
    args[1].vt = VT_I4;
    args[1].lVal = para;
    args[2].vt = VT_I4;
    args[2].lVal = list;

    DISPPARAMS params = { args, NULL, 3, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    if (FAILED(hr)) return false;
    return true;
}

bool HwpWrapper::MovePos(int move_id, int para, int pos)
{
    if (!m_pHwp) return false;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"MovePos");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    VARIANT args[3];
    VariantInit(&args[0]);
    VariantInit(&args[1]);
    VariantInit(&args[2]);

    args[0].vt = VT_I4;
    args[0].lVal = pos;
    args[1].vt = VT_I4;
    args[1].lVal = para;
    args[2].vt = VT_I4;
    args[2].lVal = move_id;

    DISPPARAMS params = { args, NULL, 3, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    if (FAILED(hr)) return false;
    return (result.vt == VT_BOOL) ? (result.boolVal != VARIANT_FALSE) : true;
}

bool HwpWrapper::InitScan(int option, int range, int spara, int spos, int epara, int epos)
{
    if (!m_pHwp) return false;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"InitScan");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    // 파라미터 순서: epos, epara, spos, spara, range, option (역순)
    VARIANT args[6];
    for (int i = 0; i < 6; i++) VariantInit(&args[i]);

    args[0].vt = VT_I4; args[0].lVal = epos;
    args[1].vt = VT_I4; args[1].lVal = epara;
    args[2].vt = VT_I4; args[2].lVal = spos;
    args[3].vt = VT_I4; args[3].lVal = spara;
    args[4].vt = VT_I4; args[4].lVal = range;
    args[5].vt = VT_I4; args[5].lVal = option;

    DISPPARAMS params = { args, NULL, 6, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    if (FAILED(hr)) return false;
    return (result.vt == VT_BOOL) ? (result.boolVal != VARIANT_FALSE) : true;
}

void HwpWrapper::ReleaseScan()
{
    if (!m_pHwp) return;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"ReleaseScan");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return;

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                   &params, &result, NULL, NULL);
    VariantClear(&result);
}

bool HwpWrapper::SelectText(int spara, int spos, int epara, int epos, int slist)
{
    if (!m_pHwp) return false;

    // 먼저 set_pos로 리스트 설정
    SetPos(slist, 0, 0);

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"SelectText");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    // 파라미터 순서: epos, epara, spos, spara (역순)
    VARIANT args[4];
    for (int i = 0; i < 4; i++) VariantInit(&args[i]);

    args[0].vt = VT_I4; args[0].lVal = epos;
    args[1].vt = VT_I4; args[1].lVal = epara;
    args[2].vt = VT_I4; args[2].lVal = spos;
    args[3].vt = VT_I4; args[3].lVal = spara;

    DISPPARAMS params = { args, NULL, 4, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    if (FAILED(hr)) return false;
    return (result.vt == VT_BOOL) ? (result.boolVal != VARIANT_FALSE) : true;
}

bool HwpWrapper::SelectTextByGetPos(const HwpPos& s_pos, const HwpPos& e_pos)
{
    // set_pos로 리스트 설정 후 SelectText 호출
    SetPos(s_pos.list, 0, 0);
    return SelectText(s_pos.para, s_pos.pos, e_pos.para, e_pos.pos, s_pos.list);
}

IDispatch* HwpWrapper::GetPosBySet()
{
    if (!m_pHwp) return nullptr;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"GetPosBySet");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return nullptr;

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        VariantClear(&result);
        return nullptr;
    }

    return result.pdispVal;
}

bool HwpWrapper::SetPosBySet(IDispatch* pDispVal)
{
    if (!m_pHwp || !pDispVal) return false;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"SetPosBySet");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_DISPATCH;
    arg.pdispVal = pDispVal;

    DISPPARAMS params = { &arg, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    VariantClear(&result);
    return SUCCEEDED(hr);
}

int HwpWrapper::GetPosBySetPy()
{
    IDispatch* pPos = GetPosBySet();
    if (!pPos) return -1;

    m_posCache.push_back(pPos);
    return static_cast<int>(m_posCache.size() - 1);
}

bool HwpWrapper::SetPosBySetPy(int idx)
{
    if (idx < 0 || idx >= static_cast<int>(m_posCache.size())) return false;
    return SetPosBySet(m_posCache[idx]);
}

void HwpWrapper::ClearPosCache()
{
    for (auto p : m_posCache) {
        if (p) p->Release();
    }
    m_posCache.clear();
}

//=============================================================================
// 창/UI 관리
//=============================================================================

void HwpWrapper::SetVisible(bool visible)
{
    if (!m_pHwp) return;

    // pyhwpx 방식: XHwpWindows.Active_XHwpWindow.Visible = visible

    // 1. XHwpWindows 속성 획득
    VARIANT xhwpWindowsVar = GetProperty(L"XHwpWindows");
    if (xhwpWindowsVar.vt != VT_DISPATCH || !xhwpWindowsVar.pdispVal) {
        VariantClear(&xhwpWindowsVar);
        return;
    }

    IDispatch* pXHwpWindows = xhwpWindowsVar.pdispVal;

    // 2. Active_XHwpWindow 속성 획득 (또는 Active)
    DISPID dispidActive;
    OLECHAR* activeName = const_cast<OLECHAR*>(L"Active_XHwpWindow");
    HRESULT hr = pXHwpWindows->GetIDsOfNames(IID_NULL, &activeName, 1,
                                              LOCALE_USER_DEFAULT, &dispidActive);
    if (FAILED(hr)) {
        // 대체: "Active" 시도
        activeName = const_cast<OLECHAR*>(L"Active");
        hr = pXHwpWindows->GetIDsOfNames(IID_NULL, &activeName, 1,
                                          LOCALE_USER_DEFAULT, &dispidActive);
    }
    if (FAILED(hr)) {
        pXHwpWindows->Release();
        return;
    }

    DISPPARAMS paramsEmpty = { NULL, NULL, 0, 0 };
    VARIANT activeWindowVar;
    VariantInit(&activeWindowVar);
    hr = pXHwpWindows->Invoke(dispidActive, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYGET, &paramsEmpty, &activeWindowVar, NULL, NULL);

    if (FAILED(hr) || activeWindowVar.vt != VT_DISPATCH || !activeWindowVar.pdispVal) {
        VariantClear(&activeWindowVar);
        pXHwpWindows->Release();
        return;
    }

    IDispatch* pActiveWindow = activeWindowVar.pdispVal;

    // 3. Visible 속성 설정
    DISPID dispidVisible;
    OLECHAR* visibleName = const_cast<OLECHAR*>(L"Visible");
    hr = pActiveWindow->GetIDsOfNames(IID_NULL, &visibleName, 1,
                                       LOCALE_USER_DEFAULT, &dispidVisible);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_BOOL;
        val.boolVal = visible ? VARIANT_TRUE : VARIANT_FALSE;

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pActiveWindow->Invoke(dispidVisible, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYPUT, &params, NULL, NULL, NULL);
    }

    m_bVisible = visible;

    // 4. visible=true일 때 추가 처리: WindowHandle로 ShowWindow 호출
    if (visible) {
        HWND hwnd = GetActiveWindowHandle(pActiveWindow);
        if (hwnd) {
            ShowWindow(hwnd, SW_SHOW);
            ShowWindow(hwnd, SW_MAXIMIZE);  // SW_MAXIMIZE = 3
        }
    }

    pActiveWindow->Release();
    pXHwpWindows->Release();
}

void HwpWrapper::MaximizeWindow()
{
    // HAction.Run("WindowMaximize")
    RunAction(L"WindowMaximize");
}

void HwpWrapper::MinimizeWindow()
{
    // HAction.Run("WindowMinimize")
    RunAction(L"WindowMinimize");
}

HWND HwpWrapper::GetActiveWindowHandle(IDispatch* pActiveWindow)
{
    if (!pActiveWindow) return NULL;

    // WindowHandle 속성 조회
    DISPID dispidHandle;
    OLECHAR* handleName = const_cast<OLECHAR*>(L"WindowHandle");
    HRESULT hr = pActiveWindow->GetIDsOfNames(IID_NULL, &handleName, 1,
                                               LOCALE_USER_DEFAULT, &dispidHandle);
    if (FAILED(hr)) return NULL;

    DISPPARAMS paramsEmpty = { NULL, NULL, 0, 0 };
    VARIANT handleResult;
    VariantInit(&handleResult);
    hr = pActiveWindow->Invoke(dispidHandle, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYGET, &paramsEmpty, &handleResult, NULL, NULL);

    if (FAILED(hr)) return NULL;

    HWND hwnd = NULL;
    if (handleResult.vt == VT_I4) {
        hwnd = (HWND)(intptr_t)handleResult.lVal;
    } else if (handleResult.vt == VT_I8) {
        hwnd = (HWND)(intptr_t)handleResult.llVal;
    } else if (handleResult.vt == VT_INT || handleResult.vt == VT_UINT) {
        hwnd = (HWND)(intptr_t)handleResult.intVal;
    }

    VariantClear(&handleResult);
    return hwnd;
}

HWND HwpWrapper::GetHwnd()
{
    if (!m_pHwp) return NULL;

    // XHwpWindows.Active_XHwpWindow.WindowHandle 방식 (pyhwpx와 동일)

    // 1. XHwpWindows 속성 획득
    VARIANT xhwpWindowsVar = GetProperty(L"XHwpWindows");
    if (xhwpWindowsVar.vt != VT_DISPATCH || !xhwpWindowsVar.pdispVal) {
        VariantClear(&xhwpWindowsVar);
        return NULL;
    }

    IDispatch* pXHwpWindows = xhwpWindowsVar.pdispVal;

    // 2. Active_XHwpWindow 속성 획득
    DISPID dispidActive;
    OLECHAR* activeName = const_cast<OLECHAR*>(L"Active_XHwpWindow");
    HRESULT hr = pXHwpWindows->GetIDsOfNames(IID_NULL, &activeName, 1,
                                              LOCALE_USER_DEFAULT, &dispidActive);
    if (FAILED(hr)) {
        // 대체: "Active" 시도
        activeName = const_cast<OLECHAR*>(L"Active");
        hr = pXHwpWindows->GetIDsOfNames(IID_NULL, &activeName, 1,
                                          LOCALE_USER_DEFAULT, &dispidActive);
    }
    if (FAILED(hr)) {
        pXHwpWindows->Release();
        return NULL;
    }

    DISPPARAMS paramsEmpty = { NULL, NULL, 0, 0 };
    VARIANT activeWindowVar;
    VariantInit(&activeWindowVar);
    hr = pXHwpWindows->Invoke(dispidActive, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYGET, &paramsEmpty, &activeWindowVar, NULL, NULL);

    if (FAILED(hr) || activeWindowVar.vt != VT_DISPATCH || !activeWindowVar.pdispVal) {
        VariantClear(&activeWindowVar);
        pXHwpWindows->Release();
        return NULL;
    }

    IDispatch* pActiveWindow = activeWindowVar.pdispVal;

    // 3. WindowHandle 획득
    HWND hwnd = GetActiveWindowHandle(pActiveWindow);

    pActiveWindow->Release();
    pXHwpWindows->Release();

    return hwnd;
}

bool HwpWrapper::ActivateOleObject()
{
    if (!m_pHwp) return false;

    IOleObject* pOleObject = nullptr;
    HRESULT hr = m_pHwp->QueryInterface(IID_IOleObject, (void**)&pOleObject);
    if (SUCCEEDED(hr) && pOleObject) {
        // OLEIVERB_SHOW = -1: 객체를 보이도록 함
        hr = pOleObject->DoVerb(OLEIVERB_SHOW, NULL, NULL, 0, NULL, NULL);
        pOleObject->Release();
        return SUCCEEDED(hr);
    }
    return false;
}

void HwpWrapper::ShowWindowWithForeground(HWND hwnd)
{
    if (!hwnd || !IsWindow(hwnd)) return;

    // 1. 창이 최소화되어 있으면 복원
    if (IsIconic(hwnd)) {
        ShowWindow(hwnd, SW_RESTORE);
    } else {
        ShowWindow(hwnd, SW_SHOW);
    }

    // 2. 포그라운드 스레드 연결 (백그라운드 프로세스에서 포그라운드 권한 획득)
    HWND foregroundHwnd = GetForegroundWindow();
    if (foregroundHwnd) {
        DWORD foregroundThreadId = GetWindowThreadProcessId(foregroundHwnd, NULL);
        DWORD currentThreadId = GetCurrentThreadId();

        if (foregroundThreadId != currentThreadId) {
            AttachThreadInput(currentThreadId, foregroundThreadId, TRUE);
        }

        // 3. 포그라운드로 가져오기
        SetForegroundWindow(hwnd);
        BringWindowToTop(hwnd);

        if (foregroundThreadId != currentThreadId) {
            AttachThreadInput(currentThreadId, foregroundThreadId, FALSE);
        }
    } else {
        // 포그라운드 창이 없으면 직접 시도
        SetForegroundWindow(hwnd);
        BringWindowToTop(hwnd);
    }

    // 4. 창 활성화
    SetActiveWindow(hwnd);

    // 5. Z-order 최상위로 이동
    SetWindowPos(hwnd, HWND_TOP, 0, 0, 0, 0,
                 SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
}

bool HwpWrapper::SetViewState(int flag)
{
    if (!m_pHwp) return false;

    VARIANT val;
    VariantInit(&val);
    val.vt = VT_I4;
    val.lVal = flag;

    return SetProperty(L"ViewProperties", val);
}

int HwpWrapper::GetViewState()
{
    VARIANT result = GetProperty(L"ViewProperties");
    if (result.vt == VT_I4) {
        return result.lVal;
    }
    return 0;
}

//=============================================================================
// 메시지 박스
//=============================================================================

int HwpWrapper::MsgBox(const std::wstring& message, int flag)
{
    if (!m_pHwp) return -1;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"MsgBox");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return -1;

    VARIANT args[2];
    VariantInit(&args[0]);
    VariantInit(&args[1]);

    args[0].vt = VT_I4;
    args[0].lVal = flag;
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(message.c_str());

    DISPPARAMS params = { args, NULL, 2, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    SysFreeString(args[1].bstrVal);

    if (FAILED(hr)) return -1;
    return (result.vt == VT_I4) ? result.lVal : -1;
}

int HwpWrapper::GetMessageBoxMode()
{
    VARIANT result = GetProperty(L"MessageBoxMode");
    if (result.vt == VT_I4) {
        return result.lVal;
    }
    return 0;
}

int HwpWrapper::SetMessageBoxMode(int mode)
{
    int old_mode = GetMessageBoxMode();

    VARIANT val;
    VariantInit(&val);
    val.vt = VT_I4;
    val.lVal = mode;

    SetProperty(L"MessageBoxMode", val);
    return old_mode;
}

//=============================================================================
// 문서 상태
//=============================================================================

bool HwpWrapper::IsEmpty()
{
    VARIANT result = GetProperty(L"IsEmpty");
    if (result.vt == VT_BOOL) {
        return result.boolVal != VARIANT_FALSE;
    }
    return true;
}

bool HwpWrapper::IsModified()
{
    VARIANT result = GetProperty(L"IsModified");
    if (result.vt == VT_BOOL) {
        return result.boolVal != VARIANT_FALSE;
    }
    return false;
}

bool HwpWrapper::IsCell()
{
    // 현재 커서가 표 셀 안에 있는지 확인
    VARIANT result = GetProperty(L"ParentCtrl");
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        // ParentCtrl이 존재하면 셀 내부
        result.pdispVal->Release();
        return true;
    }
    return false;
}

//=============================================================================
// 찾기/바꾸기
//=============================================================================

bool HwpWrapper::Find(const std::wstring& text,
                       bool forward,
                       bool match_case,
                       bool regex,
                       bool replace_mode)
{
    if (!m_pHwp || text.empty()) return false;

    HRESULT hr;
    DISPID dispid;
    VARIANT result;
    VariantInit(&result);

    // 1. HParameterSet 속성 가져오기
    IDispatch* pHParameterSet = GetHParameterSet();
    if (!pHParameterSet) return false;

    // 2. HFindReplace 속성 가져오기
    OLECHAR* findReplaceName = const_cast<OLECHAR*>(L"HFindReplace");
    hr = pHParameterSet->GetIDsOfNames(IID_NULL, &findReplaceName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    VariantInit(&result);
    hr = pHParameterSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                 &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pHFindReplace = result.pdispVal;

    // 3. HSet 속성 가져오기
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHFindReplace->Release();
        return false;
    }

    VariantInit(&result);
    hr = pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        pHFindReplace->Release();
        return false;
    }

    IDispatch* pHSet = result.pdispVal;

    // 4. HAction 가져오기
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pHSet->Release();
        pHFindReplace->Release();
        return false;
    }

    // 5. HAction.GetDefault("FindDlg", HSet) 호출하여 초기화
    OLECHAR* getDefaultName = const_cast<OLECHAR*>(L"GetDefault");
    hr = pHAction->GetIDsOfNames(IID_NULL, &getDefaultName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT getDefaultArgs[2];
        VariantInit(&getDefaultArgs[0]);
        VariantInit(&getDefaultArgs[1]);
        getDefaultArgs[0].vt = VT_DISPATCH;
        getDefaultArgs[0].pdispVal = pHSet;
        getDefaultArgs[1].vt = VT_BSTR;
        getDefaultArgs[1].bstrVal = SysAllocString(L"FindDlg");

        DISPPARAMS getDefaultParams = { getDefaultArgs, NULL, 2, 0 };
        VariantInit(&result);
        pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                         &getDefaultParams, &result, NULL, NULL);
        SysFreeString(getDefaultArgs[1].bstrVal);
    }

    // 6. 파라미터 설정
    DISPID putid = DISPID_PROPERTYPUT;

    // FindString 설정
    OLECHAR* findStringName = const_cast<OLECHAR*>(L"FindString");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &findStringName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_BSTR;
        val.bstrVal = SysAllocString(text.c_str());
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
        SysFreeString(val.bstrVal);
    }

    // MatchCase 설정
    OLECHAR* matchCaseName = const_cast<OLECHAR*>(L"MatchCase");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &matchCaseName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = match_case ? 1 : 0;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // FindRegExp 설정 (정규식)
    OLECHAR* findRegExpName = const_cast<OLECHAR*>(L"FindRegExp");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &findRegExpName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = regex ? 1 : 0;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // Direction 설정 (0=Forward, 1=Backward)
    OLECHAR* directionName = const_cast<OLECHAR*>(L"Direction");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &directionName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = forward ? 0 : 1;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // IgnoreMessage 설정 (메시지 무시)
    OLECHAR* ignoreMsgName = const_cast<OLECHAR*>(L"IgnoreMessage");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &ignoreMsgName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = 1;  // 메시지 무시
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // ReplaceMode 설정
    OLECHAR* replaceModeName = const_cast<OLECHAR*>(L"ReplaceMode");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &replaceModeName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = replace_mode ? 1 : 0;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // 7. HAction.Execute("RepeatFind", HSet) 호출
    OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
    hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHSet->Release();
        pHFindReplace->Release();
        return false;
    }

    VARIANT executeArgs[2];
    VariantInit(&executeArgs[0]);
    VariantInit(&executeArgs[1]);
    executeArgs[0].vt = VT_DISPATCH;
    executeArgs[0].pdispVal = pHSet;
    executeArgs[1].vt = VT_BSTR;
    executeArgs[1].bstrVal = SysAllocString(L"RepeatFind");

    DISPPARAMS executeParams = { executeArgs, NULL, 2, 0 };
    VariantInit(&result);
    hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &executeParams, &result, NULL, NULL);
    SysFreeString(executeArgs[1].bstrVal);

    pHSet->Release();
    pHFindReplace->Release();

    return SUCCEEDED(hr) && (result.vt == VT_BOOL ? result.boolVal != VARIANT_FALSE : true);
}

bool HwpWrapper::Replace(const std::wstring& find_text,
                          const std::wstring& replace_text,
                          bool forward,
                          bool match_case,
                          bool regex)
{
    if (!m_pHwp) return false;

    HRESULT hr;
    DISPID dispid;
    VARIANT result;
    VariantInit(&result);

    // 1. HParameterSet 속성 가져오기
    IDispatch* pHParameterSet = GetHParameterSet();
    if (!pHParameterSet) return false;

    // 2. HFindReplace 속성 가져오기
    OLECHAR* findReplaceName = const_cast<OLECHAR*>(L"HFindReplace");
    hr = pHParameterSet->GetIDsOfNames(IID_NULL, &findReplaceName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    VariantInit(&result);
    hr = pHParameterSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                 &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pHFindReplace = result.pdispVal;

    // 3. HSet 속성 가져오기
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHFindReplace->Release();
        return false;
    }

    VariantInit(&result);
    hr = pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        pHFindReplace->Release();
        return false;
    }

    IDispatch* pHSet = result.pdispVal;

    // 4. HAction 가져오기
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pHSet->Release();
        pHFindReplace->Release();
        return false;
    }

    // 5. HAction.GetDefault("FindDlg", HSet) 호출하여 초기화
    OLECHAR* getDefaultName = const_cast<OLECHAR*>(L"GetDefault");
    hr = pHAction->GetIDsOfNames(IID_NULL, &getDefaultName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT getDefaultArgs[2];
        VariantInit(&getDefaultArgs[0]);
        VariantInit(&getDefaultArgs[1]);
        getDefaultArgs[0].vt = VT_DISPATCH;
        getDefaultArgs[0].pdispVal = pHSet;
        getDefaultArgs[1].vt = VT_BSTR;
        getDefaultArgs[1].bstrVal = SysAllocString(L"FindDlg");

        DISPPARAMS getDefaultParams = { getDefaultArgs, NULL, 2, 0 };
        VariantInit(&result);
        pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                         &getDefaultParams, &result, NULL, NULL);
        SysFreeString(getDefaultArgs[1].bstrVal);
    }

    // 6. 파라미터 설정
    DISPID putid = DISPID_PROPERTYPUT;

    // FindString 설정
    OLECHAR* findStringName = const_cast<OLECHAR*>(L"FindString");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &findStringName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_BSTR;
        val.bstrVal = SysAllocString(find_text.c_str());
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
        SysFreeString(val.bstrVal);
    }

    // ReplaceString 설정
    OLECHAR* replaceStringName = const_cast<OLECHAR*>(L"ReplaceString");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &replaceStringName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_BSTR;
        val.bstrVal = SysAllocString(replace_text.c_str());
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
        SysFreeString(val.bstrVal);
    }

    // MatchCase 설정
    OLECHAR* matchCaseName = const_cast<OLECHAR*>(L"MatchCase");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &matchCaseName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = match_case ? 1 : 0;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // FindRegExp 설정 (정규식)
    OLECHAR* findRegExpName = const_cast<OLECHAR*>(L"FindRegExp");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &findRegExpName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = regex ? 1 : 0;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // Direction 설정 (0=Forward, 1=Backward)
    OLECHAR* directionName = const_cast<OLECHAR*>(L"Direction");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &directionName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = forward ? 0 : 1;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // IgnoreMessage 설정 (메시지 무시)
    OLECHAR* ignoreMsgName = const_cast<OLECHAR*>(L"IgnoreMessage");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &ignoreMsgName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = 1;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // ReplaceMode 설정
    OLECHAR* replaceModeName = const_cast<OLECHAR*>(L"ReplaceMode");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &replaceModeName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = 1;  // 바꾸기 모드
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // 7. HAction.Execute("ExecReplace", HSet) 호출
    OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
    hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHSet->Release();
        pHFindReplace->Release();
        return false;
    }

    VARIANT executeArgs[2];
    VariantInit(&executeArgs[0]);
    VariantInit(&executeArgs[1]);
    executeArgs[0].vt = VT_DISPATCH;
    executeArgs[0].pdispVal = pHSet;
    executeArgs[1].vt = VT_BSTR;
    executeArgs[1].bstrVal = SysAllocString(L"ExecReplace");

    DISPPARAMS executeParams = { executeArgs, NULL, 2, 0 };
    VariantInit(&result);
    hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &executeParams, &result, NULL, NULL);
    SysFreeString(executeArgs[1].bstrVal);

    pHSet->Release();
    pHFindReplace->Release();

    return SUCCEEDED(hr) && (result.vt == VT_BOOL ? result.boolVal != VARIANT_FALSE : true);
}

int HwpWrapper::ReplaceAll(const std::wstring& find_text,
                            const std::wstring& replace_text,
                            bool match_case,
                            bool regex)
{
    if (!m_pHwp || find_text.empty()) return 0;

    HRESULT hr;
    DISPID dispid;
    VARIANT result;
    VariantInit(&result);

    // 1. HParameterSet 속성 가져오기
    IDispatch* pHParameterSet = GetHParameterSet();
    if (!pHParameterSet) return 0;

    // 2. HFindReplace 속성 가져오기
    OLECHAR* findReplaceName = const_cast<OLECHAR*>(L"HFindReplace");
    hr = pHParameterSet->GetIDsOfNames(IID_NULL, &findReplaceName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return 0;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    VariantInit(&result);
    hr = pHParameterSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                 &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) return 0;

    IDispatch* pHFindReplace = result.pdispVal;

    // 3. HSet 속성 가져오기
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHFindReplace->Release();
        return 0;
    }

    VariantInit(&result);
    hr = pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                                &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        pHFindReplace->Release();
        return 0;
    }

    IDispatch* pHSet = result.pdispVal;

    // 4. HAction 가져오기
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pHSet->Release();
        pHFindReplace->Release();
        return 0;
    }

    // 5. HAction.GetDefault("FindDlg", HSet) 호출하여 초기화
    OLECHAR* getDefaultName = const_cast<OLECHAR*>(L"GetDefault");
    hr = pHAction->GetIDsOfNames(IID_NULL, &getDefaultName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT getDefaultArgs[2];
        VariantInit(&getDefaultArgs[0]);
        VariantInit(&getDefaultArgs[1]);
        getDefaultArgs[0].vt = VT_DISPATCH;
        getDefaultArgs[0].pdispVal = pHSet;
        getDefaultArgs[1].vt = VT_BSTR;
        getDefaultArgs[1].bstrVal = SysAllocString(L"FindDlg");

        DISPPARAMS getDefaultParams = { getDefaultArgs, NULL, 2, 0 };
        VariantInit(&result);
        pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                         &getDefaultParams, &result, NULL, NULL);
        SysFreeString(getDefaultArgs[1].bstrVal);
    }

    // 6. 파라미터 설정
    DISPID putid = DISPID_PROPERTYPUT;

    // FindString 설정
    OLECHAR* findStringName = const_cast<OLECHAR*>(L"FindString");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &findStringName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_BSTR;
        val.bstrVal = SysAllocString(find_text.c_str());
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
        SysFreeString(val.bstrVal);
    }

    // ReplaceString 설정
    OLECHAR* replaceStringName = const_cast<OLECHAR*>(L"ReplaceString");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &replaceStringName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_BSTR;
        val.bstrVal = SysAllocString(replace_text.c_str());
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
        SysFreeString(val.bstrVal);
    }

    // MatchCase 설정
    OLECHAR* matchCaseName = const_cast<OLECHAR*>(L"MatchCase");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &matchCaseName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = match_case ? 1 : 0;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // FindRegExp 설정 (정규식)
    OLECHAR* findRegExpName = const_cast<OLECHAR*>(L"FindRegExp");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &findRegExpName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = regex ? 1 : 0;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // IgnoreMessage 설정 (메시지 무시)
    OLECHAR* ignoreMsgName = const_cast<OLECHAR*>(L"IgnoreMessage");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &ignoreMsgName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = 1;
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // ReplaceMode 설정
    OLECHAR* replaceModeName = const_cast<OLECHAR*>(L"ReplaceMode");
    hr = pHFindReplace->GetIDsOfNames(IID_NULL, &replaceModeName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = 1;  // 바꾸기 모드
        DISPPARAMS params = { &val, &putid, 1, 1 };
        pHFindReplace->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                               &params, NULL, NULL, NULL);
    }

    // 7. HAction.Execute("AllReplace", HSet) 호출
    OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
    hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHSet->Release();
        pHFindReplace->Release();
        return 0;
    }

    VARIANT executeArgs[2];
    VariantInit(&executeArgs[0]);
    VariantInit(&executeArgs[1]);
    executeArgs[0].vt = VT_DISPATCH;
    executeArgs[0].pdispVal = pHSet;
    executeArgs[1].vt = VT_BSTR;
    executeArgs[1].bstrVal = SysAllocString(L"AllReplace");

    DISPPARAMS executeParams = { executeArgs, NULL, 2, 0 };
    VariantInit(&result);
    hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &executeParams, &result, NULL, NULL);
    SysFreeString(executeArgs[1].bstrVal);

    int replaceCount = 0;
    if (SUCCEEDED(hr)) {
        // 바꾼 개수 가져오기 (결과에서 또는 Count 속성에서)
        if (result.vt == VT_I4) {
            replaceCount = result.lVal;
        } else if (result.vt == VT_BOOL && result.boolVal != VARIANT_FALSE) {
            // 성공했지만 개수를 반환하지 않는 경우, Count 속성 확인
            OLECHAR* countName = const_cast<OLECHAR*>(L"Count");
            DISPID dispidCount;
            if (SUCCEEDED(pHFindReplace->GetIDsOfNames(IID_NULL, &countName, 1,
                                                        LOCALE_USER_DEFAULT, &dispidCount))) {
                VARIANT countResult;
                VariantInit(&countResult);
                if (SUCCEEDED(pHFindReplace->Invoke(dispidCount, IID_NULL, LOCALE_USER_DEFAULT,
                                                     DISPATCH_PROPERTYGET, &noParams,
                                                     &countResult, NULL, NULL))) {
                    if (countResult.vt == VT_I4) {
                        replaceCount = countResult.lVal;
                    }
                }
            }
            // Count 속성이 없으면 최소 1개는 바뀐 것으로 간주
            if (replaceCount == 0) replaceCount = 1;
        }
    }

    pHSet->Release();
    pHFindReplace->Release();

    return replaceCount;
}

//=============================================================================
// HAction 관련
//=============================================================================

bool HwpWrapper::RunAction(const std::wstring& action_name)
{
    if (!m_pHwp) return false;

    // HAction 객체 가져오기
    IDispatch* pHAction = GetHAction();
    if (!pHAction) return false;

    // Run 메서드 호출
    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"Run");
    HRESULT hr = pHAction->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    VARIANT args[1];
    VariantInit(&args[0]);
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(action_name.c_str());

    DISPPARAMS params = { args, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &params, &result, NULL, NULL);

    SysFreeString(args[0].bstrVal);

    if (FAILED(hr)) return false;
    return (result.vt == VT_BOOL) ? (result.boolVal != VARIANT_FALSE) : true;
}

IDispatch* HwpWrapper::CreateAction(const std::wstring& action_id)
{
    if (!m_pHwp) return nullptr;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"CreateAction");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return nullptr;

    VARIANT args[1];
    VariantInit(&args[0]);
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(action_id.c_str());

    DISPPARAMS params = { args, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    SysFreeString(args[0].bstrVal);

    if (FAILED(hr) || result.vt != VT_DISPATCH) return nullptr;
    return result.pdispVal;
}

IDispatch* HwpWrapper::CreateSet(const std::wstring& set_id)
{
    if (!m_pHwp) return nullptr;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"CreateSet");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return nullptr;

    VARIANT args[1];
    VariantInit(&args[0]);
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(set_id.c_str());

    DISPPARAMS params = { args, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    SysFreeString(args[0].bstrVal);

    if (FAILED(hr) || result.vt != VT_DISPATCH) return nullptr;
    return result.pdispVal;
}

bool HwpWrapper::FindCtrl()
{
    if (!m_pHwp) return false;

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"FindCtrl");
    if (dispid == DISPID_UNKNOWN) return false;

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    // FindCtrl은 찾은 컨트롤 객체를 반환 (VT_DISPATCH면 성공)
    bool success = SUCCEEDED(hr) && result.vt == VT_DISPATCH && result.pdispVal != nullptr;

    if (result.vt == VT_DISPATCH && result.pdispVal) {
        result.pdispVal->Release();
    }
    VariantClear(&result);

    return success;
}

//=============================================================================
// 컨트롤 관리
//=============================================================================

std::unique_ptr<HwpCtrl> HwpWrapper::InsertCtrl(const std::wstring& ctrl_id,
                                                  IDispatch* initparam)
{
    if (!m_pHwp) return nullptr;

    HRESULT hr;
    DISPID dispid;

    // 1. InsertCtrl 메서드의 DISPID 획득
    OLECHAR* methodName = const_cast<OLECHAR*>(L"InsertCtrl");
    hr = m_pHwp->GetIDsOfNames(IID_NULL, &methodName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return nullptr;

    // 2. InsertCtrl(CtrlID, initparam) 호출
    VARIANT args[2];
    VariantInit(&args[0]);
    VariantInit(&args[1]);

    // 인자는 역순 (initparam, CtrlID)
    if (initparam) {
        args[0].vt = VT_DISPATCH;
        args[0].pdispVal = initparam;
    } else {
        args[0].vt = VT_ERROR;
        args[0].scode = DISP_E_PARAMNOTFOUND;  // Optional parameter
    }
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(ctrl_id.c_str());

    DISPPARAMS params = { args, NULL, 2, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    SysFreeString(args[1].bstrVal);

    if (FAILED(hr)) return nullptr;

    // 3. 반환된 컨트롤 객체 래핑
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        IDispatch* pCtrl = result.pdispVal;
        // HwpCtrl 생성자에서 AddRef를 하므로 여기서 Release 필요
        auto ctrl = std::make_unique<HwpCtrl>(pCtrl, this);
        pCtrl->Release();
        return ctrl;
    }

    return nullptr;
}

bool HwpWrapper::DeleteCtrl(HwpCtrl* ctrl)
{
    if (!ctrl || !ctrl->IsValid()) return false;
    return DeleteCtrl(ctrl->GetDispatch());
}

bool HwpWrapper::DeleteCtrl(IDispatch* pCtrl)
{
    if (!m_pHwp || !pCtrl) return false;

    HRESULT hr;
    DISPID dispid;

    // 1. DeleteCtrl 메서드의 DISPID 획득
    OLECHAR* methodName = const_cast<OLECHAR*>(L"DeleteCtrl");
    hr = m_pHwp->GetIDsOfNames(IID_NULL, &methodName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    // 2. DeleteCtrl(ctrl) 호출
    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_DISPATCH;
    arg.pdispVal = pCtrl;

    DISPPARAMS params = { &arg, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    if (FAILED(hr)) return false;

    // 3. 결과 반환
    if (result.vt == VT_BOOL) return result.boolVal != VARIANT_FALSE;
    return true;  // 성공으로 간주
}

//=============================================================================
// 속성 접근
//=============================================================================

std::wstring HwpWrapper::GetVersion()
{
    VARIANT result = GetProperty(L"Version");
    if (result.vt == VT_BSTR) {
        std::wstring ver(result.bstrVal ? result.bstrVal : L"");
        SysFreeString(result.bstrVal);
        return ver;
    }
    return L"";
}

std::wstring HwpWrapper::GetBuildNumber()
{
    VARIANT result = GetProperty(L"BuildNumber");
    if (result.vt == VT_BSTR) {
        std::wstring build(result.bstrVal ? result.bstrVal : L"");
        SysFreeString(result.bstrVal);
        return build;
    }
    return L"";
}

int HwpWrapper::GetCurrentPage()
{
    VARIANT result = GetProperty(L"CurrentPage");
    if (result.vt == VT_I4) {
        return result.lVal;
    }
    return 0;
}

int HwpWrapper::GetCurrentPrintPage()
{
    VARIANT result = GetProperty(L"CurrentPrnPage");
    if (result.vt == VT_I4) {
        return result.lVal;
    }
    return 0;
}

int HwpWrapper::GetPageCount()
{
    VARIANT result = GetProperty(L"PageCount");
    if (result.vt == VT_I4) {
        return result.lVal;
    }
    return 0;
}

bool HwpWrapper::GetEditMode()
{
    VARIANT result = GetProperty(L"EditMode");
    if (result.vt == VT_BOOL) {
        return result.boolVal != VARIANT_FALSE;
    }
    return false;
}

void HwpWrapper::SetEditMode(bool mode)
{
    VARIANT val;
    VariantInit(&val);
    val.vt = VT_BOOL;
    val.boolVal = mode ? VARIANT_TRUE : VARIANT_FALSE;

    SetProperty(L"EditMode", val);
}

//=============================================================================
// 컨트롤 속성 접근
//=============================================================================

std::unique_ptr<HwpCtrl> HwpWrapper::GetCurSelectedCtrl()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"CurSelectedCtrl");
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        return std::make_unique<HwpCtrl>(result.pdispVal, this);
    }
    return nullptr;
}

std::unique_ptr<HwpCtrl> HwpWrapper::GetHeadCtrl()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"HeadCtrl");
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        return std::make_unique<HwpCtrl>(result.pdispVal, this);
    }
    return nullptr;
}

std::unique_ptr<HwpCtrl> HwpWrapper::GetLastCtrl()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"LastCtrl");
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        return std::make_unique<HwpCtrl>(result.pdispVal, this);
    }
    return nullptr;
}

std::unique_ptr<HwpCtrl> HwpWrapper::GetParentCtrl()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"ParentCtrl");
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        return std::make_unique<HwpCtrl>(result.pdispVal, this);
    }
    return nullptr;
}

std::vector<std::unique_ptr<HwpCtrl>> HwpWrapper::GetCtrlList()
{
    std::vector<std::unique_ptr<HwpCtrl>> result;
    if (!m_pHwp) return result;

    // HeadCtrl부터 시작
    auto ctrl = GetHeadCtrl();
    if (!ctrl) return result;

    // secd(섹션정의) 스킵
    ctrl = ctrl->Next();
    if (!ctrl) return result;

    // cold(단정의) 스킵
    ctrl = ctrl->Next();

    // 나머지 모든 컨트롤 순회
    while (ctrl) {
        auto next = ctrl->Next();
        result.push_back(std::move(ctrl));
        ctrl = std::move(next);
    }

    return result;
}

//=============================================================================
// 문서 컬렉션 접근
//=============================================================================

std::unique_ptr<XHwpDocuments> HwpWrapper::GetXHwpDocuments()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"XHwpDocuments");
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        return std::make_unique<XHwpDocuments>(result.pdispVal);
    }
    return nullptr;
}

std::unique_ptr<XHwpDocument> HwpWrapper::SwitchTo(int num)
{
    auto docs = GetXHwpDocuments();
    if (!docs) return nullptr;

    int count = docs->GetCount();
    if (num < 0 || num >= count) return nullptr;

    auto doc = docs->Item(num);
    if (doc) {
        doc->SetActiveDocument();
    }
    return doc;
}

std::unique_ptr<XHwpDocument> HwpWrapper::AddTab()
{
    auto docs = GetXHwpDocuments();
    if (!docs) return nullptr;

    return docs->Add(true);  // isTab = true
}

std::unique_ptr<XHwpDocument> HwpWrapper::AddDoc()
{
    auto docs = GetXHwpDocuments();
    if (!docs) return nullptr;

    return docs->Add(false);  // isTab = false (new window)
}

//=============================================================================
// COM 인터페이스 직접 접근
//=============================================================================

IDispatch* HwpWrapper::GetHParameterSet()
{
    if (m_pHParameterSet) {
        return m_pHParameterSet;
    }

    VARIANT result = GetProperty(L"HParameterSet");
    if (result.vt == VT_DISPATCH) {
        m_pHParameterSet = result.pdispVal;
        return m_pHParameterSet;
    }
    return nullptr;
}

IDispatch* HwpWrapper::GetHAction()
{
    if (m_pHAction) {
        return m_pHAction;
    }

    VARIANT result = GetProperty(L"HAction");
    if (result.vt == VT_DISPATCH) {
        m_pHAction = result.pdispVal;
        return m_pHAction;
    }
    return nullptr;
}

//=============================================================================
// COM 헬퍼 메서드
//=============================================================================

VARIANT HwpWrapper::GetProperty(const std::wstring& name)
{
    VARIANT result;
    VariantInit(&result);

    if (!m_pHwp) return result;

    // DISPID 캐시 사용 (성능 최적화)
    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, name);
    if (dispid == DISPID_UNKNOWN) return result;

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                   &params, &result, NULL, NULL);

    return result;
}

bool HwpWrapper::SetProperty(const std::wstring& name, const VARIANT& value)
{
    if (!m_pHwp) return false;

    // DISPID 캐시 사용 (성능 최적화)
    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, name);
    if (dispid == DISPID_UNKNOWN) return false;

    DISPID putid = DISPID_PROPERTYPUT;
    DISPPARAMS params = { const_cast<VARIANT*>(&value), &putid, 1, 1 };

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                        &params, NULL, NULL, NULL);

    return SUCCEEDED(hr);
}

bool HwpWrapper::InvokeMethod(const std::wstring& name)
{
    VARIANT result = InvokeMethodWithResult(name);
    VariantClear(&result);
    return true;
}

VARIANT HwpWrapper::InvokeMethodWithResult(const std::wstring& name,
                                            const std::vector<VARIANT>& args)
{
    VARIANT result;
    VariantInit(&result);

    if (!m_pHwp) return result;

    // DISPID 캐시 사용 (성능 최적화)
    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, name);
    if (dispid == DISPID_UNKNOWN) return result;

    // 인자 역순 배열 생성
    std::vector<VARIANT> reversedArgs(args.rbegin(), args.rend());

    DISPPARAMS params = {
        reversedArgs.empty() ? NULL : reversedArgs.data(),
        NULL,
        static_cast<UINT>(reversedArgs.size()),
        0
    };

    m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                   &params, &result, NULL, NULL);

    return result;
}

//=============================================================================
// 필드 작업 (Field Operations)
//=============================================================================

bool HwpWrapper::CreateField(const std::wstring& name,
                              const std::wstring& direction,
                              const std::wstring& memo)
{
    if (!m_pHwp) return false;

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"CreateField");
    if (dispid == DISPID_UNKNOWN) return false;

    // Positional parameters (역순): name, memo, direction
    // COM에서 arguments는 역순으로 전달됨
    VARIANT args[3];
    VariantInit(&args[0]);
    VariantInit(&args[1]);
    VariantInit(&args[2]);

    // 역순: [0]=name, [1]=memo, [2]=direction
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(name.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(memo.c_str());
    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(direction.c_str());

    DISPPARAMS params = { args, NULL, 3, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_METHOD, &params, &result, NULL, NULL);

    VariantClear(&args[0]);
    VariantClear(&args[1]);
    VariantClear(&args[2]);

    bool success = SUCCEEDED(hr) && (result.vt == VT_BOOL) &&
                   (result.boolVal != VARIANT_FALSE);
    VariantClear(&result);
    return success;
}

std::wstring HwpWrapper::GetFieldList(int number, int option)
{
    if (!m_pHwp) return L"";

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"GetFieldList");
    if (dispid == DISPID_UNKNOWN) return L"";

    // Positional parameters (역순): option, number
    VARIANT args[2];
    args[0].vt = VT_I4;
    args[0].lVal = option;
    args[1].vt = VT_I4;
    args[1].lVal = number;

    DISPPARAMS params = { args, NULL, 2, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    std::wstring fieldList;
    if (SUCCEEDED(hr) && result.vt == VT_BSTR && result.bstrVal) {
        fieldList = result.bstrVal;
    }
    VariantClear(&result);
    return fieldList;
}

std::wstring HwpWrapper::GetFieldText(const std::wstring& field, int idx)
{
    if (!m_pHwp) return L"";

    // 인덱스 처리: "field{{n}}" 형식
    std::wstring fieldName = field;
    if (idx > 0 && field.find(L"{{") == std::wstring::npos) {
        fieldName = field + L"{{" + std::to_wstring(idx) + L"}}";
    }

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"GetFieldText");
    if (dispid == DISPID_UNKNOWN) return L"";

    // Positional parameter: Field
    VARIANT arg;
    arg.vt = VT_BSTR;
    arg.bstrVal = SysAllocString(fieldName.c_str());

    DISPPARAMS params = { &arg, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    SysFreeString(arg.bstrVal);

    std::wstring text;
    if (SUCCEEDED(hr) && result.vt == VT_BSTR && result.bstrVal) {
        text = result.bstrVal;
    }
    VariantClear(&result);
    return text;
}

bool HwpWrapper::PutFieldText(const std::wstring& field, const std::wstring& text)
{
    if (!m_pHwp) return false;

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"PutFieldText");
    if (dispid == DISPID_UNKNOWN) return false;

    // Positional parameters (역순): text, field
    VARIANT args[2];
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(text.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(field.c_str());

    DISPPARAMS params = { args, NULL, 2, 0 };

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, NULL, NULL, NULL);

    SysFreeString(args[0].bstrVal);
    SysFreeString(args[1].bstrVal);

    return SUCCEEDED(hr);
}

bool HwpWrapper::FieldExist(const std::wstring& field)
{
    if (!m_pHwp) return false;

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"FieldExist");
    if (dispid == DISPID_UNKNOWN) return false;

    // Positional parameter: Field
    VARIANT arg;
    arg.vt = VT_BSTR;
    arg.bstrVal = SysAllocString(field.c_str());

    DISPPARAMS params = { &arg, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    SysFreeString(arg.bstrVal);

    bool exists = SUCCEEDED(hr) && (result.vt == VT_BOOL) &&
                  (result.boolVal != VARIANT_FALSE);
    VariantClear(&result);
    return exists;
}

bool HwpWrapper::MoveToField(const std::wstring& field, int idx,
                              bool text, bool start, bool select)
{
    if (!m_pHwp) return false;

    // 인덱스 처리: "field{{n}}" 형식
    std::wstring fieldName = field;
    if (field.find(L"{{") == std::wstring::npos) {
        fieldName = field + L"{{" + std::to_wstring(idx) + L"}}";
    }

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"MoveToField");
    if (dispid == DISPID_UNKNOWN) return false;

    // Positional parameters (역순): select, start, text, field
    VARIANT args[4];
    args[0].vt = VT_BOOL;
    args[0].boolVal = select ? VARIANT_TRUE : VARIANT_FALSE;
    args[1].vt = VT_BOOL;
    args[1].boolVal = start ? VARIANT_TRUE : VARIANT_FALSE;
    args[2].vt = VT_BOOL;
    args[2].boolVal = text ? VARIANT_TRUE : VARIANT_FALSE;
    args[3].vt = VT_BSTR;
    args[3].bstrVal = SysAllocString(fieldName.c_str());

    DISPPARAMS params = { args, NULL, 4, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    SysFreeString(args[3].bstrVal);

    bool success = SUCCEEDED(hr) && (result.vt == VT_BOOL) &&
                   (result.boolVal != VARIANT_FALSE);
    VariantClear(&result);
    return success;
}

bool HwpWrapper::RenameField(const std::wstring& oldname, const std::wstring& newname)
{
    if (!m_pHwp) return false;

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"RenameField");
    if (dispid == DISPID_UNKNOWN) return false;

    // Positional parameters (역순): newname, oldname
    VARIANT args[2];
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(newname.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(oldname.c_str());

    DISPPARAMS params = { args, NULL, 2, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    SysFreeString(args[0].bstrVal);
    SysFreeString(args[1].bstrVal);

    bool success = SUCCEEDED(hr) && (result.vt == VT_BOOL) &&
                   (result.boolVal != VARIANT_FALSE);
    VariantClear(&result);
    return success;
}

std::wstring HwpWrapper::GetCurFieldName(int option)
{
    if (!m_pHwp) return L"";

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"GetCurFieldName");
    if (dispid == DISPID_UNKNOWN) return L"";

    // Positional parameter: option
    VARIANT arg;
    arg.vt = VT_I4;
    arg.lVal = option;

    DISPPARAMS params = { &arg, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    std::wstring fieldName;
    if (SUCCEEDED(hr) && result.vt == VT_BSTR && result.bstrVal) {
        fieldName = result.bstrVal;
    }
    VariantClear(&result);
    return fieldName;
}

bool HwpWrapper::SetCurFieldName(const std::wstring& field,
                                  const std::wstring& direction,
                                  const std::wstring& memo,
                                  int option)
{
    if (!m_pHwp) return false;

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"SetCurFieldName");
    if (dispid == DISPID_UNKNOWN) return false;

    // Positional parameters (역순): option, memo, direction, field
    VARIANT args[4];
    args[0].vt = VT_I4;
    args[0].lVal = option;
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(memo.c_str());
    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(direction.c_str());
    args[3].vt = VT_BSTR;
    args[3].bstrVal = SysAllocString(field.c_str());

    DISPPARAMS params = { args, NULL, 4, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    SysFreeString(args[1].bstrVal);
    SysFreeString(args[2].bstrVal);
    SysFreeString(args[3].bstrVal);

    bool success = SUCCEEDED(hr);
    VariantClear(&result);
    return success;
}

int HwpWrapper::SetFieldViewOption(int option)
{
    if (!m_pHwp) return -1;

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"SetFieldViewOption");
    if (dispid == DISPID_UNKNOWN) return -1;

    // Positional parameter: option
    VARIANT arg;
    arg.vt = VT_I4;
    arg.lVal = option;

    DISPPARAMS params = { &arg, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    int prevOption = -1;
    if (SUCCEEDED(hr) && result.vt == VT_I4) {
        prevOption = result.lVal;
    }
    VariantClear(&result);
    return prevOption;
}

bool HwpWrapper::DeleteAllFields()
{
    if (!m_pHwp) return false;

    // 현재 위치 저장
    HwpPos startPos = GetPos();

    // 필드 목록 가져오기
    std::wstring fieldList = GetFieldList();
    if (fieldList.empty()) return true;

    // \x02로 분리하여 각 필드 삭제
    std::vector<std::wstring> fields;
    size_t start = 0;
    size_t end = 0;
    while ((end = fieldList.find(L'\x02', start)) != std::wstring::npos) {
        if (end > start) {
            fields.push_back(fieldList.substr(start, end - start));
        }
        start = end + 1;
    }
    if (start < fieldList.length()) {
        fields.push_back(fieldList.substr(start));
    }

    // 역순으로 필드 삭제 (rename으로 이름 제거)
    for (auto it = fields.rbegin(); it != fields.rend(); ++it) {
        RenameField(*it, L"");
    }

    // 원래 위치로 복원
    SetPos(startPos.list, startPos.para, startPos.pos);

    return true;
}

bool HwpWrapper::DeleteFieldByName(const std::wstring& field_name, int idx)
{
    if (!m_pHwp) return false;

    // 현재 위치 저장
    HwpPos startPos = GetPos();

    if (idx == -1) {
        // 모든 동일 이름 필드 삭제
        std::wstring fieldList = GetFieldList();
        if (fieldList.empty()) return true;

        // 필드 목록에서 해당 이름의 필드 찾아서 삭제
        std::vector<std::wstring> fields;
        size_t start = 0;
        size_t end = 0;
        while ((end = fieldList.find(L'\x02', start)) != std::wstring::npos) {
            if (end > start) {
                fields.push_back(fieldList.substr(start, end - start));
            }
            start = end + 1;
        }
        if (start < fieldList.length()) {
            fields.push_back(fieldList.substr(start));
        }

        // 역순으로 해당 이름의 필드만 삭제
        for (auto it = fields.rbegin(); it != fields.rend(); ++it) {
            // field_name{{n}} 형식에서 이름 부분 추출
            std::wstring fieldNamePart = *it;
            size_t bracketPos = it->find(L"{{");
            if (bracketPos != std::wstring::npos) {
                fieldNamePart = it->substr(0, bracketPos);
            }
            if (fieldNamePart == field_name) {
                RenameField(*it, L"");
            }
        }
    } else {
        // 특정 인덱스의 필드만 삭제
        std::wstring fieldWithIdx = field_name + L"{{" + std::to_wstring(idx) + L"}}";
        RenameField(fieldWithIdx, L"");
    }

    // 원래 위치로 복원
    SetPos(startPos.list, startPos.para, startPos.pos);

    return true;
}

std::map<std::wstring, std::wstring> HwpWrapper::FieldsToMap()
{
    std::map<std::wstring, std::wstring> result;
    if (!m_pHwp) return result;

    // 필드 목록 가져오기
    std::wstring fieldList = GetFieldList(1, 0);
    if (fieldList.empty()) return result;

    // \x02로 분리
    std::vector<std::wstring> fields;
    size_t start = 0;
    size_t end = 0;
    while ((end = fieldList.find(L'\x02', start)) != std::wstring::npos) {
        if (end > start) {
            fields.push_back(fieldList.substr(start, end - start));
        }
        start = end + 1;
    }
    if (start < fieldList.length()) {
        fields.push_back(fieldList.substr(start));
    }

    // 각 필드의 텍스트 가져오기
    for (const auto& field : fields) {
        std::wstring text = GetFieldText(field, 0);
        result[field] = text;
    }

    return result;
}

//=============================================================================
// 테이블 작업 (Table Operations)
//=============================================================================

bool HwpWrapper::CreateTable(int rows, int cols,
                              bool treat_as_char,
                              int width_type,
                              int height_type,
                              bool header)
{
    if (!m_pHwp) return false;

    HRESULT hr;
    DISPID dispid;
    VARIANT result;
    VariantInit(&result);

    // 1. HParameterSet 속성 가져오기
    IDispatch* pHParameterSet = GetHParameterSet();
    if (!pHParameterSet) return false;

    // 2. HTableCreation 속성 가져오기
    OLECHAR* tableCreationName = const_cast<OLECHAR*>(L"HTableCreation");
    hr = pHParameterSet->GetIDsOfNames(IID_NULL, &tableCreationName, 1,
                                        LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    VariantInit(&result);
    hr = pHParameterSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) return false;

    IDispatch* pTableCreation = result.pdispVal;

    // 3. HSet 속성 가져오기
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pTableCreation->GetIDsOfNames(IID_NULL, &hsetName, 1,
                                        LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pTableCreation->Release();
        return false;
    }

    VariantInit(&result);
    hr = pTableCreation->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &noParams, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH) {
        pTableCreation->Release();
        return false;
    }

    IDispatch* pHSet = result.pdispVal;

    // 4. HAction 가져오기
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pHSet->Release();
        pTableCreation->Release();
        return false;
    }

    // 5. HAction.GetDefault("TableCreate", HSet) 호출
    OLECHAR* getDefaultName = const_cast<OLECHAR*>(L"GetDefault");
    hr = pHAction->GetIDsOfNames(IID_NULL, &getDefaultName, 1,
                                  LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHSet->Release();
        pTableCreation->Release();
        return false;
    }

    VARIANT getDefaultArgs[2];
    VariantInit(&getDefaultArgs[0]);
    VariantInit(&getDefaultArgs[1]);
    getDefaultArgs[0].vt = VT_DISPATCH;
    getDefaultArgs[0].pdispVal = pHSet;
    getDefaultArgs[1].vt = VT_BSTR;
    getDefaultArgs[1].bstrVal = SysAllocString(L"TableCreate");

    DISPPARAMS getDefaultParams = { getDefaultArgs, NULL, 2, 0 };
    VariantInit(&result);
    hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &getDefaultParams, &result, NULL, NULL);
    SysFreeString(getDefaultArgs[1].bstrVal);

    // 6. 파라미터 설정: Rows, Cols, WidthType, HeightType
    DISPID putid = DISPID_PROPERTYPUT;

    // Rows 설정
    OLECHAR* rowsName = const_cast<OLECHAR*>(L"Rows");
    hr = pTableCreation->GetIDsOfNames(IID_NULL, &rowsName, 1,
                                        LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT rowsVal;
        VariantInit(&rowsVal);
        rowsVal.vt = VT_I4;
        rowsVal.lVal = rows;
        DISPPARAMS rowsParams = { &rowsVal, &putid, 1, 1 };
        pTableCreation->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYPUT, &rowsParams, NULL, NULL, NULL);
    }

    // Cols 설정
    OLECHAR* colsName = const_cast<OLECHAR*>(L"Cols");
    hr = pTableCreation->GetIDsOfNames(IID_NULL, &colsName, 1,
                                        LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT colsVal;
        VariantInit(&colsVal);
        colsVal.vt = VT_I4;
        colsVal.lVal = cols;
        DISPPARAMS colsParams = { &colsVal, &putid, 1, 1 };
        pTableCreation->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYPUT, &colsParams, NULL, NULL, NULL);
    }

    // WidthType 설정
    OLECHAR* widthTypeName = const_cast<OLECHAR*>(L"WidthType");
    hr = pTableCreation->GetIDsOfNames(IID_NULL, &widthTypeName, 1,
                                        LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT widthTypeVal;
        VariantInit(&widthTypeVal);
        widthTypeVal.vt = VT_I4;
        widthTypeVal.lVal = width_type;
        DISPPARAMS widthTypeParams = { &widthTypeVal, &putid, 1, 1 };
        pTableCreation->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYPUT, &widthTypeParams, NULL, NULL, NULL);
    }

    // HeightType 설정
    OLECHAR* heightTypeName = const_cast<OLECHAR*>(L"HeightType");
    hr = pTableCreation->GetIDsOfNames(IID_NULL, &heightTypeName, 1,
                                        LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT heightTypeVal;
        VariantInit(&heightTypeVal);
        heightTypeVal.vt = VT_I4;
        heightTypeVal.lVal = height_type;
        DISPPARAMS heightTypeParams = { &heightTypeVal, &putid, 1, 1 };
        pTableCreation->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYPUT, &heightTypeParams, NULL, NULL, NULL);
    }

    // TreatAsChar 설정 (CreateItemArray 대신 직접 설정)
    OLECHAR* treatAsCharName = const_cast<OLECHAR*>(L"TreatAsChar");
    hr = pTableCreation->GetIDsOfNames(IID_NULL, &treatAsCharName, 1,
                                        LOCALE_USER_DEFAULT, &dispid);
    if (SUCCEEDED(hr)) {
        VARIANT treatVal;
        VariantInit(&treatVal);
        treatVal.vt = VT_I4;
        treatVal.lVal = treat_as_char ? 1 : 0;
        DISPPARAMS treatParams = { &treatVal, &putid, 1, 1 };
        pTableCreation->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYPUT, &treatParams, NULL, NULL, NULL);
    }

    // 7. HAction.Execute("TableCreate", HSet) 호출
    OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
    hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1,
                                  LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pHSet->Release();
        pTableCreation->Release();
        return false;
    }

    VARIANT executeArgs[2];
    VariantInit(&executeArgs[0]);
    VariantInit(&executeArgs[1]);
    executeArgs[0].vt = VT_DISPATCH;
    executeArgs[0].pdispVal = pHSet;
    executeArgs[1].vt = VT_BSTR;
    executeArgs[1].bstrVal = SysAllocString(L"TableCreate");

    DISPPARAMS executeParams = { executeArgs, NULL, 2, 0 };
    VariantInit(&result);
    hr = pHAction->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                          &executeParams, &result, NULL, NULL);
    SysFreeString(executeArgs[1].bstrVal);

    bool success = SUCCEEDED(hr) && (result.vt == VT_BOOL) &&
                   (result.boolVal != VARIANT_FALSE);

    pHSet->Release();
    pTableCreation->Release();

    return success;
}

bool HwpWrapper::GetIntoNthTable(int n, bool select_cell)
{
    if (!m_pHwp) return false;

    // 문서 시작으로 이동
    MovePos(2, 0, 0);  // moveDocBegin = 2

    // HeadCtrl 가져오기
    VARIANT headCtrlVar = GetProperty(L"HeadCtrl");
    if (headCtrlVar.vt != VT_DISPATCH || !headCtrlVar.pdispVal) {
        VariantClear(&headCtrlVar);
        return false;
    }

    IDispatch* pCtrl = headCtrlVar.pdispVal;
    int tableCount = 0;
    int targetN = (n >= 0) ? n : -(n + 1);

    DISPID dispidUserDesc, dispidNext;
    DISPPARAMS noParams = { NULL, NULL, 0, 0 };

    while (pCtrl) {
        // UserDesc 속성 조회
        OLECHAR* userDescName = const_cast<OLECHAR*>(L"UserDesc");
        HRESULT hr = pCtrl->GetIDsOfNames(IID_NULL, &userDescName, 1,
                                           LOCALE_USER_DEFAULT, &dispidUserDesc);
        if (SUCCEEDED(hr)) {
            VARIANT userDescVar;
            VariantInit(&userDescVar);
            hr = pCtrl->Invoke(dispidUserDesc, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYGET, &noParams, &userDescVar, NULL, NULL);
            if (SUCCEEDED(hr) && userDescVar.vt == VT_BSTR && userDescVar.bstrVal) {
                std::wstring userDesc(userDescVar.bstrVal);
                if (userDesc == L"표") {
                    if (tableCount == targetN) {
                        // 테이블 찾음 - 테이블 안으로 이동
                        OLECHAR* getAnchorPosName = const_cast<OLECHAR*>(L"GetAnchorPos");
                        DISPID dispidGetAnchorPos;
                        hr = pCtrl->GetIDsOfNames(IID_NULL, &getAnchorPosName, 1,
                                                   LOCALE_USER_DEFAULT, &dispidGetAnchorPos);
                        if (SUCCEEDED(hr)) {
                            // GetAnchorPos(0) - 시작 위치의 ParameterSet 반환
                            VARIANT anchorArg;
                            anchorArg.vt = VT_I4;
                            anchorArg.lVal = 0;
                            DISPPARAMS anchorParams = { &anchorArg, NULL, 1, 0 };
                            VARIANT anchorResult;
                            VariantInit(&anchorResult);
                            hr = pCtrl->Invoke(dispidGetAnchorPos, IID_NULL, LOCALE_USER_DEFAULT,
                                                DISPATCH_METHOD, &anchorParams, &anchorResult, NULL, NULL);

                            // GetAnchorPos는 ParameterSet (IDispatch)을 반환
                            if (SUCCEEDED(hr) && anchorResult.vt == VT_DISPATCH && anchorResult.pdispVal) {
                                // 1. SetPosBySet으로 위치 이동 (핵심!)
                                SetPosBySet(anchorResult.pdispVal);

                                // 2. FindCtrl() - 컨트롤을 선택 상태로 만듦
                                FindCtrl();

                                // 3. 셀 안으로 진입
                                if (select_cell) {
                                    // 셀 블록 선택 상태
                                    RunAction(L"ShapeObjTableSelCell");
                                } else {
                                    // 편집 모드로 진입
                                    RunAction(L"ShapeObjTextBoxEdit");
                                }
                            }
                            VariantClear(&anchorResult);
                        }

                        VariantClear(&userDescVar);
                        pCtrl->Release();
                        return true;
                    }
                    tableCount++;
                }
                VariantClear(&userDescVar);
            }
        }

        // Next 컨트롤로 이동
        OLECHAR* nextName = const_cast<OLECHAR*>(L"Next");
        hr = pCtrl->GetIDsOfNames(IID_NULL, &nextName, 1,
                                   LOCALE_USER_DEFAULT, &dispidNext);
        if (FAILED(hr)) {
            pCtrl->Release();
            break;
        }

        VARIANT nextVar;
        VariantInit(&nextVar);
        hr = pCtrl->Invoke(dispidNext, IID_NULL, LOCALE_USER_DEFAULT,
                            DISPATCH_PROPERTYGET, &noParams, &nextVar, NULL, NULL);

        pCtrl->Release();

        if (SUCCEEDED(hr) && nextVar.vt == VT_DISPATCH && nextVar.pdispVal) {
            pCtrl = nextVar.pdispVal;
        } else {
            VariantClear(&nextVar);
            pCtrl = nullptr;
        }
    }

    return false;
}

int HwpWrapper::GetTableRowCount()
{
    if (!m_pHwp) return -1;

    // CurSelectedCtrl 또는 ParentCtrl에서 테이블 정보 조회
    VARIANT parentVar = GetProperty(L"ParentCtrl");
    if (parentVar.vt != VT_DISPATCH || !parentVar.pdispVal) {
        VariantClear(&parentVar);
        return -1;
    }

    IDispatch* pParentCtrl = parentVar.pdispVal;

    // RowCount 속성 조회
    OLECHAR* rowCountName = const_cast<OLECHAR*>(L"RowCount");
    DISPID dispid;
    HRESULT hr = pParentCtrl->GetIDsOfNames(IID_NULL, &rowCountName, 1,
                                             LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pParentCtrl->Release();
        return -1;
    }

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    hr = pParentCtrl->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &noParams, &result, NULL, NULL);

    pParentCtrl->Release();

    if (SUCCEEDED(hr) && result.vt == VT_I4) {
        return result.lVal;
    }

    return -1;
}

int HwpWrapper::GetTableColCount()
{
    if (!m_pHwp) return -1;

    // CurSelectedCtrl 또는 ParentCtrl에서 테이블 정보 조회
    VARIANT parentVar = GetProperty(L"ParentCtrl");
    if (parentVar.vt != VT_DISPATCH || !parentVar.pdispVal) {
        VariantClear(&parentVar);
        return -1;
    }

    IDispatch* pParentCtrl = parentVar.pdispVal;

    // ColCount 속성 조회
    OLECHAR* colCountName = const_cast<OLECHAR*>(L"ColCount");
    DISPID dispid;
    HRESULT hr = pParentCtrl->GetIDsOfNames(IID_NULL, &colCountName, 1,
                                             LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pParentCtrl->Release();
        return -1;
    }

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);
    hr = pParentCtrl->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &noParams, &result, NULL, NULL);

    pParentCtrl->Release();

    if (SUCCEEDED(hr) && result.vt == VT_I4) {
        return result.lVal;
    }

    return -1;
}

bool HwpWrapper::TableLeftCell()
{
    return RunAction(L"TableLeftCell");
}

bool HwpWrapper::TableRightCell()
{
    return RunAction(L"TableRightCell");
}

bool HwpWrapper::TableUpperCell()
{
    return RunAction(L"TableUpperCell");
}

bool HwpWrapper::TableLowerCell()
{
    return RunAction(L"TableLowerCell");
}

bool HwpWrapper::TableRightCellAppend()
{
    return RunAction(L"TableRightCellAppend");
}

bool HwpWrapper::TableColBegin()
{
    return RunAction(L"TableColBegin");
}

bool HwpWrapper::TableColEnd()
{
    return RunAction(L"TableColEnd");
}

bool HwpWrapper::TableColPageUp()
{
    return RunAction(L"TableColPageUp");
}

bool HwpWrapper::TableCellBlockExtendAbs()
{
    return RunAction(L"TableCellBlockExtendAbs");
}

bool HwpWrapper::Cancel()
{
    return RunAction(L"Cancel");
}

bool HwpWrapper::CellFill(int r, int g, int b)
{
    if (!m_pHwp) return false;

    // CellShape 파라미터셋 생성
    IDispatch* pSet = CreateSet(L"CellShape");
    if (!pSet) return false;

    // FillAttr 하위 파라미터셋 접근
    DISPID dispidItem = m_dispidCache.GetOrLoad(pSet, L"Item");
    if (dispidItem == DISPID_UNKNOWN) {
        pSet->Release();
        return false;
    }

    VARIANT argFillAttr;
    VariantInit(&argFillAttr);
    argFillAttr.vt = VT_BSTR;
    argFillAttr.bstrVal = SysAllocString(L"FillAttr");

    DISPPARAMS params = { &argFillAttr, NULL, 1, 0 };
    VARIANT vFillAttr;
    VariantInit(&vFillAttr);

    HRESULT hr = pSet->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYGET, &params, &vFillAttr, NULL, NULL);
    SysFreeString(argFillAttr.bstrVal);

    if (FAILED(hr) || vFillAttr.vt != VT_DISPATCH || !vFillAttr.pdispVal) {
        pSet->Release();
        return false;
    }

    IDispatch* pFillAttr = vFillAttr.pdispVal;

    // FillAttr.Type = 1 (단색)
    DISPID dispidType = m_dispidCache.GetOrLoad(pFillAttr, L"Type");
    if (dispidType != DISPID_UNKNOWN) {
        VARIANT vType;
        VariantInit(&vType);
        vType.vt = VT_I4;
        vType.lVal = 1;

        DISPID dispidPut = DISPID_PROPERTYPUT;
        DISPPARAMS paramsType = { &vType, &dispidPut, 1, 1 };
        pFillAttr->Invoke(dispidType, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_PROPERTYPUT, &paramsType, NULL, NULL, NULL);
    }

    // FaceColor 설정 (RGB -> COLORREF)
    DISPID dispidFaceColor = m_dispidCache.GetOrLoad(pFillAttr, L"FaceColor");
    if (dispidFaceColor != DISPID_UNKNOWN) {
        COLORREF color = RGB(r, g, b);
        VARIANT vColor;
        VariantInit(&vColor);
        vColor.vt = VT_I4;
        vColor.lVal = static_cast<long>(color);

        DISPID dispidPut = DISPID_PROPERTYPUT;
        DISPPARAMS paramsColor = { &vColor, &dispidPut, 1, 1 };
        pFillAttr->Invoke(dispidFaceColor, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_PROPERTYPUT, &paramsColor, NULL, NULL, NULL);
    }

    // CellShape 액션 실행
    IDispatch* pAction = CreateAction(L"CellShape");
    if (!pAction) {
        pFillAttr->Release();
        pSet->Release();
        return false;
    }

    // GetDefault 호출
    DISPID dispidGetDefault = m_dispidCache.GetOrLoad(pAction, L"GetDefault");
    if (dispidGetDefault != DISPID_UNKNOWN) {
        DISPPARAMS paramsEmpty = { NULL, NULL, 0, 0 };
        VARIANT vDefault;
        VariantInit(&vDefault);
        pAction->Invoke(dispidGetDefault, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_METHOD, &paramsEmpty, &vDefault, NULL, NULL);
        VariantClear(&vDefault);
    }

    // Execute
    DISPID dispidExecute = m_dispidCache.GetOrLoad(pAction, L"Execute");
    bool success = false;
    if (dispidExecute != DISPID_UNKNOWN) {
        VARIANT argSet;
        VariantInit(&argSet);
        argSet.vt = VT_DISPATCH;
        argSet.pdispVal = pSet;

        DISPPARAMS paramsExec = { &argSet, NULL, 1, 0 };
        VARIANT vResult;
        VariantInit(&vResult);

        hr = pAction->Invoke(dispidExecute, IID_NULL, LOCALE_USER_DEFAULT,
                             DISPATCH_METHOD, &paramsExec, &vResult, NULL, NULL);
        if (SUCCEEDED(hr) && vResult.vt == VT_BOOL) {
            success = (vResult.boolVal != VARIANT_FALSE);
        }
        VariantClear(&vResult);
    }

    pAction->Release();
    pFillAttr->Release();
    pSet->Release();

    return success;
}

bool HwpWrapper::TableFromData(
    const std::vector<std::vector<std::wstring>>& data,
    bool treat_as_char,
    bool header,
    bool header_bold,
    int cell_fill_r,
    int cell_fill_g,
    int cell_fill_b)
{
    if (data.empty() || data[0].empty()) return false;

    int rows = static_cast<int>(data.size());
    int cols = static_cast<int>(data[0].size());

    // 1. 테이블 생성
    if (!CreateTable(rows, cols, treat_as_char, 0, 0, header)) {
        return false;
    }

    // 2. 데이터 채우기
    for (int row = 0; row < rows; ++row) {
        for (int col = 0; col < static_cast<int>(data[row].size()); ++col) {
            InsertText(data[row][col]);
            if (col < cols - 1) {
                TableRightCell();
            }
        }
        if (row < rows - 1) {
            TableRightCellAppend();  // 다음 행으로 이동
        }
    }

    // 3. 헤더 행 서식 적용
    if (header_bold || cell_fill_r >= 0) {
        // 첫 번째 열로 이동
        TableColBegin();
        // 열의 맨 위로 이동
        TableColPageUp();
        // 셀 블록 선택 확장
        TableCellBlockExtendAbs();
        // 마지막 열까지 선택
        TableColEnd();

        // 볼드 적용
        if (header_bold) {
            RunAction(L"CharShapeBold");
        }

        // 배경색 적용
        if (cell_fill_r >= 0 && cell_fill_g >= 0 && cell_fill_b >= 0) {
            CellFill(cell_fill_r, cell_fill_g, cell_fill_b);
        }

        // 선택 취소
        Cancel();
    }

    return true;
}

std::wstring HwpWrapper::GetTableXml()
{
    // 현재 테이블의 XML 추출 (HWPML2X 형식)
    // Python에서 파싱하여 DataFrame으로 변환

    if (!IsCell()) {
        return L"";
    }

    // 1. 테이블 전체 선택
    // 셀 블록 시작
    TableColBegin();
    TableColPageUp();
    TableCellBlockExtendAbs();
    // 마지막 셀까지 확장
    TableColEnd();
    while (TableLowerCell()) {
        // 끝까지 이동
    }

    // 2. 선택된 테이블을 HWPML2X로 추출
    std::wstring xml = GetTextFile(L"HWPML2X", L"");

    // 3. 선택 취소
    Cancel();

    return xml;
}

//=============================================================================
// 이미지 삽입 (Image Insertion)
//=============================================================================

bool HwpWrapper::InsertPicture(const std::wstring& path,
                                bool embedded,
                                int sizeoption,
                                bool reverse,
                                bool watermark,
                                int effect,
                                int width,
                                int height)
{
    if (!m_pHwp) return false;

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"InsertPicture");
    if (dispid == DISPID_UNKNOWN) return false;

    // InsertPicture 파라미터 (역순)
    // Path, Embedded, sizeoption, Reverse, watermark, Effect, Width, Height
    VARIANT args[8];
    for (int i = 0; i < 8; i++) VariantInit(&args[i]);

    // 역순으로 인자 설정
    args[0].vt = VT_I4;    args[0].lVal = height;                                    // Height
    args[1].vt = VT_I4;    args[1].lVal = width;                                     // Width
    args[2].vt = VT_I4;    args[2].lVal = effect;                                    // Effect
    args[3].vt = VT_BOOL;  args[3].boolVal = watermark ? VARIANT_TRUE : VARIANT_FALSE; // watermark
    args[4].vt = VT_BOOL;  args[4].boolVal = reverse ? VARIANT_TRUE : VARIANT_FALSE;   // Reverse
    args[5].vt = VT_I4;    args[5].lVal = sizeoption;                                // sizeoption
    args[6].vt = VT_BOOL;  args[6].boolVal = embedded ? VARIANT_TRUE : VARIANT_FALSE;  // Embedded
    args[7].vt = VT_BSTR;  args[7].bstrVal = SysAllocString(path.c_str());           // Path

    DISPPARAMS params = { args, NULL, 8, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    SysFreeString(args[7].bstrVal);

    // InsertPicture는 컨트롤 객체 반환 (성공 시 VT_DISPATCH)
    bool success = SUCCEEDED(hr) && result.vt == VT_DISPATCH && result.pdispVal != nullptr;

    if (result.vt == VT_DISPATCH && result.pdispVal) {
        result.pdispVal->Release();
    }
    VariantClear(&result);

    return success;
}

//=============================================================================
// 스타일 관리 (CharShape/ParaShape)
//=============================================================================

std::map<std::wstring, int> HwpWrapper::GetCharShape()
{
    std::map<std::wstring, int> result;
    if (!m_pHwp) return result;

    // hwp.CharShape 속성 가져오기 (HParameterSet.HCharShape가 아님!)
    DISPID dispidCharShape;
    OLECHAR* charShapeName = const_cast<OLECHAR*>(L"CharShape");
    HRESULT hr = m_pHwp->GetIDsOfNames(IID_NULL, &charShapeName, 1,
                                        LOCALE_USER_DEFAULT, &dispidCharShape);
    if (FAILED(hr)) {
        result[L"_error"] = -30000 - (int)(hr & 0xFFFF);
        return result;
    }

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    VARIANT vCharShape;
    VariantInit(&vCharShape);
    hr = m_pHwp->Invoke(dispidCharShape, IID_NULL, LOCALE_USER_DEFAULT,
                         DISPATCH_PROPERTYGET, &noParams, &vCharShape, NULL, NULL);
    if (FAILED(hr) || vCharShape.vt != VT_DISPATCH || !vCharShape.pdispVal) {
        result[L"_error"] = -31000 - (int)(hr & 0xFFFF);
        VariantClear(&vCharShape);
        return result;
    }

    IDispatch* pCharShape = vCharShape.pdispVal;

    // 주요 CharShape 속성들
    const wchar_t* charShapeProps[] = {
        L"Height",          // 글자 크기 (pt * 100)
        L"Bold",            // 볼드
        L"Italic",          // 이탤릭
        L"UnderlineType",   // 밑줄 타입 (0=없음, 1=아래, 3=위)
        L"UnderlineShape",  // 밑줄 모양 (0-12)
        L"UnderlineColor",  // 밑줄 색상
        L"StrikeOutType",   // 취소선 타입 (True/False)
        L"StrikeOutShape",  // 취소선 모양 (0-12)
        L"StrikeOutColor",  // 취소선 색상
        L"TextColor",       // 글자색
        L"ShadeColor",      // 음영색
        L"Superscript",     // 위첨자
        L"Subscript",       // 아래첨자
        L"UseFontSpace",    // 글꼴 간격 사용
        L"UseKerning",      // 커닝 사용
        L"BorderType"       // 테두리 종류
    };

    // Item 메서드 DISPID 가져오기
    DISPID dispidItem;
    OLECHAR* itemName = const_cast<OLECHAR*>(L"Item");
    hr = pCharShape->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT, &dispidItem);
    if (SUCCEEDED(hr)) {
        for (const wchar_t* propName : charShapeProps) {
            VARIANT vPropName;
            VariantInit(&vPropName);
            vPropName.vt = VT_BSTR;
            vPropName.bstrVal = SysAllocString(propName);

            DISPPARAMS itemParams = { &vPropName, NULL, 1, 0 };
            VARIANT vValue;
            VariantInit(&vValue);

            // Item은 메서드이므로 DISPATCH_METHOD 사용
            hr = pCharShape->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                                     DISPATCH_METHOD, &itemParams, &vValue, NULL, NULL);
            if (SUCCEEDED(hr)) {
                switch (vValue.vt) {
                    case VT_I4:
                        result[propName] = vValue.lVal;
                        break;
                    case VT_I2:
                        result[propName] = vValue.iVal;
                        break;
                    case VT_I1:
                        result[propName] = vValue.cVal;
                        break;
                    case VT_UI1:
                        result[propName] = vValue.bVal;
                        break;
                    case VT_UI2:
                        result[propName] = vValue.uiVal;
                        break;
                    case VT_UI4:
                        result[propName] = static_cast<int>(vValue.ulVal);
                        break;
                    case VT_BOOL:
                        result[propName] = (vValue.boolVal != VARIANT_FALSE) ? 1 : 0;
                        break;
                    case VT_R4:
                        result[propName] = static_cast<int>(vValue.fltVal);
                        break;
                    case VT_R8:
                        result[propName] = static_cast<int>(vValue.dblVal);
                        break;
                    case VT_EMPTY:
                    case VT_NULL:
                        // 값 없음 - 스킵
                        break;
                    default:
                        // 디버그: 알 수 없는 타입 (음수로 타입 코드 저장)
                        result[propName] = -1000 - (int)vValue.vt;
                        break;
                }
            }
            VariantClear(&vValue);
            SysFreeString(vPropName.bstrVal);
        }
    } else {
        result[L"_error"] = -20000 - (int)(hr & 0xFFFF);
    }

    pCharShape->Release();
    return result;
}

bool HwpWrapper::SetCharShape(const std::map<std::wstring, int>& props)
{
    if (!m_pHwp || props.empty()) return false;

    // 1. HParameterSet.HCharShape 가져오기
    IDispatch* pHParameterSet = GetHParameterSet();
    if (!pHParameterSet) return false;

    DISPID dispidHCharShape;
    OLECHAR* charShapeName = const_cast<OLECHAR*>(L"HCharShape");
    HRESULT hr = pHParameterSet->GetIDsOfNames(IID_NULL, &charShapeName, 1,
                                                LOCALE_USER_DEFAULT, &dispidHCharShape);
    if (FAILED(hr)) return false;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    VARIANT vCharShape;
    VariantInit(&vCharShape);
    hr = pHParameterSet->Invoke(dispidHCharShape, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &noParams, &vCharShape, NULL, NULL);
    if (FAILED(hr) || vCharShape.vt != VT_DISPATCH || !vCharShape.pdispVal) {
        VariantClear(&vCharShape);
        return false;
    }

    IDispatch* pCharShape = vCharShape.pdispVal;

    // 2. HSet 가져오기 (Execute에 필요)
    DISPID dispidHSet;
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pCharShape->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispidHSet);
    if (FAILED(hr)) {
        pCharShape->Release();
        return false;
    }

    VARIANT vHSet;
    VariantInit(&vHSet);
    hr = pCharShape->Invoke(dispidHSet, IID_NULL, LOCALE_USER_DEFAULT,
                             DISPATCH_PROPERTYGET, &noParams, &vHSet, NULL, NULL);
    if (FAILED(hr) || vHSet.vt != VT_DISPATCH || !vHSet.pdispVal) {
        VariantClear(&vHSet);
        pCharShape->Release();
        return false;
    }

    IDispatch* pHSet = vHSet.pdispVal;

    // 3. 속성값 설정 (각 속성을 직접 설정) - GetDefault 호출 없이!
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pHSet->Release();
        pCharShape->Release();
        return false;
    }

    for (const auto& prop : props) {
        DISPID dispidProp;
        OLECHAR* propName = const_cast<OLECHAR*>(prop.first.c_str());
        hr = pCharShape->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispidProp);
        if (SUCCEEDED(hr)) {
            VARIANT vValue;
            VariantInit(&vValue);
            vValue.vt = VT_I4;
            vValue.lVal = prop.second;

            DISPID putid = DISPID_PROPERTYPUT;
            DISPPARAMS params = { &vValue, &putid, 1, 1 };

            pCharShape->Invoke(dispidProp, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYPUT, &params, NULL, NULL, NULL);
        }
    }

    // 4. HAction.Execute 호출
    bool success = false;
    DISPID dispidExecute;
    OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
    hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1, LOCALE_USER_DEFAULT, &dispidExecute);
    if (SUCCEEDED(hr)) {
        VARIANT args[2];
        VariantInit(&args[0]);
        VariantInit(&args[1]);
        args[0].vt = VT_DISPATCH;
        args[0].pdispVal = pHSet;
        args[1].vt = VT_BSTR;
        args[1].bstrVal = SysAllocString(L"CharShape");

        DISPPARAMS params = { args, NULL, 2, 0 };
        VARIANT vResult;
        VariantInit(&vResult);

        hr = pHAction->Invoke(dispidExecute, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_METHOD, &params, &vResult, NULL, NULL);
        if (SUCCEEDED(hr) && vResult.vt == VT_BOOL) {
            success = (vResult.boolVal != VARIANT_FALSE);
        } else if (SUCCEEDED(hr)) {
            success = true;
        }

        SysFreeString(args[1].bstrVal);
        VariantClear(&vResult);
    }

    pHSet->Release();
    pCharShape->Release();

    return success;
}

bool HwpWrapper::SetFont(const std::wstring& face_name,
                          int height,
                          int bold,
                          int italic,
                          int text_color)
{
    std::map<std::wstring, int> props;

    if (height >= 0) {
        props[L"Height"] = height;
    }
    if (bold >= 0) {
        props[L"Bold"] = bold;
    }
    if (italic >= 0) {
        props[L"Italic"] = italic;
    }
    if (text_color >= 0) {
        props[L"TextColor"] = text_color;
    }

    // 글꼴 이름 설정은 별도 처리 필요 (문자열 속성)
    if (!face_name.empty()) {
        // 간단히 HAction 방식 사용
        if (!m_pHwp) return false;

        IDispatch* pHParameterSet = GetHParameterSet();
        if (!pHParameterSet) return false;

        DISPID dispidHCharShape;
        OLECHAR* charShapeName = const_cast<OLECHAR*>(L"HCharShape");
        HRESULT hr = pHParameterSet->GetIDsOfNames(IID_NULL, &charShapeName, 1,
                                                    LOCALE_USER_DEFAULT, &dispidHCharShape);
        if (SUCCEEDED(hr)) {
            DISPPARAMS noParams = { NULL, NULL, 0, 0 };
            VARIANT vCharShape;
            VariantInit(&vCharShape);
            hr = pHParameterSet->Invoke(dispidHCharShape, IID_NULL, LOCALE_USER_DEFAULT,
                                         DISPATCH_PROPERTYGET, &noParams, &vCharShape, NULL, NULL);
            if (SUCCEEDED(hr) && vCharShape.vt == VT_DISPATCH && vCharShape.pdispVal) {
                IDispatch* pCharShape = vCharShape.pdispVal;

                DISPID dispidItem;
                OLECHAR* itemName = const_cast<OLECHAR*>(L"Item");
                hr = pCharShape->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT, &dispidItem);
                if (SUCCEEDED(hr)) {
                    // FaceNameHangul 설정
                    const wchar_t* fontProps[] = { L"FaceNameHangul", L"FaceNameLatin", L"FaceNameHanja" };
                    for (const wchar_t* fontProp : fontProps) {
                        VARIANT vPropName, vValue;
                        VariantInit(&vPropName);
                        VariantInit(&vValue);

                        vPropName.vt = VT_BSTR;
                        vPropName.bstrVal = SysAllocString(fontProp);
                        vValue.vt = VT_BSTR;
                        vValue.bstrVal = SysAllocString(face_name.c_str());

                        VARIANT args[2] = { vValue, vPropName };
                        DISPID putid = DISPID_PROPERTYPUT;
                        DISPPARAMS itemParams = { args, &putid, 2, 1 };

                        pCharShape->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                                            DISPATCH_PROPERTYPUT, &itemParams, NULL, NULL, NULL);

                        SysFreeString(vPropName.bstrVal);
                        SysFreeString(vValue.bstrVal);
                    }
                }
                pCharShape->Release();
            }
        }
    }

    if (props.empty() && face_name.empty()) {
        return true;  // 변경할 내용 없음
    }

    if (!props.empty()) {
        return SetCharShape(props);
    }

    // 글꼴만 변경한 경우 Execute 호출 필요
    if (!face_name.empty()) {
        IDispatch* pHAction = GetHAction();
        if (!pHAction) return false;

        IDispatch* pHParameterSet = GetHParameterSet();
        if (!pHParameterSet) return false;

        DISPID dispidHCharShape;
        OLECHAR* charShapeName = const_cast<OLECHAR*>(L"HCharShape");
        HRESULT hr = pHParameterSet->GetIDsOfNames(IID_NULL, &charShapeName, 1,
                                                    LOCALE_USER_DEFAULT, &dispidHCharShape);
        if (FAILED(hr)) return false;

        DISPPARAMS noParams = { NULL, NULL, 0, 0 };
        VARIANT vCharShape;
        VariantInit(&vCharShape);
        hr = pHParameterSet->Invoke(dispidHCharShape, IID_NULL, LOCALE_USER_DEFAULT,
                                     DISPATCH_PROPERTYGET, &noParams, &vCharShape, NULL, NULL);
        if (FAILED(hr) || vCharShape.vt != VT_DISPATCH) return false;

        IDispatch* pCharShape = vCharShape.pdispVal;

        DISPID dispidHSet;
        OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
        hr = pCharShape->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispidHSet);
        if (FAILED(hr)) {
            pCharShape->Release();
            return false;
        }

        VARIANT vHSet;
        VariantInit(&vHSet);
        hr = pCharShape->Invoke(dispidHSet, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &noParams, &vHSet, NULL, NULL);
        if (FAILED(hr) || vHSet.vt != VT_DISPATCH) {
            pCharShape->Release();
            return false;
        }

        IDispatch* pHSet = vHSet.pdispVal;

        // Execute
        DISPID dispidExecute;
        OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
        hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1, LOCALE_USER_DEFAULT, &dispidExecute);
        bool success = false;
        if (SUCCEEDED(hr)) {
            VARIANT args[2];
            VariantInit(&args[0]);
            VariantInit(&args[1]);
            args[0].vt = VT_DISPATCH;
            args[0].pdispVal = pHSet;
            args[1].vt = VT_BSTR;
            args[1].bstrVal = SysAllocString(L"CharShape");

            DISPPARAMS params = { args, NULL, 2, 0 };
            VARIANT vResult;
            VariantInit(&vResult);

            hr = pHAction->Invoke(dispidExecute, IID_NULL, LOCALE_USER_DEFAULT,
                                   DISPATCH_METHOD, &params, &vResult, NULL, NULL);
            success = SUCCEEDED(hr);

            SysFreeString(args[1].bstrVal);
            VariantClear(&vResult);
        }

        pHSet->Release();
        pCharShape->Release();
        return success;
    }

    return true;
}

std::map<std::wstring, int> HwpWrapper::GetParaShape()
{
    std::map<std::wstring, int> result;
    if (!m_pHwp) return result;

    // 1. HParameterSet.HParaShape 가져오기
    IDispatch* pHParameterSet = GetHParameterSet();
    if (!pHParameterSet) return result;

    DISPID dispidHParaShape;
    OLECHAR* paraShapeName = const_cast<OLECHAR*>(L"HParaShape");
    HRESULT hr = pHParameterSet->GetIDsOfNames(IID_NULL, &paraShapeName, 1,
                                                LOCALE_USER_DEFAULT, &dispidHParaShape);
    if (FAILED(hr)) return result;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    VARIANT vParaShape;
    VariantInit(&vParaShape);
    hr = pHParameterSet->Invoke(dispidHParaShape, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &noParams, &vParaShape, NULL, NULL);
    if (FAILED(hr) || vParaShape.vt != VT_DISPATCH || !vParaShape.pdispVal) {
        VariantClear(&vParaShape);
        return result;
    }

    IDispatch* pParaShape = vParaShape.pdispVal;

    // 2. HAction.GetDefault 호출
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pParaShape->Release();
        return result;
    }

    DISPID dispidHSet;
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pParaShape->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispidHSet);
    if (FAILED(hr)) {
        pParaShape->Release();
        return result;
    }

    VARIANT vHSet;
    VariantInit(&vHSet);
    hr = pParaShape->Invoke(dispidHSet, IID_NULL, LOCALE_USER_DEFAULT,
                             DISPATCH_PROPERTYGET, &noParams, &vHSet, NULL, NULL);
    if (FAILED(hr) || vHSet.vt != VT_DISPATCH || !vHSet.pdispVal) {
        VariantClear(&vHSet);
        pParaShape->Release();
        return result;
    }

    IDispatch* pHSet = vHSet.pdispVal;

    // GetDefault 호출
    DISPID dispidGetDefault;
    OLECHAR* getDefaultName = const_cast<OLECHAR*>(L"GetDefault");
    hr = pHAction->GetIDsOfNames(IID_NULL, &getDefaultName, 1, LOCALE_USER_DEFAULT, &dispidGetDefault);
    if (SUCCEEDED(hr)) {
        VARIANT args[2];
        VariantInit(&args[0]);
        VariantInit(&args[1]);
        args[0].vt = VT_DISPATCH;
        args[0].pdispVal = pHSet;
        args[1].vt = VT_BSTR;
        args[1].bstrVal = SysAllocString(L"ParaShape");

        DISPPARAMS params = { args, NULL, 2, 0 };
        VARIANT vResult;
        VariantInit(&vResult);
        pHAction->Invoke(dispidGetDefault, IID_NULL, LOCALE_USER_DEFAULT,
                         DISPATCH_METHOD, &params, &vResult, NULL, NULL);
        SysFreeString(args[1].bstrVal);
        VariantClear(&vResult);
    }

    // 3. 속성값 읽기 (주요 ParaShape 속성들)
    const wchar_t* paraShapeProps[] = {
        L"AlignType",       // 정렬 (0=양쪽, 1=왼쪽, 2=가운데, 3=오른쪽)
        L"LineSpacing",     // 줄간격 (%)
        L"LeftMargin",      // 왼쪽 여백
        L"RightMargin",     // 오른쪽 여백
        L"Indentation",     // 첫줄 들여쓰기
        L"PrevSpacing",     // 문단 앞 간격
        L"NextSpacing",     // 문단 뒤 간격
        L"LineSpacingType", // 줄간격 유형
        L"PageBreakBefore", // 이 문단 앞에서 페이지 나누기
        L"KeepWithNext",    // 다음 문단과 함께
        L"WidowOrphan"      // 외톨이줄 방지
    };

    DISPID dispidItem;
    OLECHAR* itemName = const_cast<OLECHAR*>(L"Item");
    hr = pParaShape->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT, &dispidItem);
    if (SUCCEEDED(hr)) {
        for (const wchar_t* propName : paraShapeProps) {
            VARIANT vPropName;
            VariantInit(&vPropName);
            vPropName.vt = VT_BSTR;
            vPropName.bstrVal = SysAllocString(propName);

            DISPPARAMS itemParams = { &vPropName, NULL, 1, 0 };
            VARIANT vValue;
            VariantInit(&vValue);

            hr = pParaShape->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                                     DISPATCH_PROPERTYGET, &itemParams, &vValue, NULL, NULL);
            if (SUCCEEDED(hr)) {
                if (vValue.vt == VT_I4) {
                    result[propName] = vValue.lVal;
                } else if (vValue.vt == VT_I2) {
                    result[propName] = vValue.iVal;
                } else if (vValue.vt == VT_BOOL) {
                    result[propName] = (vValue.boolVal != VARIANT_FALSE) ? 1 : 0;
                }
            }
            VariantClear(&vValue);
            SysFreeString(vPropName.bstrVal);
        }
    }

    pHSet->Release();
    pParaShape->Release();

    return result;
}

bool HwpWrapper::SetParaShape(const std::map<std::wstring, int>& props)
{
    if (!m_pHwp || props.empty()) return false;

    // 1. HParameterSet.HParaShape 가져오기
    IDispatch* pHParameterSet = GetHParameterSet();
    if (!pHParameterSet) return false;

    DISPID dispidHParaShape;
    OLECHAR* paraShapeName = const_cast<OLECHAR*>(L"HParaShape");
    HRESULT hr = pHParameterSet->GetIDsOfNames(IID_NULL, &paraShapeName, 1,
                                                LOCALE_USER_DEFAULT, &dispidHParaShape);
    if (FAILED(hr)) return false;

    DISPPARAMS noParams = { NULL, NULL, 0, 0 };
    VARIANT vParaShape;
    VariantInit(&vParaShape);
    hr = pHParameterSet->Invoke(dispidHParaShape, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &noParams, &vParaShape, NULL, NULL);
    if (FAILED(hr) || vParaShape.vt != VT_DISPATCH || !vParaShape.pdispVal) {
        VariantClear(&vParaShape);
        return false;
    }

    IDispatch* pParaShape = vParaShape.pdispVal;

    // 2. HAction.GetDefault 호출
    IDispatch* pHAction = GetHAction();
    if (!pHAction) {
        pParaShape->Release();
        return false;
    }

    DISPID dispidHSet;
    OLECHAR* hsetName = const_cast<OLECHAR*>(L"HSet");
    hr = pParaShape->GetIDsOfNames(IID_NULL, &hsetName, 1, LOCALE_USER_DEFAULT, &dispidHSet);
    if (FAILED(hr)) {
        pParaShape->Release();
        return false;
    }

    VARIANT vHSet;
    VariantInit(&vHSet);
    hr = pParaShape->Invoke(dispidHSet, IID_NULL, LOCALE_USER_DEFAULT,
                             DISPATCH_PROPERTYGET, &noParams, &vHSet, NULL, NULL);
    if (FAILED(hr) || vHSet.vt != VT_DISPATCH || !vHSet.pdispVal) {
        VariantClear(&vHSet);
        pParaShape->Release();
        return false;
    }

    IDispatch* pHSet = vHSet.pdispVal;

    // GetDefault 호출
    DISPID dispidGetDefault;
    OLECHAR* getDefaultName = const_cast<OLECHAR*>(L"GetDefault");
    hr = pHAction->GetIDsOfNames(IID_NULL, &getDefaultName, 1, LOCALE_USER_DEFAULT, &dispidGetDefault);
    if (SUCCEEDED(hr)) {
        VARIANT args[2];
        VariantInit(&args[0]);
        VariantInit(&args[1]);
        args[0].vt = VT_DISPATCH;
        args[0].pdispVal = pHSet;
        args[1].vt = VT_BSTR;
        args[1].bstrVal = SysAllocString(L"ParaShape");

        DISPPARAMS params = { args, NULL, 2, 0 };
        VARIANT vResult;
        VariantInit(&vResult);
        pHAction->Invoke(dispidGetDefault, IID_NULL, LOCALE_USER_DEFAULT,
                         DISPATCH_METHOD, &params, &vResult, NULL, NULL);
        SysFreeString(args[1].bstrVal);
        VariantClear(&vResult);
    }

    // 3. 속성값 설정
    DISPID dispidItem;
    OLECHAR* itemName = const_cast<OLECHAR*>(L"Item");
    hr = pParaShape->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT, &dispidItem);
    if (SUCCEEDED(hr)) {
        for (const auto& prop : props) {
            VARIANT vPropName, vValue;
            VariantInit(&vPropName);
            VariantInit(&vValue);

            vPropName.vt = VT_BSTR;
            vPropName.bstrVal = SysAllocString(prop.first.c_str());
            vValue.vt = VT_I4;
            vValue.lVal = prop.second;

            VARIANT args[2] = { vValue, vPropName };
            DISPID putid = DISPID_PROPERTYPUT;
            DISPPARAMS itemParams = { args, &putid, 2, 1 };

            pParaShape->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_PROPERTYPUT, &itemParams, NULL, NULL, NULL);

            SysFreeString(vPropName.bstrVal);
        }
    }

    // 4. HAction.Execute 호출
    bool success = false;
    DISPID dispidExecute;
    OLECHAR* executeName = const_cast<OLECHAR*>(L"Execute");
    hr = pHAction->GetIDsOfNames(IID_NULL, &executeName, 1, LOCALE_USER_DEFAULT, &dispidExecute);
    if (SUCCEEDED(hr)) {
        VARIANT args[2];
        VariantInit(&args[0]);
        VariantInit(&args[1]);
        args[0].vt = VT_DISPATCH;
        args[0].pdispVal = pHSet;
        args[1].vt = VT_BSTR;
        args[1].bstrVal = SysAllocString(L"ParaShape");

        DISPPARAMS params = { args, NULL, 2, 0 };
        VARIANT vResult;
        VariantInit(&vResult);

        hr = pHAction->Invoke(dispidExecute, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_METHOD, &params, &vResult, NULL, NULL);
        if (SUCCEEDED(hr) && vResult.vt == VT_BOOL) {
            success = (vResult.boolVal != VARIANT_FALSE);
        } else if (SUCCEEDED(hr)) {
            success = true;
        }

        SysFreeString(args[1].bstrVal);
        VariantClear(&vResult);
    }

    pHSet->Release();
    pParaShape->Release();

    return success;
}

bool HwpWrapper::SetPara(int align_type,
                          int line_spacing,
                          int left_margin,
                          int indentation)
{
    std::map<std::wstring, int> props;

    if (align_type >= 0) {
        props[L"AlignType"] = align_type;
    }
    if (line_spacing >= 0) {
        props[L"LineSpacing"] = line_spacing;
    }
    if (left_margin >= 0) {
        props[L"LeftMargin"] = left_margin;
    }
    if (indentation >= 0) {
        props[L"Indentation"] = indentation;
    }

    if (props.empty()) {
        return true;  // 변경할 내용 없음
    }

    return SetParaShape(props);
}

//=============================================================================
// COM 속성 접근 (Low-level)
//=============================================================================

IDispatch* HwpWrapper::GetApplication()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"Application");
    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;  // 참조 카운트 유지
    }
    return nullptr;
}

std::wstring HwpWrapper::GetCLSID()
{
    VARIANT result = GetProperty(L"CLSID");
    if (result.vt == VT_BSTR && result.bstrVal) {
        std::wstring clsid(result.bstrVal);
        SysFreeString(result.bstrVal);
        return clsid;
    }
    return L"";
}

int HwpWrapper::GetCurFieldState()
{
    VARIANT result = GetProperty(L"CurFieldState");
    if (result.vt == VT_I4) {
        return result.lVal;
    } else if (result.vt == VT_I2) {
        return result.iVal;
    }
    return 0;
}

int HwpWrapper::GetCurMetatagState()
{
    VARIANT result = GetProperty(L"CurMetatagState");
    if (result.vt == VT_I4) {
        return result.lVal;
    } else if (result.vt == VT_I2) {
        return result.iVal;
    }
    return 0;
}

IDispatch* HwpWrapper::GetEngineProperties()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"EngineProperties");
    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    return nullptr;
}

bool HwpWrapper::GetIsPrivateInfoProtected()
{
    VARIANT result = GetProperty(L"IsPrivateInfoProtected");
    if (result.vt == VT_BOOL) {
        return result.boolVal != VARIANT_FALSE;
    }
    return false;
}

bool HwpWrapper::GetIsTrackChange()
{
    VARIANT result = GetProperty(L"IsTrackChange");
    if (result.vt == VT_BOOL) {
        return result.boolVal != VARIANT_FALSE;
    }
    return false;
}

std::wstring HwpWrapper::GetDocPath()
{
    VARIANT result = GetProperty(L"Path");
    if (result.vt == VT_BSTR && result.bstrVal) {
        std::wstring path(result.bstrVal);
        SysFreeString(result.bstrVal);
        return path;
    }
    return L"";
}

int HwpWrapper::GetSelectionMode()
{
    VARIANT result = GetProperty(L"SelectionMode");
    if (result.vt == VT_I4) {
        return result.lVal;
    } else if (result.vt == VT_I2) {
        return result.iVal;
    }
    return 0;
}

std::wstring HwpWrapper::GetTitle()
{
    // pyhwpx의 get_title() 구현 참조
    // hwp.XHwpWindows.Active_XHwpWindow.Caption 접근
    if (!m_pHwp) return L"";

    // XHwpWindows 가져오기
    VARIANT vWindows = GetProperty(L"XHwpWindows");
    if (vWindows.vt != VT_DISPATCH || !vWindows.pdispVal) {
        return L"";
    }

    IDispatch* pWindows = vWindows.pdispVal;

    // Active_XHwpWindow 가져오기
    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(L"Active_XHwpWindow");
    HRESULT hr = pWindows->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pWindows->Release();
        return L"";
    }

    DISPPARAMS params = { nullptr, nullptr, 0, 0 };
    VARIANT vActiveWindow;
    VariantInit(&vActiveWindow);
    hr = pWindows->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &vActiveWindow, nullptr, nullptr);
    pWindows->Release();

    if (FAILED(hr) || vActiveWindow.vt != VT_DISPATCH || !vActiveWindow.pdispVal) {
        VariantClear(&vActiveWindow);
        return L"";
    }

    IDispatch* pActiveWindow = vActiveWindow.pdispVal;

    // Caption 가져오기
    propName = const_cast<OLECHAR*>(L"Caption");
    hr = pActiveWindow->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pActiveWindow->Release();
        return L"";
    }

    VARIANT vCaption;
    VariantInit(&vCaption);
    hr = pActiveWindow->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &vCaption, nullptr, nullptr);
    pActiveWindow->Release();

    std::wstring title;
    if (SUCCEEDED(hr) && vCaption.vt == VT_BSTR && vCaption.bstrVal) {
        title = vCaption.bstrVal;
    }
    VariantClear(&vCaption);

    return title;
}

IDispatch* HwpWrapper::GetViewProperties()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"ViewProperties");
    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    return nullptr;
}

void HwpWrapper::SetViewProperties(IDispatch* props)
{
    if (!m_pHwp || !props) return;

    VARIANT vProps;
    VariantInit(&vProps);
    vProps.vt = VT_DISPATCH;
    vProps.pdispVal = props;

    SetProperty(L"ViewProperties", vProps);
}

IDispatch* HwpWrapper::GetXHwpMessageBox()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"XHwpMessageBox");
    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    return nullptr;
}

IDispatch* HwpWrapper::GetXHwpODBC()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"XHwpODBC");
    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    return nullptr;
}

IDispatch* HwpWrapper::GetXHwpWindows()
{
    if (!m_pHwp) return nullptr;

    VARIANT result = GetProperty(L"XHwpWindows");
    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    return nullptr;
}

std::wstring HwpWrapper::GetCurrentFont()
{
    // HParameterSet.HCharShape.Item("FaceNameHangul") 접근
    if (!m_pHwp) return L"";

    // 1. HParameterSet 가져오기
    IDispatch* pHParamSet = GetHParameterSet();
    if (!pHParamSet) return L"";

    // 2. HCharShape 가져오기
    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(L"HCharShape");
    HRESULT hr = pHParamSet->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return L"";

    DISPPARAMS params = { nullptr, nullptr, 0, 0 };
    VARIANT vCharShape;
    VariantInit(&vCharShape);
    hr = pHParamSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &params, &vCharShape, nullptr, nullptr);
    if (FAILED(hr) || vCharShape.vt != VT_DISPATCH || !vCharShape.pdispVal) {
        VariantClear(&vCharShape);
        return L"";
    }

    IDispatch* pCharShape = vCharShape.pdispVal;

    // 3. HAction.GetDefault("CharShape") 호출하여 현재 설정 로드
    IDispatch* pHAction = GetHAction();
    if (pHAction) {
        OLECHAR* methodName = const_cast<OLECHAR*>(L"GetDefault");
        DISPID dispidGetDefault;
        hr = pHAction->GetIDsOfNames(IID_NULL, &methodName, 1, LOCALE_USER_DEFAULT, &dispidGetDefault);
        if (SUCCEEDED(hr)) {
            VARIANT vActName, vParamSet;
            VariantInit(&vActName);
            VariantInit(&vParamSet);
            vActName.vt = VT_BSTR;
            vActName.bstrVal = SysAllocString(L"CharShape");
            vParamSet.vt = VT_DISPATCH;
            vParamSet.pdispVal = pCharShape;

            VARIANT args[2] = { vParamSet, vActName };
            DISPPARAMS getDefaultParams = { args, nullptr, 2, 0 };
            VARIANT vResult;
            VariantInit(&vResult);
            pHAction->Invoke(dispidGetDefault, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &getDefaultParams, &vResult, nullptr, nullptr);
            SysFreeString(vActName.bstrVal);
            VariantClear(&vResult);
        }
    }

    // 4. Item("FaceNameHangul") 호출
    propName = const_cast<OLECHAR*>(L"Item");
    hr = pCharShape->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        pCharShape->Release();
        return L"";
    }

    VARIANT vPropName;
    VariantInit(&vPropName);
    vPropName.vt = VT_BSTR;
    vPropName.bstrVal = SysAllocString(L"FaceNameHangul");

    VARIANT args[1] = { vPropName };
    DISPPARAMS itemParams = { args, nullptr, 1, 0 };
    VARIANT vFontName;
    VariantInit(&vFontName);

    hr = pCharShape->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &itemParams, &vFontName, nullptr, nullptr);
    SysFreeString(vPropName.bstrVal);
    pCharShape->Release();

    std::wstring fontName;
    if (SUCCEEDED(hr) && vFontName.vt == VT_BSTR && vFontName.bstrVal) {
        fontName = vFontName.bstrVal;
    }
    VariantClear(&vFontName);

    return fontName;
}

} // namespace cpyhwpx
