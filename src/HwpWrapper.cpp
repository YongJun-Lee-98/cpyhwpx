/**
 * @file HwpWrapper.cpp
 * @brief HWP COM 객체 래퍼 클래스 구현
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 */

#include "HwpWrapper.h"
#include <stdexcept>
#include <atlbase.h>
#include <atlcom.h>

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

HwpWrapper::HwpWrapper(bool visible, bool new_instance)
    : m_pHwp(nullptr)
    , m_pHParameterSet(nullptr)
    , m_pHAction(nullptr)
    , m_bInitialized(false)
    , m_bVisible(visible)
    , m_bNewInstance(new_instance)
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
    // HAction을 통한 찾기 구현
    // TODO: CreateAction("Find") 사용
    return false;
}

bool HwpWrapper::Replace(const std::wstring& find_text,
                          const std::wstring& replace_text,
                          bool forward,
                          bool match_case,
                          bool regex)
{
    // TODO: 구현
    return false;
}

int HwpWrapper::ReplaceAll(const std::wstring& find_text,
                            const std::wstring& replace_text,
                            bool match_case,
                            bool regex)
{
    // TODO: 구현
    return 0;
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

bool HwpWrapper::SetPosBySet(IDispatch* pDispVal)
{
    if (!m_pHwp || !pDispVal) return false;

    DISPID dispid = m_dispidCache.GetOrLoad(m_pHwp, L"SetPosBySet");
    if (dispid == DISPID_UNKNOWN) return false;

    // SetPosBySet(dispVal) - positional parameter
    VARIANT arg;
    VariantInit(&arg);
    arg.vt = VT_DISPATCH;
    arg.pdispVal = pDispVal;

    DISPPARAMS params = { &arg, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    HRESULT hr = m_pHwp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                                DISPATCH_METHOD, &params, &result, NULL, NULL);

    bool success = SUCCEEDED(hr);
    VariantClear(&result);

    return success;
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

} // namespace cpyhwpx
