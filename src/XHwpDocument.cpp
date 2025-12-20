/**
 * @file XHwpDocument.cpp
 * @brief HWP 개별 문서 래퍼 클래스 구현
 */

#include "XHwpDocument.h"
#include <stdexcept>

namespace cpyhwpx {

//=============================================================================
// 생성자/소멸자
//=============================================================================

XHwpDocument::XHwpDocument(IDispatch* pDocument)
    : m_pDocument(pDocument)
{
    if (m_pDocument) {
        m_pDocument->AddRef();
    }
}

XHwpDocument::~XHwpDocument()
{
    if (m_pDocument) {
        m_pDocument->Release();
        m_pDocument = nullptr;
    }
}

XHwpDocument::XHwpDocument(XHwpDocument&& other) noexcept
    : m_pDocument(other.m_pDocument)
{
    other.m_pDocument = nullptr;
}

XHwpDocument& XHwpDocument::operator=(XHwpDocument&& other) noexcept
{
    if (this != &other) {
        if (m_pDocument) {
            m_pDocument->Release();
        }
        m_pDocument = other.m_pDocument;
        other.m_pDocument = nullptr;
    }
    return *this;
}

//=============================================================================
// 헬퍼 메서드
//=============================================================================

std::wstring XHwpDocument::GetStringProperty(const std::wstring& name) const
{
    if (!m_pDocument) return L"";

    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = m_pDocument->GetIDsOfNames(IID_NULL, &propName, 1,
                                             LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return L"";

    DISPPARAMS params = { nullptr, nullptr, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &params, &result, nullptr, nullptr);
    if (FAILED(hr)) {
        VariantClear(&result);
        return L"";
    }

    std::wstring value;
    if (result.vt == VT_BSTR && result.bstrVal) {
        value = result.bstrVal;
    }
    VariantClear(&result);
    return value;
}

int XHwpDocument::GetIntProperty(const std::wstring& name) const
{
    if (!m_pDocument) return 0;

    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = m_pDocument->GetIDsOfNames(IID_NULL, &propName, 1,
                                             LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return 0;

    DISPPARAMS params = { nullptr, nullptr, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &params, &result, nullptr, nullptr);
    if (FAILED(hr)) {
        VariantClear(&result);
        return 0;
    }

    int value = 0;
    if (result.vt == VT_I4) {
        value = result.lVal;
    } else if (result.vt == VT_I2) {
        value = result.iVal;
    } else if (result.vt == VT_BOOL) {
        value = result.boolVal ? 1 : 0;
    }
    VariantClear(&result);
    return value;
}

bool XHwpDocument::GetBoolProperty(const std::wstring& name) const
{
    if (!m_pDocument) return false;

    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = m_pDocument->GetIDsOfNames(IID_NULL, &propName, 1,
                                             LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    DISPPARAMS params = { nullptr, nullptr, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &params, &result, nullptr, nullptr);
    if (FAILED(hr)) {
        VariantClear(&result);
        return false;
    }

    bool value = false;
    if (result.vt == VT_BOOL) {
        value = (result.boolVal != VARIANT_FALSE);
    } else if (result.vt == VT_I4) {
        value = (result.lVal != 0);
    } else if (result.vt == VT_I2) {
        value = (result.iVal != 0);
    }
    VariantClear(&result);
    return value;
}

bool XHwpDocument::InvokeMethod(const std::wstring& name)
{
    if (!m_pDocument) return false;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = m_pDocument->GetIDsOfNames(IID_NULL, &methodName, 1,
                                             LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    DISPPARAMS params = { nullptr, nullptr, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_METHOD, &params, &result, nullptr, nullptr);
    VariantClear(&result);
    return SUCCEEDED(hr);
}

bool XHwpDocument::InvokeMethodInt(const std::wstring& name, int arg)
{
    if (!m_pDocument) return false;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = m_pDocument->GetIDsOfNames(IID_NULL, &methodName, 1,
                                             LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    VARIANT vArg;
    VariantInit(&vArg);
    vArg.vt = VT_I4;
    vArg.lVal = arg;

    DISPPARAMS params = { &vArg, nullptr, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_METHOD, &params, &result, nullptr, nullptr);
    VariantClear(&vArg);
    VariantClear(&result);
    return SUCCEEDED(hr);
}

bool XHwpDocument::InvokeMethodBool(const std::wstring& name, bool arg)
{
    if (!m_pDocument) return false;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = m_pDocument->GetIDsOfNames(IID_NULL, &methodName, 1,
                                             LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    VARIANT vArg;
    VariantInit(&vArg);
    vArg.vt = VT_BOOL;
    vArg.boolVal = arg ? VARIANT_TRUE : VARIANT_FALSE;

    DISPPARAMS params = { &vArg, nullptr, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_METHOD, &params, &result, nullptr, nullptr);
    VariantClear(&vArg);
    VariantClear(&result);
    return SUCCEEDED(hr);
}

//=============================================================================
// 읽기 전용 속성
//=============================================================================

int XHwpDocument::GetDocumentID() const
{
    return GetIntProperty(L"DocumentID");
}

std::wstring XHwpDocument::GetFullName() const
{
    return GetStringProperty(L"FullName");
}

std::wstring XHwpDocument::GetPath() const
{
    return GetStringProperty(L"Path");
}

std::wstring XHwpDocument::GetFormat() const
{
    return GetStringProperty(L"Format");
}

bool XHwpDocument::IsModified() const
{
    return GetBoolProperty(L"Modified");
}

int XHwpDocument::GetEditMode() const
{
    return GetIntProperty(L"EditMode");
}

//=============================================================================
// 문서 활성화
//=============================================================================

bool XHwpDocument::SetActiveDocument()
{
    return InvokeMethod(L"SetActive_XHwpDocument");
}

//=============================================================================
// 파일 작업
//=============================================================================

bool XHwpDocument::Open(const std::wstring& filename,
                         const std::wstring& format,
                         const std::wstring& arg)
{
    if (!m_pDocument) return false;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(L"Open");
    HRESULT hr = m_pDocument->GetIDsOfNames(IID_NULL, &methodName, 1,
                                             LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    // Open(filename, Format, arg) - 역순으로 인자 배열
    VARIANT args[3];
    VariantInit(&args[0]);
    VariantInit(&args[1]);
    VariantInit(&args[2]);

    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(filename.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(format.c_str());
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(arg.c_str());

    DISPPARAMS params = { args, nullptr, 3, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_METHOD, &params, &result, nullptr, nullptr);

    for (auto& a : args) VariantClear(&a);
    VariantClear(&result);
    return SUCCEEDED(hr);
}

bool XHwpDocument::Save(bool saveIfDirty)
{
    return InvokeMethodBool(L"Save", saveIfDirty);
}

bool XHwpDocument::SaveAs(const std::wstring& path,
                           const std::wstring& format,
                           const std::wstring& arg)
{
    if (!m_pDocument) return false;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(L"SaveAs");
    HRESULT hr = m_pDocument->GetIDsOfNames(IID_NULL, &methodName, 1,
                                             LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    // SaveAs(Path, Format, arg) - 역순으로 인자 배열
    VARIANT args[3];
    VariantInit(&args[0]);
    VariantInit(&args[1]);
    VariantInit(&args[2]);

    args[2].vt = VT_BSTR;
    args[2].bstrVal = SysAllocString(path.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(format.c_str());
    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(arg.c_str());

    DISPPARAMS params = { args, nullptr, 3, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_METHOD, &params, &result, nullptr, nullptr);

    for (auto& a : args) VariantClear(&a);
    VariantClear(&result);
    return SUCCEEDED(hr);
}

void XHwpDocument::Close(bool isDirty)
{
    InvokeMethodBool(L"Close", isDirty);
}

void XHwpDocument::Clear(bool option)
{
    InvokeMethodBool(L"Clear", option);
}

//=============================================================================
// 편집 작업
//=============================================================================

bool XHwpDocument::Undo(int count)
{
    return InvokeMethodInt(L"Undo", count);
}

bool XHwpDocument::Redo(int count)
{
    return InvokeMethodInt(L"Redo", count);
}

} // namespace cpyhwpx
