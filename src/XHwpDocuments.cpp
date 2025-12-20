/**
 * @file XHwpDocuments.cpp
 * @brief HWP 문서 컬렉션 래퍼 클래스 구현
 */

#include "XHwpDocuments.h"
#include <stdexcept>

namespace cpyhwpx {

//=============================================================================
// 생성자/소멸자
//=============================================================================

XHwpDocuments::XHwpDocuments(IDispatch* pDocuments)
    : m_pDocuments(pDocuments)
{
    if (m_pDocuments) {
        m_pDocuments->AddRef();
    }
}

XHwpDocuments::~XHwpDocuments()
{
    if (m_pDocuments) {
        m_pDocuments->Release();
        m_pDocuments = nullptr;
    }
}

XHwpDocuments::XHwpDocuments(XHwpDocuments&& other) noexcept
    : m_pDocuments(other.m_pDocuments)
{
    other.m_pDocuments = nullptr;
}

XHwpDocuments& XHwpDocuments::operator=(XHwpDocuments&& other) noexcept
{
    if (this != &other) {
        if (m_pDocuments) {
            m_pDocuments->Release();
        }
        m_pDocuments = other.m_pDocuments;
        other.m_pDocuments = nullptr;
    }
    return *this;
}

//=============================================================================
// 컬렉션 속성
//=============================================================================

int XHwpDocuments::GetCount() const
{
    if (!m_pDocuments) return 0;

    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(L"Count");
    HRESULT hr = m_pDocuments->GetIDsOfNames(IID_NULL, &propName, 1,
                                              LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return 0;

    DISPPARAMS params = { nullptr, nullptr, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocuments->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYGET, &params, &result, nullptr, nullptr);
    if (FAILED(hr)) {
        VariantClear(&result);
        return 0;
    }

    int count = 0;
    if (result.vt == VT_I4) {
        count = result.lVal;
    } else if (result.vt == VT_I2) {
        count = result.iVal;
    }
    VariantClear(&result);
    return count;
}

std::unique_ptr<XHwpDocument> XHwpDocuments::GetActiveDocument() const
{
    if (!m_pDocuments) return nullptr;

    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(L"Active_XHwpDocument");
    HRESULT hr = m_pDocuments->GetIDsOfNames(IID_NULL, &propName, 1,
                                              LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return nullptr;

    DISPPARAMS params = { nullptr, nullptr, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocuments->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYGET, &params, &result, nullptr, nullptr);
    if (FAILED(hr)) {
        VariantClear(&result);
        return nullptr;
    }

    std::unique_ptr<XHwpDocument> doc;
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        doc = std::make_unique<XHwpDocument>(result.pdispVal);
    }
    VariantClear(&result);
    return doc;
}

//=============================================================================
// 문서 접근
//=============================================================================

std::unique_ptr<XHwpDocument> XHwpDocuments::Item(int index) const
{
    if (!m_pDocuments) {
        throw std::runtime_error("XHwpDocuments is not valid");
    }

    int count = GetCount();

    // 음수 인덱스 지원
    if (index < 0) {
        index += count;
    }

    // 범위 검사
    if (index < 0 || index >= count) {
        throw std::out_of_range("Index out of range");
    }

    // Item 메서드 호출
    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(L"Item");
    HRESULT hr = m_pDocuments->GetIDsOfNames(IID_NULL, &methodName, 1,
                                              LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        throw std::runtime_error("Failed to get Item method");
    }

    VARIANT vIndex;
    VariantInit(&vIndex);
    vIndex.vt = VT_I4;
    vIndex.lVal = index;

    DISPPARAMS params = { &vIndex, nullptr, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    // DISPATCH_PROPERTYGET로 먼저 시도 (COM 컬렉션 표준)
    hr = m_pDocuments->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYGET | DISPATCH_METHOD, &params, &result, nullptr, nullptr);
    VariantClear(&vIndex);

    if (FAILED(hr)) {
        VariantClear(&result);
        throw std::runtime_error("Failed to invoke Item method");
    }

    std::unique_ptr<XHwpDocument> doc;
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        doc = std::make_unique<XHwpDocument>(result.pdispVal);
    }
    VariantClear(&result);
    return doc;
}

std::unique_ptr<XHwpDocument> XHwpDocuments::FindItem(int docID) const
{
    if (!m_pDocuments) return nullptr;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(L"FindItem");
    HRESULT hr = m_pDocuments->GetIDsOfNames(IID_NULL, &methodName, 1,
                                              LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return nullptr;

    VARIANT vDocID;
    VariantInit(&vDocID);
    vDocID.vt = VT_I4;
    vDocID.lVal = docID;

    DISPPARAMS params = { &vDocID, nullptr, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocuments->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_METHOD, &params, &result, nullptr, nullptr);
    VariantClear(&vDocID);

    if (FAILED(hr)) {
        VariantClear(&result);
        return nullptr;
    }

    std::unique_ptr<XHwpDocument> doc;
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        doc = std::make_unique<XHwpDocument>(result.pdispVal);
    }
    VariantClear(&result);
    return doc;
}

//=============================================================================
// 문서 관리
//=============================================================================

std::unique_ptr<XHwpDocument> XHwpDocuments::Add(bool isTab)
{
    if (!m_pDocuments) return nullptr;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(L"Add");
    HRESULT hr = m_pDocuments->GetIDsOfNames(IID_NULL, &methodName, 1,
                                              LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return nullptr;

    VARIANT vIsTab;
    VariantInit(&vIsTab);
    vIsTab.vt = VT_BOOL;
    vIsTab.boolVal = isTab ? VARIANT_TRUE : VARIANT_FALSE;

    DISPPARAMS params = { &vIsTab, nullptr, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pDocuments->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_METHOD, &params, &result, nullptr, nullptr);
    VariantClear(&vIsTab);

    if (FAILED(hr)) {
        VariantClear(&result);
        return nullptr;
    }

    std::unique_ptr<XHwpDocument> doc;
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        doc = std::make_unique<XHwpDocument>(result.pdispVal);
    }
    VariantClear(&result);
    return doc;
}

void XHwpDocuments::Close(bool isDirty)
{
    if (!m_pDocuments) return;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(L"Close");
    HRESULT hr = m_pDocuments->GetIDsOfNames(IID_NULL, &methodName, 1,
                                              LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return;

    VARIANT vIsDirty;
    VariantInit(&vIsDirty);
    vIsDirty.vt = VT_BOOL;
    vIsDirty.boolVal = isDirty ? VARIANT_TRUE : VARIANT_FALSE;

    DISPPARAMS params = { &vIsDirty, nullptr, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    m_pDocuments->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_METHOD, &params, &result, nullptr, nullptr);
    VariantClear(&vIsDirty);
    VariantClear(&result);
}

} // namespace cpyhwpx
