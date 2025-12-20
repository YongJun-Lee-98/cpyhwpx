/**
 * @file HwpParameter.cpp
 * @brief HParameterSet COM 객체 래퍼 클래스 구현
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 */

#include "HwpParameter.h"
#include "HwpWrapper.h"

namespace cpyhwpx {

//=============================================================================
// HwpParameterSet 구현
//=============================================================================

HwpParameterSet::HwpParameterSet(IDispatch* pset)
    : m_pSet(pset)
{
    if (m_pSet) {
        m_pSet->AddRef();
    }
}

HwpParameterSet::~HwpParameterSet()
{
    if (m_pSet) {
        m_pSet->Release();
        m_pSet = nullptr;
    }
}

//=============================================================================
// 파라미터 설정
//=============================================================================

bool HwpParameterSet::SetItem(const std::wstring& name, int value)
{
    VARIANT var;
    VariantInit(&var);
    var.vt = VT_I4;
    var.lVal = value;
    return SetItemInternal(name, var);
}

bool HwpParameterSet::SetItem(const std::wstring& name, double value)
{
    VARIANT var;
    VariantInit(&var);
    var.vt = VT_R8;
    var.dblVal = value;
    return SetItemInternal(name, var);
}

bool HwpParameterSet::SetItem(const std::wstring& name, const std::wstring& value)
{
    VARIANT var;
    VariantInit(&var);
    var.vt = VT_BSTR;
    var.bstrVal = SysAllocString(value.c_str());

    bool result = SetItemInternal(name, var);
    SysFreeString(var.bstrVal);
    return result;
}

bool HwpParameterSet::SetItem(const std::wstring& name, bool value)
{
    VARIANT var;
    VariantInit(&var);
    var.vt = VT_BOOL;
    var.boolVal = value ? VARIANT_TRUE : VARIANT_FALSE;
    return SetItemInternal(name, var);
}

bool HwpParameterSet::SetItem(const std::wstring& name, const VARIANT& value)
{
    VARIANT var;
    VariantInit(&var);
    VariantCopy(&var, const_cast<VARIANT*>(&value));

    bool result = SetItemInternal(name, var);
    VariantClear(&var);
    return result;
}

bool HwpParameterSet::SetItemInternal(const std::wstring& name, VARIANT& value)
{
    if (!m_pSet) return false;

    // Item 속성에 값 설정
    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(L"Item");
    HRESULT hr = m_pSet->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    // 파라미터 이름
    VARIANT nameVar;
    VariantInit(&nameVar);
    nameVar.vt = VT_BSTR;
    nameVar.bstrVal = SysAllocString(name.c_str());

    // 인자: [value, name] (역순)
    VARIANT args[2] = { value, nameVar };

    DISPID putid = DISPID_PROPERTYPUT;
    DISPPARAMS params = { args, &putid, 2, 1 };

    hr = m_pSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                        &params, NULL, NULL, NULL);

    SysFreeString(nameVar.bstrVal);

    return SUCCEEDED(hr);
}

//=============================================================================
// 파라미터 가져오기
//=============================================================================

int HwpParameterSet::GetItemInt(const std::wstring& name, int default_value)
{
    VARIANT result = GetItemInternal(name);
    if (result.vt == VT_I4) {
        return result.lVal;
    }
    if (result.vt == VT_I2) {
        return result.iVal;
    }
    VariantClear(&result);
    return default_value;
}

double HwpParameterSet::GetItemDouble(const std::wstring& name, double default_value)
{
    VARIANT result = GetItemInternal(name);
    if (result.vt == VT_R8) {
        return result.dblVal;
    }
    if (result.vt == VT_R4) {
        return result.fltVal;
    }
    VariantClear(&result);
    return default_value;
}

std::wstring HwpParameterSet::GetItemString(const std::wstring& name, const std::wstring& default_value)
{
    VARIANT result = GetItemInternal(name);
    if (result.vt == VT_BSTR && result.bstrVal) {
        std::wstring str(result.bstrVal);
        SysFreeString(result.bstrVal);
        return str;
    }
    VariantClear(&result);
    return default_value;
}

bool HwpParameterSet::GetItemBool(const std::wstring& name, bool default_value)
{
    VARIANT result = GetItemInternal(name);
    if (result.vt == VT_BOOL) {
        return result.boolVal != VARIANT_FALSE;
    }
    VariantClear(&result);
    return default_value;
}

VARIANT HwpParameterSet::GetItem(const std::wstring& name)
{
    return GetItemInternal(name);
}

VARIANT HwpParameterSet::GetItemInternal(const std::wstring& name)
{
    VARIANT result;
    VariantInit(&result);

    if (!m_pSet) return result;

    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(L"Item");
    HRESULT hr = m_pSet->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return result;

    VARIANT nameVar;
    VariantInit(&nameVar);
    nameVar.vt = VT_BSTR;
    nameVar.bstrVal = SysAllocString(name.c_str());

    DISPPARAMS params = { &nameVar, NULL, 1, 0 };
    m_pSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                   &params, &result, NULL, NULL);

    SysFreeString(nameVar.bstrVal);
    return result;
}

//=============================================================================
// 서브셋 관련
//=============================================================================

IDispatch* HwpParameterSet::CreateItemSet(const std::wstring& item_id, const std::wstring& set_id)
{
    if (!m_pSet) return nullptr;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(L"CreateItemSet");
    HRESULT hr = m_pSet->GetIDsOfNames(IID_NULL, &methodName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return nullptr;

    VARIANT args[2];
    VariantInit(&args[0]);
    VariantInit(&args[1]);

    args[0].vt = VT_BSTR;
    args[0].bstrVal = SysAllocString(set_id.c_str());
    args[1].vt = VT_BSTR;
    args[1].bstrVal = SysAllocString(item_id.c_str());

    DISPPARAMS params = { args, NULL, 2, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                        &params, &result, NULL, NULL);

    SysFreeString(args[0].bstrVal);
    SysFreeString(args[1].bstrVal);

    if (SUCCEEDED(hr) && result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    return nullptr;
}

IDispatch* HwpParameterSet::GetItemSet(const std::wstring& item_id)
{
    VARIANT result = GetItemInternal(item_id);
    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    VariantClear(&result);
    return nullptr;
}

IDispatch* HwpParameterSet::GetItemByIndex(int index)
{
    if (!m_pSet) return nullptr;

    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(L"Item");
    HRESULT hr = m_pSet->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return nullptr;

    VARIANT indexVar;
    VariantInit(&indexVar);
    indexVar.vt = VT_I4;
    indexVar.lVal = index;

    DISPPARAMS params = { &indexVar, NULL, 1, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                        &params, &result, NULL, NULL);

    if (SUCCEEDED(hr) && result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    VariantClear(&result);
    return nullptr;
}

//=============================================================================
// 유틸리티
//=============================================================================

std::wstring HwpParameterSet::GetSetID() const
{
    if (!m_pSet) return L"";

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"SetID");
    HRESULT hr = m_pSet->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return L"";

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                        &params, &result, NULL, NULL);

    if (SUCCEEDED(hr) && result.vt == VT_BSTR && result.bstrVal) {
        std::wstring id(result.bstrVal);
        SysFreeString(result.bstrVal);
        return id;
    }
    return L"";
}

int HwpParameterSet::GetCount()
{
    if (!m_pSet) return 0;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"Count");
    HRESULT hr = m_pSet->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return 0;

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = m_pSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                        &params, &result, NULL, NULL);

    if (SUCCEEDED(hr) && result.vt == VT_I4) {
        return result.lVal;
    }
    return 0;
}

void HwpParameterSet::Clear()
{
    if (!m_pSet) return;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"Clear");
    HRESULT hr = m_pSet->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return;

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    m_pSet->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                   &params, NULL, NULL, NULL);
}

//=============================================================================
// HwpParameterSetHelper 구현
//=============================================================================

std::unique_ptr<HwpParameterSet> HwpParameterSetHelper::CreateFindReplace(
    HwpWrapper* hwp,
    const std::wstring& find_text,
    const std::wstring& replace_text,
    bool match_case,
    bool whole_word,
    bool regex)
{
    if (!hwp) return nullptr;

    IDispatch* pSet = hwp->CreateSet(L"FindReplace");
    if (!pSet) return nullptr;

    auto set = std::make_unique<HwpParameterSet>(pSet);

    set->SetItem(L"FindString", find_text);
    if (!replace_text.empty()) {
        set->SetItem(L"ReplaceString", replace_text);
    }
    set->SetItem(L"MatchCase", match_case);
    set->SetItem(L"WholeWord", whole_word);
    set->SetItem(L"UseRegexp", regex);

    pSet->Release();
    return set;
}

std::unique_ptr<HwpParameterSet> HwpParameterSetHelper::CreateTable(
    HwpWrapper* hwp,
    int rows,
    int cols,
    HwpUnit width,
    HwpUnit row_height)
{
    if (!hwp) return nullptr;

    IDispatch* pSet = hwp->CreateSet(L"Table");
    if (!pSet) return nullptr;

    auto set = std::make_unique<HwpParameterSet>(pSet);

    set->SetItem(L"Rows", rows);
    set->SetItem(L"Cols", cols);
    if (width > 0) {
        set->SetItem(L"Width", width);
    }
    if (row_height > 0) {
        set->SetItem(L"RowHeight", row_height);
    }

    pSet->Release();
    return set;
}

std::unique_ptr<HwpParameterSet> HwpParameterSetHelper::CreateCharShape(
    HwpWrapper* hwp,
    const std::wstring& face_name,
    int height,
    COLORREF color)
{
    if (!hwp) return nullptr;

    IDispatch* pSet = hwp->CreateSet(L"CharShape");
    if (!pSet) return nullptr;

    auto set = std::make_unique<HwpParameterSet>(pSet);

    if (!face_name.empty()) {
        set->SetItem(L"FaceNameHangul", face_name);
        set->SetItem(L"FaceNameLatin", face_name);
    }
    if (height > 0) {
        set->SetItem(L"Height", height);
    }
    if (color != 0) {
        set->SetItem(L"TextColor", static_cast<int>(color));
    }

    pSet->Release();
    return set;
}

std::unique_ptr<HwpParameterSet> HwpParameterSetHelper::CreateParaShape(
    HwpWrapper* hwp,
    HAlign align,
    int line_spacing)
{
    if (!hwp) return nullptr;

    IDispatch* pSet = hwp->CreateSet(L"ParaShape");
    if (!pSet) return nullptr;

    auto set = std::make_unique<HwpParameterSet>(pSet);

    set->SetItem(L"Align", static_cast<int>(align));
    set->SetItem(L"LineSpacing", line_spacing);

    pSet->Release();
    return set;
}

std::unique_ptr<HwpParameterSet> HwpParameterSetHelper::CreateInsertPicture(
    HwpWrapper* hwp,
    const std::wstring& path,
    bool embedded,
    int size_option)
{
    if (!hwp) return nullptr;

    IDispatch* pSet = hwp->CreateSet(L"InsertPicture");
    if (!pSet) return nullptr;

    auto set = std::make_unique<HwpParameterSet>(pSet);

    set->SetItem(L"Path", path);
    set->SetItem(L"Embedded", embedded);
    set->SetItem(L"SizeOption", size_option);

    pSet->Release();
    return set;
}

std::unique_ptr<HwpParameterSet> HwpParameterSetHelper::CreateCellShape(
    HwpWrapper* hwp,
    COLORREF bg_color,
    VAlign valign)
{
    if (!hwp) return nullptr;

    IDispatch* pSet = hwp->CreateSet(L"CellShape");
    if (!pSet) return nullptr;

    auto set = std::make_unique<HwpParameterSet>(pSet);

    set->SetItem(L"BackColor", static_cast<int>(bg_color));
    set->SetItem(L"VAlign", static_cast<int>(valign));

    pSet->Release();
    return set;
}

std::unique_ptr<HwpParameterSet> HwpParameterSetHelper::CreateBorderLine(
    HwpWrapper* hwp,
    LineStyle style,
    HwpUnit width,
    COLORREF color)
{
    if (!hwp) return nullptr;

    IDispatch* pSet = hwp->CreateSet(L"BorderLine");
    if (!pSet) return nullptr;

    auto set = std::make_unique<HwpParameterSet>(pSet);

    set->SetItem(L"Style", static_cast<int>(style));
    set->SetItem(L"Width", width);
    set->SetItem(L"Color", static_cast<int>(color));

    pSet->Release();
    return set;
}

} // namespace cpyhwpx
