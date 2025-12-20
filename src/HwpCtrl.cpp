/**
 * @file HwpCtrl.cpp
 * @brief HWP 컨트롤 래퍼 클래스 구현
 *
 * pyhwpx → cpyhwpx 포팅 프로젝트
 */

#include "HwpCtrl.h"
#include "HwpWrapper.h"

namespace cpyhwpx {

//=============================================================================
// 생성자/소멸자
//=============================================================================

HwpCtrl::HwpCtrl(IDispatch* ctrl, HwpWrapper* hwp)
    : m_pCtrl(ctrl)
    , m_pHwp(hwp)
{
    if (m_pCtrl) {
        m_pCtrl->AddRef();
    }
}

HwpCtrl::~HwpCtrl()
{
    if (m_pCtrl) {
        m_pCtrl->Release();
        m_pCtrl = nullptr;
    }
}

HwpCtrl::HwpCtrl(HwpCtrl&& other) noexcept
    : m_pCtrl(other.m_pCtrl)
    , m_pHwp(other.m_pHwp)
{
    other.m_pCtrl = nullptr;
    other.m_pHwp = nullptr;
}

HwpCtrl& HwpCtrl::operator=(HwpCtrl&& other) noexcept
{
    if (this != &other) {
        if (m_pCtrl) {
            m_pCtrl->Release();
        }
        m_pCtrl = other.m_pCtrl;
        m_pHwp = other.m_pHwp;
        other.m_pCtrl = nullptr;
        other.m_pHwp = nullptr;
    }
    return *this;
}

//=============================================================================
// 컨트롤 정보
//=============================================================================

std::wstring HwpCtrl::GetCtrlID() const
{
    VARIANT result = GetProperty(L"CtrlID");
    if (result.vt == VT_BSTR && result.bstrVal) {
        std::wstring id(result.bstrVal);
        SysFreeString(result.bstrVal);
        return id;
    }
    return L"";
}

CtrlType HwpCtrl::GetCtrlType() const
{
    return CtrlIDToType(GetCtrlID());
}

int HwpCtrl::GetCtrlCh() const
{
    VARIANT result = GetProperty(L"CtrlCh");
    if (result.vt == VT_I4) {
        return result.lVal;
    }
    return 0;
}

CtrlType HwpCtrl::CtrlIDToType(const std::wstring& id)
{
    if (id == L"tbl") return CtrlType::Table;
    if (id == L"eqed") return CtrlType::Equation;
    if (id == L"pic") return CtrlType::Picture;
    if (id == L"ole") return CtrlType::OLE;
    if (id == L"lin") return CtrlType::Line;
    if (id == L"rec") return CtrlType::Rectangle;
    if (id == L"ell") return CtrlType::Ellipse;
    if (id == L"arc") return CtrlType::Arc;
    if (id == L"pol") return CtrlType::Polygon;
    if (id == L"cur") return CtrlType::Curve;
    if (id == L"$txt") return CtrlType::TextBox;
    if (id == L"$con") return CtrlType::Container;
    if (id == L"gso") return CtrlType::GenShapeObj;
    if (id == L"head") return CtrlType::Header;
    if (id == L"foot") return CtrlType::Footer;
    if (id == L"fn") return CtrlType::Footnote;
    if (id == L"en") return CtrlType::Endnote;
    if (id == L"atno") return CtrlType::AutoNum;
    if (id == L"pgno") return CtrlType::PageNum;
    if (id == L"%fld") return CtrlType::Field;
    if (id == L"bokm") return CtrlType::Bookmark;
    if (id == L"btn") return CtrlType::Button;
    if (id == L"rdo") return CtrlType::RadioButton;
    if (id == L"chk") return CtrlType::CheckButton;
    if (id == L"cbo") return CtrlType::ComboBox;
    if (id == L"edt") return CtrlType::Edit;
    if (id == L"lst") return CtrlType::ListBox;
    if (id == L"scr") return CtrlType::ScrollBar;
    if (id == L"vid") return CtrlType::Video;
    if (id == L"cold") return CtrlType::ColumnDef;
    return CtrlType::Unknown;
}

//=============================================================================
// 탐색
//=============================================================================

std::unique_ptr<HwpCtrl> HwpCtrl::Next() const
{
    VARIANT result = GetProperty(L"Next");
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        return std::make_unique<HwpCtrl>(result.pdispVal, m_pHwp);
    }
    return nullptr;
}

std::unique_ptr<HwpCtrl> HwpCtrl::Prev() const
{
    VARIANT result = GetProperty(L"Prev");
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        return std::make_unique<HwpCtrl>(result.pdispVal, m_pHwp);
    }
    return nullptr;
}

std::unique_ptr<HwpCtrl> HwpCtrl::Parent() const
{
    VARIANT result = GetProperty(L"Parent");
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        return std::make_unique<HwpCtrl>(result.pdispVal, m_pHwp);
    }
    return nullptr;
}

std::unique_ptr<HwpCtrl> HwpCtrl::FirstChild() const
{
    VARIANT result = GetProperty(L"FirstChild");
    if (result.vt == VT_DISPATCH && result.pdispVal) {
        return std::make_unique<HwpCtrl>(result.pdispVal, m_pHwp);
    }
    return nullptr;
}

//=============================================================================
// 속성 접근
//=============================================================================

IDispatch* HwpCtrl::GetProperties() const
{
    VARIANT result = GetProperty(L"Properties");
    if (result.vt == VT_DISPATCH) {
        return result.pdispVal;
    }
    return nullptr;
}

bool HwpCtrl::SetProperties(IDispatch* pset)
{
    if (!pset) return false;

    VARIANT val;
    VariantInit(&val);
    val.vt = VT_DISPATCH;
    val.pdispVal = pset;

    return SetProperty(L"Properties", val);
}

std::wstring HwpCtrl::GetUserData() const
{
    VARIANT result = GetProperty(L"UserData");
    if (result.vt == VT_BSTR && result.bstrVal) {
        std::wstring data(result.bstrVal);
        SysFreeString(result.bstrVal);
        return data;
    }
    return L"";
}

bool HwpCtrl::SetUserData(const std::wstring& data)
{
    VARIANT val;
    VariantInit(&val);
    val.vt = VT_BSTR;
    val.bstrVal = SysAllocString(data.c_str());

    bool success = SetProperty(L"UserData", val);
    SysFreeString(val.bstrVal);
    return success;
}

int HwpCtrl::GetInstID() const
{
    VARIANT result = GetProperty(L"InstID");
    if (result.vt == VT_I4) {
        return result.lVal;
    }
    return 0;
}

//=============================================================================
// 표 전용 속성
//=============================================================================

bool HwpCtrl::IsTable() const
{
    return GetCtrlID() == L"tbl";
}

int HwpCtrl::GetRowCount() const
{
    if (!IsTable()) return 0;

    IDispatch* props = GetProperties();
    if (!props) return 0;

    // Properties에서 RowCount 가져오기
    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"RowCount");
    HRESULT hr = props->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        props->Release();
        return 0;
    }

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = props->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                       &params, &result, NULL, NULL);
    props->Release();

    if (SUCCEEDED(hr) && result.vt == VT_I4) {
        return result.lVal;
    }
    return 0;
}

int HwpCtrl::GetColCount() const
{
    if (!IsTable()) return 0;

    IDispatch* props = GetProperties();
    if (!props) return 0;

    DISPID dispid;
    OLECHAR* name = const_cast<OLECHAR*>(L"ColCount");
    HRESULT hr = props->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) {
        props->Release();
        return 0;
    }

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    VARIANT result;
    VariantInit(&result);

    hr = props->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                       &params, &result, NULL, NULL);
    props->Release();

    if (SUCCEEDED(hr) && result.vt == VT_I4) {
        return result.lVal;
    }
    return 0;
}

std::unique_ptr<HwpCtrl> HwpCtrl::GetCell(int row, int col) const
{
    if (!IsTable()) return nullptr;

    // GetCellAddr 메서드 호출
    VARIANT args[2];
    VariantInit(&args[0]);
    VariantInit(&args[1]);

    args[0].vt = VT_I4;
    args[0].lVal = col;
    args[1].vt = VT_I4;
    args[1].lVal = row;

    VARIANT result = InvokeMethod(L"GetCellAddr", { args[1], args[0] });
    // TODO: 셀 컨트롤 반환 구현

    return nullptr;
}

bool HwpCtrl::SelectTable()
{
    if (!IsTable()) return false;
    // 표 선택 로직
    return MoveTo() && Select();
}

bool HwpCtrl::GoToCell(const std::wstring& addr)
{
    if (!IsTable()) return false;
    // TODO: 셀 주소 파싱 및 이동 구현
    return false;
}

//=============================================================================
// 그림 전용 속성
//=============================================================================

bool HwpCtrl::IsPicture() const
{
    std::wstring id = GetCtrlID();
    return (id == L"pic" || id == L"gso");
}

std::wstring HwpCtrl::GetPicturePath() const
{
    if (!IsPicture()) return L"";

    // BinData에서 경로 가져오기
    IDispatch* props = GetProperties();
    if (!props) return L"";

    // TODO: 그림 경로 추출 구현
    props->Release();
    return L"";
}

std::tuple<HwpUnit, HwpUnit> HwpCtrl::GetPictureSize() const
{
    if (!IsPicture()) return std::make_tuple(0, 0);
    return GetSize();
}

//=============================================================================
// 위치/크기
//=============================================================================

std::tuple<HwpUnit, HwpUnit> HwpCtrl::GetPosition() const
{
    IDispatch* props = GetProperties();
    if (!props) return std::make_tuple(0, 0);

    // X 좌표
    DISPID dispidX;
    OLECHAR* nameX = const_cast<OLECHAR*>(L"XPos");
    HRESULT hr = props->GetIDsOfNames(IID_NULL, &nameX, 1, LOCALE_USER_DEFAULT, &dispidX);

    HwpUnit x = 0, y = 0;

    if (SUCCEEDED(hr)) {
        DISPPARAMS params = { NULL, NULL, 0, 0 };
        VARIANT result;
        VariantInit(&result);

        hr = props->Invoke(dispidX, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                           &params, &result, NULL, NULL);
        if (SUCCEEDED(hr) && result.vt == VT_I4) {
            x = result.lVal;
        }
    }

    // Y 좌표
    DISPID dispidY;
    OLECHAR* nameY = const_cast<OLECHAR*>(L"YPos");
    hr = props->GetIDsOfNames(IID_NULL, &nameY, 1, LOCALE_USER_DEFAULT, &dispidY);

    if (SUCCEEDED(hr)) {
        DISPPARAMS params = { NULL, NULL, 0, 0 };
        VARIANT result;
        VariantInit(&result);

        hr = props->Invoke(dispidY, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                           &params, &result, NULL, NULL);
        if (SUCCEEDED(hr) && result.vt == VT_I4) {
            y = result.lVal;
        }
    }

    props->Release();
    return std::make_tuple(x, y);
}

std::tuple<HwpUnit, HwpUnit> HwpCtrl::GetSize() const
{
    IDispatch* props = GetProperties();
    if (!props) return std::make_tuple(0, 0);

    HwpUnit width = 0, height = 0;

    // Width
    DISPID dispidW;
    OLECHAR* nameW = const_cast<OLECHAR*>(L"Width");
    HRESULT hr = props->GetIDsOfNames(IID_NULL, &nameW, 1, LOCALE_USER_DEFAULT, &dispidW);

    if (SUCCEEDED(hr)) {
        DISPPARAMS params = { NULL, NULL, 0, 0 };
        VARIANT result;
        VariantInit(&result);

        hr = props->Invoke(dispidW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                           &params, &result, NULL, NULL);
        if (SUCCEEDED(hr) && result.vt == VT_I4) {
            width = result.lVal;
        }
    }

    // Height
    DISPID dispidH;
    OLECHAR* nameH = const_cast<OLECHAR*>(L"Height");
    hr = props->GetIDsOfNames(IID_NULL, &nameH, 1, LOCALE_USER_DEFAULT, &dispidH);

    if (SUCCEEDED(hr)) {
        DISPPARAMS params = { NULL, NULL, 0, 0 };
        VARIANT result;
        VariantInit(&result);

        hr = props->Invoke(dispidH, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                           &params, &result, NULL, NULL);
        if (SUCCEEDED(hr) && result.vt == VT_I4) {
            height = result.lVal;
        }
    }

    props->Release();
    return std::make_tuple(width, height);
}

bool HwpCtrl::SetSize(HwpUnit width, HwpUnit height)
{
    IDispatch* props = GetProperties();
    if (!props) return false;

    bool success = true;

    // Width 설정
    DISPID dispidW;
    OLECHAR* nameW = const_cast<OLECHAR*>(L"Width");
    HRESULT hr = props->GetIDsOfNames(IID_NULL, &nameW, 1, LOCALE_USER_DEFAULT, &dispidW);

    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = width;

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS params = { &val, &putid, 1, 1 };

        hr = props->Invoke(dispidW, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                           &params, NULL, NULL, NULL);
        if (FAILED(hr)) success = false;
    }

    // Height 설정
    DISPID dispidH;
    OLECHAR* nameH = const_cast<OLECHAR*>(L"Height");
    hr = props->GetIDsOfNames(IID_NULL, &nameH, 1, LOCALE_USER_DEFAULT, &dispidH);

    if (SUCCEEDED(hr)) {
        VARIANT val;
        VariantInit(&val);
        val.vt = VT_I4;
        val.lVal = height;

        DISPID putid = DISPID_PROPERTYPUT;
        DISPPARAMS params = { &val, &putid, 1, 1 };

        hr = props->Invoke(dispidH, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                           &params, NULL, NULL, NULL);
        if (FAILED(hr)) success = false;
    }

    props->Release();
    return success;
}

//=============================================================================
// 컨트롤 조작
//=============================================================================

bool HwpCtrl::MoveTo()
{
    if (!m_pHwp || !m_pCtrl) return false;

    // SetPosBySet 사용하여 컨트롤 위치로 이동
    // TODO: 구현
    return false;
}

bool HwpCtrl::Select()
{
    if (!m_pHwp || !m_pCtrl) return false;

    // 컨트롤 선택
    // TODO: 구현
    return false;
}

bool HwpCtrl::Delete()
{
    if (!m_pHwp || !m_pCtrl) return false;

    // 컨트롤 삭제
    // TODO: 구현
    return false;
}

bool HwpCtrl::Copy()
{
    if (!m_pHwp || !m_pCtrl) return false;

    // 컨트롤 복사
    // TODO: 구현
    return false;
}

//=============================================================================
// COM 헬퍼 메서드
//=============================================================================

VARIANT HwpCtrl::GetProperty(const std::wstring& name) const
{
    VARIANT result;
    VariantInit(&result);

    if (!m_pCtrl) return result;

    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = m_pCtrl->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return result;

    DISPPARAMS params = { NULL, NULL, 0, 0 };
    m_pCtrl->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
                    &params, &result, NULL, NULL);

    return result;
}

bool HwpCtrl::SetProperty(const std::wstring& name, const VARIANT& value)
{
    if (!m_pCtrl) return false;

    DISPID dispid;
    OLECHAR* propName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = m_pCtrl->GetIDsOfNames(IID_NULL, &propName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return false;

    DISPID putid = DISPID_PROPERTYPUT;
    DISPPARAMS params = { const_cast<VARIANT*>(&value), &putid, 1, 1 };

    hr = m_pCtrl->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
                         &params, NULL, NULL, NULL);

    return SUCCEEDED(hr);
}

VARIANT HwpCtrl::InvokeMethod(const std::wstring& name,
                               const std::vector<VARIANT>& args) const
{
    VARIANT result;
    VariantInit(&result);

    if (!m_pCtrl) return result;

    DISPID dispid;
    OLECHAR* methodName = const_cast<OLECHAR*>(name.c_str());
    HRESULT hr = m_pCtrl->GetIDsOfNames(IID_NULL, &methodName, 1, LOCALE_USER_DEFAULT, &dispid);
    if (FAILED(hr)) return result;

    std::vector<VARIANT> reversedArgs(args.rbegin(), args.rend());

    DISPPARAMS params = {
        reversedArgs.empty() ? NULL : reversedArgs.data(),
        NULL,
        static_cast<UINT>(reversedArgs.size()),
        0
    };

    m_pCtrl->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                    &params, &result, NULL, NULL);

    return result;
}

} // namespace cpyhwpx
