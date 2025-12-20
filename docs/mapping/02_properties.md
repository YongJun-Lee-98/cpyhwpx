# Hwp 클래스 속성 매핑

## 요약
- 총 속성: 35개
- 포팅 완료: 0개 (0%)
- 우선순위: High

---

## Hwp 클래스 속성 (core.py:1016~1610)

### 읽기 전용 속성 (Read-Only)

| # | Python 속성 | 타입 | 역할 | C++ Getter | 상태 |
|---|------------|------|------|-----------|------|
| 1 | `Application` | HwpApplication | 저수준 API 접근 | `GetApplication()` | TODO |
| 2 | `CLSID` | str | 클래스 ID | `GetCLSID()` | TODO |
| 3 | `coclass_clsid` | str | coclass CLSID | `GetCoclassCLSID()` | TODO |
| 4 | `CurFieldState` | int | 현재 필드 상태 | `GetCurFieldState()` | TODO |
| 5 | `CurMetatagState` | int | 현재 메타태그 상태 (한글2024+) | `GetCurMetatagState()` | TODO |
| 6 | `CurSelectedCtrl` | Ctrl | 현재 선택된 컨트롤 | `GetCurSelectedCtrl()` | TODO |
| 7 | `EngineProperties` | Any | 엔진 속성 | `GetEngineProperties()` | TODO |
| 8 | `HAction` | HAction | 액션 실행 인터페이스 | `GetHAction()` | TODO |
| 9 | `HeadCtrl` | Ctrl | 문서 첫 번째 컨트롤 | `GetHeadCtrl()` | TODO |
| 10 | `HParameterSet` | HParameterSet | 파라미터셋 접근 | `GetHParameterSet()` | TODO |
| 11 | `IsEmpty` | bool | 빈 문서 여부 | `IsEmpty()` | TODO |
| 12 | `IsModified` | bool | 수정 여부 | `IsModified()` | TODO |
| 13 | `IsPrivateInfoProtected` | bool | 개인정보 보호 여부 | `IsPrivateInfoProtected()` | TODO |
| 14 | `IsTrackChange` | bool | 변경 추적 여부 | `IsTrackChange()` | TODO |
| 15 | `IsTrackChangePassword` | bool | 변경 추적 암호 여부 | `IsTrackChangePassword()` | TODO |
| 16 | `LastCtrl` | Ctrl | 문서 마지막 컨트롤 | `GetLastCtrl()` | TODO |
| 17 | `PageCount` | int | 총 페이지 수 | `GetPageCount()` | TODO |
| 18 | `ParentCtrl` | Ctrl | 상위 컨트롤 | `GetParentCtrl()` | TODO |
| 19 | `Path` | str | 문서 경로 | `GetPath()` | TODO |
| 20 | `SelectionMode` | int | 선택 모드 | `GetSelectionMode()` | TODO |
| 21 | `Title` | str | 창 제목 | `GetTitle()` | TODO |
| 22 | `Version` | List[int] | 한/글 버전 | `GetVersion()` | TODO |
| 23 | `XHwpDocuments` | XHwpDocuments | 문서 컬렉션 | `GetXHwpDocuments()` | TODO |
| 24 | `XHwpMessageBox` | XHwpMessageBox | 메시지박스 객체 | `GetXHwpMessageBox()` | TODO |
| 25 | `XHwpODBC` | XHwpODBC | ODBC 객체 | `GetXHwpODBC()` | TODO |
| 26 | `XHwpWindows` | XHwpWindows | 창 관리 객체 | `GetXHwpWindows()` | TODO |
| 27 | `ctrl_list` | List[Ctrl] | 모든 컨트롤 리스트 | `GetCtrlList()` | TODO |
| 28 | `current_page` | int | 현재 페이지 번호 | `GetCurrentPage()` | TODO |
| 29 | `current_printpage` | int | 현재 인쇄 페이지 번호 | `GetCurrentPrintPage()` | TODO |
| 30 | `current_font` | str | 현재 폰트 | `GetCurrentFont()` | TODO |

### 읽기/쓰기 속성 (Read/Write)

| # | Python 속성 | 타입 | 역할 | C++ Getter/Setter | 상태 |
|---|------------|------|------|-------------------|------|
| 1 | `CellShape` | Any | 셀 모양 파라미터셋 | `GetCellShape()` / `SetCellShape()` | TODO |
| 2 | `CharShape` | Any | 글자 모양 파라미터셋 | `GetCharShape()` / `SetCharShape()` | TODO |
| 3 | `EditMode` | int | 편집 모드 (0: 읽기전용, 1: 편집) | `GetEditMode()` / `SetEditMode()` | TODO |
| 4 | `ParaShape` | Any | 문단 모양 파라미터셋 | `GetParaShape()` / `SetParaShape()` | TODO |
| 5 | `ViewProperties` | Any | 보기 속성 | `GetViewProperties()` / `SetViewProperties()` | TODO |

---

## 상세 설명

### Application (core.py:1016~1032)
저수준 HwpApplication 객체에 직접 접근

```python
@property
def Application(self) -> "Hwp.Application":
    return self.hwp.Application
```

**C++ 구현:**
```cpp
IDispatch* Hwp::GetApplication() {
    IDispatch* pApp = nullptr;
    VARIANT result;
    VariantInit(&result);
    if (InvokeMethod(m_pHwp, L"Application", &result, DISPATCH_PROPERTYGET)) {
        pApp = result.pdispVal;
    }
    return pApp;
}
```

### CellShape (core.py:1034~1058)
표 셀 모양 파라미터셋

```python
@property
def CellShape(self) -> Any:
    return self.hwp.CellShape

@CellShape.setter
def CellShape(self, prop: Any) -> None:
    self.hwp.CellShape = prop
```

**사용 예:**
```python
# 현재 셀 높이 조회
height = hwp.CellShape.Item("Height")
print(hwp.HwpUnitToMili(height))  # mm 단위로 출력
```

### CharShape (core.py:1060~1123)
글자 모양 파라미터셋

```python
@property
def CharShape(self) -> "Hwp.CharShape":
    return self.hwp.CharShape

@CharShape.setter
def CharShape(self, prop: Any) -> None:
    self.hwp.CharShape = prop
```

**사용 예:**
```python
# 글자 크기 조회
height = hwp.CharShape.Item("Height")
print(hwp.HwpUnitToPoint(height))  # 포인트 단위로 출력

# 글자 모양 변경
prop = hwp.CharShape
prop.SetItem("Height", hwp.PointToHwpUnit(20))
hwp.CharShape = prop
```

### CurFieldState (core.py:1151~1174)
현재 캐럿이 들어있는 영역 상태

**반환값:**
- 0: 본문, 캡션, 주석 (필드 외부)
- 1: 셀 안
- 4: 글상자 안
- 17: 셀필드 안
- 18: 누름틀 안

```python
@property
def CurFieldState(self) -> int:
    return self.hwp.CurFieldState
```

### CurMetatagState (core.py:1176~1196)
현재 캐럿이 들어있는 메타태그 상태 (한글2024+)

**반환값:**
- 1: 셀 메타태그 영역
- 4: 메타태그가 부여된 글상자/그리기개체 내부
- 8: 메타태그가 부여된 이미지/글맵시/글상자 선택 상태
- 16: 메타태그가 부여된 표 선택 상태
- 32: 메타태그 영역 외부
- 40: 메타태그 없는 컨트롤 선택 상태 (8+32)
- 64: 본문 메타태그 영역

### HeadCtrl / LastCtrl (core.py:1267~1392)
문서의 첫 번째/마지막 컨트롤

```python
@property
def HeadCtrl(self) -> Ctrl:
    return Ctrl(self.hwp.HeadCtrl)

@property
def LastCtrl(self) -> Ctrl:
    return Ctrl(self.hwp.LastCtrl)
```

**사용 예:**
```python
# 모든 컨트롤 순회
ctrl = hwp.HeadCtrl
while ctrl:
    print(ctrl.UserDesc)
    ctrl = ctrl.Next

# 역순 순회 (삭제 시 사용)
ctrl = hwp.LastCtrl
while ctrl:
    if ctrl.UserDesc == "그림":
        hwp.DeleteCtrl(ctrl)
    ctrl = ctrl.Prev
```

### ctrl_list (core.py:1548~1566)
문서 내 모든 컨트롤 리스트 (secd, cold 제외)

```python
@property
def ctrl_list(self) -> list:
    c_list = []
    ctrl = self.hwp.HeadCtrl.Next.Next  # secd, cold 건너뛰기
    while ctrl:
        c_list.append(ctrl)
        ctrl = ctrl.Next
    return [Ctrl(i) for i in c_list]
```

**C++ 구현:**
```cpp
std::vector<Ctrl*> Hwp::GetCtrlList() {
    std::vector<Ctrl*> result;
    Ctrl* ctrl = GetHeadCtrl();
    if (!ctrl) return result;

    ctrl = ctrl->GetNext();  // secd 건너뛰기
    if (!ctrl) return result;

    ctrl = ctrl->GetNext();  // cold 건너뛰기

    while (ctrl) {
        result.push_back(ctrl);
        ctrl = ctrl->GetNext();
    }
    return result;
}
```

### current_page / current_printpage (core.py:1568~1597)
현재 페이지 번호

```python
@property
def current_page(self) -> int:
    # 페이지 인덱스 (새 쪽번호와 무관)
    return self.hwp.XHwpDocuments.Active_XHwpDocument.XHwpDocumentInfo.CurrentPage + 1

@property
def current_printpage(self) -> int:
    # 인쇄 페이지 번호 (새 쪽번호 반영)
    return self.hwp.XHwpDocuments.Active_XHwpDocument.XHwpDocumentInfo.CurrentPrintPage
```

### ParaShape (core.py:1411~1438)
문단 모양 파라미터셋

```python
@property
def ParaShape(self):
    return self.hwp.ParaShape

@ParaShape.setter
def ParaShape(self, prop):
    self.hwp.ParaShape = prop
```

**사용 예:**
```python
# 줄간격 조회
linespacing = hwp.ParaShape.Item("LineSpacing")
print(f"줄간격: {linespacing}%")

# 줄간격 200%로 변경
hwp.SelectAll()
prop = hwp.ParaShape
prop.SetItem("LineSpacing", 200)
hwp.ParaShape = prop
```

---

## C++ 클래스 구조

```cpp
// Hwp.h - 속성 선언부
class Hwp : public ParamHelpers, public RunMethods {
public:
    // Read-Only Properties
    IDispatch* GetApplication();
    std::wstring GetCLSID();
    int GetCurFieldState();
    int GetCurMetatagState();
    Ctrl* GetCurSelectedCtrl();
    IDispatch* GetHAction();
    Ctrl* GetHeadCtrl();
    IDispatch* GetHParameterSet();
    bool IsEmpty();
    bool IsModified();
    Ctrl* GetLastCtrl();
    int GetPageCount();
    Ctrl* GetParentCtrl();
    std::wstring GetPath();
    int GetSelectionMode();
    std::wstring GetTitle();
    std::vector<int> GetVersion();
    XHwpDocuments* GetXHwpDocuments();
    XHwpWindows* GetXHwpWindows();
    std::vector<Ctrl*> GetCtrlList();
    int GetCurrentPage();
    int GetCurrentPrintPage();
    std::wstring GetCurrentFont();

    // Read/Write Properties
    IDispatch* GetCellShape();
    void SetCellShape(IDispatch* prop);
    IDispatch* GetCharShape();
    void SetCharShape(IDispatch* prop);
    int GetEditMode();
    void SetEditMode(int mode);
    IDispatch* GetParaShape();
    void SetParaShape(IDispatch* prop);
    IDispatch* GetViewProperties();
    void SetViewProperties(IDispatch* prop);

    // ...
};
```

---

## 구현 우선순위

1. **Critical**: HAction, HParameterSet (액션 실행에 필수)
2. **Critical**: HeadCtrl, LastCtrl, ctrl_list (컨트롤 순회)
3. **High**: CharShape, ParaShape, CellShape (서식 조작)
4. **High**: Path, PageCount, IsEmpty, IsModified (문서 정보)
5. **Medium**: CurFieldState, CurSelectedCtrl (위치/선택 정보)
6. **Low**: 기타 속성
