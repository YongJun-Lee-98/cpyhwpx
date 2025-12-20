# 개체(Shape/Object) 매핑

## 요약
- 총 함수: 약 60개 (core.py ~15개 + Run 액션 ~45개)
- 포팅 완료: 0개 (0%)
- 우선순위: Medium (Phase 7)

---

## 개체 삽입/삭제 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 1 | `insert_ctrl()` | (ctrl_id, initparam) | Ctrl | 컨트롤 삽입 | `InsertCtrl()` | TODO |
| 2 | `InsertCtrl()` | (ctrl_id, initparam) | Ctrl | 저수준 컨트롤 삽입 | `InsertCtrl()` | TODO |
| 3 | `insert_picture()` | (path, treat_as_char, ...) | Ctrl | 그림 삽입 | `InsertPicture()` | TODO |
| 4 | `InsertPicture()` | (path, treat_as_char, ...) | Ctrl | 저수준 그림 삽입 | `InsertPicture()` | TODO |
| 5 | `insert_random_picture()` | (x, y) | Ctrl | 랜덤 그림 삽입 | `InsertRandomPicture()` | TODO |
| 6 | `delete_ctrl()` | (ctrl) | bool | 컨트롤 삭제 | `DeleteCtrl()` | TODO |
| 7 | `DeleteCtrl()` | (ctrl) | bool | 저수준 컨트롤 삭제 | `DeleteCtrl()` | TODO |

## 개체 선택 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 8 | `SelectCtrl()` | (ctrllist, option) | bool | 컨트롤 선택 (한글2024+) | `SelectCtrl()` | TODO |
| 9 | `select_ctrl()` | (ctrl, anchor_type, option) | bool | 컨트롤 선택 | `SelectCtrl()` | TODO |

## 이미지/페이지 생성 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 10 | `create_page_image()` | (path, pgno, resolution, ...) | bool | 페이지를 이미지로 저장 | `CreatePageImage()` | TODO |

## 수식 관련 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 11 | `EquationCreate()` | (thread) | bool | 수식 편집기 실행 | `EquationCreate()` | TODO |
| 12 | `EquationClose()` | (save, delay) | bool | 수식 편집기 닫기 | `EquationClose()` | TODO |
| 13 | `EquationModify()` | (thread) | bool | 수식 편집 모드 | `EquationModify()` | TODO |

---

## 개체 관련 Run 액션 (HAction.Run)

### 개체 정렬

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 1 | `ShapeObjAlignLeft()` | 왼쪽 정렬 | `ShapeObjAlignLeft()` | TODO |
| 2 | `ShapeObjAlignCenter()` | 가운데 정렬 | `ShapeObjAlignCenter()` | TODO |
| 3 | `ShapeObjAlignRight()` | 오른쪽 정렬 | `ShapeObjAlignRight()` | TODO |
| 4 | `ShapeObjAlignTop()` | 위쪽 정렬 | `ShapeObjAlignTop()` | TODO |
| 5 | `ShapeObjAlignMiddle()` | 중간 정렬 | `ShapeObjAlignMiddle()` | TODO |
| 6 | `ShapeObjAlignBottom()` | 아래쪽 정렬 | `ShapeObjAlignBottom()` | TODO |
| 7 | `ShapeObjAlignWidth()` | 너비 맞춤 | `ShapeObjAlignWidth()` | TODO |
| 8 | `ShapeObjAlignHeight()` | 높이 맞춤 | `ShapeObjAlignHeight()` | TODO |
| 9 | `ShapeObjAlignSize()` | 크기 맞춤 | `ShapeObjAlignSize()` | TODO |
| 10 | `ShapeObjAlignHorzSpacing()` | 가로 간격 균등 | `ShapeObjAlignHorzSpacing()` | TODO |
| 11 | `ShapeObjAlignVertSpacing()` | 세로 간격 균등 | `ShapeObjAlignVertSpacing()` | TODO |

### 개체 순서

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 12 | `ShapeObjBringToFront()` | 맨 앞으로 | `ShapeObjBringToFront()` | TODO |
| 13 | `ShapeObjBringForward()` | 앞으로 | `ShapeObjBringForward()` | TODO |
| 14 | `ShapeObjBringInFrontOfText()` | 글 앞으로 | `ShapeObjBringInFrontOfText()` | TODO |
| 15 | `ShapeObjCtrlSendBehindText()` | 글 뒤로 | `ShapeObjCtrlSendBehindText()` | TODO |

### 개체 그룹

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 16 | `ShapeObjGroup()` | 그룹 | `ShapeObjGroup()` | TODO |
| 17 | `ShapeObjUngroup()` | 그룹 해제 | `ShapeObjUngroup()` | TODO |

### 개체 변환

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 18 | `ShapeObjHorzFlip()` | 좌우 대칭 | `ShapeObjHorzFlip()` | TODO |
| 19 | `ShapeObjHorzFlipOrgState()` | 좌우 대칭 원상복구 | `ShapeObjHorzFlipOrgState()` | TODO |
| 20 | `ShapeObjNorm()` | 원래 상태 | `ShapeObjNorm()` | TODO |

### 개체 이동

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 21 | `ShapeObjMoveUp()` | 위로 이동 | `ShapeObjMoveUp()` | TODO |
| 22 | `ShapeObjMoveDown()` | 아래로 이동 | `ShapeObjMoveDown()` | TODO |
| 23 | `ShapeObjMoveLeft()` | 왼쪽으로 이동 | `ShapeObjMoveLeft()` | TODO |
| 24 | `ShapeObjMoveRight()` | 오른쪽으로 이동 | `ShapeObjMoveRight()` | TODO |

### 개체 크기 조절

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 25 | `ShapeObjResizeUp()` | 위쪽으로 크기 조절 | `ShapeObjResizeUp()` | TODO |
| 26 | `ShapeObjResizeDown()` | 아래쪽으로 크기 조절 | `ShapeObjResizeDown()` | TODO |
| 27 | `ShapeObjResizeLeft()` | 왼쪽으로 크기 조절 | `ShapeObjResizeLeft()` | TODO |
| 28 | `ShapeObjResizeRight()` | 오른쪽으로 크기 조절 | `ShapeObjResizeRight()` | TODO |

### 개체 속성

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 29 | `ShapeObjFillProperty()` | 채우기 속성 | `ShapeObjFillProperty()` | TODO |
| 30 | `ShapeObjLineProperty()` | 선 속성 | `ShapeObjLineProperty()` | TODO |
| 31 | `ShapeObjLineStyleOther()` | 선 스타일 | `ShapeObjLineStyleOther()` | TODO |
| 32 | `ShapeObjLineWidthOther()` | 선 두께 | `ShapeObjLineWidthOther()` | TODO |
| 33 | `ShapeObjLock()` | 개체 잠금 | `ShapeObjLock()` | TODO |

### 캡션/글상자

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 34 | `ShapeObjAttachCaption()` | 캡션 붙이기 | `ShapeObjAttachCaption()` | TODO |
| 35 | `ShapeObjDetachCaption()` | 캡션 분리 | `ShapeObjDetachCaption()` | TODO |
| 36 | `ShapeObjInsertCaptionNum()` | 캡션 번호 삽입 | `ShapeObjInsertCaptionNum()` | TODO |
| 37 | `ShapeObjAttachTextBox()` | 글상자 붙이기 | `ShapeObjAttachTextBox()` | TODO |
| 38 | `ShapeObjDetachTextBox()` | 글상자 분리 | `ShapeObjDetachTextBox()` | TODO |

### 개체 선택/탐색

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 39 | `ShapeObjNextObject()` | 다음 개체 | `ShapeObjNextObject()` | TODO |
| 40 | `ShapeObjPrevObject()` | 이전 개체 | `ShapeObjPrevObject()` | TODO |
| 41 | `ShapeObjTableSelCell()` | 표 셀 선택 | `ShapeObjTableSelCell()` | TODO |
| 42 | `ShapeObjTextBoxEdit()` | 글상자 편집 | `ShapeObjTextBoxEdit()` | TODO |

### 기타

| # | Run 액션 | 역할 | C++ 함수명 | 상태 |
|---|---------|------|-----------|------|
| 43 | `ShapeObjGuideLine()` | 안내선 | `ShapeObjGuideLine()` | TODO |

---

## 상세 설명

### insert_ctrl()
현재 캐럿 위치에 컨트롤을 삽입한다.

```python
def insert_ctrl(self, ctrl_id: str, initparam: Any) -> Ctrl:
    """
    Args:
        ctrl_id: 컨트롤 ID (tbl, gso, eqed 등)
        initparam: 초기 속성 (ParameterSet)

    Returns:
        생성된 Ctrl 객체

    Examples:
        >>> # 3행 5열 표 삽입
        >>> tbset = hwp.create_set("TableCreation")
        >>> tbset.SetItem("Rows", 3)
        >>> tbset.SetItem("Cols", 5)
        >>> table = hwp.insert_ctrl("tbl", tbset)
    """
    return self.hwp.InsertCtrl(CtrlID=ctrl_id, initparam=initparam)
```

**컨트롤 ID 목록:**
```
"tbl"  : 표
"gso"  : 그리기 개체 (그림 포함)
"eqed" : 수식
"cold" : 단
"secd" : 구역
"fn"   : 각주
"en"   : 미주
"%clk" : 누름틀
"%bmk" : 블록 책갈피
"%hlk" : 하이퍼링크
"bokm" : 책갈피
"tcmt" : 주석
```

### insert_picture()
그림을 삽입한다.

```python
def insert_picture(self,
    path: str,
    treat_as_char: bool = True,
    embedded: bool = True,
    sizeoption: int = 0,
    reverse: bool = False,
    watermark: bool = False,
    effect: int = 0,
    width: int = 0,
    height: int = 0
) -> Ctrl:
    """
    Args:
        path: 이미지 파일 경로 (URL도 가능)
        treat_as_char: 글자처럼 취급 여부
        embedded: 문서에 포함 여부
        sizeoption:
            - 0 (realSize): 원래 크기
            - 1 (specificSize): 지정 크기 (width, height 필수)
            - 2 (cellSize): 셀 크기에 맞춤 (종횡비 무시)
            - 3 (cellSizeWithSameRatio): 셀 크기에 맞춤 (종횡비 유지)
        reverse: 반전 여부
        watermark: 워터마크 효과
        effect:
            - 0: 그대로
            - 1: 그레이스케일
            - 2: 흑백
        width: 너비 (mm 단위)
        height: 높이 (mm 단위)

    Returns:
        삽입된 그림 Ctrl 객체

    Examples:
        >>> ctrl = hwp.insert_picture("C:/image.jpg")
        >>> pset = ctrl.Properties
        >>> pset.SetItem("TreatAsChar", False)
        >>> pset.SetItem("TextWrap", 2)  # 글 뒤로
        >>> ctrl.Properties = pset
    """
```

**C++ 구현:**
```cpp
Ctrl* Hwp::InsertPicture(const std::wstring& path,
                         bool treatAsChar, bool embedded,
                         int sizeOption, bool reverse,
                         bool watermark, int effect,
                         int width, int height) {
    VARIANT vPath, vEmbedded, vSizeOpt, vReverse, vWatermark, vEffect, vWidth, vHeight;
    // VARIANT 초기화 및 값 설정

    VARIANT result;
    DISPPARAMS params = {...};

    HRESULT hr = m_pHwp->Invoke(DISPID_INSERTPICTURE, IID_NULL,
                                 LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                                 &params, &result, NULL, NULL);

    if (SUCCEEDED(hr)) {
        IDispatch* pCtrl = result.pdispVal;
        // Properties 설정
        IDispatch* pProps = GetProperty(pCtrl, L"Properties");
        SetProperty(pProps, L"TreatAsChar", treatAsChar);
        SetProperty(pCtrl, L"Properties", pProps);
        return new Ctrl(pCtrl);
    }
    return nullptr;
}
```

### delete_ctrl()
컨트롤을 삭제한다.

```python
def delete_ctrl(self, ctrl: Ctrl) -> bool:
    """
    Args:
        ctrl: 삭제할 컨트롤 객체

    Returns:
        성공시 True

    Examples:
        >>> ctrl = hwp.HeadCtrl.Next.Next
        >>> if ctrl.UserDesc == "표":
        ...     hwp.delete_ctrl(ctrl)
    """
    return self.hwp.DeleteCtrl(ctrl=ctrl._com_obj)
```

### SelectCtrl() / select_ctrl()
컨트롤을 선택한다.

```python
# 한글2024 이상
def SelectCtrl(self, ctrllist: Union[str, Ctrl, List[Ctrl]], option: int = 1) -> bool:
    """
    Args:
        ctrllist: 선택할 컨트롤 (인스턴스 ID 또는 Ctrl 객체)
        option:
            - 0: 추가 선택
            - 1: 기존 해제 후 선택 (기본값)

    Examples:
        >>> hwp.SelectCtrl(hwp.LastCtrl.GetCtrlInstID(), 0)
    """

# 모든 버전
def select_ctrl(self, ctrl: Union[str, Ctrl, List[Ctrl]],
                anchor_type: int = 0, option: int = 1) -> bool:
    """
    Args:
        ctrl: 선택할 컨트롤
        anchor_type:
            - 0: 바로 상위 리스트 기준 (기본값)
            - 1: 탑레벨 리스트 기준
            - 2: 루트 리스트 기준
        option: (한글2024+) 추가 선택(0) 또는 교체(1)
    """
```

### create_page_image()
페이지를 이미지 파일로 저장한다.

```python
def create_page_image(self,
    path: str,
    pgno: int = -1,
    resolution: int = 300,
    depth: int = 24,
    format: str = "bmp"
) -> bool:
    """
    Args:
        path: 저장할 파일 경로
        pgno: 페이지 번호
            - -1: 전체 페이지 (기본값)
            - 0: 현재 페이지
            - 1~N: 특정 페이지
        resolution: 해상도 (DPI)
        depth: 색상 심도 (1, 4, 8, 24)
        format: 파일 형식 ("bmp", "gif")

    Returns:
        성공시 True

    Examples:
        >>> hwp.create_page_image("C:/output.bmp")  # 전체 페이지
        >>> hwp.create_page_image("C:/page1.jpg", pgno=1)  # 1페이지만
    """
```

### ShapeObjAttachCaption()
개체에 캡션을 붙인다.

```python
def ShapeObjAttachCaption(self, text: str = "", add_num: bool = True) -> bool:
    """
    Args:
        text: 캡션 텍스트
        add_num: 번호 자동 추가 여부

    Returns:
        성공시 True
    """
    pset = self.hwp.HParameterSet.HCaptionNum
    self.hwp.HAction.GetDefault("ShapeObjAttachCaption", pset.HSet)
    pset.string = text
    pset.add_number = add_num
    return self.hwp.HAction.Execute("ShapeObjAttachCaption", pset.HSet)
```

---

## C++ 클래스 구조

```cpp
// Hwp.h - 개체 관련 메서드
class Hwp : public ParamHelpers, public RunMethods {
public:
    // 개체 삽입/삭제
    Ctrl* InsertCtrl(const std::wstring& ctrlId, IDispatch* initParam);
    Ctrl* InsertPicture(const std::wstring& path,
                        bool treatAsChar = true, bool embedded = true,
                        int sizeOption = 0, bool reverse = false,
                        bool watermark = false, int effect = 0,
                        int width = 0, int height = 0);
    bool DeleteCtrl(Ctrl* ctrl);

    // 개체 선택
    bool SelectCtrl(const std::wstring& ctrlList, int option = 1);
    bool SelectCtrl(Ctrl* ctrl, int anchorType = 0, int option = 1);
    bool SelectCtrl(const std::vector<Ctrl*>& ctrls, int option = 1);

    // 이미지/페이지 생성
    bool CreatePageImage(const std::wstring& path, int pgno = -1,
                         int resolution = 300, int depth = 24,
                         const std::wstring& format = L"bmp");

    // 수식
    bool EquationCreate(bool thread = false);
    bool EquationClose(bool save = false, double delay = 0.1);
    bool EquationModify(bool thread = false);

    // Run 액션 - 정렬
    bool ShapeObjAlignLeft();
    bool ShapeObjAlignCenter();
    bool ShapeObjAlignRight();
    bool ShapeObjAlignTop();
    bool ShapeObjAlignMiddle();
    bool ShapeObjAlignBottom();

    // Run 액션 - 순서
    bool ShapeObjBringToFront();
    bool ShapeObjBringForward();
    bool ShapeObjBringInFrontOfText();
    bool ShapeObjCtrlSendBehindText();

    // Run 액션 - 그룹
    bool ShapeObjGroup();
    bool ShapeObjUngroup();

    // Run 액션 - 변환
    bool ShapeObjHorzFlip();
    bool ShapeObjNorm();

    // Run 액션 - 이동/크기
    bool ShapeObjMoveUp();
    bool ShapeObjMoveDown();
    bool ShapeObjMoveLeft();
    bool ShapeObjMoveRight();
    bool ShapeObjResizeUp();
    bool ShapeObjResizeDown();
    bool ShapeObjResizeLeft();
    bool ShapeObjResizeRight();

    // Run 액션 - 캡션/글상자
    bool ShapeObjAttachCaption(const std::wstring& text = L"",
                                bool addNum = true);
    bool ShapeObjDetachCaption();
    bool ShapeObjAttachTextBox();
    bool ShapeObjDetachTextBox();

    // Run 액션 - 선택/탐색
    bool ShapeObjNextObject();
    bool ShapeObjPrevObject();
    bool ShapeObjTableSelCell();
    bool ShapeObjTextBoxEdit();
};
```

---

## TextWrap 상수 (글자 배치)

```cpp
enum TextWrap {
    wrapSquare = 0,        // 어울림 사각형
    wrapTight = 1,         // 어울림 빈틈없음
    wrapThrough = 2,       // 글 뒤로
    wrapTopAndBottom = 3,  // 위/아래
    wrapInFront = 4,       // 글 앞으로
    wrapBehind = 5,        // 자리 차지
    wrapNone = 6           // 글자처럼 취급
};
```

---

## 구현 우선순위

1. **Critical**: InsertPicture(), InsertCtrl() - 개체 삽입
2. **High**: DeleteCtrl(), select_ctrl() - 개체 관리
3. **High**: ShapeObjTableSelCell(), ShapeObjTextBoxEdit() - 표/글상자 진입
4. **Medium**: CreatePageImage() - 이미지 생성
5. **Medium**: ShapeObjGroup/Ungroup - 그룹화
6. **Low**: 정렬, 순서, 캡션 관련 Run 액션
7. **Low**: 수식 관련 메서드
