# 텍스트 편집 매핑

## 요약
- 총 함수: 약 35개
- 포팅 완료: 0개 (0%)
- 우선순위: High (Phase 3)

---

## 텍스트 삽입/추출 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 1 | `insert_text()` | (text) | bool | 캐럿 위치에 텍스트 삽입 | `InsertText()` | TODO |
| 2 | `get_text()` | () | Tuple[int, str] | 텍스트 추출 (InitScan 필요) | `GetText()` | TODO |
| 3 | `GetText()` | () | Tuple[int, str] | 저수준 텍스트 추출 | `GetText()` | TODO |
| 4 | `get_text_file()` | (format, option) | str | 문서 전체/선택 영역 문자열 추출 | `GetTextFile()` | TODO |
| 5 | `GetTextFile()` | (format, option) | str | 저수준 텍스트 파일 추출 | `GetTextFile()` | TODO |
| 6 | `set_text_file()` | (data, format, option) | int | 문자열을 문서에 삽입 | `SetTextFile()` | TODO |
| 7 | `SetTextFile()` | (data, format, option) | int | 저수준 텍스트 파일 삽입 | `SetTextFile()` | TODO |
| 8 | `get_selected_text()` | (as_, keep_select) | str/list | 선택 영역 텍스트 추출 | `GetSelectedText()` | TODO |

## 캐럿 위치 제어 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 9 | `get_pos()` | () | Tuple[int,int,int] | 캐럿 위치 조회 | `GetPos()` | TODO |
| 10 | `GetPos()` | () | Tuple[int,int,int] | 저수준 위치 조회 | `GetPos()` | TODO |
| 11 | `get_pos_by_set()` | () | HParameterSet | 캐럿 위치를 ParameterSet으로 조회 | `GetPosBySet()` | TODO |
| 12 | `GetPosBySet()` | () | HParameterSet | 저수준 위치 조회 | `GetPosBySet()` | TODO |
| 13 | `set_pos()` | (List, para, pos) | bool | 캐럿 위치 설정 | `SetPos()` | TODO |
| 14 | `SetPos()` | (List, para, pos) | bool | 저수준 위치 설정 | `SetPos()` | TODO |
| 15 | `set_pos_by_set()` | (disp_val) | bool | ParameterSet으로 위치 설정 | `SetPosBySet()` | TODO |
| 16 | `SetPosBySet()` | (disp_val) | bool | 저수준 위치 설정 | `SetPosBySet()` | TODO |
| 17 | `move_pos()` | (move_id, para, pos) | bool | 캐럿 이동 (다양한 옵션) | `MovePos()` | TODO |
| 18 | `MovePos()` | (move_id, para, pos) | bool | 저수준 캐럿 이동 | `MovePos()` | TODO |
| 19 | `move_to_field()` | (field, idx, text, start, select) | bool | 필드로 캐럿 이동 | `MoveToField()` | TODO |
| 20 | `MoveToField()` | (field, idx, text, start, select) | bool | 저수준 필드 이동 | `MoveToField()` | TODO |

## 텍스트 선택 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 21 | `select_text()` | (spara, spos, epara, epos, slist) | bool | 범위 블록 선택 | `SelectText()` | TODO |
| 22 | `SelectText()` | (spara, spos, epara, epos, slist) | bool | 저수준 텍스트 선택 | `SelectText()` | TODO |
| 23 | `select_text_by_get_pos()` | (s_getpos, e_getpos) | bool | get_pos 튜플로 선택 | `SelectTextByGetPos()` | TODO |

## 문서 검색 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 24 | `init_scan()` | (option, range, spara, spos, epara, epos) | bool | 텍스트 검색 초기화 | `InitScan()` | TODO |
| 25 | `InitScan()` | (option, range, spara, spos, epara, epos) | bool | 저수준 검색 초기화 | `InitScan()` | TODO |
| 26 | `release_scan()` | () | None | 검색 정보 해제 | `ReleaseScan()` | TODO |
| 27 | `ReleaseScan()` | () | None | 저수준 검색 해제 | `ReleaseScan()` | TODO |

## 찾기/바꾸기 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 28 | `find()` | (src, direction, regex, ...) | bool | 텍스트 찾기 | `Find()` | TODO |
| 29 | `find_forward()` | (src, regex) | bool | 아래쪽으로 찾기 | `FindForward()` | TODO |
| 30 | `find_backward()` | (src, regex) | bool | 위쪽으로 찾기 | `FindBackward()` | TODO |
| 31 | `find_replace()` | (src, dst, regex, ...) | bool | 찾아 바꾸기 | `FindReplace()` | TODO |
| 32 | `find_replace_all()` | (src, dst, regex, ...) | int | 모두 바꾸기 | `FindReplaceAll()` | TODO |

## 클립보드 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 33 | `paste()` | (option) | None | 붙여넣기 (확장) | `Paste()` | TODO |

---

## 상세 설명

### insert_text()
캐럿 위치에 텍스트를 삽입한다.

```python
def insert_text(self, text):
    """
    Args:
        text: 삽입할 문자열

    Returns:
        성공시 True, 실패시 False
    """
    param = self.hwp.HParameterSet.HInsertText
    self.hwp.HAction.GetDefault("InsertText", param.HSet)
    param.Text = text
    return self.hwp.HAction.Execute("InsertText", param.HSet)
```

**C++ 구현:**
```cpp
bool Hwp::InsertText(const std::wstring& text) {
    // HParameterSet.HInsertText 획득
    IDispatch* pParamSet = GetHParameterSet(L"HInsertText");
    if (!pParamSet) return false;

    // HSet 획득
    VARIANT vHSet;
    GetProperty(pParamSet, L"HSet", &vHSet);

    // GetDefault 호출
    InvokeMethod(m_pHAction, L"GetDefault",
                 L"InsertText", vHSet.pdispVal);

    // Text 속성 설정
    SetProperty(pParamSet, L"Text", text);

    // Execute 호출
    VARIANT result;
    InvokeMethod(m_pHAction, L"Execute",
                 L"InsertText", vHSet.pdispVal, &result);

    pParamSet->Release();
    return result.boolVal;
}
```

### get_pos() / set_pos()
캐럿 위치 조회/설정

```python
def get_pos(self) -> Tuple[int, int, int]:
    """
    Returns:
        (List, Para, Pos) 튜플
        - List: 문서 내 list ID (본문이 0)
        - Para: 문단 ID (0부터 시작)
        - Pos: 문단 내 글자 위치 (0부터 시작)
    """
    return self.hwp.GetPos()

def set_pos(self, List: int, para: int, pos: int) -> bool:
    """
    Args:
        List: 문서 내 list ID
        para: 문단 ID (음수/범위 초과시 문서 시작으로)
        pos: 글자 위치 (-1이면 문단 끝으로)

    Returns:
        성공시 True, 실패시 False
    """
    self.hwp.SetPos(List=List, Para=para, pos=pos)
    return (List, para) == self.get_pos()[:2]
```

**C++ 구현:**
```cpp
std::tuple<int, int, int> Hwp::GetPos() {
    VARIANT result;
    DISPPARAMS params = {NULL, NULL, 0, 0};
    m_pHwp->Invoke(DISPID_GETPOS, IID_NULL, LOCALE_USER_DEFAULT,
                   DISPATCH_METHOD, &params, &result, NULL, NULL);

    // SafeArray에서 3개 값 추출
    SAFEARRAY* psa = result.parray;
    long indices[1];
    int list, para, pos;

    indices[0] = 0; SafeArrayGetElement(psa, indices, &list);
    indices[0] = 1; SafeArrayGetElement(psa, indices, &para);
    indices[0] = 2; SafeArrayGetElement(psa, indices, &pos);

    return std::make_tuple(list, para, pos);
}

bool Hwp::SetPos(int list, int para, int pos) {
    VARIANT vList, vPara, vPos;
    vList.vt = VT_I4; vList.lVal = list;
    vPara.vt = VT_I4; vPara.lVal = para;
    vPos.vt = VT_I4; vPos.lVal = pos;

    DISPPARAMS params = {...};
    return SUCCEEDED(m_pHwp->Invoke(DISPID_SETPOS, ...));
}
```

### move_pos()
다양한 옵션으로 캐럿을 이동한다.

```python
def move_pos(self, move_id: int = 1, para: int = 0, pos: int = 0) -> bool:
    """
    Args:
        move_id:
            - 0: 루트 리스트의 특정 위치 (moveMain)
            - 1: 현재 리스트의 특정 위치 (moveCurList, 기본값)
            - 2: 문서 시작 (moveTopOfFile)
            - 3: 문서 끝 (moveBottomOfFile)
            - 4: 현재 리스트 시작 (moveTopOfList)
            - 5: 현재 리스트 끝 (moveBottomOfList)
            - 6: 문단 시작 (moveStartOfPara)
            - 7: 문단 끝 (moveEndOfPara)
            - 8: 단어 시작 (moveStartOfWord)
            - 9: 단어 끝 (moveEndOfWord)
            - 10: 다음 문단 시작 (moveNextPara)
            - 11: 이전 문단 끝 (movePrevPara)
            - 12: 한 글자 뒤로 (moveNextPos)
            - 13: 한 글자 앞으로 (movePrevPos)
            - 14: 한 글자 뒤로 확장 (moveNextPosEx)
            - 15: 한 글자 앞으로 확장 (movePrevPosEx)
            - 16: 한 글자 뒤로 현재 리스트만 (moveNextChar)
            - 17: 한 글자 앞으로 현재 리스트만 (movePrevChar)
            - 18: 한 단어 뒤로 (moveNextWord)
            - 19: 한 단어 앞으로 (movePrevWord)
            - 20: 한 줄 아래 (moveNextLine)
            - 21: 한 줄 위 (movePrevLine)
            - 22: 줄 시작 (moveStartOfLine)
            - 23: 줄 끝 (moveEndOfLine)
            - 24: 상위 레벨 (moveParentList)
            - 25: 탑레벨 (moveTopLevelList)
            - 26: 루트 리스트 (moveRootList)
            - 27: 현재 캐럿 위치 (moveCurrentCaret)
            - 100~107: 셀 이동 (moveLeftOfCell ~ moveBottomOfCell)
            - 200: 스크린 좌표로 이동 (moveScrPos)
            - 201: GetText() 실행 후 위치로 (moveScanPos)
        para: 문단 번호 (move_id=0,1일 때 사용)
        pos: 문자 위치 (move_id=0,1일 때 사용)

    Returns:
        성공시 True, 실패시 False
    """
    return self.hwp.MovePos(moveID=move_id, Para=para, pos=pos)
```

**move_id 상수 정의 (C++):**
```cpp
enum MoveID {
    moveMain = 0,
    moveCurList = 1,
    moveTopOfFile = 2,
    moveBottomOfFile = 3,
    moveTopOfList = 4,
    moveBottomOfList = 5,
    moveStartOfPara = 6,
    moveEndOfPara = 7,
    moveStartOfWord = 8,
    moveEndOfWord = 9,
    moveNextPara = 10,
    movePrevPara = 11,
    moveNextPos = 12,
    movePrevPos = 13,
    moveNextPosEx = 14,
    movePrevPosEx = 15,
    moveNextChar = 16,
    movePrevChar = 17,
    moveNextWord = 18,
    movePrevWord = 19,
    moveNextLine = 20,
    movePrevLine = 21,
    moveStartOfLine = 22,
    moveEndOfLine = 23,
    moveParentList = 24,
    moveTopLevelList = 25,
    moveRootList = 26,
    moveCurrentCaret = 27,
    moveLeftOfCell = 100,
    moveRightOfCell = 101,
    moveUpOfCell = 102,
    moveDownOfCell = 103,
    moveStartOfCell = 104,
    moveEndOfCell = 105,
    moveTopOfCell = 106,
    moveBottomOfCell = 107,
    moveScrPos = 200,
    moveScanPos = 201
};
```

### init_scan() / get_text() / release_scan()
문서 텍스트 순차 추출

```python
def init_scan(self, option: int = 0x07, range: int = 0x77,
              spara: int = 0, spos: int = 0,
              epara: int = -1, epos: int = -1) -> bool:
    """
    Args:
        option: 찾을 대상 (기본값 0x07 = 모든 컨트롤)
            - 0x00: 본문만 (maskNormal)
            - 0x01: char 타입 컨트롤 (maskChar)
            - 0x02: inline 타입 컨트롤 (maskInline)
            - 0x04: extended 타입 컨트롤 (maskCtrl)

        range: 검색 범위 (기본값 0x77 = 문서 시작~끝)
            시작 위치:
            - 0x00: 캐럿 위치 (scanSposCurrent)
            - 0x10: 특정 위치 (scanSposSpecified)
            - 0x70: 문서 시작 (scanSposDocument)
            끝 위치:
            - 0x00: 캐럿 위치 (scanEposCurrent)
            - 0x01: 특정 위치 (scanEposSpecified)
            - 0x07: 문서 끝 (scanEposDocument)
            기타:
            - 0xFF: 블록 범위 제한 (scanWithinSelection)
            - 0x100: 역방향 (scanBackward)

    Returns:
        성공시 True, 실패시 False
    """
    return self.hwp.InitScan(option=option, Range=range,
                              spara=spara, spos=spos,
                              epara=epara, epos=epos)

def get_text(self) -> Tuple[int, str]:
    """
    Returns:
        (state, text) 튜플
        state 의미:
        - 0: 텍스트 정보 없음
        - 1: 리스트 끝
        - 2: 일반 텍스트
        - 3: 다음 문단
        - 4: 제어문자 내부로 들어감
        - 5: 제어문자 빠져나옴
        - 101: 초기화 안됨
        - 102: 텍스트 변환 실패
    """
    return self.hwp.GetText()

def release_scan(self) -> None:
    """검색 정보 해제 (반드시 호출)"""
    return self.hwp.ReleaseScan()
```

**사용 예:**
```python
hwp.init_scan()
while True:
    state, text = hwp.get_text()
    print(state, text)
    if state <= 1:
        break
hwp.release_scan()
```

### get_text_file() / set_text_file()
문서 내용을 문자열로 가져오기/삽입하기

```python
def get_text_file(self,
    format: Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"] = "UNICODE",
    option: str = "saveblock:true"
) -> str:
    """
    Args:
        format:
            - "HWP": 한/글 native (BASE64 인코딩)
            - "HWPML2X": 한/글 XML 형식
            - "HTML": HTML 형식
            - "UNICODE": 유니코드 텍스트 (기본값)
            - "TEXT": 일반 텍스트
        option: "saveblock:true"면 선택 블록만

    Returns:
        문서 내용 문자열
    """
    return self.hwp.GetTextFile(Format=format, option=option)

def set_text_file(self,
    data: str,
    format: Literal["HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"] = "HWPML2X",
    option: str = "insertfile"
) -> int:
    """
    Args:
        data: 삽입할 문자열
        format: 파일 형식
        option: "insertfile"이면 현재 커서 이후에 삽입

    Returns:
        성공시 1, 실패시 0
    """
    return self.hwp.SetTextFile(data=data, Format=format, option=option)
```

### find()
텍스트 찾기

```python
def find(self,
    src: str = "",
    direction: Literal["Forward", "Backward", "AllDoc"] = "Forward",
    regex: bool = False,
    TextColor: Optional[int] = None,
    Height: Optional[int | float] = None,
    MatchCase: int = 1,
    SeveralWords: int = 0,
    UseWildCards: int = 0,
    WholeWordOnly: int = 0,
    AutoSpell: int = 1,
    HanjaFromHangul: int = 1,
    AllWordForms: int = 0,
    FindStyle: str = "",
    ReplaceStyle: str = "",
    FindJaso: int = 0,
    FindType: int = 1
) -> bool:
    """
    Args:
        src: 찾을 문자열
        direction: 탐색 방향
            - "Forward": 아래쪽 (기본값)
            - "Backward": 위쪽
            - "AllDoc": 문서 전체 (끝에서 처음으로 순환)
        regex: 정규식 사용 여부
        TextColor: 글자색 (RGBColor)
        Height: 글자 크기 (포인트)
        MatchCase: 대소문자 구별 (1)
        SeveralWords: 여러 단어 찾기 (0)
        UseWildCards: 와일드카드 사용 (0)
        WholeWordOnly: 온전한 낱말 (0)
        AutoSpell: 토씨 자동 교정 (1)
        HanjaFromHangul: 한글로 한자 찾기 (1)
        AllWordForms: 띄어쓰기 무시 (0)
        FindStyle: 찾을 스타일
        ReplaceStyle: 바꿀 스타일
        FindJaso: 자소 단위 찾기 (0)
        FindType: 다시 찾기 타입 (1)

    Returns:
        찾으면 True, 없으면 False
    """
```

### find_replace() / find_replace_all()
찾아 바꾸기

```python
def find_replace(self, src, dst, regex=False,
    direction: Literal["Backward", "Forward", "AllDoc"] = "Forward",
    MatchCase=1, AllWordForms=0, SeveralWords=1, UseWildCards=0,
    WholeWordOnly=0, AutoSpell=1, IgnoreFindString=0,
    IgnoreReplaceString=0, ReplaceMode=1, HanjaFromHangul=1,
    FindJaso=0, FindStyle="", ReplaceStyle="", FindType=1
) -> bool:
    """
    Args:
        src: 찾을 문자열
        dst: 바꿀 문자열
        regex: True면 파이썬 re.sub 사용 (전체 문서 순회)
        direction: 탐색 방향

    Returns:
        성공시 True
    """

def find_replace_all(self, src, dst, regex=False, ...) -> int:
    """모든 항목 바꾸기. 바꾼 횟수 반환"""
```

### select_text()
범위 텍스트 선택

```python
def select_text(self,
    spara: Union[int, list, tuple] = 0,
    spos: int = 0,
    epara: int = 0,
    epos: int = 0,
    slist: int = 0
) -> bool:
    """
    Args:
        spara: 시작 문단 번호 (또는 get_selected_pos() 반환값)
        spos: 시작 문자 위치
        epara: 끝 문단 번호
        epos: 끝 문자 위치 (포함되지 않음, -1이면 문단 끝)
        slist: list ID

    Returns:
        성공시 True

    Examples:
        # 본문 세 번째 문단 전체 선택
        hwp.select_text(2, 0, 2, -1, 0)
    """
```

### get_selected_text()
선택 영역 텍스트 추출

```python
def get_selected_text(self, as_: Literal["list", "str"] = "str",
                      keep_select: bool = False):
    """
    Args:
        as_: 반환 형식 ("str" 또는 "list")
        keep_select: 선택 상태 유지 여부

    Returns:
        선택한 텍스트 (str 또는 list)
    """
```

### paste()
붙여넣기 확장

```python
def paste(self, option: Literal[0, 1, 2, 3, 4, 5, 6] = 4):
    """
    Args:
        option:
            - 0: 왼쪽에 끼워넣기
            - 1: 오른쪽에 끼워넣기
            - 2: 위쪽에 끼워넣기
            - 3: 아래쪽에 끼워넣기
            - 4: 덮어쓰기 (기본값)
            - 5: 내용만 덮어쓰기
            - 6: 셀 안에 표로 넣기
    """
```

---

## C++ 클래스 구조

```cpp
// Hwp.h - 텍스트 편집 메서드
class Hwp : public ParamHelpers, public RunMethods {
public:
    // 텍스트 삽입/추출
    bool InsertText(const std::wstring& text);
    std::tuple<int, std::wstring> GetText();
    std::wstring GetTextFile(const std::wstring& format = L"UNICODE",
                             const std::wstring& option = L"saveblock:true");
    int SetTextFile(const std::wstring& data,
                    const std::wstring& format = L"HWPML2X",
                    const std::wstring& option = L"insertfile");
    std::wstring GetSelectedText(bool asList = false,
                                  bool keepSelect = false);

    // 캐럿 위치
    std::tuple<int, int, int> GetPos();
    IDispatch* GetPosBySet();
    bool SetPos(int list, int para, int pos);
    bool SetPosBySet(IDispatch* dispVal);
    bool MovePos(int moveId = 1, int para = 0, int pos = 0);
    bool MoveToField(const std::wstring& field, int idx = 0,
                     bool text = true, bool start = true,
                     bool select = false);

    // 텍스트 선택
    bool SelectText(int spara = 0, int spos = 0,
                    int epara = 0, int epos = 0, int slist = 0);
    bool SelectTextByGetPos(std::tuple<int,int,int> start,
                            std::tuple<int,int,int> end);

    // 문서 검색
    bool InitScan(int option = 0x07, int range = 0x77,
                  int spara = 0, int spos = 0,
                  int epara = -1, int epos = -1);
    void ReleaseScan();

    // 찾기/바꾸기
    bool Find(const std::wstring& src,
              const std::wstring& direction = L"Forward",
              bool regex = false);
    bool FindForward(const std::wstring& src, bool regex = false);
    bool FindBackward(const std::wstring& src, bool regex = false);
    bool FindReplace(const std::wstring& src, const std::wstring& dst,
                     bool regex = false);
    int FindReplaceAll(const std::wstring& src, const std::wstring& dst,
                       bool regex = false);

    // 클립보드
    void Paste(int option = 4);
};
```

---

## 구현 우선순위

1. **Critical**: InsertText(), GetText(), GetTextFile(), SetTextFile()
2. **Critical**: GetPos(), SetPos(), MovePos() - 커서 제어 필수
3. **High**: InitScan(), ReleaseScan() - 텍스트 순회에 필수
4. **High**: SelectText(), GetSelectedText() - 블록 선택
5. **Medium**: Find(), FindReplace(), FindReplaceAll()
6. **Medium**: MoveToField() - 필드 기능과 연계
7. **Low**: Paste() - Run 액션으로 대체 가능
