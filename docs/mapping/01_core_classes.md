# 핵심 클래스 매핑

## 요약
- 총 클래스: 4개 (Ctrl, XHwpDocuments, XHwpDocument, Hwp)
- 포팅 완료: 0개 (0%)
- 우선순위: Critical

---

## 1. Ctrl 클래스 (core.py:398~671)

컨트롤(표, 그림, 글상자, 각주/미주 등) 래퍼 클래스

### 생성자
| Python | C++ | 설명 |
|--------|-----|------|
| `__init__(self, com_obj)` | `Ctrl(IDispatch* pComObj)` | COM 객체로 초기화 |

### 메서드
| # | Python 메서드 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|--------------|---------|--------|------|-----------|------|
| 1 | `GetCtrlInstID()` | () | str | 컨트롤 고유 ID (한글2024+) | `GetCtrlInstID()` | TODO |
| 2 | `GetAnchorPos()` | (type_: int = 0) | HParameterSet | 앵커 위치 조회 | `GetAnchorPos()` | TODO |

### 속성 (Properties)
| # | Python 속성 | 타입 | R/W | 역할 | C++ 함수명 | 상태 |
|---|------------|------|-----|------|-----------|------|
| 1 | `CtrlCh` | int | R | 컨트롤 문자 (1~31) | `GetCtrlCh()` | TODO |
| 2 | `CtrlID` | str | R | 컨트롤 ID (tbl, gso 등) | `GetCtrlID()` | TODO |
| 3 | `HasList` | bool | R | 리스트 포함 여부 | `GetHasList()` | TODO |
| 4 | `Next` | Ctrl | R | 다음 컨트롤 | `GetNext()` | TODO |
| 5 | `Prev` | Ctrl | R | 이전 컨트롤 | `GetPrev()` | TODO |
| 6 | `Properties` | Any | R/W | 컨트롤 속성 파라미터셋 | `GetProperties()` / `SetProperties()` | TODO |
| 7 | `UserDesc` | str | R | 컨트롤 종류 문자열 | `GetUserDesc()` | TODO |

### CtrlID 목록
```
"cold" : 단
"secd" : 구역
"fn"   : 각주
"en"   : 미주
"tbl"  : 표
"eqed" : 수식
"gso"  : 그리기 개체 (그림 포함)
"atno" : 번호 넣기
"nwno" : 새 번호로
"pgct" : 페이지 번호 제어
"pghd" : 감추기
"pgnp" : 쪽 번호 위치
"head" : 머리말
"foot" : 꼬리말
"%dte" : 날짜/시간 필드
"%ddt" : 파일 작성 날짜/시간
"%pat" : 문서 경로
"%bmk" : 블록 책갈피
"%mmg" : 메일 머지
"%xrf" : 상호 참조
"%fmu" : 계산식
"%clk" : 누름틀
"%smr" : 문서 요약 정보
"%usr" : 사용자 정보
"%hlk" : 하이퍼링크
"bokm" : 책갈피
"idxm" : 찾아보기
"tdut" : 덧말
"tcmt" : 주석
```

---

## 2. XHwpDocuments 클래스 (core.py:672~771)

열린 문서 컬렉션 래퍼 클래스

### 생성자
| Python | C++ | 설명 |
|--------|-----|------|
| `__init__(self, com_obj)` | `XHwpDocuments(IDispatch* pComObj)` | COM 객체로 초기화 |

### 특수 메서드
| # | Python | C++ | 역할 |
|---|--------|-----|------|
| 1 | `__getitem__(index)` | `operator[]` 또는 `Item(int index)` | 인덱스로 문서 접근 |
| 2 | `__iter__()` | `begin()`, `end()` | 반복자 |
| 3 | `__len__()` | `size()` 또는 `Count()` | 문서 개수 |

### 속성 (Properties)
| # | Python 속성 | 타입 | R/W | 역할 | C++ 함수명 | 상태 |
|---|------------|------|-----|------|-----------|------|
| 1 | `Active_XHwpDocument` | XHwpDocument | R | 현재 활성 문서 | `GetActiveDocument()` | TODO |
| 2 | `Application` | Any | R | 애플리케이션 객체 | `GetApplication()` | TODO |
| 3 | `CLSID` | str | R | 클래스 ID | `GetCLSID()` | TODO |
| 4 | `Count` | int | R | 문서 개수 | `GetCount()` | TODO |

### 메서드
| # | Python 메서드 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|--------------|---------|--------|------|-----------|------|
| 1 | `Add()` | (isTab: bool = False) | XHwpDocument | 새 문서 추가 | `Add()` | TODO |
| 2 | `Close()` | (isDirty: bool = False) | None | 문서창 닫기 | `Close()` | TODO |
| 3 | `FindItem()` | (lDocID: int) | XHwpDocument | ID로 문서 찾기 | `FindItem()` | TODO |

---

## 3. XHwpDocument 클래스 (core.py:774~930)

단일 문서 래퍼 클래스

### 생성자
| Python | C++ | 설명 |
|--------|-----|------|
| `__init__(self, com_obj)` | `XHwpDocument(IDispatch* pComObj)` | COM 객체로 초기화 |

### 속성 (Properties)
| # | Python 속성 | 타입 | R/W | 역할 | C++ 함수명 | 상태 |
|---|------------|------|-----|------|-----------|------|
| 1 | `Application` | Any | R | 애플리케이션 객체 | `GetApplication()` | TODO |
| 2 | `CLSID` | str | R | 클래스 ID | `GetCLSID()` | TODO |
| 3 | `DocumentID` | int | R | 문서 고유 ID | `GetDocumentID()` | TODO |
| 4 | `EditMode` | int | R | 편집 모드 | `GetEditMode()` | TODO |
| 5 | `Format` | str | R | 문서 형식 | `GetFormat()` | TODO |
| 6 | `FullName` | str | R | 전체 경로 | `GetFullName()` | TODO |
| 7 | `Modified` | int | R | 수정 여부 | `GetModified()` | TODO |
| 8 | `Path` | str | R | 문서 경로 | `GetPath()` | TODO |
| 9 | `XHwpCharacterShape` | Any | R | 글자모양 설정 | `GetXHwpCharacterShape()` | TODO |
| 10 | `XHwpDocumentInfo` | Any | R | 문서 정보 | `GetXHwpDocumentInfo()` | TODO |
| 11 | `XHwpFind` | Any | R | 찾기 기능 | `GetXHwpFind()` | TODO |
| 12 | `XHwpFormCheckButtons` | Any | R | 체크박스 양식 | `GetXHwpFormCheckButtons()` | TODO |
| 13 | `XHwpFormComboBoxs` | Any | R | 콤보박스 양식 | `GetXHwpFormComboBoxs()` | TODO |
| 14 | `XHwpFormEdits` | Any | R | 편집 양식 | `GetXHwpFormEdits()` | TODO |
| 15 | `XHwpFormPushButtons` | Any | R | 버튼 양식 | `GetXHwpFormPushButtons()` | TODO |
| 16 | `XHwpFormRadioButtons` | Any | R | 라디오버튼 양식 | `GetXHwpFormRadioButtons()` | TODO |
| 17 | `XHwpParagraphShape` | Any | R | 문단 모양 | `GetXHwpParagraphShape()` | TODO |
| 18 | `XHwpPrint` | Any | R | 인쇄 제어 | `GetXHwpPrint()` | TODO |
| 19 | `XHwpRange` | Any | R | 범위 설정 | `GetXHwpRange()` | TODO |
| 20 | `XHwpSelection` | Any | R | 선택 영역 | `GetXHwpSelection()` | TODO |
| 21 | `XHwpSendMail` | Any | R | 메일 보내기 | `GetXHwpSendMail()` | TODO |
| 22 | `XHwpSummaryInfo` | Any | R | 요약 정보 | `GetXHwpSummaryInfo()` | TODO |

### 메서드
| # | Python 메서드 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|--------------|---------|--------|------|-----------|------|
| 1 | `Clear()` | (option: bool = False) | None | 문서 내용 삭제 | `Clear()` | TODO |
| 2 | `Close()` | (isDirty: bool = False) | None | 문서 닫기 | `Close()` | TODO |
| 3 | `Open()` | (filename: str, Format: str, arg: str) | bool | 파일 열기 | `Open()` | TODO |
| 4 | `Redo()` | (Count: int) | bool | 다시 실행 | `Redo()` | TODO |
| 5 | `Save()` | (save_if_dirty: bool) | bool | 저장 | `Save()` | TODO |
| 6 | `SaveAs()` | (Path: str, Format: str, arg: str) | bool | 다른 이름 저장 | `SaveAs()` | TODO |
| 7 | `SendBrowser()` | () | bool | 브라우저로 보내기 | `SendBrowser()` | TODO |
| 8 | `SetActive_XHwpDocument()` | () | None | 활성 문서 설정 | `SetActive()` | TODO |
| 9 | `Undo()` | (Count: int) | bool | 실행 취소 | `Undo()` | TODO |

---

## 4. Hwp 클래스 (core.py:933~)

메인 자동화 클래스 (ParamHelpers, RunMethods 상속)

### 생성자
| Python | C++ | 설명 |
|--------|-----|------|
| `__init__(new, visible, register_module, on_quit)` | `Hwp(bool new_, bool visible, bool register_module)` | 한/글 인스턴스 생성 |

**파라미터:**
- `new` (bool): True면 새 인스턴스, False면 기존 인스턴스 연결
- `visible` (bool): True면 창 표시, False면 백그라운드
- `register_module` (bool): True면 보안모듈 자동 등록
- `on_quit` (bool): True면 소멸자에서 자동 종료

### 소멸자
```python
def __del__(self):
    if self.on_quit:
        self.quit(save=False)
    pythoncom.CoUninitialize()
```

**C++ 구현:**
```cpp
~Hwp() {
    if (m_bOnQuit) {
        Quit(false);
    }
    CoUninitialize();
}
```

### 주요 속성 (Properties) - 02_properties.md에서 상세 다룸
(약 30개 이상의 속성)

### 주요 메서드 - 카테고리별 문서에서 상세 다룸
(약 200개 이상의 메서드)

---

## C++ 클래스 구조 설계

```cpp
// Ctrl.h
class Ctrl {
public:
    Ctrl(IDispatch* pComObj);
    ~Ctrl();

    // Methods
    std::wstring GetCtrlInstID();
    IDispatch* GetAnchorPos(int type_ = 0);

    // Properties (getter)
    int GetCtrlCh() const;
    std::wstring GetCtrlID() const;
    bool GetHasList() const;
    Ctrl* GetNext();
    Ctrl* GetPrev();
    IDispatch* GetProperties();
    void SetProperties(IDispatch* prop);
    std::wstring GetUserDesc() const;

private:
    IDispatch* m_pComObj;
};

// XHwpDocuments.h
class XHwpDocuments {
public:
    XHwpDocuments(IDispatch* pComObj);
    ~XHwpDocuments();

    // Collection interface
    XHwpDocument* operator[](int index);
    int size() const;

    // Properties
    XHwpDocument* GetActiveDocument();
    IDispatch* GetApplication();
    int GetCount() const;

    // Methods
    XHwpDocument* Add(bool isTab = false);
    void Close(bool isDirty = false);
    XHwpDocument* FindItem(int lDocID);

private:
    IDispatch* m_pComObj;
};

// XHwpDocument.h
class XHwpDocument {
public:
    XHwpDocument(IDispatch* pComObj);
    ~XHwpDocument();

    // Properties
    int GetDocumentID() const;
    std::wstring GetFullName() const;
    std::wstring GetPath() const;
    bool GetModified() const;
    // ... 기타 속성

    // Methods
    void Clear(bool option = false);
    void Close(bool isDirty = false);
    bool Open(const std::wstring& filename, const std::wstring& format, const std::wstring& arg);
    bool Save(bool saveIfDirty);
    bool SaveAs(const std::wstring& path, const std::wstring& format, const std::wstring& arg);
    bool Undo(int count);
    bool Redo(int count);

private:
    IDispatch* m_pComObj;
};

// Hwp.h
class Hwp : public ParamHelpers, public RunMethods {
public:
    Hwp(bool new_ = false, bool visible = true, bool registerModule = true);
    ~Hwp();

    // 내부 COM 객체 접근
    IDispatch* GetHwpObject() { return m_pHwp; }

    // 속성 (02_properties.md)
    // 메서드 (03~08 문서)
    // Run 액션 (09_run_actions.md)
    // 파라미터 헬퍼 (10_param_helpers.md)

private:
    IDispatch* m_pHwp;
    bool m_bInitialized;
    bool m_bOnQuit;
};
```

---

## 구현 우선순위

1. **Critical**: Hwp 클래스 기본 구조 (생성자, 소멸자, COM 초기화)
2. **High**: Ctrl 클래스 (컨트롤 순회에 필수)
3. **High**: XHwpDocument 클래스 (파일 I/O)
4. **Medium**: XHwpDocuments 클래스 (다중 문서 관리)
