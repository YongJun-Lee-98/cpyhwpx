# 파일 I/O 매핑

## 요약
- 총 함수: 약 25개
- 포팅 완료: 0개 (0%)
- 우선순위: High (Phase 2)

---

## 파일 열기/저장 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 1 | `open()` | (path, format="", arg="") | bool | 파일 열기 | `Open()` | TODO |
| 2 | `Open()` | (filename, Format, arg) | bool | 저수준 파일 열기 | `Open()` | TODO |
| 3 | `save()` | (save_if_dirty=True) | bool | 저장 | `Save()` | TODO |
| 4 | `Save()` | (save_if_dirty) | bool | 저수준 저장 | `Save()` | TODO |
| 5 | `save_as()` | (path, format="HWP", arg="") | bool | 다른 이름으로 저장 | `SaveAs()` | TODO |
| 6 | `SaveAs()` | (Path, Format, arg) | bool | 저수준 다른 이름 저장 | `SaveAs()` | TODO |
| 7 | `clear()` | (option=1) | None | 새 문서로 초기화 | `Clear()` | TODO |
| 8 | `Clear()` | (option) | None | 저수준 새 문서 | `Clear()` | TODO |
| 9 | `close()` | (is_dirty=False) | None | 문서 닫기 | `Close()` | TODO |
| 10 | `FileClose()` | () | bool | 파일 닫기 액션 | `FileClose()` | TODO |
| 11 | `quit()` | (save=False) | None | 한/글 종료 | `Quit()` | TODO |
| 12 | `Quit()` | (save) | None | 저수준 종료 | `Quit()` | TODO |

## 파일 변환/삽입 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 13 | `insert_file()` | (filename, format, option) | bool | 파일 삽입 | `InsertFile()` | TODO |
| 14 | `InsertFile()` | (filename, Format, arg) | bool | 저수준 파일 삽입 | `InsertFile()` | TODO |
| 15 | `save_block_as()` | (path, format, attributes) | bool | 선택 영역 저장 | `SaveBlockAs()` | TODO |
| 16 | `get_text_file()` | (format, option) | str | 텍스트 파일로 가져오기 | `GetTextFile()` | TODO |
| 17 | `GetTextFile()` | (format, option) | str | 저수준 텍스트 가져오기 | `GetTextFile()` | TODO |
| 18 | `set_text_file()` | (filename, format, option) | bool | 텍스트 파일에서 삽입 | `SetTextFile()` | TODO |
| 19 | `SetTextFile()` | (data, format, option) | bool | 저수준 텍스트 삽입 | `SetTextFile()` | TODO |
| 20 | `open_pdf()` | (pdf_path, this_window) | bool | PDF 열기 | `OpenPDF()` | TODO |

## 문서 정보 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 21 | `get_file_info()` | (filename) | dict | 파일 정보 조회 | `GetFileInfo()` | TODO |
| 22 | `GetFileInfo()` | (filename) | ParameterSet | 저수준 파일 정보 | `GetFileInfo()` | TODO |
| 23 | `is_empty()` | () | bool | 빈 문서 여부 | `IsEmpty()` | TODO |
| 24 | `is_modified()` | () | bool | 수정 여부 | `IsModified()` | TODO |

## 보안/모듈 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 25 | `register_module()` | () | bool | 보안 모듈 등록 | `RegisterModule()` | TODO |
| 26 | `RegisterModule()` | (ModuleType, ModuleName) | bool | 저수준 모듈 등록 | `RegisterModule()` | TODO |

---

## 상세 설명

### open() / Open()
파일 열기

```python
def open(self, path: str, format: str = "", arg: str = "") -> bool:
    """
    Args:
        path: 파일 경로
        format: 파일 형식 ("HWP", "HWPX", "HWP30", "HTML", "TEXT", "UNICODE", "DOC", "DOCX" 등)
        arg: 추가 인자 (예: "lock:false" - 잠금 해제)

    Returns:
        성공시 True, 실패시 False
    """
    return self.hwp.Open(path, format, arg)
```

**C++ 구현:**
```cpp
bool Hwp::Open(const std::wstring& path,
               const std::wstring& format,
               const std::wstring& arg) {
    VARIANT vPath, vFormat, vArg, result;
    VariantInit(&vPath); VariantInit(&vFormat);
    VariantInit(&vArg); VariantInit(&result);

    vPath.vt = VT_BSTR;
    vPath.bstrVal = SysAllocString(path.c_str());
    vFormat.vt = VT_BSTR;
    vFormat.bstrVal = SysAllocString(format.c_str());
    vArg.vt = VT_BSTR;
    vArg.bstrVal = SysAllocString(arg.c_str());

    DISPPARAMS params = { ... };
    HRESULT hr = m_pHwp->Invoke(DISPID_OPEN, ..., &result);

    // Cleanup
    SysFreeString(vPath.bstrVal);
    SysFreeString(vFormat.bstrVal);
    SysFreeString(vArg.bstrVal);

    return (hr == S_OK && result.boolVal);
}
```

### save_as() / SaveAs()
다른 이름으로 저장

```python
def save_as(self, path: str, format: str = "HWP", arg: str = "") -> bool:
    """
    Args:
        path: 저장 경로
        format: 저장 형식
            - "HWP": 한/글 2007 이후 (.hwp)
            - "HWPX": 한/글 2022 이후 (.hwpx)
            - "HWP30": 한/글 3.0 (.hwp)
            - "HTML": HTML (.html)
            - "TEXT": 텍스트 (.txt)
            - "UNICODE": 유니코드 텍스트
            - "DOC": MS Word (.doc)
            - "DOCX": MS Word 2007 이후 (.docx)
            - "PDF": PDF
        arg: 추가 인자

    Returns:
        성공시 True
    """
```

### clear() / Clear()
새 문서로 초기화

```python
def clear(self, option: int = 1) -> None:
    """
    Args:
        option:
            - 0: 저장 확인 대화상자 표시
            - 1: 저장하지 않고 초기화 (기본값)
    """
    return self.hwp.Clear(option)
```

### register_module() / RegisterModule()
보안 모듈 등록 (필수)

```python
def register_module(self) -> bool:
    """
    FilePathCheckerModule.dll을 등록하여 파일 열기/저장 시
    보안 팝업이 뜨지 않도록 함.

    Returns:
        성공시 True
    """
    module_path = os.path.join(pyinstaller_path, "FilePathCheckerModule.dll")
    return self.hwp.RegisterModule("FilePathCheckDLL", module_path)
```

**C++ 구현:**
```cpp
bool Hwp::RegisterModule() {
    // DLL 경로 구하기
    wchar_t modulePath[MAX_PATH];
    GetModuleFileNameW(NULL, modulePath, MAX_PATH);
    std::wstring dllPath = std::wstring(modulePath);
    dllPath = dllPath.substr(0, dllPath.find_last_of(L"\\"));
    dllPath += L"\\FilePathCheckerModule.dll";

    VARIANT vType, vPath, result;
    // ... COM 호출
    return (hr == S_OK && result.boolVal);
}
```

### GetTextFile() / SetTextFile()
텍스트 형식으로 가져오기/삽입

```python
def GetTextFile(self, format: str = "TEXT", option: str = "") -> str:
    """
    Args:
        format: "TEXT", "UNICODE", "HWPML2X", "HTML" 등
        option: "saveblock" - 선택 영역만

    Returns:
        텍스트 또는 XML 문자열
    """

def SetTextFile(self, data: str, format: str = "TEXT", option: str = "") -> bool:
    """
    Args:
        data: 삽입할 텍스트/XML
        format: "TEXT", "UNICODE", "HWPML2X", "HTML" 등
        option: "insertfile" - 파일 삽입 모드
    """
```

### save_block_as()
선택 영역을 파일로 저장

```python
def save_block_as(self, path: str, format: str = "HWP", attributes: str = "") -> bool:
    """
    Args:
        path: 저장 경로
        format: 저장 형식
        attributes: 추가 속성

    Returns:
        성공시 True
    """
```

---

## 지원 파일 형식

| Format | 확장자 | 설명 |
|--------|-------|------|
| HWP | .hwp | 한/글 2007 이후 기본 형식 |
| HWPX | .hwpx | 한/글 2022 이후 XML 형식 |
| HWP30 | .hwp | 한/글 3.0 형식 |
| TEXT | .txt | 텍스트 (ANSI) |
| UNICODE | .txt | 텍스트 (UTF-16) |
| HTML | .html | HTML |
| HWPML2X | .xml | 한/글 XML |
| DOC | .doc | MS Word 97-2003 |
| DOCX | .docx | MS Word 2007 이후 |
| PDF | .pdf | PDF |
| RTF | .rtf | Rich Text Format |

---

## C++ 클래스 구조

```cpp
// Hwp.h - 파일 I/O 관련 메서드
class Hwp {
public:
    // 파일 열기/저장
    bool Open(const std::wstring& path,
              const std::wstring& format = L"",
              const std::wstring& arg = L"");
    bool Save(bool saveIfDirty = true);
    bool SaveAs(const std::wstring& path,
                const std::wstring& format = L"HWP",
                const std::wstring& arg = L"");
    void Clear(int option = 1);
    void Close(bool isDirty = false);
    void Quit(bool save = false);

    // 파일 변환/삽입
    bool InsertFile(const std::wstring& filename,
                    const std::wstring& format,
                    const std::wstring& option);
    bool SaveBlockAs(const std::wstring& path,
                     const std::wstring& format,
                     const std::wstring& attributes);
    std::wstring GetTextFile(const std::wstring& format = L"TEXT",
                             const std::wstring& option = L"");
    bool SetTextFile(const std::wstring& data,
                     const std::wstring& format = L"TEXT",
                     const std::wstring& option = L"");

    // 문서 정보
    bool IsEmpty();
    bool IsModified();

    // 보안 모듈
    bool RegisterModule();
};
```

---

## 구현 우선순위

1. **Critical**: RegisterModule() - 보안 팝업 제거 필수
2. **Critical**: Open(), Save(), SaveAs() - 기본 파일 조작
3. **High**: Clear(), Close(), Quit() - 생명주기 관리
4. **Medium**: GetTextFile(), SetTextFile() - 텍스트 변환
5. **Low**: InsertFile(), SaveBlockAs() - 고급 기능
