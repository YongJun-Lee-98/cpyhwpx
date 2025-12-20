# 필드/메타태그 매핑

## 요약
- 총 함수: 약 25개
- 포팅 완료: 0개 (0%)
- 우선순위: Medium (Phase 5)

---

## 필드 메서드

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 1 | `create_field()` | (name, direction, memo) | bool | 누름틀 필드 생성 | `CreateField()` | TODO |
| 2 | `CreateField()` | (name, direction, memo) | bool | 저수준 필드 생성 | `CreateField()` | TODO |
| 3 | `get_field_list()` | (number, option) | str | 필드 목록 조회 | `GetFieldList()` | TODO |
| 4 | `GetFieldList()` | (number, option) | str | 저수준 필드 목록 | `GetFieldList()` | TODO |
| 5 | `get_field_text()` | (field, idx) | str | 필드 텍스트 조회 | `GetFieldText()` | TODO |
| 6 | `GetFieldText()` | (field, idx) | str | 저수준 텍스트 조회 | `GetFieldText()` | TODO |
| 7 | `put_field_text()` | (field, text, idx) | None | 필드에 텍스트 삽입 | `PutFieldText()` | TODO |
| 8 | `PutFieldText()` | (field, text, idx) | None | 저수준 텍스트 삽입 | `PutFieldText()` | TODO |
| 9 | `rename_field()` | (oldname, newname) | bool | 필드 이름 변경 | `RenameField()` | TODO |
| 10 | `RenameField()` | (oldname, newname) | bool | 저수준 이름 변경 | `RenameField()` | TODO |
| 11 | `set_cur_field_name()` | (field, direction, memo, option) | bool | 현재 셀 필드명 설정 | `SetCurFieldName()` | TODO |
| 12 | `set_field_view_option()` | (option) | int | 필드 보기 옵션 | `SetFieldViewOption()` | TODO |
| 13 | `move_to_field()` | (field, idx, text, start, select) | bool | 필드로 이동 | `MoveToField()` | TODO |
| 14 | `MoveToField()` | (field, idx, text, start, select) | bool | 저수준 필드 이동 | `MoveToField()` | TODO |
| 15 | `get_field_info()` | () | List[dict] | 모든 필드 정보 | `GetFieldInfo()` | TODO |
| 16 | `fields_to_dict()` | () | dict | 필드를 딕셔너리로 | `FieldsToDict()` | TODO |
| 17 | `set_field_by_bracket()` | () | None | 중괄호를 필드로 변환 | `SetFieldByBracket()` | TODO |
| 18 | `delete_all_fields()` | () | bool | 모든 필드 삭제 | `DeleteAllFields()` | TODO |
| 19 | `delete_field_by_name()` | (field_name, idx) | bool | 이름으로 필드 삭제 | `DeleteFieldByName()` | TODO |

## 메타태그 메서드 (한글2024+)

| # | Python 함수 | 시그니처 | 반환값 | 역할 | C++ 함수명 | 상태 |
|---|------------|---------|--------|------|-----------|------|
| 20 | `get_metatag_list()` | (number, option) | str | 메타태그 목록 | `GetMetatagList()` | TODO |
| 21 | `GetMetatagList()` | (number, option) | str | 저수준 목록 조회 | `GetMetatagList()` | TODO |
| 22 | `get_metatag_name_text()` | (tag) | str | 메타태그 텍스트 조회 | `GetMetatagNameText()` | TODO |
| 23 | `GetMetatagNameText()` | (tag) | str | 저수준 텍스트 조회 | `GetMetatagNameText()` | TODO |
| 24 | `put_metatag_name_text()` | (tag, text) | bool | 메타태그에 텍스트 삽입 | `PutMetatagNameText()` | TODO |
| 25 | `PutMetatagNameText()` | (tag, text) | bool | 저수준 텍스트 삽입 | `PutMetatagNameText()` | TODO |
| 26 | `rename_metatag()` | (oldtag, newtag) | bool | 메타태그 이름 변경 | `RenameMetatag()` | TODO |
| 27 | `RenameMetatag()` | (oldtag, newtag) | bool | 저수준 이름 변경 | `RenameMetatag()` | TODO |

---

## 상세 설명

### create_field()
캐럿 위치에 누름틀 필드를 생성한다.

```python
def create_field(self, name: str, direction: str = "", memo: str = "") -> bool:
    """
    Args:
        name: 필드 이름 (필수)
        direction: 안내문/지시문 (입력 전 표시됨)
        memo: 필드 설명/도움말

    Returns:
        성공시 True, 실패시 False

    Examples:
        >>> hwp.create_field(name="name", direction="이름", memo="이름 입력 필드")
        >>> hwp.put_field_text("name", "홍길동")
    """
    return self.hwp.CreateField(Direction=direction, memo=memo, name=name)
```

**C++ 구현:**
```cpp
bool Hwp::CreateField(const std::wstring& name,
                      const std::wstring& direction,
                      const std::wstring& memo) {
    VARIANT vName, vDirection, vMemo, result;
    // ... VARIANT 초기화 및 값 설정

    DISPID namedArgs[] = {DISPID_DIRECTION, DISPID_MEMO, DISPID_NAME};
    VARIANT args[] = {vDirection, vMemo, vName};
    DISPPARAMS params = {args, namedArgs, 3, 3};

    HRESULT hr = m_pHwp->Invoke(DISPID_CREATEFIELD, IID_NULL,
                                 LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                                 &params, &result, NULL, NULL);

    return (hr == S_OK && result.boolVal);
}
```

### get_field_list()
문서 내 모든 필드 목록을 조회한다.

```python
def get_field_list(self, number: int = 1, option: int = 0) -> str:
    """
    Args:
        number: 필드 식별 방식
            - 0 (hwpFieldPlain): 순서대로 나열
                예: "title\x02body\x02title\x02body"
            - 1 (hwpFieldNumber): 일련번호 포함
                예: "title{{0}}\x02body{{0}}\x02title{{1}}\x02body{{1}}"
            - 2 (hwpFieldCount): 개수 포함
                예: "title{{2}}\x02body{{2}}"

        option: 필드 유형 필터
            - 0x00: 모든 필드
            - 0x01 (hwpFieldCell): 셀 필드만
            - 0x02 (hwpFieldClickHere): 누름틀만
            - 0x04 (hwpFieldSelection): 선택 영역 내 필드만

    Returns:
        필드 목록 (0x02로 구분)
        예: "필드1\x02필드2\x02필드3"
    """
    return self.hwp.GetFieldList(Number=number, option=option)
```

### get_field_text()
필드의 텍스트를 조회한다.

```python
def get_field_text(self, field: Union[str, list, tuple, set], idx: int = 0) -> str:
    """
    Args:
        field: 필드 이름(들)
            - str: 단일 필드 또는 0x02로 구분된 여러 필드
            - list/tuple/set: 여러 필드 이름
            - 동일 이름의 n번째 필드: "필드명{{n}}"

        idx: 동일 이름 필드의 인덱스
            (field에 {{n}}이 없을 때 사용)

    Returns:
        필드 텍스트 (여러 개면 0x02로 구분)

    Examples:
        >>> hwp.get_field_text("title")  # 첫 번째 title 필드
        >>> hwp.get_field_text("title{{1}}")  # 두 번째 title 필드
        >>> hwp.get_field_text("title", idx=1)  # 동일
        >>> hwp.get_field_text(["title", "body"])  # 여러 필드
    """
```

### put_field_text()
필드에 텍스트를 삽입한다.

```python
def put_field_text(self,
    field: Any = "",
    text: Union[str, list, tuple, pd.Series] = "",
    idx=None
) -> None:
    """
    Args:
        field: 필드 이름(들)
            - str: 필드 이름 (0x02로 구분하여 여러 개 가능)
            - dict: {필드명: 텍스트} 형태 (text 파라미터 무시)
            - list/tuple: 필드 이름 리스트
            - DataFrame: 컬럼=필드명, 값=텍스트

        text: 삽입할 텍스트
            - str: 단일 텍스트
            - list/tuple: 여러 텍스트

        idx: 동일 이름 필드의 인덱스

    Returns:
        None

    Examples:
        >>> # 단일 필드에 텍스트 삽입
        >>> hwp.put_field_text("name", "홍길동")

        >>> # 여러 필드에 텍스트 삽입
        >>> hwp.put_field_text(["name", "age"], ["홍길동", "30"])

        >>> # 딕셔너리로 삽입
        >>> hwp.put_field_text({"name": "홍길동", "age": "30"})

        >>> # 동일 이름 필드 여러 개에 각각 다른 값 삽입
        >>> hwp.put_field_text("title", ["제목1", "제목2", "제목3"])

        >>> # DataFrame으로 삽입
        >>> df = pd.DataFrame({"name": ["홍길동"], "age": ["30"]})
        >>> hwp.put_field_text(df)
    """
```

**C++ 구현 (단순 버전):**
```cpp
void Hwp::PutFieldText(const std::wstring& field,
                       const std::wstring& text) {
    VARIANT vField, vText;
    vField.vt = VT_BSTR;
    vField.bstrVal = SysAllocString(field.c_str());
    vText.vt = VT_BSTR;
    vText.bstrVal = SysAllocString(text.c_str());

    DISPPARAMS params = {...};
    m_pHwp->Invoke(DISPID_PUTFIELDTEXT, ...);

    SysFreeString(vField.bstrVal);
    SysFreeString(vText.bstrVal);
}
```

### rename_field()
필드 이름을 변경한다.

```python
def rename_field(self, oldname: str, newname: str) -> bool:
    """
    Args:
        oldname: 기존 필드 이름 (0x02로 구분하여 여러 개 가능)
        newname: 새 필드 이름 (oldname과 동일한 개수)

    Returns:
        성공시 True

    Examples:
        >>> hwp.create_field("asdf")
        >>> hwp.rename_field("asdf", "zxcv")
        >>> hwp.put_field_text("zxcv", "Hello!")

        >>> # 여러 개 동시 변경
        >>> hwp.rename_field("title{{0}}\x02title{{1}}", "tt1\x02tt2")
    """
    return self.hwp.RenameField(oldname=oldname, newname=newname)
```

### set_cur_field_name()
현재 셀에 필드명을 설정한다. (표 안에서만 사용)

```python
def set_cur_field_name(self,
    field: str = "",
    direction: str = "",
    memo: str = "",
    option: int = 0
) -> bool:
    """
    Args:
        field: 필드 이름
        direction: 안내문
        memo: 메모
        option:
            - 0: 모든 필드
            - 1 (hwpFieldCell): 셀 필드
            - 2 (hwpFieldClickHere): 누름틀 필드

    Returns:
        성공시 True

    Examples:
        >>> hwp.create_table(5, 5)
        >>> hwp.TableCellBlockExtendAbs()
        >>> hwp.TableCellBlockExtend()  # 전체 셀 선택
        >>> hwp.set_cur_field_name("target", option=1)  # 모든 셀에 필드명 설정
        >>> hwp.put_field_text("target", list(range(1, 26)))
    """
```

### move_to_field()
지정한 필드로 캐럿을 이동한다.

```python
def move_to_field(self,
    field: str,
    idx: int = 0,
    text: bool = True,
    start: bool = True,
    select: bool = False
) -> bool:
    """
    Args:
        field: 필드 이름 (또는 "필드명{{n}}" 형식)
        idx: 동일 이름 필드의 인덱스
        text: 누름틀 내부 텍스트로 이동할지 여부
        start: 필드 처음(True) 또는 끝(False)으로 이동
        select: 필드 내용을 블록 선택할지 여부

    Returns:
        성공시 True

    Examples:
        >>> hwp.move_to_field("name")  # name 필드로 이동
        >>> hwp.move_to_field("title", idx=2)  # 세 번째 title로 이동
        >>> hwp.move_to_field("data", select=True)  # 선택 상태로 이동
    """
```

### 메타태그 관련 메서드 (한글2024+)
메타태그는 한글2024부터 지원되는 기능이다.

```python
def get_metatag_list(self, number: int, option: int) -> str:
    """메타태그 목록 조회 (get_field_list와 유사)"""
    return self.hwp.GetMetatagList(Number=number, option=option)

def get_metatag_name_text(self, tag: str) -> str:
    """메타태그 텍스트 조회"""
    return self.hwp.GetMetatagNameText(tag=tag)

def put_metatag_name_text(self, tag: str, text: str) -> bool:
    """메타태그에 텍스트 삽입"""
    return self.hwp.PutMetatagNameText(tag=tag, Text=text)

def rename_metatag(self, oldtag: str, newtag: str) -> bool:
    """메타태그 이름 변경"""
    return self.hwp.RenameMetatag(oldtag=oldtag, newtag=newtag)
```

---

## 필드 관련 상수

```cpp
// 필드 번호 형식 (get_field_list의 number 파라미터)
enum FieldNumberFormat {
    hwpFieldPlain = 0,   // 단순 나열
    hwpFieldNumber = 1,  // 일련번호 포함
    hwpFieldCount = 2    // 개수 포함
};

// 필드 옵션 (get_field_list의 option 파라미터)
enum FieldOption {
    hwpFieldAll = 0x00,        // 모든 필드
    hwpFieldCell = 0x01,       // 셀 필드만
    hwpFieldClickHere = 0x02,  // 누름틀만
    hwpFieldSelection = 0x04   // 선택 영역 내 필드만
};
```

---

## C++ 클래스 구조

```cpp
// Hwp.h - 필드/메타태그 메서드
class Hwp : public ParamHelpers, public RunMethods {
public:
    // 필드 생성/삭제
    bool CreateField(const std::wstring& name,
                     const std::wstring& direction = L"",
                     const std::wstring& memo = L"");
    bool DeleteAllFields();
    bool DeleteFieldByName(const std::wstring& fieldName, int idx = -1);

    // 필드 목록/텍스트
    std::wstring GetFieldList(int number = 1, int option = 0);
    std::wstring GetFieldText(const std::wstring& field, int idx = 0);
    void PutFieldText(const std::wstring& field, const std::wstring& text);
    void PutFieldText(const std::map<std::wstring, std::wstring>& fieldTextMap);

    // 필드 관리
    bool RenameField(const std::wstring& oldname, const std::wstring& newname);
    bool SetCurFieldName(const std::wstring& field,
                         const std::wstring& direction = L"",
                         const std::wstring& memo = L"",
                         int option = 0);
    int SetFieldViewOption(int option);

    // 필드 이동
    bool MoveToField(const std::wstring& field, int idx = 0,
                     bool text = true, bool start = true, bool select = false);

    // 필드 정보
    std::vector<std::map<std::wstring, std::wstring>> GetFieldInfo();
    std::map<std::wstring, std::vector<std::wstring>> FieldsToDict();

    // 메타태그 (한글2024+)
    std::wstring GetMetatagList(int number, int option);
    std::wstring GetMetatagNameText(const std::wstring& tag);
    bool PutMetatagNameText(const std::wstring& tag, const std::wstring& text);
    bool RenameMetatag(const std::wstring& oldtag, const std::wstring& newtag);
};
```

---

## 사용 패턴

### 패턴 1: 템플릿 문서 채우기
```python
# 템플릿 문서 열기
hwp.open("template.hwp")

# 필드에 데이터 삽입
hwp.put_field_text({
    "name": "홍길동",
    "date": "2024-01-01",
    "content": "안녕하세요."
})

# 저장
hwp.save_as("output.hwp")
```

### 패턴 2: 반복 필드에 데이터 삽입
```python
# 동일 이름의 여러 필드에 각각 다른 값 삽입
names = ["홍길동", "김철수", "이영희"]
hwp.put_field_text("name", names)

# 또는 인덱스 지정
for i, name in enumerate(names):
    hwp.put_field_text(f"name{{{{{i}}}}}", name)
```

### 패턴 3: 표 셀에 필드 설정
```python
# 표 생성 후 전체 셀에 필드 설정
hwp.create_table(3, 3)
hwp.TableCellBlockExtendAbs()
hwp.TableCellBlockExtend()
hwp.set_cur_field_name("data", option=1)

# 각 셀에 데이터 삽입
hwp.put_field_text("data", list(range(1, 10)))
```

---

## 구현 우선순위

1. **Critical**: CreateField(), GetFieldList(), GetFieldText(), PutFieldText()
2. **High**: MoveToField(), RenameField()
3. **High**: SetCurFieldName() - 표 셀 필드 설정
4. **Medium**: DeleteAllFields(), DeleteFieldByName()
5. **Low**: 메타태그 관련 (한글2024+ 전용)
6. **Low**: FieldsToDict(), SetFieldByBracket() - 고급 기능
