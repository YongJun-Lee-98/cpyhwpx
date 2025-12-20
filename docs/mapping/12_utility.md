# 유틸리티 함수 매핑

## 개요

| 항목 | 내용 |
|------|------|
| 소스 파일 | `pyhwpx/core.py` |
| 총 함수 수 | **~60개** |
| 포팅 완료 | 0개 (0%) |
| 우선순위 | Medium |

## 설명

다른 카테고리에 포함되지 않는 유틸리티 함수들을 모아놓은 문서입니다.
윈도우 관리, 문서 탭 관리, 위치 조회, 메시지 처리 등의 기능을 포함합니다.

---

## 1. 윈도우/UI 관리

### 1.1 창 제어

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `maximize_window()` | `()` | `None` | 창 최대화 |
| 2 | `minimize_window()` | `()` | `None` | 창 최소화 |
| 3 | `set_visible()` | `(visible: bool)` | `None` | 창 표시/숨김 |
| 4 | `set_viewstate()` | `(flag: int)` | `bool` | 뷰 상태 설정 |
| 5 | `get_viewstate()` | `()` | `int` | 뷰 상태 조회 |

#### set_viewstate() 플래그 값

| 값 | 상수 | 설명 |
|----|------|------|
| 0 | VS_NORMAL | 일반 보기 |
| 1 | VS_MULTICOLUMN | 다단 보기 |
| 2 | VS_MULTIPAGE | 다중 페이지 |
| 3 | VS_MASTERPAGE | 바탕쪽 보기 |
| 4 | VS_PAGE | 쪽 보기 |
| 5 | VS_TEXT | 텍스트 보기 |
| 6 | VS_OUTLINE | 개요 보기 |

### 1.2 메시지 박스

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `msgbox()` | `(string, flag=0)` | `int` | 메시지 박스 표시 |
| 2 | `get_message_box_mode()` | `()` | `int` | 메시지 박스 모드 조회 |
| 3 | `set_message_box_mode()` | `(mode: int)` | `int` | 메시지 박스 모드 설정 |

#### set_message_box_mode() 값

```python
# 0xF0000 : 모든 다이얼로그 표시 (기본값)
# 0x10000 : 확인 버튼 자동 클릭
# 0x20000 : 취소 버튼 자동 클릭
# 0x1000  : 예 버튼 자동 클릭
# 0x2000  : 아니오 버튼 자동 클릭
```

---

## 2. 문서 탭 관리

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `doc_list()` | `()` | `List[XHwpDocument]` | 열린 문서 목록 |
| 2 | `switch_to()` | `(num: int)` | `Optional[XHwpDocument]` | 특정 탭으로 전환 |
| 3 | `add_tab()` | `()` | `XHwpDocument` | 새 탭 추가 |
| 4 | `add_doc()` | `()` | `XHwpDocument` | 새 문서 추가 |

### 사용 예시 (Python)
```python
# 열린 문서 목록 조회
docs = hwp.doc_list()
print(f"열린 문서 수: {len(docs)}")

# 두 번째 문서로 전환
hwp.switch_to(1)

# 새 탭 추가
new_doc = hwp.add_tab()
```

---

## 3. 위치(Position) 관리

### 3.1 위치 조회/설정

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `get_pos()` | `()` | `Tuple[int, int, int]` | 현재 위치 (list, para, pos) |
| 2 | `set_pos()` | `(List, para, pos)` | `bool` | 위치 설정 |
| 3 | `get_pos_by_set()` | `()` | `HParameterSet` | 위치 정보 (파라미터셋) |
| 4 | `set_pos_by_set()` | `(disp_val)` | `bool` | 위치 설정 (파라미터셋) |
| 5 | `move_pos()` | `(move_id, para, pos)` | `bool` | 위치 이동 |

### 3.2 선택 영역 위치

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `get_selected_pos()` | `()` | `Tuple[7개]` | 선택 영역 위치 |
| 2 | `get_selected_pos_by_set()` | `(sset, eset)` | `bool` | 선택 영역 (파라미터셋) |
| 3 | `get_selected_range()` | `()` | `List[str]` | 선택 영역 범위 |
| 4 | `select_text_by_get_pos()` | `(s_getpos, e_getpos)` | `bool` | 위치로 텍스트 선택 |

### move_pos() move_id 값

| 값 | 상수 | 설명 |
|----|------|------|
| 0 | MOVEID_NONE | 이동 없음 |
| 1 | MOVEID_TOP | 문서 처음 |
| 2 | MOVEID_BOTTOM | 문서 끝 |
| 3 | MOVEID_PARA_BEGIN | 문단 처음 |
| 4 | MOVEID_PARA_END | 문단 끝 |
| 5 | MOVEID_LINE_BEGIN | 줄 처음 |
| 6 | MOVEID_LINE_END | 줄 끝 |
| 7 | MOVEID_WORD_BEGIN | 단어 처음 |
| 8 | MOVEID_WORD_END | 단어 끝 |

---

## 4. 페이지 관리

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `current_page` | `(property)` | `int` | 현재 페이지 번호 (0부터) |
| 2 | `current_printpage` | `(property)` | `int` | 현재 인쇄 페이지 (1부터) |
| 3 | `goto_page()` | `(page_index)` | `Tuple[int, int]` | 페이지로 이동 |
| 4 | `goto_printpage()` | `(page_num)` | `bool` | 인쇄 페이지로 이동 |
| 5 | `get_page_text()` | `(pgno, option)` | `str` | 페이지 텍스트 조회 |
| 6 | `is_empty_page()` | `(pgno, ...)` | `bool` | 빈 페이지 여부 |

---

## 5. 문서 상태 확인

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `is_empty()` | `()` | `bool` | 빈 문서 여부 |
| 2 | `is_modified()` | `()` | `bool` | 수정 여부 |
| 3 | `is_cell()` | `()` | `bool` | 셀 안 여부 |
| 4 | `is_empty_para()` | `()` | `bool` | 빈 문단 여부 |
| 5 | `is_action_enable()` | `(action_id)` | `bool` | 액션 활성화 여부 |
| 6 | `is_command_lock()` | `(action_id)` | `bool` | 명령 잠금 여부 |

---

## 6. 단위 변환

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `point_to_hwp_unit()` | `(point)` | `int` | 포인트 → HWP 단위 |
| 2 | `mm_to_hwp_unit()` | `(mm)` | `int` | mm → HWP 단위 (암시적) |

### 단위 변환 공식
```python
# HwpUnit = 1/7200 inch
# 1 inch = 25.4 mm = 72 point

# point → HwpUnit
hwp_unit = int(point * 100)

# mm → HwpUnit
hwp_unit = int(mm * 7200 / 25.4)

# inch → HwpUnit
hwp_unit = int(inch * 7200)
```

---

## 7. 컨트롤/액션 관리

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `ctrl_list` | `(property)` | `list` | 문서 내 모든 컨트롤 |
| 2 | `get_ctrl_by_ctrl_id()` | `(ctrl_id)` | `Ctrl` | ID로 컨트롤 조회 |
| 3 | `move_to_ctrl()` | `(ctrl, option)` | `bool` | 컨트롤로 이동 |
| 4 | `delete_ctrl()` | `(ctrl)` | `bool` | 컨트롤 삭제 |
| 5 | `insert_ctrl()` | `(ctrl_id, initparam)` | `Ctrl` | 컨트롤 삽입 |
| 6 | `find_ctrl()` | `()` | `Any` | 컨트롤 찾기 |

### 컨트롤 타입 (ctrl_id)

| 값 | 설명 |
|----|------|
| `tbl` | 표 |
| `eqed` | 수식 |
| `gso` | 그리기 개체 |
| `pic` | 그림 |
| `ole` | OLE 개체 |
| `btn` | 버튼 |
| `cold` | 단정의 |
| `head` | 머리말 |
| `foot` | 꼬리말 |
| `fn` | 각주 |
| `en` | 미주 |

---

## 8. 액션/명령 관리

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `create_action()` | `(actidstr)` | `Any` | 액션 생성 |
| 2 | `create_set()` | `(setidstr)` | `HParameterSet` | 파라미터셋 생성 |
| 3 | `release_action()` | `(action)` | `None` | 액션 해제 |
| 4 | `replace_action()` | `(old_id, new_id)` | `bool` | 액션 교체 |
| 5 | `lock_command()` | `(act_id, is_lock)` | `None` | 명령 잠금 |

---

## 9. 클립보드

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `paste()` | `(option=4)` | `bool` | 붙여넣기 |
| 2 | `clipboard_to_pyfunc()` | `()` | `str` | 클립보드를 파이썬 함수로 |

### paste() option 값

| 값 | 설명 |
|----|------|
| 0 | 기본 붙여넣기 |
| 1 | 특수 붙여넣기 |
| 2 | 서식 없이 |
| 3 | HTML |
| 4 | 텍스트만 |
| 5 | 그림으로 |
| 6 | OLE |

---

## 10. PDF 관련

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `open_pdf()` | `(pdf_path, this_window=1)` | `bool` | PDF 열기 |
| 2 | `save_pdf_as_image()` | `(path, img_format)` | `bool` | PDF를 이미지로 저장 |

---

## 11. 스크립트 관련

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `get_script_source()` | `(filename)` | `str` | 스크립트 소스 조회 |

---

## 12. 개인정보 처리

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `set_private_info_password()` | `(password)` | `bool` | 개인정보 비밀번호 설정 |
| 2 | `find_private_info()` | `(private_type, private_string)` | `int` | 개인정보 찾기 |

---

## 13. 기타 유틸리티

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `count()` | `(word)` | `int` | 단어 개수 |
| 2 | `clear()` | `(option=1)` | `None` | 문서 초기화 |
| 3 | `close()` | `(is_dirty, interval)` | `bool` | 문서 닫기 |
| 4 | `quit()` | `(save=False)` | `None` | 프로그램 종료 |
| 5 | `release_scan()` | `()` | `None` | 스캔 해제 |
| 6 | `scan_font()` | `()` | `None` | 폰트 스캔 |
| 7 | `key_indicator()` | `()` | `tuple` | 키 인디케이터 |
| 8 | `check_xobject()` | `(bstring)` | `bool` | X객체 확인 |
| 9 | `get_title()` | `()` | `str` | 제목 조회 |
| 10 | `set_title()` | `(title)` | `bool` | 제목 설정 |
| 11 | `get_heading_string()` | `()` | `str` | 헤딩 문자열 조회 |
| 12 | `init_hparameterset()` | `()` | `None` | HParameterSet 초기화 |
| 13 | `create_id()` | `(creation_id)` | `Any` | ID 생성 |
| 14 | `create_mode()` | `(creation_mode)` | `Any` | 모드 생성 |
| 15 | `insert_lorem()` | `(para_num)` | `bool` | Lorem ipsum 삽입 |
| 16 | `insert_random_picture()` | `(x, y)` | `Ctrl` | 랜덤 그림 삽입 |

---

## 14. 하이퍼링크

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `insert_hyperlink()` | `(hypertext, description)` | `bool` | 하이퍼링크 삽입 |

---

## 15. 수식 (MathML)

| # | Python 메서드 | 시그니처 | 반환값 | 설명 |
|---|--------------|---------|--------|------|
| 1 | `export_mathml()` | `(mml_path, delay)` | `None` | MathML 내보내기 |
| 2 | `import_mathml()` | `(mml_path, delay)` | `None` | MathML 가져오기 |

---

## C++ 구현 예시

### Utils.h
```cpp
#pragma once
#include "HwpWrapper.h"

class HwpUtils {
public:
    HwpUtils(HwpWrapper* hwp) : m_pHwp(hwp) {}

    // 윈도우 관리
    void MaximizeWindow();
    void MinimizeWindow();
    void SetVisible(bool visible);
    bool SetViewState(int flag);
    int GetViewState();

    // 메시지 박스
    int MsgBox(const std::wstring& message, int flag = 0);
    int GetMessageBoxMode();
    int SetMessageBoxMode(int mode);

    // 위치 관리
    std::tuple<int, int, int> GetPos();
    bool SetPos(int listId, int para, int pos);
    bool MovePos(int moveId, int para = 0, int pos = 0);

    // 문서 상태
    bool IsEmpty();
    bool IsModified();
    bool IsCell();

    // 단위 변환
    static int PointToHwpUnit(double point);
    static int MmToHwpUnit(double mm);
    static int InchToHwpUnit(double inch);

private:
    HwpWrapper* m_pHwp;
};
```

### Utils.cpp (부분)
```cpp
#include "Utils.h"

void HwpUtils::MaximizeWindow() {
    // ShowWindow(hwnd, SW_MAXIMIZE) 또는
    // COM 인터페이스 통해 구현
}

int HwpUtils::PointToHwpUnit(double point) {
    return static_cast<int>(point * 100);
}

int HwpUtils::MmToHwpUnit(double mm) {
    return static_cast<int>(mm * 7200 / 25.4);
}

int HwpUtils::InchToHwpUnit(double inch) {
    return static_cast<int>(inch * 7200);
}

std::tuple<int, int, int> HwpUtils::GetPos() {
    // hwp.GetPos() 호출
    int list, para, pos;
    // ... COM 호출
    return std::make_tuple(list, para, pos);
}
```

### pybind11 바인딩
```cpp
PYBIND11_MODULE(cpyhwpx, m) {
    py::class_<HwpWrapper>(m, "Hwp")
        // 윈도우 관리
        .def("maximize_window", &HwpWrapper::MaximizeWindow)
        .def("minimize_window", &HwpWrapper::MinimizeWindow)
        .def("set_visible", &HwpWrapper::SetVisible)
        .def("set_viewstate", &HwpWrapper::SetViewState)
        .def("get_viewstate", &HwpWrapper::GetViewState)

        // 메시지 박스
        .def("msgbox", &HwpWrapper::MsgBox,
             py::arg("string"), py::arg("flag") = 0)
        .def("get_message_box_mode", &HwpWrapper::GetMessageBoxMode)
        .def("set_message_box_mode", &HwpWrapper::SetMessageBoxMode)

        // 위치 관리
        .def("get_pos", &HwpWrapper::GetPos)
        .def("set_pos", &HwpWrapper::SetPos)
        .def("move_pos", &HwpWrapper::MovePos,
             py::arg("move_id") = 1,
             py::arg("para") = 0,
             py::arg("pos") = 0)

        // 문서 상태
        .def("is_empty", &HwpWrapper::IsEmpty)
        .def("is_modified", &HwpWrapper::IsModified)
        .def("is_cell", &HwpWrapper::IsCell)

        // 기타
        .def("count", &HwpWrapper::Count)
        .def("clear", &HwpWrapper::Clear, py::arg("option") = 1)
        .def("quit", &HwpWrapper::Quit, py::arg("save") = false)
        ;

    // 정적 유틸리티 함수
    m.def("point_to_hwp_unit", &HwpUtils::PointToHwpUnit);
    m.def("mm_to_hwp_unit", &HwpUtils::MmToHwpUnit);
    m.def("inch_to_hwp_unit", &HwpUtils::InchToHwpUnit);
}
```

---

## 포팅 우선순위

### High (자주 사용)
- `get_pos()`, `set_pos()`, `move_pos()`
- `is_empty()`, `is_modified()`
- `set_visible()`
- `msgbox()`, `set_message_box_mode()`
- `quit()`, `close()`, `clear()`

### Medium
- `goto_page()`, `goto_printpage()`
- `doc_list()`, `switch_to()`, `add_tab()`
- `ctrl_list`, `get_ctrl_by_ctrl_id()`
- `paste()`
- 단위 변환 함수들

### Low
- `clipboard_to_pyfunc()`
- PDF 관련 함수
- MathML 관련 함수
- 개인정보 관련 함수

---

## 진행 상황

| 카테고리 | 함수 수 | 완료 | 진행률 |
|---------|--------|------|--------|
| 윈도우/UI | 5 | 0 | 0% |
| 문서 탭 | 4 | 0 | 0% |
| 위치 관리 | 8 | 0 | 0% |
| 페이지 관리 | 6 | 0 | 0% |
| 상태 확인 | 6 | 0 | 0% |
| 단위 변환 | 2 | 0 | 0% |
| 컨트롤 관리 | 6 | 0 | 0% |
| 액션 관리 | 5 | 0 | 0% |
| 클립보드 | 2 | 0 | 0% |
| PDF | 2 | 0 | 0% |
| 기타 | ~16 | 0 | 0% |
| **총계** | **~60** | **0** | **0%** |
