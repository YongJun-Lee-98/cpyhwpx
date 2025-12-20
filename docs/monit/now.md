# cpyhwpx 구현 현황 분석

> **분석 일자**: 2025-12-20 (이미지 삽입 기능 추가)

## 요약

| 구분 | 문서화 | 구현됨 | 구현률 |
|-----|--------|--------|--------|
| **총 API** | ~1,300+ | ~224 | **~17%** |
| Core 클래스 | 4 | 2 (Hwp, Ctrl) | 50% |
| 속성 | 35 | 6 | 17% |
| 파일 I/O | 26 | 10 | 38% |
| 텍스트 편집 | 35 | 10 | 29% |
| **테이블 작업** | 65+ | **21** | **32%** ✅ |
| **필드/메타태그** | 27 | **13** | **48%** ✅ |
| **이미지/도형** | 60+ | **7** | **12%** ✅ |
| 스타일/포맷팅 | 70+ | 10 | 14% |
| Run 액션 | 684 | 38 | 6% |
| 파라미터 헬퍼 | 90 | 7 | 8% |
| 폰트 프리셋 | 111 | 3 | 3% |
| 유틸리티 | 60+ | 15 | 25% |

---

## 상세 비교

### 1. Core 클래스

#### ✅ 구현됨
| 클래스 | 상태 |
|--------|------|
| `Hwp` (HwpWrapper) | ✅ 기본 구현 완료 |
| `Ctrl` (HwpCtrl) | ✅ 기본 구현 완료 |

#### ❌ 미구현
| 클래스 | 상태 |
|--------|------|
| `XHwpDocuments` | ❌ 미구현 |
| `XHwpDocument` | ❌ 미구현 |

---

### 2. Hwp 클래스 속성 (35개 문서화)

#### ✅ 구현됨 (6개)
- `version` - HWP 버전
- `build_number` - 빌드 번호
- `current_page` - 현재 페이지
- `page_count` - 총 페이지 수
- `edit_mode` - 편집 모드 (읽기/쓰기)
- `is_initialized` - 초기화 여부

#### ❌ 미구현 (29개)
- `Application` - Low-level API 접근
- `CLSID` - 클래스 ID
- `CurFieldState` - 현재 필드 상태
- `CurMetatagState` - 현재 메타태그 상태
- `CurSelectedCtrl` - 현재 선택된 컨트롤
- `EngineProperties` - 엔진 속성
- `HeadCtrl` - 첫 번째 컨트롤
- `LastCtrl` - 마지막 컨트롤
- `ParentCtrl` - 부모 컨트롤
- `IsEmpty` - 빈 문서 여부 (메서드로만 존재)
- `IsModified` - 수정 여부 (메서드로만 존재)
- `IsPrivateInfoProtected` - 개인정보 보호
- `IsTrackChange` - 변경 추적
- `Path` - 문서 경로
- `SelectionMode` - 선택 모드
- `Title` - 창 제목
- `XHwpDocuments` - 문서 컬렉션
- `XHwpMessageBox` - 메시지 박스 객체
- `XHwpODBC` - ODBC 객체
- `XHwpWindows` - 창 관리 객체
- `ctrl_list` - 모든 컨트롤 목록
- `current_printpage` - 현재 인쇄 페이지
- `current_font` - 현재 폰트
- `CellShape` - 셀 모양 파라미터셋
- `CharShape` - 문자 모양 파라미터셋
- `ParaShape` - 문단 모양 파라미터셋
- `ViewProperties` - 뷰 속성
- `HAction` - 액션 인터페이스
- `HParameterSet` - 파라미터셋 인터페이스

---

### 3. 파일 I/O (26개 문서화)

#### ✅ 구현됨 (10개)
- `open(path, format, arg)` ✅
- `save(save_if_dirty)` ✅
- `save_as(path, format, arg)` ✅
- `clear(option)` ✅
- `close(is_dirty)` ✅
- `quit(save)` ✅
- `register_module()` ✅
- `is_empty()` ✅
- `is_modified()` ✅
- `initialize()` ✅

#### ❌ 미구현 (16개)
- `insert_file(filename, format, option)`
- `InsertFile(filename, Format, arg)`
- `save_block_as(path, format, attributes)`
- `get_text_file(format, option)`
- `set_text_file(data, format, option)`
- `open_pdf(pdf_path, this_window)`
- `get_file_info(filename)`
- `FileClose()` (액션으로는 존재)

---

### 4. 텍스트 편집 (35개 문서화)

#### ✅ 구현됨 (8개)
- `insert_text(text)` ✅
- `get_text()` ✅
- `get_selected_text(keep_select)` ✅
- `get_pos()` ✅
- `set_pos(list, para, pos)` ✅
- `move_pos(move_id, para, pos)` ✅
- `find(text, ...)` ✅
- `replace(find_text, replace_text, ...)` ✅
- `replace_all(find_text, replace_text, ...)` ✅

#### ❌ 미구현 (26개)
- `get_pos_by_set()` / `GetPosBySet()`
- `set_pos_by_set(disp_val)` / `SetPosBySet()`
- `move_to_field(field, idx, text, start, select)`
- `select_text(spara, spos, epara, epos, slist)`
- `select_text_by_get_pos(s_getpos, e_getpos)`
- `init_scan(option, range, ...)`
- `release_scan()`
- `find_forward(src, regex)`
- `find_backward(src, regex)`
- `find_replace(src, dst, ...)`
- `paste(option)`

---

### 5. 테이블 작업 (65+개 문서화)

#### ✅ 구현됨 (21개)

**테이블 생성/탐색:**
- `create_table(rows, cols, treat_as_char, width_type, height_type, header)` ✅ NEW
- `get_into_nth_table(n, select_cell)` ✅ NEW
- `find_ctrl()` ✅ NEW

**셀 이동:**
- `table_left_cell()` ✅ NEW
- `table_right_cell()` ✅ NEW
- `table_upper_cell()` ✅ NEW
- `table_lower_cell()` ✅ NEW
- `table_right_cell_append()` ✅ NEW

**테이블 정보:**
- `get_table_row_count()` ✅ NEW
- `get_table_col_count()` ✅ NEW
- `is_cell()` ✅

**HwpActionHelper (기존):**
- `TableCellBlock()` ✅
- `TableColBegin()` / `TableColEnd()` ✅
- `TableRowBegin()` / `TableRowEnd()` ✅
- `TableAppendRow()` / `TableAppendColumn()` ✅
- `TableDeleteRow()` / `TableDeleteColumn()` ✅
- `TableMergeCell()` / `TableSplitCell()` ✅

#### ❌ 미구현 (44+개)
- `table_from_data(data, transpose, ...)`
- `get_row_height()` / `get_col_width()`
- `set_row_height()` / `set_col_width()`
- `cell_fill(face_color)`
- `table_to_string()` / `table_to_csv()` / `table_to_df()`
- 테이블 셀 정렬 (9개)
- 테이블 테두리 (12개)
- 테이블 리사이즈 (10개)
- 테이블 자동 함수 (10개)

---

### 6. 필드/메타태그 (27개 문서화)

#### ✅ 구현됨 (13개)
- `create_field(name, direction, memo)` ✅
- `get_field_list(number, option)` ✅
- `get_field_text(field, idx)` ✅
- `put_field_text(field, text)` ✅
- `field_exist(field)` ✅
- `move_to_field(field, idx, text, start, select)` ✅
- `rename_field(oldname, newname)` ✅
- `get_cur_field_name(option)` ✅
- `set_cur_field_name(field, direction, memo, option)` ✅
- `set_field_view_option(option)` ✅
- `delete_all_fields()` ✅
- `delete_field_by_name(field_name, idx)` ✅
- `fields_to_map()` ✅

#### ❌ 미구현 (14개)
- `get_field_info()` - HWPML2X 파싱 필요
- `set_field_by_bracket()` - 중괄호를 필드로 변환
- `get_metatag_list()` (HWP2024+)
- `get_metatag_name_text(tag)`
- `put_metatag_name_text(tag, text)`
- `rename_metatag(oldtag, newtag)`
- 기타 메타태그 관련 메서드들

---

### 7. 이미지/도형 객체 (60+개 문서화)

#### ✅ 구현됨 (7개)

**이미지 삽입:**
- `insert_picture(path, embedded, sizeoption, reverse, watermark, effect, width, height)` ✅ NEW

**HwpActionHelper:**
- `ShapeObjSelect()` ✅
- `ShapeObjDelete()` ✅
- `ShapeObjCopy()` ✅
- `ShapeObjCut()` ✅
- `ShapeObjBringToFront()` ✅
- `ShapeObjSendToBack()` ✅

#### ❌ 미구현 (53+개)
- `insert_ctrl(ctrl_id, initparam)`
- `delete_ctrl(ctrl)`
- `create_page_image(path, pgno, ...)`
- `EquationCreate()` / `EquationClose()` / `EquationModify()`
- 도형 정렬 (11개)
- 도형 순서 (4개)
- 도형 그룹화 (2개)
- 도형 변환 (3개)
- 도형 이동/리사이즈 (8개)
- 캡션/텍스트박스 (5개)

---

### 8. 스타일/포맷팅 (70+개 문서화)

#### ✅ 구현됨 (10개, HwpActionHelper)
- `CharShapeBold()` ✅
- `CharShapeItalic()` ✅
- `CharShapeUnderline()` ✅
- `CharShapeStrikeout()` ✅
- `CharShapeSuperscript()` ✅
- `CharShapeSubscript()` ✅
- `ParagraphShapeAlignLeft()` ✅
- `ParagraphShapeAlignCenter()` ✅
- `ParagraphShapeAlignRight()` ✅
- `ParagraphShapeAlignJustify()` ✅

#### ❌ 미구현 (60+개)
- `get_charshape()` / `get_charshape_as_dict()`
- `set_charshape(pset)`
- `get_parashape()` / `get_parashape_as_dict()`
- `set_parashape(pset)`
- `import_style(sty_filepath)`
- 문자 효과 (5개)
- 폰트 크기 (3개)
- 자간/장평 (6개)
- 텍스트 색상 (8개)
- 문단 여백 (6개)
- 줄간격 (2개)

---

### 9. Run 액션 (684개 문서화)

#### ✅ 구현됨 (38개, HwpActionHelper)
```
BreakPara, BreakPage, BreakSection, BreakColumn,
SelectAll, Cancel,
MoveLeft, MoveRight, MoveUp, MoveDown,
MoveLineBegin, MoveLineEnd, MoveDocBegin, MoveDocEnd,
MoveParaBegin, MoveParaEnd, MoveWordBegin, MoveWordEnd,
MovePageUp, MovePageDown,
Delete, DeleteBack, Cut, Copy, Paste, Undo, Redo,
TableCellBlock, TableColBegin, TableColEnd, ...
WindowMaximize, WindowMinimize, ViewZoomIn, ViewZoomOut,
FilePrint, FilePrintPreview, FileClose, FileQuit
```

#### ❌ 미구현 (646개)
- 대부분의 Run 액션들

---

### 10. 파라미터 헬퍼 (90개 문서화)

#### ✅ 구현됨 (7개)
- `CreateFindReplace()` ✅
- `CreateTable()` ✅
- `CreateCharShape()` ✅
- `CreateParaShape()` ✅
- `CreateInsertPicture()` ✅
- `CreateCellShape()` ✅
- `CreateBorderLine()` ✅

#### ❌ 미구현 (83개)
- `mili_to_hwp_unit()` / `hwp_unit_to_mili()` (Units 모듈에 일부 구현)
- `HwpLineType()` / `HwpLineWidth()`
- `BorderShape()`
- `HAlign()` / `VAlign()` (Enum으로만 존재)
- `HeadType()` / `NumberFormat()`
- `LunarToSolar()` / `SolarToLunar()`
- 기타 80+ 변환 헬퍼들

---

### 11. 폰트 프리셋 (111개 문서화)

#### ✅ 구현됨 (3개)
- `malgun_gothic()` ✅
- `nanum_gothic()` ✅
- `nanum_myeongjo()` ✅

#### ❌ 미구현 (108개)
- 나머지 108개 폰트 프리셋

---

### 12. 유틸리티 (60+개 문서화)

#### ✅ 구현됨 (15개)
**Utils 서브모듈:**
- `addr_to_tuple()`, `tuple_to_addr()` ✅
- `parse_range()`, `expand_range()` ✅
- `trim()`, `split()`, `join()` ✅
- `file_exists()`, `get_extension()` ✅
- `hex_to_colorref()`, `colorref_to_hex()` ✅

**Units 서브모듈:**
- `from_mm()`, `from_cm()`, `from_inch()`, `from_point()` ✅
- `to_mm()`, `to_cm()`, `to_inch()`, `to_point()` ✅

#### ❌ 미구현 (45+개)
- `doc_list()` / `switch_to(num)` / `add_tab()` / `add_doc()`
- `set_viewstate()` / `get_viewstate()` (일부 구현)
- `msgbox()` (일부 구현)
- 기타 문서/탭 관리 유틸리티

---

## 우선순위별 미구현 기능

### ✅ 완료
1. ~~**필드 작업** - 문서 자동화의 핵심~~ ✅ (13개 구현, 48%)
2. ~~**테이블 생성/탐색** - 보고서 생성 필수~~ ✅ (10개 NEW 구현, 32%)
   - `create_table()`, `get_into_nth_table()`, 셀 이동 5개
3. ~~**이미지 삽입** - 문서 작성 필수~~ ✅ NEW
   - `insert_picture()` - 8개 파라미터 지원 (embedded, sizeoption, reverse, watermark, effect, width, height)

### 🔴 높음 (즉시 필요)
1. **파일 작업 확장** - 문서 처리
   - `insert_file()`, `get_text_file()`, `set_text_file()`
2. **테이블 데이터** - 대량 데이터 입력
   - `table_from_data()`, `table_to_df()`
3. **추가 컨트롤 작업** - 도형/객체 관리
   - `insert_ctrl()`, `delete_ctrl()`

### 🟡 중간 (기능 확장)
1. **XHwpDocuments/XHwpDocument** - 다중 문서 관리
2. **스타일 관리** - `get_charshape()`, `set_charshape()`
3. **추가 속성** - `CurSelectedCtrl`, `HeadCtrl`, `ctrl_list`

### 🟢 낮음 (추후 구현)
1. 나머지 Run 액션들 (646개)
2. 폰트 프리셋 (108개)
3. 파라미터 헬퍼 (83개)
4. 메타태그 (HWP2024+)

---

## 소스 파일 위치

| 파일 | 역할 | 라인 수 |
|------|------|--------|
| `src/HwpWrapper.h` | 메인 래퍼 헤더 | ~443 |
| `src/HwpWrapper.cpp` | 메인 래퍼 구현 | ~1064 |
| `src/HwpCtrl.h/cpp` | 컨트롤 래퍼 | - |
| `src/HwpAction.h/cpp` | 액션 헬퍼 | - |
| `src/HwpParameter.h/cpp` | 파라미터셋 래퍼 | - |
| `src/bindings.cpp` | Python 바인딩 | ~402 |
| `src/FontDefs.h/cpp` | 폰트 프리셋 | - |
| `src/Utils.h/cpp` | 유틸리티 | - |

## 문서 파일 위치

| 파일 | 내용 |
|------|------|
| `docs/mapping/01_core_classes.md` | 코어 클래스 |
| `docs/mapping/02_properties.md` | 속성 35개 |
| `docs/mapping/03_file_io.md` | 파일 I/O 26개 |
| `docs/mapping/04_text_editing.md` | 텍스트 편집 35개 |
| `docs/mapping/05_table_operations.md` | 테이블 65+개 |
| `docs/mapping/06_field_metatag.md` | 필드/메타태그 27개 |
| `docs/mapping/07_shape_objects.md` | 도형 60+개 |
| `docs/mapping/08_style_formatting.md` | 스타일 70+개 |
| `docs/mapping/09_run_actions.md` | Run 액션 684개 |
| `docs/mapping/10_param_helpers.md` | 파라미터 헬퍼 90개 |
| `docs/mapping/11_fonts.md` | 폰트 111개 |
| `docs/mapping/12_utility.md` | 유틸리티 60+개 |
