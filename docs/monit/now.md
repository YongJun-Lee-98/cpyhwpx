# cpyhwpx 구현 현황 분석

> **분석 일자**: 2025-12-21 (Run 액션 102개 추가: Auto/ViewOption/Macro/Quick/MasterPage/Picture/Note/FormObj/Window)

## 요약

| 구분 | 문서화 | 구현됨 | 구현률 |
|-----|--------|--------|--------|
| **총 API** | ~1,300+ | **~1,071** | **~82%** ✅ |
| Core 클래스 | 4 | 4 | 100% ✅ |
| 속성 | 35 | **22** | **63%** ✅ |
| **파일 I/O** | 26 | **19** | **73%** ✅ |
| **보안 모듈** | 4 | **4** | **100%** ✅ |
| **텍스트 편집** | 35 | **30** | **86%** ✅ |
| **테이블 작업** | 65+ | **71** | **100%+** ✅ (actions 포함) |
| **필드/메타태그** | 27 | **23** | **85%** ✅ |
| **이미지/도형** | 60+ | **51** | **85%** ✅ (actions 포함) |
| **스타일/포맷팅** | 70+ | **45** | **64%** ✅ (actions 포함) |
| **Run 액션** | 684 | **558** | **82%** ✅ |
| **파라미터 헬퍼** | 110 | **108** | **98%** ✅ |
| **폰트 프리셋** | 111 | **111** | **100%** ✅ NEW |
| **유틸리티** | 60+ | **28** | **47%** ✅ |

---

## 상세 비교

### 1. Core 클래스

#### ✅ 구현됨
| 클래스 | 상태 |
|--------|------|
| `Hwp` (HwpWrapper) | ✅ 기본 구현 완료 |
| `Ctrl` (HwpCtrl) | ✅ 기본 구현 완료 |
| `XHwpDocuments` | ✅ NEW - 문서 컬렉션 관리 |
| `XHwpDocument` | ✅ NEW - 개별 문서 관리 |

---

### 2. Hwp 클래스 속성 (35개 문서화)

#### ✅ 구현됨 (22개)
- `version` - HWP 버전
- `build_number` - 빌드 번호
- `current_page` - 현재 페이지
- `page_count` - 총 페이지 수
- `edit_mode` - 편집 모드 (읽기/쓰기)
- `is_initialized` - 초기화 여부
- `XHwpDocuments` - 문서 컬렉션
- `head_ctrl` - 첫 번째 컨트롤
- `last_ctrl` - 마지막 컨트롤
- `parent_ctrl` - 부모 컨트롤
- `cur_selected_ctrl` - 현재 선택된 컨트롤
- `ctrl_list` - 모든 컨트롤 목록
- `CLSID` - 클래스 ID ✅ NEW
- `CurFieldState` - 현재 필드 상태 (0=본문, 1=셀, 4=글상자) ✅ NEW
- `CurMetatagState` - 현재 메타태그 상태 ✅ NEW
- `IsPrivateInfoProtected` - 개인정보 보호 ✅ NEW
- `IsTrackChange` - 변경 추적 ✅ NEW
- `Path` - 문서 경로 ✅ NEW
- `SelectionMode` - 선택 모드 (0=일반, 1=블록) ✅ NEW
- `Title` - 창 제목 ✅ NEW
- `current_printpage` - 현재 인쇄 페이지 ✅ NEW
- `current_font` - 현재 폰트 ✅ NEW

#### ⚠️ 메서드로 구현됨 (4개)
- `IsEmpty` → `is_empty()` 메서드
- `IsModified` → `is_modified()` 메서드
- `CharShape` → `get_charshape()`/`set_charshape()` 메서드
- `ParaShape` → `get_parashape()`/`set_parashape()` 메서드

#### ❌ 미구현 (9개) - IDispatch* 반환으로 pybind11 미지원
- `Application` - Low-level API 접근
- `EngineProperties` - 엔진 속성
- `ViewProperties` - 뷰 속성
- `XHwpMessageBox` - 메시지 박스 객체
- `XHwpODBC` - ODBC 객체
- `XHwpWindows` - 창 관리 객체
- `HAction` - 액션 인터페이스 (내부적으로 사용됨)
- `HParameterSet` - 파라미터셋 인터페이스 (내부적으로 사용됨)
- `CellShape` - 셀 모양 파라미터셋

---

### 3. 파일 I/O (26개 문서화)

#### ✅ 구현됨 (19개)
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
- `insert_file(filename, keep_section, keep_charshape, keep_parashape, keep_style, move_doc_end)` ✅
- `get_text_file(format, option)` ✅ - 문서 텍스트 추출 (HWP/HWPML2X/HTML/UNICODE/TEXT)
- `set_text_file(data, format, option)` ✅ - 텍스트 데이터 삽입
- `save_block_as(path, format, attributes)` ✅ - 선택 블록 저장
- `get_file_info(filename)` ✅ - 파일 정보 조회
- `lock_command(act_id, is_lock)` ✅ NEW - 명령 잠금/해제 (Undo, Redo 등)
- `create_page_image(path, pgno, resolution, depth, format)` ✅ NEW - 페이지 이미지 생성 (BMP/GIF)
- `print_document()` ✅ NEW - 인쇄 다이얼로그 표시
- `mail_merge()` ✅ NEW - 메일 머지 다이얼로그 표시

#### ⚠️ 대체 가능 (4개)
- `get_html_text()` → `get_text_file(format="HTML")` 사용
- `set_html_text(html)` → `set_text_file(data, format="HTML")` 사용
- `get_hwpml_text()` → `get_text_file(format="HWPML2X")` 사용
- `convert_to_pdf(path)` → `save_as(path, format="PDF")` 사용

#### ❌ HWP API 예외 발생 (2개) - pyhwpx에서도 동일 오류
- `export_style(path)` - COM 예외 발생 (`pywintypes.com_error: -2147417851`)
- `import_style(path)` - export_style 의존

#### ⏸️ 테스트 필요 (1개)
- `open_pdf(pdf_path, this_window)` - PDF 열기 (구현됨, 테스트 필요)

---

### 3.1. 보안 모듈 자동 등록 (NEW)

> pyhwpx의 FilePathCheckerModule.dll 자동 등록 기능을 C++로 포팅

#### ✅ 구현됨 (4개, 100%)

**정적 메서드:**
- `Hwp.check_registry_key(key_name)` ✅ NEW - 레지스트리 등록 여부 확인
- `Hwp.find_dll_path()` ✅ NEW - DLL 파일 경로 자동 감지
- `Hwp.register_to_registry(dll_path, key_name)` ✅ NEW - DLL을 레지스트리에 등록

**인스턴스 메서드:**
- `hwp.auto_register_module(module_type, module_data)` ✅ NEW - 자동 등록 (check + regedit + COM API)

#### 레지스트리 경로
```
HKEY_CURRENT_USER\Software\HNC\HwpAutomation\Modules
  └─ FilePathCheckerModule = "...\FilePathCheckerModule.dll"

(대체 경로)
HKEY_CURRENT_USER\Software\Hnc\HwpUserAction\Modules
```

#### 사용 예시
```python
import cpyhwpx

# 정적 메서드로 상태 확인
print(cpyhwpx.Hwp.check_registry_key())  # True/False
print(cpyhwpx.Hwp.find_dll_path())       # DLL 경로

# HWP 초기화 후 자동 등록
hwp = cpyhwpx.Hwp(visible=True)
hwp.initialize()
hwp.auto_register_module()  # 레지스트리 확인/등록 + COM API 호출
```

---

### 4. 텍스트 편집 (35개 문서화)

#### ✅ 구현됨 (30개)
- `insert_text(text)` ✅
- `get_text()` ✅
- `get_selected_text(keep_select)` ✅
- `get_pos()` ✅
- `set_pos(list, para, pos)` ✅
- `move_pos(move_id, para, pos)` ✅
- `find(text, ...)` ✅
- `replace(find_text, replace_text, ...)` ✅
- `replace_all(find_text, replace_text, ...)` ✅
- `find_forward(src, regex)` ✅ - 아래 방향 찾기
- `find_backward(src, regex)` ✅ - 위 방향 찾기
- `find_replace(src, dst, regex, direction)` ✅ - 모두 찾아바꾸기
- `paste(option)` ✅ - 붙여넣기 확장
- `init_scan(option, range, ...)` ✅ - 텍스트 스캔 초기화
- `release_scan()` ✅ - 스캔 해제
- `select_text(spara, spos, epara, epos, slist)` ✅ - 범위 지정 텍스트 선택
- `get_pos_by_set()` ✅ - 위치 저장 (인덱스 반환)
- `set_pos_by_set(idx)` ✅ - 위치 복원 (인덱스 사용)
- `select_text_by_get_pos(s_getpos, e_getpos)` ✅ - GetPos 튜플로 선택
- `clear_pos_cache()` ✅ - 위치 캐시 정리
- `insert(path, format, arg, move_doc_end)` ✅ NEW - 파일 끼워넣기
- `insert_background_picture(path, ...)` ✅ NEW - 배경이미지 삽입
- `move_to_metatag(tag, text, start, select)` ✅ NEW - 메타태그로 이동
- `clear_field_text()` ✅ NEW - 필드 텍스트 지우기
- `insert_hyperlink(hypertext, description)` ✅ NEW - 하이퍼링크 삽입
- `insert_memo(text, memo_type)` ✅ NEW - 메모 삽입
- `compose_chars(chars, char_size, ...)` ✅ NEW - 원 문자 조합
- `move_to_ctrl(ctrl, option)` ✅ NEW - 컨트롤로 캐럿 이동
- `select_ctrl(ctrl, anchor_type, option)` ✅ NEW - 컨트롤 선택
- `move_all_caption(location, align)` ⚠️ NEW - 캡션 위치 일괄 변경 (테스트 실패)

#### ❌ 미구현 (2개)

| 메서드 | 미구현 사유 |
|--------|------------|
| `insert_lorem(para_num)` | Lorem Ipsum 외부 API 의존 (네트워크 필요), Python 전용 라이브러리 사용 |
| `clipboard_to_pyfunc()` | Python 전용 기능 (클립보드 매크로→Python 함수 변환), 보안 문제로 C++ 미지원 |

---

### 5. 테이블 작업 (65+개 문서화)

#### ✅ 구현됨 (71개, actions 포함)

**테이블 생성/탐색:**
- `create_table(rows, cols, treat_as_char, width_type, height_type, header)` ✅
- `get_into_nth_table(n, select_cell)` ✅
- `find_ctrl()` ✅
- `table_from_data(data, treat_as_char, header, header_bold, cell_fill_r, cell_fill_g, cell_fill_b)` ✅ NEW

**셀 이동:**
- `table_left_cell()` ✅
- `table_right_cell()` ✅
- `table_upper_cell()` ✅
- `table_lower_cell()` ✅
- `table_right_cell_append()` ✅
- `table_col_begin()` ✅ NEW
- `table_col_end()` ✅ NEW
- `table_col_page_up()` ✅ NEW
- `table_cell_block_extend_abs()` ✅ NEW
- `cancel()` ✅ NEW

**테이블 정보:**
- `get_table_row_count()` ✅
- `get_table_col_count()` ✅
- `is_cell()` ✅

**셀 서식:**
- `cell_fill(r, g, b)` ✅ NEW - 셀 배경색 채우기

**테이블 데이터 추출 (NEW):**
- `get_table_xml()` ✅ NEW - 테이블 XML 추출 (HWPML2X 형식)
- `cpyhwpx_utils.table_to_df(hwp, header)` ✅ NEW - DataFrame 변환
- `cpyhwpx_utils.table_to_csv(hwp, path)` ✅ NEW - CSV 저장
- `cpyhwpx_utils.table_to_string(hwp)` ✅ NEW - 문자열 변환

**HwpActionHelper (기존):**
- `TableCellBlock()` ✅
- `TableColBegin()` / `TableColEnd()` ✅
- `TableRowBegin()` / `TableRowEnd()` ✅
- `TableAppendRow()` / `TableAppendColumn()` ✅
- `TableDeleteRow()` / `TableDeleteColumn()` ✅
- `TableMergeCell()` / `TableSplitCell()` ✅
- `TableColPageUp()` ✅ NEW
- `TableCellBlockExtendAbs()` ✅ NEW

**HwpActionHelper 추가 (actions 서브모듈):**
- `TableCellAlignLeftTop`, `TableCellAlignCenterTop`, `TableCellAlignRightTop` ✅
- `TableCellAlignLeftCenter`, `TableCellAlignCenterCenter`, `TableCellAlignRightCenter` ✅
- `TableCellAlignLeftBottom`, `TableCellAlignCenterBottom`, `TableCellAlignRightBottom` ✅
- `TableVAlignTop`, `TableVAlignCenter`, `TableVAlignBottom` ✅
- `TableCellBorderAll`, `TableCellBorderOutside`, `TableCellBorderInside` 등 12개 ✅
- `TableResizeUp`, `TableResizeDown`, `TableResizeLeft`, `TableResizeRight` 등 16개 ✅
- `TableFormulaSumAuto`, `TableFormulaAvgHor`, `TableFormulaProVer` 등 9개 ✅

#### ❌ 미구현 (4개)

| 메서드 | 미구현 사유 |
|--------|------------|
| `get_row_height()` | HWP API에서 직접 조회 불가, CellShape 파라미터셋을 통해 간접 접근 필요 |
| `get_col_width()` | HWP API에서 직접 조회 불가, CellShape 파라미터셋을 통해 간접 접근 필요 |
| `set_row_height()` | `TableResizeUp`/`TableResizeDown` 액션으로 대체 가능 |
| `set_col_width()` | `TableResizeLeft`/`TableResizeRight` 액션으로 대체 가능 |

---

### 6. 필드/메타태그 (27개 문서화)

#### ✅ 구현됨 (23개)
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
- `modify_field_properties(field, remove, add)` ✅ NEW - 필드 속성 수정
- `find_private_info(type, string)` ✅ NEW - 개인정보 찾기
- `get_field_info()` ✅ NEW - 필드 정보 리스트 (HWPML2X 파싱)
- `set_field_by_bracket()` ⚠️ NEW - 중괄호를 필드로 변환 (테스트 실패)
- `get_cur_metatag_name()` ✅ NEW - 현재 메타태그명
- `get_metatag_list(number, option)` ✅ NEW - 메타태그 목록
- `get_metatag_name_text(tag)` ✅ NEW - 메타태그 텍스트 가져오기
- `put_metatag_name_text(tag, text)` ✅ NEW - 메타태그 텍스트 입력
- `rename_metatag(oldtag, newtag)` ✅ NEW - 메타태그명 변경
- `modify_metatag_properties(tag, remove, add)` ✅ NEW - 메타태그 속성 수정

#### ❌ 미구현 (4개)

| 메서드 | 미구현 사유 |
|--------|------------|
| `metatag_exist(tag)` | pyhwpx 미구현, `get_metatag_list()`로 목록 조회 후 확인 가능 |
| `delete_metatag(tag)` | pyhwpx 미구현, HWP2024+ API 필요 (메타태그 삭제 기능 제한) |
| `set_metatag_view_option(option)` | pyhwpx 미구현, `ViewOption` 액션으로 대체 가능 |
| `get_all_metatags()` | pyhwpx 미구현, `get_metatag_list()`로 대체 가능 |

---

### 7. 이미지/도형 객체 (60+개 문서화)

#### ✅ 구현됨 (51개, actions 포함)

**이미지 삽입:**
- `insert_picture(path, embedded, sizeoption, reverse, watermark, effect, width, height)` ✅
- `create_page_image(path, pgno, resolution, depth, format)` ✅ NEW - 페이지 이미지 생성

**컨트롤 관리:**
- `insert_ctrl(ctrl_id, initparam)` ✅ NEW - 컨트롤 삽입 (tbl/pic/gso/eqed 등)
- `delete_ctrl(ctrl)` ✅ NEW - 컨트롤 삭제

**HwpActionHelper:**
- `ShapeObjSelect()` ✅
- `ShapeObjDelete()` ✅
- `ShapeObjCopy()` ✅
- `ShapeObjCut()` ✅
- `ShapeObjBringToFront()` ✅
- `ShapeObjSendToBack()` ✅

**HwpActionHelper 추가 (actions 서브모듈):**
- 도형 정렬: `ShapeObjAlignLeft`, `ShapeObjAlignCenter`, `ShapeObjAlignRight`, `ShapeObjAlignTop`, `ShapeObjAlignMiddle`, `ShapeObjAlignBottom` 등 11개 ✅
- 도형 순서: `ShapeObjBringForward`, `ShapeObjSendBack`, `ShapeObjBringInFrontOfText`, `ShapeObjCtrlSendBehindText` ✅
- 그룹화: `ShapeObjGroup`, `ShapeObjUngroup`, `ShapeObjLock`, `ShapeObjUnlockAll` ✅
- 변환: `ShapeObjHorzFlip`, `ShapeObjVertFlip`, `ShapeObjRotater`, `ShapeObjRightAngleRotater` ✅
- 이동/크기: `ShapeObjMoveUp`, `ShapeObjMoveDown`, `ShapeObjResizeLeft`, `ShapeObjResizeRight` 등 8개 ✅
- 캡션: `ShapeObjAttachCaption`, `ShapeObjDetachCaption`, `ShapeObjInsertCaptionNum` ✅
- 글상자: `ShapeObjAttachTextBox`, `ShapeObjDetachTextBox`, `ShapeObjToggleTextBox`, `ShapeObjTextBoxEdit` ✅
- 속성: `ShapeObjFillProperty`, `ShapeObjLineProperty`, `ShapeObjWrapSquare`, `ShapeObjWrapTopAndBottom` ✅

#### ❌ 미구현 (4개)

| 메서드 | 미구현 사유 |
|--------|------------|
| `EquationCreate()` | 수식 편집기 UI 필요, 자동화 어려움 (대화형 입력 필요) |
| `EquationClose()` | `EquationCreate()` 의존, 수식 편집기 세션 관리 필요 |
| `EquationModify()` | 수식 편집기 내부 접근 불가, MathML 직접 편집 미지원 |
| `EquationEdit()` | 수식 편집기 UI 필요, `insert_ctrl("eqed")` 후 수동 편집 필요 |

---

### 8. 스타일/포맷팅 (70+개 문서화)

#### ✅ 구현됨 (45+개)

**글자 모양 API:**
- `get_charshape()` ✅ NEW - 글자모양 조회
- `set_charshape(props)` ✅ NEW - 글자모양 설정 (UnderlineType, StrikeOutType 등 포함)
- `set_font(face, height)` ✅ NEW - 간편 글꼴 설정

**문단 모양 API:**
- `get_parashape()` ✅ NEW - 문단모양 조회
- `set_parashape(props)` ✅ NEW - 문단모양 설정

**HwpActionHelper 기본:**
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

**HwpActionHelper 추가 (actions 서브모듈):**
- 문자 효과: `CharShapeOutline`, `CharShapeShadow`, `CharShapeEmboss`, `CharShapeEngrave`, `CharShapeCenterline` ✅
- 폰트 크기: `CharShapeHeight`, `CharShapeHeightIncrease`, `CharShapeHeightDecrease` ✅
- 자간/장평: `CharShapeSpacing`, `CharShapeSpacingIncrease`, `CharShapeSpacingDecrease`, `CharShapeWidth`, `CharShapeWidthIncrease`, `CharShapeWidthDecrease` ✅
- 텍스트 색상: `CharShapeTextColorRed`, `CharShapeTextColorBlue`, `CharShapeTextColorGreen`, `CharShapeTextColorYellow`, `CharShapeTextColorViolet`, `CharShapeTextColorBluish`, `CharShapeTextColorBlack`, `CharShapeTextColorWhite` ✅
- 문단 정렬: `ParagraphShapeAlignDistribute`, `ParagraphShapeAlignDivision` ✅
- 문단 여백: `ParagraphShapeIncreaseMargin`, `ParagraphShapeDecreaseMargin`, `ParagraphShapeIncreaseLeftMargin`, `ParagraphShapeDecreaseLeftMargin`, `ParagraphShapeIncreaseRightMargin`, `ParagraphShapeDecreaseRightMargin` ✅
- 줄간격: `ParagraphShapeIncreaseLineSpacing`, `ParagraphShapeDecreaseLineSpacing` ✅
- 들여쓰기: `ParagraphShapeIndentPositive`, `ParagraphShapeIndentNegative`, `ParagraphShapeIndentAtCaret` ✅

#### ❌ 미구현/API 문제 (25+개)

| 메서드 | 미구현 사유 |
|--------|------------|
| `export_style(sty_filepath)` | HWP COM API 예외 발생 (`pywintypes.com_error: -2147417851`), pyhwpx 동일 |
| `import_style(sty_filepath)` | `export_style()` 의존, 동일 COM 예외 발생 |
| `get_style_list()` | pyhwpx 미구현, HParameterSet 직접 접근 필요 |
| `set_style_by_name(name)` | pyhwpx 미구현, CharShape/ParaShape 액션 사용 권장 |
| `create_style(name, ...)` | pyhwpx 미구현, 복잡한 파라미터셋 구성 필요 |
| `delete_style(name)` | pyhwpx 미구현, 스타일 삭제 API 제한 |
| `modify_style(name, ...)` | pyhwpx 미구현, 스타일 수정 API 제한 |
| `apply_template(path)` | 템플릿 파일 적용, `insert_file()` + 스타일 병합 필요 |
| `get_default_style()` | pyhwpx 미구현, HParameterSet 직접 접근 필요 |
| `set_default_style(props)` | pyhwpx 미구현, 초기 스타일 설정 API 제한 |
| `copy_style(src, dst)` | pyhwpx 미구현, 스타일 복사 API 제한 |
| `style_to_charshape(style)` | pyhwpx 미구현, 스타일→글자모양 변환 필요 |
| `style_to_parashape(style)` | pyhwpx 미구현, 스타일→문단모양 변환 필요 |
| `get_heading_style(level)` | pyhwpx 미구현, 제목 스타일 조회 API 제한 |
| `set_heading_style(level, ...)` | pyhwpx 미구현, 제목 스타일 설정 API 제한 |
| `get_bullet_style()` | pyhwpx 미구현, 글머리표 스타일 조회 API 제한 |
| `set_bullet_style(...)` | pyhwpx 미구현, 글머리표 스타일 설정 API 제한 |
| `get_numbering_style()` | pyhwpx 미구현, 번호 매기기 스타일 조회 API 제한 |
| `set_numbering_style(...)` | pyhwpx 미구현, 번호 매기기 스타일 설정 API 제한 |
| `get_outline_style()` | pyhwpx 미구현, 개요 스타일 조회 API 제한 |
| `set_outline_style(...)` | pyhwpx 미구현, 개요 스타일 설정 API 제한 |
| `apply_quick_style(idx)` | pyhwpx 미구현, 빠른 스타일 적용 API 제한 |
| `get_page_style()` | pyhwpx 미구현, 페이지 스타일 조회 API 제한 |
| `set_page_style(...)` | pyhwpx 미구현, 페이지 스타일 설정 (HAction 필요) |
| `clone_formatting()` | pyhwpx 미구현, 서식 복사 (HAction 필요) |

---

### 9. Run 액션 (684개 문서화)

#### ✅ 구현됨 (558개, cpyhwpx.actions 서브모듈) - NEW +102개

**Break (6개):**
- `BreakPara`, `BreakPage`, `BreakSection`, `BreakColumn`, `BreakLine`, `BreakColDef`

**Move (60+개):**
- 기본: `MoveLeft`, `MoveRight`, `MoveUp`, `MoveDown`, `MoveDocBegin`, `MoveDocEnd`
- 확장: `MoveLineUp`, `MoveLineDown`, `MoveNextWord`, `MovePrevWord`, `MovePageBegin`, `MovePageEnd`
- 선택: `MoveSelDown`, `MoveSelLineEnd`, `MoveSelDocBegin`, `MoveSelNextWord` 등 20+개

**CharShape (30+개):**
- 기본: `CharShapeBold`, `CharShapeItalic`, `CharShapeUnderline`, `CharShapeStrikeout`
- 크기: `CharShapeHeight`, `CharShapeHeightIncrease`, `CharShapeHeightDecrease`
- 색상: `CharShapeTextColorRed`, `CharShapeTextColorBlue`, `CharShapeTextColorGreen` 등 8개
- 기타: `CharShapeSpacing`, `CharShapeWidth`, `CharShapeOutline`, `CharShapeShadow` 등

**ParagraphShape (20개):**
- 정렬: `AlignLeft`, `AlignCenter`, `AlignRight`, `AlignJustify`, `AlignDistribute`
- 여백: `IncreaseMargin`, `DecreaseMargin`, `IncreaseLeftMargin`, `DecreaseRightMargin` 등
- 들여쓰기: `IndentPositive`, `IndentNegative`, `IndentAtCaret`

**Table (70+개):**
- 이동: `TableLeftCell`, `TableRightCell`, `TableUpperCell`, `TableLowerCell`
- 블록: `TableCellBlock`, `TableCellBlockRow`, `TableCellBlockCol`, `TableCellBlockExtend`
- 정렬: `TableCellAlignLeftTop`, `TableCellAlignCenterCenter` 등 9개
- 수직정렬: `TableVAlignTop`, `TableVAlignCenter`, `TableVAlignBottom`
- 테두리: `TableCellBorderAll`, `TableCellBorderOutside`, `TableCellBorderInside` 등 12개
- 크기조절: `TableResizeUp`, `TableResizeDown`, `TableResizeCellLeft` 등 16개
- 수식: `TableFormulaSumAuto`, `TableFormulaAvgHor`, `TableFormulaProVer` 등 9개
- 도구: `TableAutoFill`, `TableDrawPen`, `TableEraser` 등

**ShapeObj (50+개):**
- 정렬: `ShapeObjAlignLeft`, `ShapeObjAlignCenter`, `ShapeObjAlignMiddle` 등 11개
- 순서: `ShapeObjBringForward`, `ShapeObjSendBack`, `ShapeObjBringToFront`, `ShapeObjSendToBack`
- 뒤집기/회전: `ShapeObjHorzFlip`, `ShapeObjVertFlip`, `ShapeObjRotater`, `ShapeObjRightAngleRotater`
- 이동/크기: `ShapeObjMoveUp`, `ShapeObjMoveDown`, `ShapeObjResizeLeft` 등 8개
- 그룹화: `ShapeObjGroup`, `ShapeObjUngroup`, `ShapeObjLock`, `ShapeObjUnlockAll`
- 캡션/글상자: `ShapeObjAttachCaption`, `ShapeObjAttachTextBox`, `ShapeObjTextBoxEdit` 등
- 속성: `ShapeObjFillProperty`, `ShapeObjLineProperty`, `ShapeObjWrapSquare` 등

**File (18개):**
- `FileNew`, `FileNewTab`, `FileOpen`, `FileSave`, `FileSaveAs`, `FileClose`, `FileQuit`
- `FilePrint`, `FilePrintPreview`, `FilePreview`, `FileFind`
- 버전비교: `FileNextVersionDiff`, `FilePrevVersionDiff`, `FileVersionDiffSyncScroll` 등

**Insert (23개):**
- 번호: `InsertAutoNum`, `InsertPageNum`, `InsertCpNo`, `InsertTpNo`
- 필드: `InsertFootnote`, `InsertEndnote`, `InsertFieldMemo`, `InsertFieldDateTime`
- 공백: `InsertSpace`, `InsertTab`, `InsertLine`, `InsertNonBreakingSpace`
- 날짜: `InsertDateCode`, `InsertStringDateTime`, `InsertLastSaveDate`

**Delete (13개):**
- `Delete`, `DeleteBack`, `DeleteLine`, `DeleteWord`, `DeleteWordBack`
- `DeleteField`, `DeleteFieldMemo`, `DeletePage`, `DeletePrivateInfoMark` 등

**TrackChange (13개):**
- `TrackChangeApply`, `TrackChangeApplyAll`, `TrackChangeCancel`, `TrackChangeCancelAll`
- `TrackChangeNext`, `TrackChangePrev`, `TrackChangeAuthor` 등

**ViewOption (11개):**
- `ViewZoomIn`, `ViewZoomOut`, `ViewZoomFitPage`, `ViewZoomFitWidth`
- `ViewOptionCtrlMark`, `ViewOptionParaMark`, `ViewOptionMemo`, `ViewOptionPicture` 등

**Window/Frame (20개):**
- `WindowMaximize`, `WindowMinimize`, `WindowMinimizeAll`, `WindowNextTab`, `WindowPrevTab`
- `FrameFullScreen`, `FrameHRuler`, `FrameVRuler`, `FrameStatusBar`
- `SplitHorz`, `SplitVert`, `SplitAll`, `NoSplit`
- `WindowAlignCascade`, `WindowAlignTileHorz`, `WindowAlignTileVert`

**기타 카테고리:**
- HeaderFooter (4개), Note/Memo (10개+8개 NEW), MasterPage (6개+5개 NEW), Picture (5개+8개 NEW)
- Input (4개), FormObj (7개+2개 NEW), Auto (4개+17개 NEW), DrawObj (4개)
- Quick (3개+21개 NEW), Macro (3개+22개 NEW), ViewOption (11개+14개 NEW)
- Window/Frame (20개+5개 NEW)

**사용법:**
```python
from cpyhwpx import actions

# 액션 실행
actions.MoveDocEnd(hwp)
actions.CharShapeBold(hwp)
actions.TableFormulaSumAuto(hwp)

# 직접 Run 호출
actions.Run(hwp, "CustomAction")
```

#### ❌ 미구현 Run 액션 (126개)

**사유**: pyhwpx 문서에는 684개로 기재되어 있으나, 실제 HWP API에서 지원하는 액션 수와 차이가 있음

| 카테고리 | 액션명 | 미구현 사유 |
|---------|--------|------------|
| Macro | `ScrMacroPlay9` | 스크립트 매크로 9번 슬롯 (거의 미사용) |
| Macro | `ScrMacroPlay10` | 스크립트 매크로 10번 슬롯 (거의 미사용) |
| Macro | `ScrMacroPlay11` | 스크립트 매크로 11번 슬롯 (거의 미사용) |
| 기타 | `HwpWSDic` | 웹 사전 연동 (네트워크 의존) |
| 기타 | 나머지 ~122개 | HWP 내부 액션 (대화상자/UI 전용, 자동화 불필요) |

**참고**: `Jajun`, `ChangeSkin`, `SoftKeyboard`는 이미 구현됨 (HwpAction.cpp:838-840)

---

### 10. 파라미터 헬퍼 (110개 문서화)

#### ✅ 구현됨 (108개) NEW

**파라미터셋 생성 (7개):**
- `CreateFindReplace()` ✅
- `CreateTable()` ✅
- `CreateCharShape()` ✅
- `CreateParaShape()` ✅
- `CreateInsertPicture()` ✅
- `CreateCellShape()` ✅
- `CreateBorderLine()` ✅

**정렬 관련 (5개) NEW:**
- `h_align(h_align)` ✅ - 수평 정렬 (Left, Center, Right, Justify)
- `v_align(v_align)` ✅ - 수직 정렬 (Top, Center, Bottom)
- `text_align(text_align)` ✅ - 텍스트 정렬
- `para_head_align(para_head_align)` ✅ - 문단 머리 정렬
- `text_art_align(text_art_align)` ✅ - 글맵시 정렬

**선/테두리 관련 (5개) NEW:**
- `hwp_line_type(line_type)` ✅ - 선 종류 (Solid, Dash, Dot 등)
- `hwp_line_width(line_width)` ✅ - 선 두께 (0.1mm~)
- `border_shape(border_type)` ✅ - 테두리 모양
- `end_style(end_style)` ✅ - 끝 스타일
- `end_size(end_size)` ✅ - 끝 크기

**서식 관련 (7개) NEW:**
- `number_format(num_format)` ✅ - 번호 형식
- `head_type(heading_type)` ✅ - 머리말 유형
- `font_type(font_type)` ✅ - 글꼴 유형
- `strike_out(strike_out_type)` ✅ - 취소선 유형
- `hwp_underline_type(underline_type)` ✅ - 밑줄 유형
- `hwp_underline_shape(underline_shape)` ✅ - 밑줄 모양
- `style_type(style_type)` ✅ - 스타일 유형

**검색/효과 (3개) NEW:**
- `find_dir(find_dir)` ✅ - 찾기 방향
- `pic_effect(pic_effect)` ✅ - 그림 효과
- `hwp_zoom_type(zoom_type)` ✅ - 줌 유형

**페이지/인쇄 (7개) NEW:**
- `page_num_position(pagenum_pos)` ✅ - 페이지 번호 위치
- `page_type(page_type)` ✅ - 페이지 유형
- `print_range(print_range)` ✅ - 인쇄 범위
- `print_type(print_method)` ✅ - 인쇄 방법
- `print_device(print_device)` ✅ - 인쇄 장치
- `print_paper(print_paper)` ✅ - 인쇄 용지
- `side_type(side_type)` ✅ - 측면 유형

**채우기/그라데이션 (5개) NEW:**
- `brush_type(brush_type)` ✅ - 브러시 유형
- `fill_area_type(fill_area)` ✅ - 채우기 영역
- `gradation(gradation)` ✅ - 그라데이션
- `hatch_style(hatch_style)` ✅ - 해치 스타일
- `watermark_brush(watermark_brush)` ✅ - 워터마크 브러시

**표 관련 (7개) NEW:**
- `table_format(table_format)` ✅ - 표 형식
- `table_break(page_break)` ✅ - 표 나누기
- `table_target(table_target)` ✅ - 표 대상
- `table_swap_type(tableswap)` ✅ - 표 교환 유형
- `cell_apply(cell_apply)` ✅ - 셀 적용
- `grid_method(grid_method)` ✅ - 그리드 방법
- `grid_view_line(grid_view_line)` ✅ - 그리드 보기 선

**텍스트 흐름/배치 (5개) NEW:**
- `text_dir(text_direction)` ✅ - 텍스트 방향
- `text_wrap_type(text_wrap)` ✅ - 텍스트 감싸기
- `text_flow_type(text_flow)` ✅ - 텍스트 흐름
- `line_wrap_type(line_wrap)` ✅ - 줄 감싸기
- `line_spacing_method(line_spacing)` ✅ - 줄 간격 방법

**도형/이미지 (7개) NEW:**
- `arc_type(arc_type)` ✅ - 호 유형
- `draw_aspect(draw_aspect)` ✅ - 그리기 종횡비
- `draw_fill_image(fillimage)` ✅ - 그리기 이미지 채우기
- `draw_shadow_type(shadow_type)` ✅ - 그리기 그림자 유형
- `char_shadow_type(shadow_type)` ✅ - 글자 그림자 유형
- `image_format(image_format)` ✅ - 이미지 형식
- `placement_type(restart)` ✅ - 배치 유형

**위치/크기 관련 (4개) NEW:**
- `horz_rel(horz_rel)` ✅ - 수평 상대 위치
- `vert_rel(vert_rel)` ✅ - 수직 상대 위치
- `height_rel(height_rel)` ✅ - 높이 상대 비율
- `width_rel(width_rel)` ✅ - 너비 상대 비율

**개요/번호 (4개) NEW:**
- `auto_num_type(autonum)` ✅ - 자동 번호 유형
- `numbering(numbering)` ✅ - 번호 매기기
- `hwp_outline_style(hwp_outline_style)` ✅ - 개요 스타일
- `hwp_outline_type(hwp_outline_type)` ✅ - 개요 유형

**열/단 정의 (3개) NEW:**
- `col_def_type(col_def_type)` ✅ - 열 정의 유형
- `col_layout_type(col_layout_type)` ✅ - 열 레이아웃 유형
- `gutter_method(gutter_type)` ✅ - 거터 방법

**기타 옵션 (20개) NEW:**
- `break_word_latin()`, `canonical()`, `convert_pua_hangul_to_unicode()` ✅
- `crooked_slash()`, `dbf_code_type()`, `delimiter()`, `ds_mark()` ✅
- `encrypt()`, `handler()`, `hash()`, `hiding()` ✅
- `macro_state()`, `mail_type()`, `present_effect()`, `signature()` ✅
- `slash()`, `sort_delimiter()`, `subt_pos()`, `view_flag()` ✅

**사용자 정보 (2개) NEW:**
- `get_user_info(user_info_id)` ✅ - 사용자 정보 가져오기
- `set_user_info(user_info_id, value)` ✅ - 사용자 정보 설정

**메타태그/DRM (2개) NEW:**
- `set_cur_metatag_name(tag)` ✅ - 현재 메타태그 이름 설정
- `set_drm_authority(authority)` ✅ - DRM 권한 설정

**번역 (1개) NEW:**
- `get_translate_lang_list(cur_lang)` ✅ - 번역 언어 목록

**음력/양력 변환 (2개) NEW:**
- `lunar_to_solar_by_set(l_year, l_month, l_day, l_leap)` ✅ - 음력→양력
- `solar_to_lunar_by_set(s_year, s_month, s_day)` ✅ - 양력→음력

**단위 변환 확장 (3개) NEW:**
- `hwp_unit_to_inch(hwp_unit)` ✅ - HwpUnit→인치
- `hwp_unit_to_point(hwp_unit)` ✅ - HwpUnit→포인트
- `point_to_hwp_unit(point)` ✅ - 포인트→HwpUnit

#### ❌ 미구현 (2개)
- `LunarToSolar()` - 단순 문자열 반환 (LunarToSolarBySet으로 대체)
- `SolarToLunar()` - 단순 문자열 반환 (SolarToLunarBySet으로 대체)

---

### 11. 폰트 프리셋 (111개 문서화)

#### ✅ 구현됨 (111개, 100%) NEW

**프리셋 조회 API:**
- `cpyhwpx.FontDefs.get_preset(name)` ✅ - 프리셋 이름으로 폰트 정보 가져오기
- `cpyhwpx.FontDefs.has_preset(name)` ✅ - 프리셋 존재 여부 확인
- `cpyhwpx.FontDefs.get_preset_names()` ✅ - 모든 프리셋 이름 목록

**맑은 고딕 (2개):**
- `MalgunGothic`, `MalgunGothicBold`

**전통 폰트 (8개):**
- `Gulim`, `GulimChe`, `Dotum`, `DotumChe`
- `Batang`, `BatangChe`, `Gungsuh`, `GungsuhChe`

**나눔 계열 (21개):**
- `NanumGothic`, `NanumGothicBold`, `NanumGothicLight`, `NanumGothicExtraBold`
- `NanumMyeongjo`, `NanumMyeongjoBold`, `NanumMyeongjoExtraBold`
- `NanumPenScript`, `NanumBrushScript`, `NanumBarunGothic`, `NanumBarunPen`
- `NanumSquare`, `NanumSquareBold`, `NanumSquareRound` 등

**한컴 계열 (9개):**
- `HCRBatang`, `HCRDotum`, `HYHeadline`, `HYGothic`, `HYPost`
- `HYGungso`, `HYPMokgak`, `HYGraphic`, `HYYeopseo`

**D2Coding (2개):**
- `D2Coding`, `D2CodingBold`

**영문 폰트 (14개):**
- `Arial`, `TimesNewRoman`, `CourierNew`, `Verdana`
- `Georgia`, `Tahoma`, `Consolas` (각각 Bold 포함)

**본고딕/본명조 (10개):**
- `SourceHanSans` 5개 (Light/Medium/Bold/Heavy 포함)
- `SourceHanSerif` 5개

**스포카 한 산스 (4개):**
- `SpoqaHanSans`, `SpoqaHanSansBold`, `SpoqaHanSansLight`, `SpoqaHanSansThin`

**Pretendard (9개):**
- `Pretendard`, `PretendardBold`, `PretendardLight`, `PretendardMedium`
- `PretendardSemiBold`, `PretendardExtraBold`, `PretendardThin`, `PretendardExtraLight`, `PretendardBlack`

**SUIT (9개):**
- `SUIT`, `SUITBold`, `SUITLight`, `SUITMedium`
- `SUITSemiBold`, `SUITExtraBold`, `SUITThin`, `SUITExtraLight`, `SUITHeavy`

**KoPub (6개):**
- `KoPubBatang`, `KoPubBatangBold`, `KoPubBatangLight`
- `KoPubDotum`, `KoPubDotumBold`, `KoPubDotumLight`

**마루부리 (5개):**
- `MaruBuri`, `MaruBuriBold`, `MaruBuriLight`, `MaruBuriSemiBold`, `MaruBuriExtraLight`

**보한글 레거시 폰트 (20개) NEW:**
- `BHGothic(고딕)`, `BHMyeongjo(명조)`, `BHSaemmul(샘물)`, `BHPilgi(필기)`
- `BHSinMyeongjo(신명조)`, `BHGyeonMyeongjo(견명조)`, `BHJungGothic(중고딕)`, `BHGyeonGothic(견고딕)`
- `BHGraphic(그래픽)`, `BHGungseo(궁서)`, `BHGaneunGonghan(가는공한)`, `BHJungganGonghan(중간공한)`
- `BHGulgeunGonghan(굵은공한)`, `BHGaneunHan(가는한)`, `BHJungganHan(중간한)`, `BHGulgeunHan(굵은한)`
- `BHPenHeullim(펜흘림)`, `BHHeadline(헤드라인)`, `BHGaneunHeadline(가는헤드라인)`, `BHTaeNamu(태나무)`

**휴먼 폰트 (9개) NEW:**
- `HumanMyeongjo(휴먼명조)`, `HumanGothic(휴먼고딕)`, `HumanYetche(휴먼옛체)`
- `HumanGaneunSaemche(휴먼가는샘체)`, `HumanJungganSaemche(휴먼중간샘체)`, `HumanGulgeunSaemche(휴먼굵은샘체)`
- `HumanGaneunPamche(휴먼가는팸체)`, `HumanJungganPamche(휴먼중간팸체)`, `HumanGulgeunPamche(휴먼굵은팸체)`

**양재 폰트 (10개) NEW:**
- `YangJaeDaunMyeongjoM(양재다운명조M)`, `YangJaeBonmokgakM(양재본목각M)`, `YangJaeSoseul(양재소슬)`
- `YangJaeTeunteunB(양재튼튼B)`, `YangJaeChamsutB(양재참숯B)`, `YangJaeDulgi(양재둘기)`
- `YangJaeMaehwa(양재매화)`, `YangJaeShanel(양재샤넬)`, `YangJaeWadang(양재와당)`, `YangJaeInitial(양재이니셜)`

**신명 폰트 (10개) NEW:**
- `SMSeMyeongjo(신명세명조)`, `SMJungMyeongjo(신명중명조)`, `SMTaeMyeongjo(신명태명조)`
- `SMGyeonMyeongjo(신명견명조)`, `SMSinmunMyeongjo(신명신문명조)`
- `SMSeGothic(신명세고딕)`, `SMJungGothic(신명중고딕)`, `SMTaeGothic(신명태고딕)`
- `SMGyeonGothic(신명견고딕)`, `SMGungseo(신명궁서)`

**특수 한자 폰트 (#접두사) (5개) NEW:**
- `HanjaSeMyeongjo(#세명조)`, `HanjaJungMyeongjo(#중명조)`, `HanjaTaeMyeongjo(#태명조)`
- `HanjaGyeonMyeongjo(#견명조)`, `HanjaJungGothic(#중고딕)`

**사용법:**
```python
import cpyhwpx

# 프리셋 조회
names = cpyhwpx.FontDefs.get_preset_names()
print(f"총 {len(names)}개 프리셋")

# 특정 프리셋 가져오기 (한글 또는 영문명)
preset = cpyhwpx.FontDefs.get_preset("휴먼고딕")
preset = cpyhwpx.FontDefs.get_preset("HumanGothic")

# 프리셋 정보 (객체 속성으로 접근)
print(f"FaceNameHangul: {preset.FaceNameHangul}")
print(f"FontTypeHangul: {preset.FontTypeHangul}")

# 폰트 적용
hwp.set_font(preset.FaceNameHangul, 12)
```

---

### 12. 유틸리티 (60+개 문서화)

#### ✅ 구현됨 (28개)

**문서/탭 관리:**
- `switch_to(num)` ✅ - 문서 전환
- `add_tab()` ✅ - 새 탭 추가
- `add_doc()` ✅ - 새 문서 추가
- `doc_list` ✅ NEW - 문서 컬렉션 (XHwpDocuments 별칭)

**창/UI 관리 (NEW):**
- `set_visible(visible)` ✅ NEW - 창 표시/숨김
- `set_viewstate(flag)` ✅ NEW - 뷰 상태 설정 (0=조판부호~6=줄표시)
- `get_viewstate()` ✅ NEW - 뷰 상태 가져오기
- `msgbox(message, flag)` ✅ NEW - 메시지 박스 표시
- `get_message_box_mode()` ✅ NEW - 메시지 박스 모드 가져오기
- `set_message_box_mode(mode)` ✅ NEW - 메시지 박스 모드 설정

**상태 조회 (NEW):**
- `key_indicator()` ✅ NEW - 키 인디케이터 (구역, 페이지, 줄, 위치 등)
- `goto_page(page_index)` ✅ NEW - 페이지로 이동

**단위 변환 (NEW):**
- `mili_to_hwp_unit(mili)` ✅ NEW - 밀리미터→HwpUnit 변환
- `hwp_unit_to_mili(hwp_unit)` ✅ NEW - HwpUnit→밀리미터 변환 (정적)

**Utils 서브모듈:**
- `addr_to_tuple()`, `tuple_to_addr()` ✅
- `parse_range()`, `expand_range()` ✅
- `trim()`, `split()`, `join()` ✅
- `file_exists()`, `get_extension()` ✅
- `hex_to_colorref()`, `colorref_to_hex()` ✅

**Units 서브모듈:**
- `from_mm()`, `from_cm()`, `from_inch()`, `from_point()` ✅
- `to_mm()`, `to_cm()`, `to_inch()`, `to_point()` ✅

#### ❌ 미구현 (32+개)

| 메서드 | 미구현 사유 |
|--------|------------|
| `maximize_window()` | win32gui 의존, `WindowMaximize` 액션으로 대체 가능 |
| `minimize_window()` | win32gui 의존, `WindowMinimize` 액션으로 대체 가능 |
| `restore_window()` | win32gui 의존, `WindowRestore` 액션으로 대체 가능 |
| `get_window_handle()` | win32gui 의존, COM 객체 핸들 직접 접근 불가 |
| `set_window_position(x, y)` | win32gui 의존, 자동화에 불필요 |
| `set_window_size(w, h)` | win32gui 의존, 자동화에 불필요 |
| `goto_printpage(page)` | pyhwpx에서 `goto_page()`로 통합 |
| `get_page_text(page)` | `get_text_file()` + 페이지 범위 지정으로 대체 가능 |
| `is_empty_page(page)` | 선택적 기능, 거의 미사용 |
| `is_empty_para()` | 선택적 기능, `get_text()`로 확인 가능 |
| `is_action_enable(action)` | HAction 저수준 API 필요 |
| `get_ctrl_by_ctrl_id(id)` | `find_ctrl()`로 대체 가능 |
| `replace_action(old, new)` | 액션 교체 기능, 거의 미사용 |
| `release_action(action)` | 수동 메모리 관리, 자동 해제 권장 |
| `save_pdf_as_image(path)` | `create_page_image()`로 대체 가능 |
| `export_mathml(path)` | 수식 MathML 내보내기, 고급 기능 |
| `import_mathml(path)` | 수식 MathML 가져오기, 고급 기능 |
| `set_private_info_password(pw)` | 개인정보 보호, 선택적 기능 |
| `set_title(title)` | `GetTitle()` 구현됨, Setter 미구현 (읽기 전용) |
| `get_hwp_version_info()` | `version` 속성으로 대체 가능 |
| `get_install_path()` | 레지스트리 조회 필요, 자동화에 불필요 |
| `check_spell(text)` | 맞춤법 검사 API 의존, 고급 기능 |
| `convert_encoding(text, enc)` | Python 내장 기능으로 대체 가능 |
| `compress_file(path)` | 압축 라이브러리 의존, 자동화에 불필요 |
| `decompress_file(path)` | 압축 라이브러리 의존, 자동화에 불필요 |
| `get_clipboard_text()` | win32clipboard 의존, Python 전용 |
| `set_clipboard_text(text)` | win32clipboard 의존, Python 전용 |
| `get_clipboard_image()` | win32clipboard/PIL 의존, Python 전용 |
| `wait_for_idle(timeout)` | HWP 상태 폴링 필요, 복잡한 구현 |
| `sleep_hwp(ms)` | Python `time.sleep()`으로 대체 가능 |
| `run_script(js_code)` | JavaScript 실행, 보안 문제 |
| `execute_macro(name)` | 매크로 실행, `MacroPlay` 액션으로 대체 가능 |

---

## 우선순위별 미구현 기능

### ✅ 완료
1. ~~**필드 작업** - 문서 자동화의 핵심~~ ✅ (13개 구현, 48%)
2. ~~**테이블 생성/탐색** - 보고서 생성 필수~~ ✅ (10개 구현, 32%)
   - `create_table()`, `get_into_nth_table()`, 셀 이동 5개
3. ~~**이미지 삽입** - 문서 작성 필수~~ ✅
   - `insert_picture()` - 8개 파라미터 지원
4. ~~**파일 삽입** - 문서 병합~~ ✅
   - `insert_file()` - 6개 파라미터 지원
5. ~~**보안 모듈 자동 등록** - 파일 승인 다이얼로그 우회~~ ✅
   - `check_registry_key()`, `find_dll_path()`, `register_to_registry()`, `auto_register_module()`
6. ~~**파일 작업 확장** - 문서 텍스트 추출/삽입~~ ✅ NEW
   - `get_text_file(format, option)` - HWP/HWPML2X/HTML/UNICODE/TEXT 지원
   - `set_text_file(data, format, option)` - 텍스트 데이터 삽입
7. ~~**컨트롤 관리** - 도형/객체 생성/삭제~~ ✅ NEW
   - `insert_ctrl(ctrl_id, initparam)` - tbl/pic/gso/eqed 등 컨트롤 삽입
   - `delete_ctrl(ctrl)` - 컨트롤 삭제
8. ~~**테이블 데이터** - 대량 데이터 입력~~ ✅ NEW
   - `table_from_data(data, ...)` - 2D 리스트로 테이블 생성
   - `cell_fill(r, g, b)` - 셀 배경색 채우기
   - `cpyhwpx_utils.table_from_data()` - DataFrame/CSV/dict 지원 wrapper
9. ~~**보안 모듈 생성자 자동등록** - pyhwpx 호환~~ ✅ NEW
   - `Hwp(register_module=True)` - 기본값으로 자동 등록
10. ~~**테이블 출력** - 테이블 데이터 추출~~ ✅ NEW
    - `get_table_xml()` - HWPML2X XML 추출
    - `cpyhwpx_utils.table_to_df()`, `table_to_csv()`, `table_to_string()`
11. ~~**XHwpDocuments/XHwpDocument** - 다중 문서 관리~~ ✅ NEW
    - `XHwpDocuments`, `XHwpDocument` 클래스
    - `switch_to()`, `add_tab()`, `add_doc()` 편의 메서드
12. ~~**글자모양 관리** - 스타일 조회/설정~~ ✅ NEW
    - `get_charshape()`, `set_charshape()`, `set_font()`
    - UnderlineType, StrikeOutType 등 하위 속성 지원
13. ~~**텍스트 검색** - 찾기/바꾸기~~ ✅ NEW
    - `find(text, forward, match_case, regex, replace_mode)` - 텍스트 찾기
    - `replace(find_text, replace_text, forward, match_case, regex)` - 바꾸기
    - `replace_all(find_text, replace_text, match_case, regex)` - 모두 바꾸기
14. ~~**유틸리티 확장** - 창/UI 관리, 상태 조회, 단위 변환~~ ✅ NEW
    - `doc_list`, `set_visible()`, `msgbox()`, `get/set_viewstate()`
    - `get/set_message_box_mode()`, `key_indicator()`, `goto_page()`
    - `mili_to_hwp_unit()`, `hwp_unit_to_mili()`

### 🟡 중간 (기능 확장)
1. ~~**추가 속성** - `CurSelectedCtrl`, `HeadCtrl`, `ctrl_list`~~ ✅ NEW (5개 구현)
   - `head_ctrl`, `last_ctrl`, `parent_ctrl`, `cur_selected_ctrl`, `ctrl_list`
2. ~~**문단모양 관리** - `get_parashape()`, `set_parashape()`~~ ✅ 구현됨
3. ~~**텍스트 편집 확장**~~ ✅ NEW (6개 구현)
   - `init_scan()`, `release_scan()`, `select_text()`, `get_pos_by_set()`, `set_pos_by_set()`, `select_text_by_get_pos()`
4. ~~**파일 I/O 확장**~~ ✅ NEW (2개 구현)
   - `save_block_as()`, `get_file_info()`

### 🟢 낮음 (추후 구현)
1. ~~나머지 Run 액션들 (646개)~~ → ✅ 558개 구현 완료 (82%)
2. 폰트 프리셋 (108개)
3. 파라미터 헬퍼 (83개)
4. 메타태그 (HWP2024+)

---

## 소스 파일 위치

| 파일 | 역할 | 라인 수 |
|------|------|--------|
| `src/HwpWrapper.h` | 메인 래퍼 헤더 | ~480 |
| `src/HwpWrapper.cpp` | 메인 래퍼 구현 | ~1200 |
| `src/HwpCtrl.h/cpp` | 컨트롤 래퍼 | - |
| `src/HwpAction.h/cpp` | 액션 헬퍼 | **~880** (558개 액션) |
| `src/HwpParameter.h/cpp` | 파라미터셋 래퍼 | - |
| `src/XHwpDocument.h/cpp` | 개별 문서 래퍼 | - |
| `src/XHwpDocuments.h/cpp` | 문서 컬렉션 래퍼 | - |
| `src/bindings.cpp` | Python 바인딩 | **~1850** (actions 포함) |
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
