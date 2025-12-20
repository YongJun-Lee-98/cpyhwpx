# cpyhwpx 구현 현황 분석

> **분석 일자**: 2025-12-21 (우선순위 API 14개 추가)

## 요약

| 구분 | 문서화 | 구현됨 | 구현률 |
|-----|--------|--------|--------|
| **총 API** | ~1,300+ | **~764** | **~59%** ✅ |
| Core 클래스 | 4 | 4 | 100% ✅ |
| 속성 | 35 | **12** | **34%** ✅ NEW |
| **파일 I/O** | 26 | **15** | **58%** ✅ NEW |
| **보안 모듈** | 4 | **4** | **100%** ✅ |
| 텍스트 편집 | 35 | **16** | **46%** ✅ NEW |
| **테이블 작업** | 65+ | **71** | **100%+** ✅ (actions 포함) |
| **필드/메타태그** | 27 | **13** | **48%** ✅ |
| **이미지/도형** | 60+ | **50** | **83%** ✅ (actions 포함) |
| **스타일/포맷팅** | 70+ | **45** | **64%** ✅ (actions 포함) |
| **Run 액션** | 684 | **456** | **67%** ✅ |
| 파라미터 헬퍼 | 90 | 7 | 8% |
| 폰트 프리셋 | 111 | 3 | 3% |
| 유틸리티 | 60+ | 18 | 30% |

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

#### ✅ 구현됨 (12개)
- `version` - HWP 버전
- `build_number` - 빌드 번호
- `current_page` - 현재 페이지
- `page_count` - 총 페이지 수
- `edit_mode` - 편집 모드 (읽기/쓰기)
- `is_initialized` - 초기화 여부
- `XHwpDocuments` - 문서 컬렉션
- `head_ctrl` - 첫 번째 컨트롤 ✅ NEW
- `last_ctrl` - 마지막 컨트롤 ✅ NEW
- `parent_ctrl` - 부모 컨트롤 ✅ NEW
- `cur_selected_ctrl` - 현재 선택된 컨트롤 ✅ NEW
- `ctrl_list` - 모든 컨트롤 목록 ✅ NEW

#### ❌ 미구현 (23개)
- `Application` - Low-level API 접근
- `CLSID` - 클래스 ID
- `CurFieldState` - 현재 필드 상태
- `CurMetatagState` - 현재 메타태그 상태
- `EngineProperties` - 엔진 속성
- `IsEmpty` - 빈 문서 여부 (메서드로만 존재)
- `IsModified` - 수정 여부 (메서드로만 존재)
- `IsPrivateInfoProtected` - 개인정보 보호
- `IsTrackChange` - 변경 추적
- `Path` - 문서 경로
- `SelectionMode` - 선택 모드
- `Title` - 창 제목
- `XHwpMessageBox` - 메시지 박스 객체
- `XHwpODBC` - ODBC 객체
- `XHwpWindows` - 창 관리 객체
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

#### ✅ 구현됨 (15개)
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
- `save_block_as(path, format, attributes)` ✅ NEW - 선택 블록 저장
- `get_file_info(filename)` ✅ NEW - 파일 정보 조회

#### ❌ 미구현 (11개)
- `open_pdf(pdf_path, this_window)` - PDF 열기 (구현됨, 테스트 스킵)
- `FileClose()` (액션으로는 존재)

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

#### ✅ 구현됨 (16개)
- `insert_text(text)` ✅
- `get_text()` ✅
- `get_selected_text(keep_select)` ✅
- `get_pos()` ✅
- `set_pos(list, para, pos)` ✅
- `move_pos(move_id, para, pos)` ✅
- `find(text, ...)` ✅
- `replace(find_text, replace_text, ...)` ✅
- `replace_all(find_text, replace_text, ...)` ✅
- `init_scan(option, range, ...)` ✅ NEW - 텍스트 스캔 초기화
- `release_scan()` ✅ NEW - 스캔 해제
- `select_text(spara, spos, epara, epos, slist)` ✅ NEW - 범위 지정 텍스트 선택
- `get_pos_by_set()` ✅ NEW - 위치 저장 (인덱스 반환)
- `set_pos_by_set(idx)` ✅ NEW - 위치 복원 (인덱스 사용)
- `select_text_by_get_pos(s_getpos, e_getpos)` ✅ NEW - GetPos 튜플로 선택
- `clear_pos_cache()` ✅ NEW - 위치 캐시 정리

#### ❌ 미구현 (19개)
- `move_to_field(field, idx, text, start, select)`
- `find_forward(src, regex)`
- `find_backward(src, regex)`
- `find_replace(src, dst, ...)`
- `paste(option)`

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
- `get_row_height()` / `get_col_width()`
- `set_row_height()` / `set_col_width()`

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

#### ✅ 구현됨 (50개, actions 포함)

**이미지 삽입:**
- `insert_picture(path, embedded, sizeoption, reverse, watermark, effect, width, height)` ✅

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

#### ❌ 미구현 (5개)
- `create_page_image(path, pgno, ...)`
- `EquationCreate()` / `EquationClose()` / `EquationModify()`

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

#### ❌ 미구현 (25+개)
- `import_style(sty_filepath)`
- 기타 고급 스타일 기능들

---

### 9. Run 액션 (684개 문서화)

#### ✅ 구현됨 (456개, cpyhwpx.actions 서브모듈) - NEW

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
- HeaderFooter (4개), Note/Memo (10개), MasterPage (6개), Picture (5개)
- Input (4개), FormObj (7개), Auto (4개), DrawObj (4개)
- Quick (3개), Macro (3개), Misc (7개)

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

#### ❌ 미구현 (228개)
- 일부 특수 Run 액션들

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

#### ✅ 구현됨 (18개)

**문서/탭 관리 (NEW):**
- `switch_to(num)` ✅ NEW - 문서 전환
- `add_tab()` ✅ NEW - 새 탭 추가
- `add_doc()` ✅ NEW - 새 문서 추가

**Utils 서브모듈:**
- `addr_to_tuple()`, `tuple_to_addr()` ✅
- `parse_range()`, `expand_range()` ✅
- `trim()`, `split()`, `join()` ✅
- `file_exists()`, `get_extension()` ✅
- `hex_to_colorref()`, `colorref_to_hex()` ✅

**Units 서브모듈:**
- `from_mm()`, `from_cm()`, `from_inch()`, `from_point()` ✅
- `to_mm()`, `to_cm()`, `to_inch()`, `to_point()` ✅

#### ❌ 미구현 (42+개)
- `doc_list()`
- `set_viewstate()` / `get_viewstate()` (일부 구현)
- `msgbox()` (일부 구현)
- 기타 문서/탭 관리 유틸리티

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

### 🟡 중간 (기능 확장)
1. ~~**추가 속성** - `CurSelectedCtrl`, `HeadCtrl`, `ctrl_list`~~ ✅ NEW (5개 구현)
   - `head_ctrl`, `last_ctrl`, `parent_ctrl`, `cur_selected_ctrl`, `ctrl_list`
2. ~~**문단모양 관리** - `get_parashape()`, `set_parashape()`~~ ✅ 구현됨
3. ~~**텍스트 편집 확장**~~ ✅ NEW (6개 구현)
   - `init_scan()`, `release_scan()`, `select_text()`, `get_pos_by_set()`, `set_pos_by_set()`, `select_text_by_get_pos()`
4. ~~**파일 I/O 확장**~~ ✅ NEW (2개 구현)
   - `save_block_as()`, `get_file_info()`

### 🟢 낮음 (추후 구현)
1. ~~나머지 Run 액션들 (646개)~~ → ✅ 456개 구현 완료 (67%)
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
| `src/HwpAction.h/cpp` | 액션 헬퍼 | **~680** (456개 액션) |
| `src/HwpParameter.h/cpp` | 파라미터셋 래퍼 | - |
| `src/XHwpDocument.h/cpp` | 개별 문서 래퍼 | - |
| `src/XHwpDocuments.h/cpp` | 문서 컬렉션 래퍼 | - |
| `src/bindings.cpp` | Python 바인딩 | **~1250** (actions 포함) |
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
