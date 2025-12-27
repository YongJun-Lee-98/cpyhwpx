# cpyhwpx 미확인 함수 목록

이 문서는 pyhwpx에서 대응하는 docstring을 찾지 못했거나, 추가 문서화가 필요한 함수들의 목록입니다.

## 파라미터 헬퍼 (Parameter Helpers)

아래 함수들은 한/글 API의 열거형 값을 문자열로 변환하는 헬퍼 함수들입니다.
pyhwpx의 `param_helpers.py`에 대응하지만 상세 docstring이 없습니다.

### 정렬 관련
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `h_align(h_align)` | 수평 정렬 변환 | Left, Center, Right, Justify |
| `v_align(v_align)` | 수직 정렬 변환 | Top, Center, Bottom |
| `text_align(text_align)` | 텍스트 정렬 변환 | |
| `para_head_align(para_head_align)` | 문단 머리 정렬 변환 | |
| `text_art_align(text_art_align)` | 글맵시 정렬 변환 | |

### 선/테두리 관련
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `hwp_line_type(line_type)` | 선 종류 변환 | Solid, Dash, Dot 등 |
| `hwp_line_width(line_width)` | 선 두께 변환 | 0.1mm, 0.12mm 등 |
| `border_shape(border_type)` | 테두리 모양 변환 | |
| `end_style(end_style)` | 끝 스타일 변환 | |
| `end_size(end_size)` | 끝 크기 변환 | |

### 서식 관련
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `number_format(num_format)` | 번호 형식 변환 | |
| `head_type(heading_type)` | 머리말 유형 변환 | |
| `font_type(font_type)` | 글꼴 유형 변환 | |
| `strike_out(strike_out_type)` | 취소선 유형 변환 | |
| `hwp_underline_type(underline_type)` | 밑줄 유형 변환 | |
| `hwp_underline_shape(underline_shape)` | 밑줄 모양 변환 | |
| `style_type(style_type)` | 스타일 유형 변환 | |

### 검색/효과
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `find_dir(find_dir)` | 찾기 방향 변환 | |
| `pic_effect(pic_effect)` | 그림 효과 변환 | |
| `hwp_zoom_type(zoom_type)` | 줌 유형 변환 | |

### 페이지/인쇄
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `page_num_position(pagenum_pos)` | 페이지 번호 위치 변환 | |
| `page_type(page_type)` | 페이지 유형 변환 | |
| `print_range(print_range)` | 인쇄 범위 변환 | |
| `print_type(print_method)` | 인쇄 방법 변환 | |
| `print_device(print_device)` | 인쇄 장치 변환 | |
| `print_paper(print_paper)` | 인쇄 용지 변환 | |
| `side_type(side_type)` | 측면 유형 변환 | |

### 채우기/그라데이션
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `brush_type(brush_type)` | 브러시 유형 변환 | |
| `fill_area_type(fill_area)` | 채우기 영역 변환 | |
| `gradation(gradation)` | 그라데이션 변환 | |
| `hatch_style(hatch_style)` | 해치 스타일 변환 | |
| `watermark_brush(watermark_brush)` | 워터마크 브러시 변환 | |

### 표 관련
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `table_format(table_format)` | 표 형식 변환 | |
| `table_break(page_break)` | 표 나누기 변환 | |
| `table_target(table_target)` | 표 대상 변환 | |
| `table_swap_type(tableswap)` | 표 교환 유형 변환 | |
| `cell_apply(cell_apply)` | 셀 적용 변환 | |
| `grid_method(grid_method)` | 그리드 방법 변환 | |
| `grid_view_line(grid_view_line)` | 그리드 보기 선 변환 | |

### 텍스트 흐름/배치
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `text_dir(text_direction)` | 텍스트 방향 변환 | |
| `text_wrap_type(text_wrap)` | 텍스트 감싸기 변환 | |
| `text_flow_type(text_flow)` | 텍스트 흐름 변환 | |
| `line_wrap_type(line_wrap)` | 줄 감싸기 변환 | |
| `line_spacing_method(line_spacing)` | 줄 간격 방법 변환 | |

### 도형/이미지
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `arc_type(arc_type)` | 호 유형 변환 | |
| `draw_aspect(draw_aspect)` | 그리기 종횡비 변환 | |
| `draw_fill_image(fillimage)` | 그리기 이미지 채우기 변환 | |
| `draw_shadow_type(shadow_type)` | 그리기 그림자 유형 변환 | |
| `char_shadow_type(shadow_type)` | 글자 그림자 유형 변환 | |
| `image_format(image_format)` | 이미지 형식 변환 | |
| `placement_type(restart)` | 배치 유형 변환 | |

### 위치/크기 관련
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `horz_rel(horz_rel)` | 수평 상대 위치 변환 | |
| `vert_rel(vert_rel)` | 수직 상대 위치 변환 | |
| `height_rel(height_rel)` | 높이 상대 비율 변환 | |
| `width_rel(width_rel)` | 너비 상대 비율 변환 | |

### 개요/번호
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `auto_num_type(autonum)` | 자동 번호 유형 변환 | |
| `numbering(numbering)` | 번호 매기기 변환 | |
| `hwp_outline_style(hwp_outline_style)` | 개요 스타일 변환 | |
| `hwp_outline_type(hwp_outline_type)` | 개요 유형 변환 | |

### 열/단 정의
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `col_def_type(col_def_type)` | 열 정의 유형 변환 | |
| `col_layout_type(col_layout_type)` | 열 레이아웃 유형 변환 | |
| `gutter_method(gutter_type)` | 거터 방법 변환 | |

### 기타 옵션
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `break_word_latin(break_latin_word)` | 라틴어 단어 나누기 변환 | |
| `canonical(canonical)` | 표준 형식 변환 | |
| `convert_pua_hangul_to_unicode(reverse)` | PUA 한글을 유니코드로 변환 | |
| `crooked_slash(crooked_slash)` | 비뚤어진 슬래시 변환 | |
| `dbf_code_type(dbf_code)` | DBF 코드 유형 변환 | |
| `delimiter(delimiter)` | 구분 문자 변환 | |
| `ds_mark(diac_sym_mark)` | 발음 기호 표시 변환 | |
| `encrypt(encrypt)` | 암호화 변환 | |
| `handler(handler)` | 핸들러 변환 | |
| `hash(hash)` | 해시 변환 | |
| `hiding(hiding)` | 숨기기 변환 | |
| `macro_state(macro_state)` | 매크로 상태 변환 | |
| `mail_type(mail_type)` | 메일 유형 변환 | |
| `present_effect(prsnteffect)` | 프레젠테이션 효과 변환 | |
| `signature(signature)` | 서명 변환 | |
| `slash(slash)` | 슬래시 변환 | |
| `sort_delimiter(sort_delimiter)` | 정렬 구분자 변환 | |
| `subt_pos(subt_pos)` | 자막 위치 변환 | |
| `view_flag(view_flag)` | 보기 플래그 변환 | |

---

## cpyhwpx 고유 함수

아래 함수들은 cpyhwpx에서 새로 추가되었거나 pyhwpx와 다른 구현을 가진 함수들입니다.

### 보안 모듈 관련 (cpyhwpx 고유)
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `check_registry_key(key_name)` | 보안 모듈 레지스트리 등록 여부 확인 | 정적 메서드 |
| `register_to_registry(dll_path, key_name)` | 보안 모듈 DLL을 레지스트리에 등록 | 정적 메서드 |
| `find_dll_path()` | DLL 파일 경로 자동 감지 | 정적 메서드 |
| `auto_register_module(module_type, module_data)` | 보안 모듈 자동 등록 | |

### 위치 캐시 관련 (cpyhwpx 확장)
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `get_pos_by_set()` | 위치 저장 후 인덱스 반환 | pyhwpx는 ParameterSet 반환 |
| `set_pos_by_set(idx)` | 인덱스로 위치 복원 | pyhwpx는 ParameterSet 사용 |
| `clear_pos_cache()` | 위치 캐시 정리 | cpyhwpx 고유 |

### 메타태그 관련
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `get_cur_metatag_name()` | 현재 메타태그명 조회 | 한글2024+ |
| `get_metatag_list(number, option)` | 메타태그 목록 조회 | |
| `get_metatag_name_text(tag)` | 메타태그 텍스트 조회 | |
| `put_metatag_name_text(tag, text)` | 메타태그 텍스트 설정 | |
| `rename_metatag(oldtag, newtag)` | 메타태그 이름 변경 | |
| `modify_metatag_properties(tag, remove, add)` | 메타태그 속성 수정 | |

### 필드/문서 확장
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `modify_field_properties(field, remove, add)` | 필드 속성 수정 | |
| `find_private_info(private_type, private_string)` | 개인정보 찾기 | |
| `get_field_info()` | 필드 정보 리스트 (HWPML2X 파싱) | |
| `set_field_by_bracket()` | 중괄호 구문을 필드로 변환 | {{name}}, [[name]] |
| `get_table_xml()` | 테이블 XML 추출 | HWPML2X 형식 |

### 테이블 이동 (간략 docstring)
| 함수명 | 설명 | 비고 |
|--------|------|------|
| `table_left_cell()` | 왼쪽 셀로 이동 | |
| `table_right_cell()` | 오른쪽 셀로 이동 | |
| `table_upper_cell()` | 위쪽 셀로 이동 | |
| `table_lower_cell()` | 아래쪽 셀로 이동 | |
| `table_right_cell_append()` | 오른쪽 셀로 이동 (행 끝 시 다음 행) | |
| `table_col_begin()` | 첫 번째 열로 이동 | |
| `table_col_end()` | 마지막 열로 이동 | |
| `table_col_page_up()` | 열의 맨 위로 이동 | |
| `table_cell_block_extend_abs()` | 셀 블록 선택 확장 | |

---

## 참고사항

1. 파라미터 헬퍼 함수들은 `pyhwpx/param_helpers.py`에 대응하며, 한/글 API의 열거형 값을 문자열에서 정수로 변환합니다.

2. 각 함수의 정확한 열거형 값은 한/글 개발자 문서 또는 `HwpTypes.h` 파일을 참조하세요.

3. cpyhwpx 고유 함수들은 C++ 구현의 특성상 pyhwpx와 다른 방식으로 동작할 수 있습니다.

---

## 문서화 우선순위

### 높음 (자주 사용)
- 파라미터 헬퍼 중 정렬, 선/테두리 관련
- 보안 모듈 관련 함수

### 중간 (필요시 사용)
- 메타태그 관련 함수 (한글2024+ 기능)
- 필드/문서 확장 함수

### 낮음 (특수 용도)
- 나머지 파라미터 헬퍼 함수들
