#!/usr/bin/env python3
"""
cpyhwpx API docstring 추출 스크립트

bindings.cpp에서 함수명, 인자, docstring을 파싱하여 JSON으로 저장합니다.
"""

import re
import json
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Any, Optional


# 카테고리 정의
CATEGORIES = {
    'init': {
        'name': '초기화/종료',
        'patterns': ['initialize', 'register_module', 'is_initialized', 'quit', 'auto_register'],
        'description': 'HWP 객체 생성, 초기화, 종료 관련'
    },
    'file_io': {
        'name': '파일 I/O',
        'patterns': ['open', 'save', 'close', 'clear', 'insert_file', 'get_text_file',
                     'set_text_file', 'export', 'import', 'mail_merge', 'print_'],
        'description': '문서 열기, 저장, 내보내기 관련'
    },
    'text_editing': {
        'name': '텍스트 편집',
        'patterns': ['insert_text', 'get_text', 'get_selected', 'find', 'replace',
                     'select', 'init_scan', 'release_scan', 'clipboard'],
        'description': '텍스트 삽입, 검색, 치환 관련'
    },
    'cursor': {
        'name': '커서/위치',
        'patterns': ['get_pos', 'set_pos', 'move_pos', 'move_to', 'goto_', 'key_indicator'],
        'description': '커서 이동, 위치 조회 관련'
    },
    'table': {
        'name': '표',
        'patterns': ['table_', 'cell_', 'create_table', 'get_row', 'get_col', 'set_row',
                     'set_col', 'is_cell'],
        'description': '표 생성, 셀 편집 관련'
    },
    'field': {
        'name': '필드/누름틀',
        'patterns': ['field_', 'create_field', 'get_field', 'put_field', 'click_here'],
        'description': '필드, 누름틀 관련'
    },
    'style': {
        'name': '스타일/서식',
        'patterns': ['charshape', 'parashape', 'font', 'set_style', 'get_style'],
        'description': '글자 모양, 문단 모양, 스타일 관련'
    },
    'shape': {
        'name': '도형/이미지',
        'patterns': ['picture', 'shape_obj', 'draw_', 'insert_picture', 'background'],
        'description': '그림, 도형, 개체 관련'
    },
    'page': {
        'name': '페이지/구역',
        'patterns': ['page_', 'section', 'header', 'footer', 'create_page'],
        'description': '페이지, 구역, 머리말/꼬리말 관련'
    },
    'action': {
        'name': 'HAction',
        'patterns': ['run', 'execute', 'get_default', 'create_action'],
        'description': 'HAction 실행 관련'
    },
    'param': {
        'name': '파라미터 헬퍼',
        'patterns': ['h_align', 'v_align', 'line_type', 'rgb_color', 'brush_type',
                     'number_format', 'mili_to', 'hwp_unit'],
        'description': '단위 변환, 파라미터 값 헬퍼'
    },
    'property': {
        'name': '속성',
        'patterns': ['get_version', 'get_cur', 'version', 'path', 'is_'],
        'description': '문서/객체 속성 조회'
    },
    'ctrl': {
        'name': '컨트롤',
        'patterns': ['ctrl', 'get_ctrl', 'cur_ctrl', 'head_ctrl'],
        'description': '컨트롤 객체 관련'
    },
}


def categorize_function(func_name: str) -> str:
    """함수명으로 카테고리 분류"""
    func_lower = func_name.lower()

    for cat_id, cat_info in CATEGORIES.items():
        for pattern in cat_info['patterns']:
            if pattern in func_lower:
                return cat_id

    return 'misc'


def parse_docstring(docstring: str) -> Dict[str, Any]:
    """Google-style docstring 파싱"""
    result = {
        'description': '',
        'args': [],
        'returns': '',
        'examples': [],
        'requires': [],
        'before': [],
        'notes': []
    }

    if not docstring:
        return result

    lines = docstring.strip().split('\n')
    current_section = 'description'
    current_arg = None
    current_content = []

    for line in lines:
        stripped = line.strip()

        # 빈 줄 처리
        if not stripped:
            if current_section == 'description' and current_content:
                result['description'] = ' '.join(current_content).strip()
                current_content = []
            continue

        # 섹션 헤더 확인
        if stripped == 'Args:':
            if current_section == 'description':
                result['description'] = ' '.join(current_content).strip()
            current_section = 'args'
            current_content = []
            current_arg = None
            continue
        elif stripped == 'Returns:':
            current_section = 'returns'
            current_content = []
            continue
        elif stripped == 'Examples:':
            current_section = 'examples'
            current_content = []
            continue
        elif stripped.startswith('@requires:'):
            req = stripped[10:].strip()
            if req:
                result['requires'].append(req)
            continue
        elif stripped.startswith('@before:'):
            before = stripped[8:].strip()
            if before:
                result['before'].append(before)
            continue
        elif stripped.startswith('@note:') or stripped.startswith('Note:'):
            note = stripped.split(':', 1)[1].strip() if ':' in stripped else ''
            if note:
                result['notes'].append(note)
            continue

        # 섹션별 처리
        if current_section == 'description':
            current_content.append(stripped)
        elif current_section == 'args':
            # 인자 정의 (name: description 형식)
            if ':' in stripped and not stripped.startswith('-') and not stripped.startswith('>>>'):
                parts = stripped.split(':', 1)
                arg_name = parts[0].strip()
                arg_desc = parts[1].strip() if len(parts) > 1 else ''
                current_arg = {
                    'name': arg_name,
                    'description': arg_desc,
                    'options': []
                }
                result['args'].append(current_arg)
            elif stripped.startswith('-') and current_arg:
                # 옵션 설명
                current_arg['options'].append(stripped[1:].strip())
                current_arg['description'] += ' ' + stripped
        elif current_section == 'returns':
            current_content.append(stripped)
        elif current_section == 'examples':
            result['examples'].append(stripped)

    # 마지막 섹션 처리
    if current_section == 'description' and current_content:
        result['description'] = ' '.join(current_content).strip()
    elif current_section == 'returns' and current_content:
        result['returns'] = ' '.join(current_content).strip()

    return result


def extract_py_args(block: str) -> List[Dict[str, Any]]:
    """py::arg() 추출"""
    args = []
    pattern = r'py::arg\s*\(\s*"([^"]+)"\s*\)(?:\s*=\s*([^,\)]+))?'

    for match in re.finditer(pattern, block):
        arg_name = match.group(1)
        default = match.group(2)

        # 기본값 정리
        if default:
            default = default.strip()
            # L"..." -> "..."
            if default.startswith('L"'):
                default = default[1:]
            # true/false -> True/False
            if default == 'true':
                default = 'True'
            elif default == 'false':
                default = 'False'

        args.append({
            'name': arg_name,
            'default': default
        })

    return args


def extract_functions(content: str) -> List[Dict[str, Any]]:
    """bindings.cpp에서 함수 정의 추출"""
    functions = []

    # 패턴 1: R"doc(...)doc" 형식의 상세 docstring
    pattern_rdoc = re.compile(
        r'\.def(?:_static)?\s*\(\s*"([^"]+)"'  # 함수명
        r'(.*?)'                                # 중간 내용 (py::arg 등)
        r'R"doc\((.*?)\)doc"\s*\)',             # docstring
        re.DOTALL
    )

    # 패턴 2: 짧은 문자열 형식
    pattern_short = re.compile(
        r'\.def(?:_static)?\s*\(\s*"([^"]+)"'  # 함수명
        r'([^)]*?)'                             # 중간 내용
        r',\s*"([^"]+)"\s*\)',                  # 짧은 설명
        re.DOTALL
    )

    seen_funcs = set()

    # R"doc" 형식 추출
    for match in pattern_rdoc.finditer(content):
        func_name = match.group(1)
        middle = match.group(2)
        docstring = match.group(3)

        if func_name in seen_funcs:
            continue
        seen_funcs.add(func_name)

        # docstring 파싱
        parsed = parse_docstring(docstring)

        # py::arg 추출
        py_args = extract_py_args(middle)

        # Args 병합
        if py_args:
            for py_arg in py_args:
                # 기존 args에서 찾기
                found = False
                for arg in parsed['args']:
                    if arg['name'] == py_arg['name']:
                        if py_arg['default']:
                            arg['default'] = py_arg['default']
                        found = True
                        break
                if not found:
                    parsed['args'].append({
                        'name': py_arg['name'],
                        'description': '',
                        'default': py_arg.get('default'),
                        'options': []
                    })

        functions.append({
            'name': func_name,
            'category': categorize_function(func_name),
            'description': parsed['description'],
            'args': parsed['args'],
            'returns': parsed['returns'],
            'examples': parsed['examples'],
            'requires': parsed['requires'],
            'before': parsed['before'],
            'notes': parsed['notes'],
            'is_detailed': True
        })

    # 짧은 형식 추출
    for match in pattern_short.finditer(content):
        func_name = match.group(1)
        middle = match.group(2)
        short_desc = match.group(3)

        if func_name in seen_funcs:
            continue
        seen_funcs.add(func_name)

        # py::arg 추출
        py_args = extract_py_args(middle)

        functions.append({
            'name': func_name,
            'category': categorize_function(func_name),
            'description': short_desc,
            'args': [{'name': a['name'], 'description': '', 'default': a.get('default'), 'options': []} for a in py_args],
            'returns': '',
            'examples': [],
            'requires': [],
            'before': [],
            'notes': [],
            'is_detailed': False
        })

    return functions


def build_api_json(functions: List[Dict[str, Any]]) -> Dict[str, Any]:
    """API JSON 데이터 구성"""
    # 카테고리별 개수 계산
    category_counts = {}
    for func in functions:
        cat = func['category']
        category_counts[cat] = category_counts.get(cat, 0) + 1

    # 카테고리 정보 구성
    categories = []
    for cat_id, cat_info in CATEGORIES.items():
        if cat_id in category_counts:
            categories.append({
                'id': cat_id,
                'name': cat_info['name'],
                'description': cat_info['description'],
                'count': category_counts[cat_id]
            })

    # misc 카테고리 추가
    if 'misc' in category_counts:
        categories.append({
            'id': 'misc',
            'name': '기타',
            'description': '분류되지 않은 함수',
            'count': category_counts['misc']
        })

    # 카테고리 정렬 (개수 내림차순)
    categories.sort(key=lambda x: x['count'], reverse=True)

    # 함수 정렬 (이름순)
    functions.sort(key=lambda x: x['name'])

    return {
        'version': '1.0.0',
        'generated': datetime.now().isoformat(),
        'total_functions': len(functions),
        'categories': categories,
        'functions': functions
    }


def main():
    """메인 함수"""
    # 경로 설정
    script_dir = Path(__file__).parent
    project_root = script_dir.parent
    bindings_path = project_root / 'src' / 'bindings.cpp'
    output_path = project_root / 'docs' / 'api' / 'data' / 'api.json'

    print(f"Parsing: {bindings_path}")

    # bindings.cpp 읽기
    with open(bindings_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 함수 추출
    functions = extract_functions(content)
    print(f"Extracted {len(functions)} functions")

    # 카테고리별 통계
    categories = {}
    for func in functions:
        cat = func['category']
        categories[cat] = categories.get(cat, 0) + 1

    print("\nCategories:")
    for cat, count in sorted(categories.items(), key=lambda x: -x[1]):
        cat_name = CATEGORIES.get(cat, {}).get('name', cat)
        print(f"  {cat_name}: {count}")

    # JSON 생성
    api_data = build_api_json(functions)

    # 출력 디렉토리 생성
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # JSON 저장
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(api_data, f, ensure_ascii=False, indent=2)

    print(f"\nSaved to: {output_path}")
    print(f"Total functions: {api_data['total_functions']}")


if __name__ == '__main__':
    main()
