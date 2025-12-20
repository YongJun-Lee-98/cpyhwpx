# -*- coding: utf-8 -*-
"""
cpyhwpx_utils.py
pyhwpx 호환 유틸리티 함수들

DataFrame, CSV, Excel 등 다양한 데이터 형식을 지원하는 wrapper 함수 제공
"""

from typing import Union, Tuple, List, Any
import os

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False


def table_from_data(
    hwp,
    data: Union[Any, dict, list, str],  # DataFrame, dict, list, CSV/Excel path
    transpose: bool = False,
    header0: str = "",
    treat_as_char: bool = False,
    header: bool = True,
    index: bool = True,
    cell_fill: Union[bool, Tuple[int, int, int]] = False,
    header_bold: bool = True,
) -> bool:
    """
    pyhwpx 호환 table_from_data wrapper

    다양한 데이터 형식을 2D 리스트로 변환 후 cpyhwpx.table_from_data 호출

    Args:
        hwp: cpyhwpx.Hwp 인스턴스
        data: 테이블 데이터 (DataFrame, dict, list, CSV/Excel 경로)
        transpose: 행/열 전환
        header0: index=True일 때 (1,1) 셀에 들어갈 텍스트
        treat_as_char: 글자처럼 취급
        header: 1행을 제목행으로 설정
        index: DataFrame 인덱스를 첫 열로 포함
        cell_fill: 헤더 배경색 (True=회색, (r,g,b)=RGB)
        header_bold: 헤더에 볼드 적용

    Returns:
        성공 여부

    Examples:
        >>> import cpyhwpx
        >>> from cpyhwpx_utils import table_from_data
        >>>
        >>> hwp = cpyhwpx.Hwp(visible=True)
        >>> hwp.initialize()
        >>>
        >>> # 2D 리스트
        >>> data = [["이름", "나이"], ["홍길동", "30"]]
        >>> table_from_data(hwp, data)
        >>>
        >>> # DataFrame
        >>> import pandas as pd
        >>> df = pd.DataFrame({"이름": ["홍길동"], "나이": [30]})
        >>> table_from_data(hwp, df, cell_fill=True)
        >>>
        >>> # CSV 파일
        >>> table_from_data(hwp, "data.csv", header_bold=True)
    """
    rows: List[List[str]] = []

    # 데이터 타입에 따른 처리
    if isinstance(data, list):
        if data and isinstance(data[0], list):
            # 이미 2D 리스트
            rows = [[str(cell) for cell in row] for row in data]
        elif HAS_PANDAS:
            # 1D 리스트 -> DataFrame 변환
            df = pd.DataFrame(data)
            rows = _dataframe_to_rows(df, transpose, header0, index)
        else:
            # pandas 없으면 단순 변환
            rows = [[str(item) for item in data]]

    elif isinstance(data, dict):
        if HAS_PANDAS:
            df = pd.DataFrame(data)
            rows = _dataframe_to_rows(df, transpose, header0, index)
        else:
            # pandas 없으면 간단한 dict 변환
            keys = list(data.keys())
            rows = [[str(k) for k in keys]]
            max_len = max(len(v) if isinstance(v, (list, tuple)) else 1 for v in data.values())
            for i in range(max_len):
                row = []
                for k in keys:
                    v = data[k]
                    if isinstance(v, (list, tuple)):
                        row.append(str(v[i]) if i < len(v) else "")
                    else:
                        row.append(str(v) if i == 0 else "")
                rows.append(row)

    elif isinstance(data, str):
        # 파일 경로
        if not HAS_PANDAS:
            raise ImportError("CSV/Excel 파일 로드에 pandas가 필요합니다")

        if not os.path.isfile(data):
            raise FileNotFoundError(f"파일을 찾을 수 없습니다: {data}")

        if data.endswith('.csv'):
            df = pd.read_csv(data)
        elif '.xls' in data:
            df = pd.read_excel(data)
        elif data.endswith('.json'):
            df = pd.read_json(data)
        else:
            raise ValueError(f"지원하지 않는 파일 형식: {data}")

        rows = _dataframe_to_rows(df, transpose, header0, index)

    elif HAS_PANDAS and isinstance(data, pd.DataFrame):
        rows = _dataframe_to_rows(data, transpose, header0, index)

    else:
        raise TypeError(f"지원하지 않는 데이터 타입: {type(data)}")

    if not rows:
        return False

    # cell_fill 처리
    r, g, b = (-1, -1, -1)
    if cell_fill:
        if isinstance(cell_fill, tuple) and len(cell_fill) >= 3:
            r, g, b = cell_fill[0], cell_fill[1], cell_fill[2]
        else:
            r, g, b = (217, 217, 217)  # 기본 회색

    # C++ 함수 호출
    return hwp.table_from_data(rows, treat_as_char, header, header_bold, r, g, b)


def _dataframe_to_rows(
    df,
    transpose: bool,
    header0: str,
    index: bool
) -> List[List[str]]:
    """DataFrame을 2D 문자열 리스트로 변환"""
    if transpose:
        df = df.T

    rows = []

    if index:
        # 헤더 행: header0 + 컬럼명
        header_row = [header0] + [str(c) for c in df.columns]
        rows.append(header_row)

        # 데이터 행: 인덱스 + 값
        for idx, row in df.iterrows():
            data_row = [str(idx)] + [str(v) for v in row]
            rows.append(data_row)
    else:
        # 헤더 행: 컬럼명만
        rows.append([str(c) for c in df.columns])

        # 데이터 행: 값만
        for _, row in df.iterrows():
            rows.append([str(v) for v in row])

    return rows


def table_to_df(
    hwp,
    header: bool = True,
) -> Any:  # Returns pd.DataFrame if pandas available, else List[List[str]]
    """
    pyhwpx 호환 table_to_df wrapper

    현재 커서 위치의 테이블 데이터를 DataFrame으로 반환

    Args:
        hwp: cpyhwpx.Hwp 인스턴스
        header: 첫 번째 행을 헤더로 사용할지 여부

    Returns:
        pandas.DataFrame (pandas 있을 경우) 또는 2D 리스트

    Examples:
        >>> import cpyhwpx
        >>> from cpyhwpx_utils import table_to_df
        >>>
        >>> hwp = cpyhwpx.Hwp(visible=True)
        >>> hwp.open("test.hwp")
        >>>
        >>> # 테이블 안에 커서 위치시킨 후
        >>> df = table_to_df(hwp)
        >>> print(df)

    Note:
        HWPML2X XML 파싱 방식으로 병합 셀도 정확히 처리합니다.
    """
    # XML 추출 후 파싱
    rows = _parse_table_xml(hwp.get_table_xml())

    if not rows:
        if HAS_PANDAS:
            return pd.DataFrame()
        return []

    if HAS_PANDAS:
        if header and len(rows) > 0:
            # 첫 행을 컬럼명으로 사용
            columns = rows[0]
            data = rows[1:] if len(rows) > 1 else []
            return pd.DataFrame(data, columns=columns)
        else:
            return pd.DataFrame(rows)
    else:
        return rows


def _parse_table_xml(xml: str) -> List[List[str]]:
    """HWPML2X XML을 파싱하여 2D 리스트로 변환"""
    if not xml:
        return []

    try:
        from xml.etree import ElementTree as ET

        # XML 파싱
        root = ET.fromstring(xml)

        rows = []

        # 테이블 행 찾기 (ROW 태그)
        for row_elem in root.iter():
            if row_elem.tag == 'ROW':
                row = []
                # 셀 찾기 (CELL 태그)
                for cell_elem in row_elem:
                    if cell_elem.tag == 'CELL':
                        # 셀 텍스트 추출
                        cell_text = _extract_cell_text(cell_elem)
                        row.append(cell_text)
                if row:
                    rows.append(row)

        return rows

    except Exception:
        return []


def _extract_cell_text(cell_element) -> str:
    """CELL 엘리먼트에서 텍스트 추출 (CHAR 태그 내용)"""
    texts = []

    # CHAR 태그에서 텍스트 추출
    for elem in cell_element.iter():
        if elem.tag == 'CHAR' and elem.text:
            texts.append(elem.text)

    return ''.join(texts)


def table_to_csv(
    hwp,
    path: str,
    header: bool = True,
    encoding: str = "utf-8-sig",
) -> bool:
    """
    현재 커서 위치의 테이블 데이터를 CSV 파일로 저장

    Args:
        hwp: cpyhwpx.Hwp 인스턴스
        path: 저장할 CSV 파일 경로
        header: 첫 번째 행을 헤더로 사용할지 여부
        encoding: 파일 인코딩 (기본: utf-8-sig for Excel 호환)

    Returns:
        성공 여부
    """
    rows = _parse_table_xml(hwp.get_table_xml())

    if not rows:
        return False

    try:
        import csv
        with open(path, 'w', newline='', encoding=encoding) as f:
            writer = csv.writer(f)
            for row in rows:
                writer.writerow(row)
        return True
    except Exception:
        return False


def table_to_string(
    hwp,
    delimiter: str = "\t",
    line_ending: str = "\n",
) -> str:
    """
    현재 커서 위치의 테이블 데이터를 문자열로 반환

    Args:
        hwp: cpyhwpx.Hwp 인스턴스
        delimiter: 셀 구분자 (기본: 탭)
        line_ending: 행 구분자 (기본: 줄바꿈)

    Returns:
        테이블 데이터 문자열
    """
    rows = _parse_table_xml(hwp.get_table_xml())

    if not rows:
        return ""

    lines = [delimiter.join(row) for row in rows]
    return line_ending.join(lines)
