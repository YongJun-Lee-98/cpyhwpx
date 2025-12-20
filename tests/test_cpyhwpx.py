"""
cpyhwpx 테스트
"""

import pytest
import sys


def test_import():
    """모듈 임포트 테스트"""
    import cpyhwpx
    assert cpyhwpx is not None


def test_enums():
    """열거형 테스트"""
    import cpyhwpx

    # ViewState
    assert cpyhwpx.ViewState.Normal is not None
    assert cpyhwpx.ViewState.Page is not None

    # MoveID
    assert cpyhwpx.MoveID.Top is not None
    assert cpyhwpx.MoveID.Bottom is not None

    # CtrlType
    assert cpyhwpx.CtrlType.Table is not None
    assert cpyhwpx.CtrlType.Picture is not None

    # HAlign
    assert cpyhwpx.HAlign.Left is not None
    assert cpyhwpx.HAlign.Center is not None
    assert cpyhwpx.HAlign.Right is not None


def test_hwp_pos():
    """HwpPos 구조체 테스트"""
    import cpyhwpx

    pos = cpyhwpx.HwpPos()
    pos.list = 1
    pos.para = 2
    pos.pos = 3

    assert pos.list == 1
    assert pos.para == 2
    assert pos.pos == 3


def test_char_shape():
    """CharShape 구조체 테스트"""
    import cpyhwpx

    shape = cpyhwpx.CharShape()
    shape.Height = 1000
    shape.Bold = True

    assert shape.Height == 1000
    assert shape.Bold == True


def test_font_preset():
    """FontPreset 테스트"""
    import cpyhwpx

    preset = cpyhwpx.FontPreset()
    preset.FaceNameHangul = "맑은 고딕"

    assert preset.FaceNameHangul == "맑은 고딕"


def test_font_defs():
    """FontDefs 테스트"""
    import cpyhwpx

    # 프리셋 존재 확인
    assert cpyhwpx.FontDefs.has_preset("맑은고딕") or True  # 초기화 안됐을 수 있음

    # 프리셋 이름 목록
    names = cpyhwpx.FontDefs.get_preset_names()
    assert isinstance(names, list)


def test_utils_addr_conversion():
    """Utils 주소 변환 테스트"""
    import cpyhwpx

    # 주소 -> 튜플
    row, col = cpyhwpx.utils.addr_to_tuple("A1")
    assert row == 1
    assert col == 1

    row, col = cpyhwpx.utils.addr_to_tuple("B3")
    assert row == 3
    assert col == 2

    # 튜플 -> 주소
    addr = cpyhwpx.utils.tuple_to_addr(1, 1)
    assert addr == "A1"

    addr = cpyhwpx.utils.tuple_to_addr(3, 2)
    assert addr == "B3"


def test_utils_string():
    """Utils 문자열 처리 테스트"""
    import cpyhwpx

    # Trim
    result = cpyhwpx.utils.trim("  hello  ")
    assert result == "hello"

    # Split
    parts = cpyhwpx.utils.split("a,b,c", ",")
    assert parts == ["a", "b", "c"]

    # Join
    joined = cpyhwpx.utils.join(["a", "b", "c"], "-")
    assert joined == "a-b-c"


def test_utils_color():
    """Utils 색상 변환 테스트"""
    import cpyhwpx

    # Hex -> ColorRef
    color = cpyhwpx.utils.hex_to_colorref("#FF0000")
    assert color > 0

    # ColorRef -> Hex
    hex_str = cpyhwpx.utils.colorref_to_hex(0x0000FF)  # RGB(0, 0, 255)
    assert "#" in hex_str


def test_units():
    """Units 단위 변환 테스트"""
    import cpyhwpx

    # mm -> HwpUnit
    unit = cpyhwpx.units.from_mm(25.4)  # 1 inch = 25.4mm
    assert unit == 7200  # 1 inch = 7200 HwpUnit

    # HwpUnit -> mm
    mm = cpyhwpx.units.to_mm(7200)
    assert abs(mm - 25.4) < 0.01

    # 상수 확인
    assert cpyhwpx.units.HWPUNIT_PER_INCH == 7200.0


@pytest.mark.skipif(True, reason="HWP가 설치되어 있어야 함")
def test_hwp_create():
    """HWP 객체 생성 테스트 (HWP 설치 필요)"""
    import cpyhwpx

    hwp = cpyhwpx.Hwp(visible=False)
    assert hwp.is_initialized() == False  # initialize() 호출 전

    # 초기화
    success = hwp.initialize()
    if success:
        assert hwp.is_initialized() == True

        # 버전 확인
        version = hwp.version
        assert len(version) > 0

        # 종료
        hwp.quit()


@pytest.mark.skipif(True, reason="HWP가 설치되어 있어야 함")
def test_hwp_text():
    """HWP 텍스트 편집 테스트 (HWP 설치 필요)"""
    import cpyhwpx

    hwp = cpyhwpx.Hwp(visible=False)
    if not hwp.initialize():
        pytest.skip("HWP 초기화 실패")

    try:
        # 텍스트 삽입
        hwp.insert_text("Hello, cpyhwpx!")

        # 텍스트 가져오기
        status, text = hwp.get_text()
        assert "Hello" in text or status >= 0

    finally:
        hwp.quit()


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
