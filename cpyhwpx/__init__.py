# -*- coding: utf-8 -*-
"""
cpyhwpx - C++ HWP 자동화 라이브러리 (pyhwpx C++ 포팅)

64-bit Python과 32-bit Python 모두 지원
- 네이티브 모듈 직접 로드 (64-bit/32-bit 빌드 필요)
- 64-bit 네이티브 모듈 없으면 32-bit 브릿지 폴백

사용법:
    import cpyhwpx

    hwp = cpyhwpx.Hwp()
    hwp.initialize()
    hwp.insert_text("Hello, HWP!")
    hwp.save_as("output.hwp")
    hwp.quit()
"""

import sys
import struct
import os

__version__ = "1.1.0"
__author__ = "cpyhwpx"

# 아키텍처 감지
def _is_64bit():
    """64-bit Python인지 확인"""
    return struct.calcsize("P") * 8 == 64

def _is_32bit():
    """32-bit Python인지 확인"""
    return struct.calcsize("P") * 8 == 32

# 네이티브 모듈 변수 초기화
_native_module = None
_using_bridge = False

# 먼저 네이티브 모듈 직접 로드 시도 (64비트/32비트 모두)
try:
    import importlib.util
    _pyd_path = os.path.join(os.path.dirname(__file__), '_native', 'cpyhwpx.pyd')
    _spec = importlib.util.spec_from_file_location("cpyhwpx", _pyd_path)
    _native_module = importlib.util.module_from_spec(_spec)
    _spec.loader.exec_module(_native_module)
    Hwp = _native_module.Hwp
    _using_bridge = False
except Exception as e:
    # 네이티브 로드 실패
    if _is_64bit():
        # 64비트에서 네이티브 로드 실패 시 브릿지 폴백
        try:
            from .bridge import HwpBridgeClient as Hwp
            _using_bridge = True
        except Exception as bridge_e:
            raise ImportError(
                f"cpyhwpx.pyd(64비트)를 로드할 수 없고, 브릿지도 실패했습니다. "
                f"네이티브 에러: {e}, 브릿지 에러: {bridge_e}"
            )
    else:
        # 32비트에서 네이티브 로드 실패 시 에러
        raise ImportError(
            f"cpyhwpx.pyd를 로드할 수 없습니다. "
            f"32-bit Python에서 _native/cpyhwpx.pyd 파일을 확인하세요. 에러: {e}"
        )

# 유틸리티 및 타입 노출 (네이티브 모듈 로드 성공 시)
if _native_module is not None:
    try:
        ViewState = getattr(_native_module, 'ViewState', None)
        MoveID = getattr(_native_module, 'MoveID', None)
        CtrlType = getattr(_native_module, 'CtrlType', None)
        FileFormat = getattr(_native_module, 'FileFormat', None)
        LineStyle = getattr(_native_module, 'LineStyle', None)
        HAlign = getattr(_native_module, 'HAlign', None)
        VAlign = getattr(_native_module, 'VAlign', None)
        HwpPos = getattr(_native_module, 'HwpPos', None)
        CharShape = getattr(_native_module, 'CharShape', None)
        ParaShape = getattr(_native_module, 'ParaShape', None)
        FontPreset = getattr(_native_module, 'FontPreset', None)
        utils = getattr(_native_module, 'utils', None)
        units = getattr(_native_module, 'units', None)
        FontDefs = getattr(_native_module, 'FontDefs', None)
    except Exception:
        ViewState = None
        MoveID = None
        CtrlType = None
        FileFormat = None
        LineStyle = None
        HAlign = None
        VAlign = None
        HwpPos = None
        CharShape = None
        ParaShape = None
        FontPreset = None
        utils = None
        units = None
        FontDefs = None
else:
    # 64-bit에서는 네이티브 타입 접근 불가
    ViewState = None
    MoveID = None
    CtrlType = None
    FileFormat = None
    LineStyle = None
    HAlign = None
    VAlign = None
    HwpPos = None
    CharShape = None
    ParaShape = None
    FontPreset = None
    utils = None
    units = None
    FontDefs = None

# 아키텍처 정보 출력용
def get_architecture_info():
    """현재 아키텍처 정보 반환"""
    bits = struct.calcsize("P") * 8
    return {
        "python_bits": bits,
        "using_bridge": _using_bridge,
        "version": __version__
    }

# 모듈 레벨 docstring
__doc__ = f"""
cpyhwpx - C++ HWP 자동화 라이브러리

버전: {__version__}
아키텍처: {struct.calcsize("P") * 8}-bit Python
브릿지 사용: {_using_bridge if '_using_bridge' in dir() else 'N/A'}

사용법:
    import cpyhwpx
    hwp = cpyhwpx.Hwp()
    hwp.initialize()
    hwp.insert_text("Hello!")
    hwp.quit()
"""

__all__ = [
    'Hwp',
    'get_architecture_info',
    '__version__',
    # Types (may be None in 64-bit)
    'ViewState', 'MoveID', 'CtrlType', 'FileFormat',
    'LineStyle', 'HAlign', 'VAlign',
    'HwpPos', 'CharShape', 'ParaShape', 'FontPreset',
    'utils', 'units', 'FontDefs'
]
