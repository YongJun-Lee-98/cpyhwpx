# -*- coding: utf-8 -*-
"""
cpyhwpx - C++ HWP 자동화 라이브러리 (pyhwpx C++ 포팅)

pyhwpx 라이브러리의 인터페이스를 유지하면서 C++로 구현하여 성능을 개선했습니다.

사용법:
    import cpyhwpx

    hwp = cpyhwpx.Hwp()
    hwp.initialize()
    hwp.insert_text("Hello, HWP!")
    hwp.save_as("output.hwp")
    hwp.quit()
"""

import os
import importlib.util

__version__ = "1.1.0"
__author__ = "cpyhwpx"

# 네이티브 모듈 로드
_pyd_path = os.path.join(os.path.dirname(__file__), '_native', 'cpyhwpx.pyd')
_spec = importlib.util.spec_from_file_location("cpyhwpx", _pyd_path)
_native_module = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_native_module)

# Hwp 클래스 노출
Hwp = _native_module.Hwp

# 타입 및 유틸리티 노출
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

# 아키텍처 정보
def get_architecture_info():
    """현재 아키텍처 정보 반환"""
    import struct
    return {
        "python_bits": struct.calcsize("P") * 8,
        "version": __version__
    }

__all__ = [
    'Hwp',
    'get_architecture_info',
    '__version__',
    # Types
    'ViewState', 'MoveID', 'CtrlType', 'FileFormat',
    'LineStyle', 'HAlign', 'VAlign',
    'HwpPos', 'CharShape', 'ParaShape', 'FontPreset',
    'utils', 'units', 'FontDefs'
]
