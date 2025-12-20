# -*- coding: utf-8 -*-
import sys
import os

# cpyhwpx 패키지 폴더를 경로에 추가 (bridge.py를 찾기 위함)
pkg_path = os.path.dirname(os.path.abspath(__file__))
if pkg_path not in sys.path:
    sys.path.insert(0, pkg_path)

# cpyhwpx 패키지의 부모 폴더도 추가 (패키지 임포트 지원)
parent_path = os.path.dirname(pkg_path)
if parent_path not in sys.path:
    sys.path.insert(0, parent_path)

from bridge import HwpBridgeServer

if __name__ == '__main__':
    port = int(sys.argv[1])
    server = HwpBridgeServer(port)
    server.start()
