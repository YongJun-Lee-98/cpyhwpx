# -*- coding: utf-8 -*-
"""
cpyhwpx 32-bit/64-bit 브릿지
subprocess + multiprocessing.connection을 사용한 간단한 IPC 브릿지
"""

import os
import sys
import subprocess
import pickle
import struct
import time
import threading
import socket
import winreg
from multiprocessing.connection import Listener, Client


# =============================================================================
# 레지스트리 헬퍼 함수
# =============================================================================

REG_PATHS = [
    r"Software\HNC\HwpUserAction\Modules",
    r"Software\Hnc\HwpUserAction\Modules",
]


def _check_registry_key(key_name="FilePathCheckerModule"):
    """레지스트리에 보안모듈이 등록되어 있는지 확인"""
    for reg_path in REG_PATHS:
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_READ)
            value, _ = winreg.QueryValueEx(key, key_name)
            winreg.CloseKey(key)
            return value  # 등록된 DLL 경로 반환
        except FileNotFoundError:
            continue
        except Exception:
            continue
    return None


def _register_to_registry(dll_path, key_name="FilePathCheckerModule"):
    """레지스트리에 보안모듈 등록"""
    for reg_path in REG_PATHS:
        try:
            # 상위 키 열기/생성
            try:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, reg_path, 0, winreg.KEY_WRITE)
            except FileNotFoundError:
                key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, reg_path)

            # 값 설정
            winreg.SetValueEx(key, key_name, 0, winreg.REG_SZ, dll_path)
            winreg.CloseKey(key)
            return True
        except Exception:
            continue
    return False


def _get_default_dll_path():
    """기본 DLL 경로 반환"""
    # _native 폴더 내의 DLL 찾기
    native_path = os.path.join(os.path.dirname(__file__), '_native')
    dll_path = os.path.join(native_path, 'FilePathCheckerModule.dll')
    if os.path.exists(dll_path):
        return dll_path
    return None


def _find_python32():
    """32-bit Python 인터프리터 경로 찾기"""
    import subprocess as sp
    try:
        result = sp.run(
            ['py', '-3.12-32', '-c', 'import sys; print(sys.executable)'],
            capture_output=True,
            text=True,
            timeout=10
        )
        if result.returncode == 0:
            return result.stdout.strip()
    except:
        pass

    # 일반적인 설치 경로 확인
    possible_paths = [
        r'C:\Users\leep0\AppData\Local\Programs\Python\Python312-32\python.exe',
        r'C:\Python312-32\python.exe',
        r'C:\Program Files (x86)\Python312-32\python.exe',
    ]

    for path in possible_paths:
        if os.path.exists(path):
            return path

    return None


class HwpBridgeServer:
    """32-bit Python에서 실행되는 HWP 서버"""

    def __init__(self, port):
        self.port = port
        self._hwp = None
        self._cpyhwpx = None

    def start(self):
        """서버 시작"""
        # cpyhwpx.pyd 모듈 직접 로드 (importlib.util 사용)
        import importlib.util
        _pyd_path = os.path.join(os.path.dirname(__file__), '_native', 'cpyhwpx.pyd')
        _spec = importlib.util.spec_from_file_location("cpyhwpx", _pyd_path)
        cpyhwpx = importlib.util.module_from_spec(_spec)
        _spec.loader.exec_module(cpyhwpx)
        self._cpyhwpx = cpyhwpx

        # 연결 대기
        address = ('localhost', self.port)
        listener = Listener(address, authkey=b'cpyhwpx-bridge')

        print(f"[Server] Listening on port {self.port}...")
        sys.stdout.flush()

        while True:
            try:
                conn = listener.accept()
                print(f"[Server] Connection accepted")
                sys.stdout.flush()

                while True:
                    try:
                        msg = conn.recv()
                        if msg is None:
                            break

                        method, args, kwargs = msg
                        result = self._handle_request(method, args, kwargs)
                        conn.send(('ok', result))
                    except EOFError:
                        break
                    except Exception as e:
                        conn.send(('error', str(e)))

                conn.close()
            except Exception as e:
                print(f"[Server] Error: {e}")
                sys.stdout.flush()

    def _handle_request(self, method, args, kwargs):
        """요청 처리"""
        print(f"[Server] Method: {method}, Args: {args}, Kwargs: {kwargs}")
        sys.stdout.flush()

        try:
            if method == 'create_hwp':
                self._hwp = self._cpyhwpx.Hwp(*args, **kwargs)
                result = True
            elif method == 'destroy_hwp':
                self._hwp = None
                result = True
            elif method == 'shutdown':
                result = 'shutdown'
            elif hasattr(self._hwp, method):
                attr = getattr(self._hwp, method)
                if callable(attr):
                    result = attr(*args, **kwargs)
                else:
                    result = attr
            else:
                raise AttributeError(f"Unknown method: {method}")

            print(f"[Server] Result: {result}")
            sys.stdout.flush()
            return result
        except Exception as e:
            print(f"[Server] Error in {method}: {e}")
            sys.stdout.flush()
            raise


class HwpBridgeClient:
    """64-bit Python에서 32-bit 서버와 통신하는 클라이언트"""

    def __init__(self, visible=True, new_instance=True):
        self._port = self._find_free_port()
        self._process = None
        self._conn = None

        # 32-bit Python 찾기
        python32 = _find_python32()
        if not python32:
            raise RuntimeError("32-bit Python을 찾을 수 없습니다.")

        # 서버 스크립트 경로
        server_script = os.path.join(os.path.dirname(__file__), '_server_runner.py')

        # 서버 스크립트 생성
        self._create_server_script(server_script)

        # 서버 프로세스 시작
        self._process = subprocess.Popen(
            [python32, server_script, str(self._port)],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
        )

        # 서버 연결 대기
        time.sleep(1)
        self._connect()

        # Hwp 인스턴스 생성
        self._call('create_hwp', visible, new_instance)

    def _find_free_port(self):
        """사용 가능한 포트 찾기"""
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            s.bind(('', 0))
            return s.getsockname()[1]

    def _create_server_script(self, path):
        """서버 실행 스크립트 생성"""
        script = '''# -*- coding: utf-8 -*-
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
'''
        with open(path, 'w', encoding='utf-8') as f:
            f.write(script)

    def _connect(self, timeout=10):
        """서버에 연결"""
        address = ('localhost', self._port)
        start = time.time()
        while time.time() - start < timeout:
            try:
                self._conn = Client(address, authkey=b'cpyhwpx-bridge')
                return
            except:
                time.sleep(0.1)
        raise ConnectionError(f"서버 연결 실패 (포트: {self._port})")

    def _call(self, method, *args, **kwargs):
        """서버에 메서드 호출 요청"""
        self._conn.send((method, args, kwargs))
        status, result = self._conn.recv()
        if status == 'error':
            raise RuntimeError(result)
        return result

    def __del__(self):
        """정리"""
        try:
            if self._conn:
                self._conn.send(('shutdown', (), {}))
                self._conn.close()
            if self._process:
                self._process.terminate()
        except:
            pass

    # =========================================================================
    # HWP API 메서드들
    # =========================================================================

    def initialize(self):
        return self._call('initialize')

    def register_module(self, module_type="FilePathCheckDLL", module_data=None):
        """
        보안 모듈 등록

        Args:
            module_type: 모듈 유형 (기본값: "FilePathCheckDLL")
            module_data: DLL 경로 또는 레지스트리 키 이름
                - None: 기본 DLL 사용 (_native/FilePathCheckerModule.dll)
                - 파일 경로: 레지스트리에 등록 후 사용
                - 키 이름: 이미 등록된 레지스트리 키 사용

        Returns:
            bool: 성공 여부
        """
        key_name = "FilePathCheckerModule"

        # module_data가 None이면 기본 DLL 사용
        if module_data is None:
            dll_path = _get_default_dll_path()
            if dll_path:
                module_data = dll_path
            else:
                # 레지스트리에 이미 등록되어 있는지 확인
                if _check_registry_key(key_name):
                    return self._call('register_module', module_type, key_name)
                return False

        # module_data가 파일 경로인 경우 (경로 구분자 포함)
        if '\\' in module_data or '/' in module_data:
            # 파일 존재 확인
            if not os.path.exists(module_data):
                return False

            # 레지스트리에 등록
            if not _register_to_registry(module_data, key_name):
                return False

            # 레지스트리 키 이름으로 호출
            return self._call('register_module', module_type, key_name)
        else:
            # module_data가 레지스트리 키 이름인 경우
            return self._call('register_module', module_type, module_data)

    def quit(self, save=False):
        return self._call('quit', save)

    def is_initialized(self):
        return self._call('is_initialized')

    def open(self, filename, format_="", arg=""):
        return self._call('open', filename, format_, arg)

    def save(self, save_if_dirty=True):
        return self._call('save', save_if_dirty)

    def save_as(self, filename, format_="HWP", arg=""):
        return self._call('save_as', filename, format_, arg)

    def clear(self, option=1):
        return self._call('clear', option)

    def close(self, is_dirty=False):
        return self._call('close', is_dirty)

    def insert_text(self, text):
        return self._call('insert_text', text)

    def get_text(self):
        return self._call('get_text')

    def get_selected_text(self, keep_select=False):
        return self._call('get_selected_text', keep_select)

    def get_pos(self):
        return self._call('get_pos')

    def set_pos(self, list_id, para, pos):
        return self._call('set_pos', list_id, para, pos)

    def move_pos(self, move_id, para=0, pos=0):
        return self._call('move_pos', move_id, para, pos)

    def set_visible(self, visible):
        return self._call('set_visible', visible)

    def maximize_window(self):
        return self._call('maximize_window')

    def minimize_window(self):
        return self._call('minimize_window')

    def is_empty(self):
        return self._call('is_empty')

    def is_modified(self):
        return self._call('is_modified')

    def is_cell(self):
        return self._call('is_cell')

    def find(self, text, forward=True, match_case=False, regex=False, replace_mode=False):
        return self._call('find', text, forward, match_case, regex, replace_mode)

    def replace(self, find_text, replace_text, forward=True, match_case=False, regex=False):
        return self._call('replace', find_text, replace_text, forward, match_case, regex)

    def replace_all(self, find_text, replace_text, match_case=False, regex=False):
        return self._call('replace_all', find_text, replace_text, match_case, regex)

    def run(self, action_name):
        return self._call('run', action_name)

    @property
    def version(self):
        return self._call('version')

    @property
    def build_number(self):
        return self._call('build_number')

    @property
    def current_page(self):
        return self._call('current_page')

    @property
    def page_count(self):
        return self._call('page_count')

    @property
    def edit_mode(self):
        return self._call('edit_mode')

    @edit_mode.setter
    def edit_mode(self, mode):
        self._call('edit_mode', mode)


# Alias
Hwp = HwpBridgeClient
