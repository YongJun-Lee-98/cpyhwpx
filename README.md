# cpyhwpx

**pyhwpx를 C++로 포팅한 Python 확장 모듈**

한/글(HWP) 자동화를 위한 C++ Python 확장 모듈입니다.
기존 pyhwpx 라이브러리의 인터페이스를 유지하면서 C++로 구현하여 성능을 개선했습니다.

## 요구 사항

- Windows 10/11 (64-bit)
- Python 3.8 이상 (64-bit)
- 한/글(HWP) 설치
- Visual Studio 2019 이상 (빌드 시)

## 설치

### pip 설치

```bash
pip install cpyhwpx
```

### 소스에서 빌드

```bash
# 의존성 설치
pip install pybind11 setuptools wheel

# 빌드 및 설치
pip install .
```

### CMake로 빌드

```bash
mkdir build
cd build
cmake .. -A x64
cmake --build . --config Release
```

## 사용법

### 기본 사용

```python
import cpyhwpx

# HWP 객체 생성
hwp = cpyhwpx.Hwp()

# 초기화
hwp.initialize()

# 보안 모듈 등록 (선택사항)
hwp.register_module("FilePathCheckDLL", "path/to/module.dll")

# 파일 열기
hwp.open("test.hwp")

# 텍스트 삽입
hwp.insert_text("Hello, cpyhwpx!")

# 저장
hwp.save_as("output.hwp")

# 종료
hwp.quit()
```

### 유틸리티 함수

```python
import cpyhwpx

# 셀 주소 변환
row, col = cpyhwpx.utils.addr_to_tuple("A1")  # (1, 1)
addr = cpyhwpx.utils.tuple_to_addr(1, 1)      # "A1"

# 단위 변환
unit = cpyhwpx.units.from_mm(25.4)  # 7200 (1 inch)
mm = cpyhwpx.units.to_mm(7200)       # 25.4
```

### 폰트 프리셋

```python
import cpyhwpx

# 폰트 프리셋 가져오기
font = cpyhwpx.FontDefs.get_preset("맑은고딕")
print(font.FaceNameHangul)  # "맑은 고딕"

# 사용 가능한 프리셋 목록
names = cpyhwpx.FontDefs.get_preset_names()
```

## API 참조

### Hwp 클래스

| 메서드 | 설명 |
|--------|------|
| `initialize()` | HWP COM 객체 초기화 |
| `register_module(type, path)` | 보안 모듈 등록 |
| `open(filename, format, arg)` | 파일 열기 |
| `save(save_if_dirty)` | 파일 저장 |
| `save_as(filename, format, arg)` | 다른 이름으로 저장 |
| `close(is_dirty)` | 문서 닫기 |
| `quit(save)` | HWP 종료 |
| `insert_text(text)` | 텍스트 삽입 |
| `get_text()` | 텍스트 가져오기 |
| `get_pos()` | 현재 위치 가져오기 |
| `set_pos(list, para, pos)` | 위치 설정 |
| `move_pos(move_id, para, pos)` | 위치 이동 |
| `run(action_name)` | HAction.Run() 실행 |
| `find(text, ...)` | 텍스트 찾기 |
| `replace(find, replace, ...)` | 텍스트 바꾸기 |

### 열거형

| 열거형 | 값 |
|--------|-----|
| `ViewState` | Normal, Page, Text, Outline, ... |
| `MoveID` | Top, Bottom, ParaBegin, ParaEnd, ... |
| `CtrlType` | Table, Picture, Equation, ... |
| `HAlign` | Left, Center, Right, Justify |
| `VAlign` | Top, Center, Bottom |
| `LineStyle` | Solid, Dash, Dot, ... |

### 단위 상수

| 상수 | 값 | 설명 |
|------|-----|------|
| `HWPUNIT_PER_MM` | 283.46 | 1mm당 HwpUnit |
| `HWPUNIT_PER_CM` | 2834.6 | 1cm당 HwpUnit |
| `HWPUNIT_PER_INCH` | 7200 | 1인치당 HwpUnit |
| `HWPUNIT_PER_PT` | 100 | 1포인트당 HwpUnit |

## pyhwpx와의 호환성

cpyhwpx는 pyhwpx와 동일한 인터페이스를 제공합니다.
기존 pyhwpx 코드를 쉽게 마이그레이션할 수 있습니다.

```python
# pyhwpx
from pyhwpx import Hwp
hwp = Hwp()

# cpyhwpx (동일한 인터페이스)
import cpyhwpx
hwp = cpyhwpx.Hwp()
```

## 라이선스

MIT License with Additional Restrictions

- 자유롭게 사용, 수정, 배포 가능
- **원본 그대로 상업적 판매 금지**
- 수정/통합한 파생 저작물 판매는 허용

## 참고

- [pyhwpx](https://github.com/martiniifun/pyhwpx) - 원본 Python 라이브러리 (cpyhwpx는 이를 C++로 포팅)
- [한/글 자동화 API](https://www.hancom.com) - 한컴오피스 공식 문서
