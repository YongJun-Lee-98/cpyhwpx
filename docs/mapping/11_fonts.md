# 폰트 정의 매핑

## 개요

| 항목 | 내용 |
|------|------|
| 소스 파일 | `pyhwpx/fonts.py` |
| 총 프리셋 수 | **111개** |
| 데이터 구조 | `fonts` 딕셔너리 |
| 포팅 완료 | 0개 (0%) |
| 우선순위 | Low |

## 설명

`fonts.py`는 HWP에서 사용할 수 있는 폰트 프리셋들을 정의한 딕셔너리입니다.
각 프리셋은 언어별 글꼴 이름과 폰트 타입을 지정합니다.

### 데이터 구조

```python
fonts = {
    "프리셋명": {
        "FaceNameHangul": "한글 글꼴명",
        "FaceNameLatin": "영문 글꼴명",
        "FaceNameHanja": "한자 글꼴명",
        "FaceNameJapanese": "일본어 글꼴명",
        "FaceNameOther": "기타 글꼴명",
        "FaceNameSymbol": "기호 글꼴명",
        "FaceNameUser": "사용자 글꼴명",
        "FontTypeHangul": 2,      # 폰트 타입 (TTF=2)
        "FontTypeHanja": 2,
        "FontTypeJapanese": 2,
        "FontTypeLatin": 2,
        "FontTypeOther": 2,
        "FontTypeSymbol": 2,
        "FontTypeUser": 2
    },
    # ...
}
```

### 폰트 타입 값

| 값 | 의미 |
|----|------|
| 0 | Unknown |
| 1 | Raster |
| 2 | TrueType (TTF) |
| 3 | Vector |

---

## 폰트 프리셋 목록

### 1. 기본 한글 폰트 (6개)

| # | 프리셋명 | 한글 글꼴 | 영문 글꼴 | 한자 글꼴 |
|---|---------|----------|----------|----------|
| 1 | `명조` | 명조 | 명조 | 명조 |
| 2 | `고딕` | 고딕 | 고딕 | 명조 |
| 3 | `샘물` | 샘물 | 산세리프 | 한양중고딕 |
| 4 | `필기` | 필기 | 필기 | 한양중고딕 |
| 5 | `신명조` | 한양신명조 | 한양신명조 | 한양신명조 |
| 6 | `견명조` | 한양견명조 | 한양견명조 | 한양신명조 |

### 2. 명조 계열 (10개)

| # | 프리셋명 | 한글 글꼴 | 비고 |
|---|---------|----------|------|
| 1 | `신명조` | 한양신명조 | 기본 신명조 |
| 2 | `신명조 약자` | 한양신명조 | 한자: 신명조 약자 |
| 3 | `신명조 간자` | 한양신명조 | 한자: 신명조 간자 |
| 4 | `견명조` | 한양견명조 | 굵은 명조 |
| 5 | `#세명조` | #세명조 | 구 HWP 폰트 |
| 6 | `#중명조` | #중명조 | 구 HWP 폰트 |
| 7 | `#신명조` | #신명조 | 구 HWP 폰트 |
| 8 | `#태명조` | #태명조 | 구 HWP 폰트 |
| 9 | `#견명조` | #견명조 | 구 HWP 폰트 |
| 10 | `#필명조` | #필명조 | 구 HWP 폰트 |

### 3. 고딕 계열 (12개)

| # | 프리셋명 | 한글 글꼴 | 비고 |
|---|---------|----------|------|
| 1 | `중고딕` | 한양중고딕 | 기본 중고딕 |
| 2 | `중고딕 약자` | 한양중고딕 | 한자: 중고딕 약자 |
| 3 | `중고딕 간자` | 한양중고딕 | 한자: 중고딕 간자 |
| 4 | `견고딕` | 한양견고딕 | 굵은 고딕 |
| 5 | `#세고딕` | #세고딕 | 구 HWP 폰트 |
| 6 | `#중고딕` | #중고딕 | 구 HWP 폰트 |
| 7 | `#신중고딕` | #신중고딕 | 구 HWP 폰트 |
| 8 | `#태고딕` | #태고딕 | 구 HWP 폰트 |
| 9 | `#견고딕` | #견고딕 | 구 HWP 폰트 |
| 10 | `#신세고딕` | #신세고딕 | 구 HWP 폰트 |
| 11 | `#신문견고` | #신문견고 | 구 HWP 폰트 |
| 12 | `#환견고딕` | #환견고딕 | 구 HWP 폰트 |

### 4. 궁서 계열 (5개)

| # | 프리셋명 | 한글 글꼴 | 비고 |
|---|---------|----------|------|
| 1 | `궁서` | 한양궁서 | 기본 궁서 |
| 2 | `해서 약자` | 한양궁서 | 한자: 해서 약자 |
| 3 | `해서 간자` | 한양궁서 | 한자: 해서 간자 |
| 4 | `#궁서` | #궁서 | 구 HWP 폰트 |
| 5 | `#필궁서` | #필궁서 | 구 HWP 폰트 |

### 5. 그래픽/디자인 폰트 (8개)

| # | 프리셋명 | 한글 글꼴 | 비고 |
|---|---------|----------|------|
| 1 | `그래픽` | 한양그래픽 | 기본 그래픽 |
| 2 | `헤드라인` | 태 헤드라인T | 제목용 |
| 3 | `가는헤드라인` | 태 가는 헤드라인T | 가는 제목 |
| 4 | `#그래픽` | #그래픽 | 구 HWP 폰트 |
| 5 | `#신그래픽` | #신그래픽 | 구 HWP 폰트 |
| 6 | `#태그래픽` | #태그래픽 | 구 HWP 폰트 |
| 7 | `#공작` | #공작 | 구 HWP 폰트 |
| 8 | `#빅` | #빅 | 구 HWP 폰트 |

### 6. 손글씨/펜 폰트 (6개)

| # | 프리셋명 | 한글 글꼴 | 비고 |
|---|---------|----------|------|
| 1 | `펜흘림` | 펜흘림 | 펜글씨 |
| 2 | `타이프` | 타이프 | 타자기 스타일 |
| 3 | `#필기` | #필기 | 구 HWP 폰트 |
| 4 | `#산뜻필기` | #산뜻필기 | 구 HWP 폰트 |
| 5 | `#줄필기` | #줄필기 | 구 HWP 폰트 |
| 6 | `#판본` | #판본 | 구 HWP 폰트 |

### 7. 공한 계열 (6개)

| # | 프리셋명 | 한글 글꼴 | 비고 |
|---|---------|----------|------|
| 1 | `가는공한` | 가는공한 | 가는 공한 |
| 2 | `중간공한` | 중간공한 | 중간 공한 |
| 3 | `굵은공한` | 굵은공한 | 굵은 공한 |
| 4 | `가는한` | 가는한 | 가는 한 |
| 5 | `중간한` | 중간한 | 중간 한 |
| 6 | `굵은한` | 굵은한 | 굵은 한 |

### 8. 과일/채소 폰트 (7개)

| # | 프리셋명 | 한글 글꼴 | 비고 |
|---|---------|----------|------|
| 1 | `복숭아` | 복숭아 | 귀여운 폰트 |
| 2 | `옥수수` | 옥수수 | 귀여운 폰트 |
| 3 | `오이` | 오이 | 귀여운 폰트 |
| 4 | `가지` | 가지 | 귀여운 폰트 |
| 5 | `강낭콩` | 강낭콩 | 귀여운 폰트 |
| 6 | `딸기` | 딸기 | 귀여운 폰트 |
| 7 | `#수암A` / `#수암B` | #수암A / #수암B | 구 HWP 폰트 |

### 9. 나루 계열 (4개)

| # | 프리셋명 | 한글 글꼴 | 비고 |
|---|---------|----------|------|
| 1 | `#세나루` | #세나루 | 구 HWP 폰트 |
| 2 | `#신세나루` | #신세나루 | 구 HWP 폰트 |
| 3 | `#디나루` | #디나루 | 구 HWP 폰트 |
| 4 | `#신디나루` | #신디나루 | 구 HWP 폰트 |

### 10. 시스템 폰트 (3개)

| # | 프리셋명 | 한글 글꼴 | 비고 |
|---|---------|----------|------|
| 1 | `시스템` | 시스템 | 기본 시스템 |
| 2 | `시스템 약자` | 시스템 | 한자: 시스템 약자 |
| 3 | `시스템 간자` | 시스템 | 한자: 시스템 간자 |

### 11. HY 폰트 (1개)

| # | 프리셋명 | 한글 글꼴 | 비고 |
|---|---------|----------|------|
| 1 | `HY둥근고딕` | HY둥근고딕 | 둥근 고딕 |

### 12. 기타 구 HWP 폰트 (40+개)

구 HWP 버전 호환용 폰트들 (`#` 접두사):
- `#굵은필기`, `#꽃분첩`, `#도시`, `#동물나라`, `#들꽃`, `#박하`
- `#바다`, `#샘물`, `#세명조`, `#세송`, `#옻칠`, `#팔분`
- `#필운`, `#해운`, `#흘림`, `#매직` 등

---

## 전체 프리셋 목록 (알파벳순)

```
#HY신명조, #견고딕, #견명조, #공작, #그래픽, #궁서, #굵은필기
#꽃분첩, #디나루, #도시, #동물나라, #들꽃, #매직, #바다
#박하, #빅, #샘물, #산뜻필기, #세고딕, #세나루, #세명조, #세송
#수암A, #수암B, #신그래픽, #신디나루, #신명조, #신문견고
#신세고딕, #신세나루, #신중고딕, #옻칠, #줄필기, #중고딕, #중명조
#태고딕, #태그래픽, #태명조, #판본, #팔분, #필궁서, #필기, #필명조
#필운, #해운, #환견고딕, #흘림
HY둥근고딕
가는공한, 가는한, 가는헤드라인, 가지, 강낭콩, 견고딕, 견명조
고딕, 그래픽, 굵은공한, 굵은한, 궁서
딸기
명조
복숭아
샘물, 시스템, 시스템 간자, 시스템 약자, 신명조, 신명조 간자, 신명조 약자
오이, 옥수수
중간공한, 중간한, 중고딕, 중고딕 간자, 중고딕 약자
타이프
펜흘림, 필기
해서 간자, 해서 약자, 헤드라인
```

---

## C++ 구현

### FontDefs.h
```cpp
#pragma once
#include <string>
#include <map>
#include <Windows.h>

struct FontPreset {
    std::wstring FaceNameHangul;
    std::wstring FaceNameLatin;
    std::wstring FaceNameHanja;
    std::wstring FaceNameJapanese;
    std::wstring FaceNameOther;
    std::wstring FaceNameSymbol;
    std::wstring FaceNameUser;
    int FontTypeHangul;
    int FontTypeHanja;
    int FontTypeJapanese;
    int FontTypeLatin;
    int FontTypeOther;
    int FontTypeSymbol;
    int FontTypeUser;
};

// 폰트 프리셋 딕셔너리
extern std::map<std::wstring, FontPreset> g_FontPresets;

// 초기화 함수
void InitializeFontPresets();

// 프리셋 조회
const FontPreset* GetFontPreset(const std::wstring& name);
```

### FontDefs.cpp
```cpp
#include "FontDefs.h"

std::map<std::wstring, FontPreset> g_FontPresets;

void InitializeFontPresets() {
    // 명조
    g_FontPresets[L"명조"] = {
        L"명조", L"명조", L"명조", L"명조", L"명조", L"명조", L"명조",
        2, 2, 2, 2, 2, 2, 2
    };

    // 고딕
    g_FontPresets[L"고딕"] = {
        L"고딕", L"고딕", L"명조", L"고딕", L"명조", L"명조", L"명조",
        2, 2, 2, 2, 2, 2, 2
    };

    // 신명조
    g_FontPresets[L"신명조"] = {
        L"한양신명조", L"한양신명조", L"한양신명조", L"한양신명조",
        L"한양신명조", L"한양신명조", L"명조",
        2, 2, 2, 2, 2, 2, 2
    };

    // 중고딕
    g_FontPresets[L"중고딕"] = {
        L"한양중고딕", L"한양중고딕", L"한양중고딕", L"한양중고딕",
        L"한양신명조", L"한양중고딕", L"명조",
        2, 2, 2, 2, 2, 2, 2
    };

    // ... 111개 프리셋 정의
}

const FontPreset* GetFontPreset(const std::wstring& name) {
    auto it = g_FontPresets.find(name);
    if (it != g_FontPresets.end()) {
        return &it->second;
    }
    return nullptr;
}
```

### pybind11 바인딩
```cpp
#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
namespace py = pybind11;

PYBIND11_MODULE(cpyhwpx, m) {
    // FontPreset 구조체 노출
    py::class_<FontPreset>(m, "FontPreset")
        .def_readonly("FaceNameHangul", &FontPreset::FaceNameHangul)
        .def_readonly("FaceNameLatin", &FontPreset::FaceNameLatin)
        .def_readonly("FaceNameHanja", &FontPreset::FaceNameHanja)
        .def_readonly("FaceNameJapanese", &FontPreset::FaceNameJapanese)
        .def_readonly("FaceNameOther", &FontPreset::FaceNameOther)
        .def_readonly("FaceNameSymbol", &FontPreset::FaceNameSymbol)
        .def_readonly("FaceNameUser", &FontPreset::FaceNameUser)
        .def_readonly("FontTypeHangul", &FontPreset::FontTypeHangul)
        // ... 나머지 필드
        ;

    // fonts 딕셔너리 노출
    m.def("get_font_preset", &GetFontPreset,
          py::return_value_policy::reference);

    // 전체 폰트 목록
    m.def("get_all_font_presets", []() {
        return g_FontPresets;
    });
}
```

---

## 사용 예시

### Python (pyhwpx)
```python
from pyhwpx.fonts import fonts

# 폰트 프리셋 조회
myeongjo = fonts["명조"]
print(myeongjo["FaceNameHangul"])  # "명조"

# 글자 모양에 적용
hwp.set_charshape(FaceNameHangul=myeongjo["FaceNameHangul"],
                  FontTypeHangul=myeongjo["FontTypeHangul"])
```

### C++ (cpyhwpx - 목표)
```python
import cpyhwpx

# 폰트 프리셋 조회
preset = cpyhwpx.get_font_preset("명조")
print(preset.FaceNameHangul)  # "명조"

# 글자 모양에 적용
hwp.set_charshape(FaceNameHangul=preset.FaceNameHangul)
```

---

## 포팅 체크리스트

- [ ] FontPreset 구조체 정의
- [ ] g_FontPresets 맵 정의
- [ ] 111개 프리셋 데이터 입력
- [ ] InitializeFontPresets() 구현
- [ ] GetFontPreset() 구현
- [ ] pybind11 바인딩
- [ ] 테스트

---

## 참고사항

1. **#접두사 폰트**: `#`으로 시작하는 폰트는 구 HWP 버전 (한/글 2.x, 3.x) 호환용 폰트입니다.
2. **FontType=2**: 대부분의 폰트가 TrueType(TTF)으로 설정되어 있습니다.
3. **약자/간자**: 일부 폰트는 한자의 약자/간자체 버전이 별도로 제공됩니다.
4. **시스템 의존성**: 일부 폰트는 시스템에 설치되어 있어야 정상 작동합니다.
