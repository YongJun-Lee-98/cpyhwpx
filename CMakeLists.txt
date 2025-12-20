cmake_minimum_required(VERSION 3.15)
project(cpyhwpx LANGUAGES CXX)

#==============================================================================
# 빌드 설정
#==============================================================================

# C++17 표준 사용
set(CMAKE_CXX_STANDARD 17)
set(CMAKE_CXX_STANDARD_REQUIRED ON)
set(CMAKE_CXX_EXTENSIONS OFF)

# 32bit 빌드 강제 (HWP가 32bit이므로)
if(CMAKE_GENERATOR_PLATFORM)
    message(STATUS "Generator platform: ${CMAKE_GENERATOR_PLATFORM}")
else()
    set(CMAKE_GENERATOR_PLATFORM "Win32" CACHE STRING "Build platform" FORCE)
endif()

# Release 빌드 기본값
if(NOT CMAKE_BUILD_TYPE)
    set(CMAKE_BUILD_TYPE Release)
endif()

# Windows 설정
if(WIN32)
    add_definitions(-D_UNICODE -DUNICODE)
    add_definitions(-DWIN32_LEAN_AND_MEAN)
    add_definitions(-DNOMINMAX)
endif()

#==============================================================================
# pybind11 설정
#==============================================================================

# pybind11 찾기 (pip install pybind11 또는 서브모듈)
find_package(pybind11 CONFIG)

if(NOT pybind11_FOUND)
    # 서브모듈로 pybind11 사용
    if(EXISTS "${CMAKE_CURRENT_SOURCE_DIR}/extern/pybind11/CMakeLists.txt")
        add_subdirectory(extern/pybind11)
    else()
        message(FATAL_ERROR "pybind11 not found. Install via: pip install pybind11")
    endif()
endif()

#==============================================================================
# 소스 파일
#==============================================================================

set(CPYHWPX_SOURCES
    src/HwpWrapper.cpp
    src/HwpCtrl.cpp
    src/HwpAction.cpp
    src/HwpParameter.cpp
    src/XHwpDocument.cpp
    src/XHwpDocuments.cpp
    src/Utils.cpp
    src/FontDefs.cpp
    src/bindings.cpp
)

set(CPYHWPX_HEADERS
    src/HwpTypes.h
    src/HwpWrapper.h
    src/HwpCtrl.h
    src/HwpAction.h
    src/HwpParameter.h
    src/XHwpDocument.h
    src/XHwpDocuments.h
    src/Utils.h
    src/FontDefs.h
)

#==============================================================================
# Python 모듈 빌드
#==============================================================================

pybind11_add_module(cpyhwpx ${CPYHWPX_SOURCES} ${CPYHWPX_HEADERS})

# 헤더 포함 경로
target_include_directories(cpyhwpx PRIVATE ${CMAKE_CURRENT_SOURCE_DIR}/src)

# COM 라이브러리 링크
target_link_libraries(cpyhwpx PRIVATE
    ole32
    oleaut32
    uuid
    comsuppw
)

# Windows 라이브러리 링크
if(WIN32)
    target_link_libraries(cpyhwpx PRIVATE
        advapi32
        shell32
    )
endif()

# MSVC 컴파일러 설정
if(MSVC)
    target_compile_options(cpyhwpx PRIVATE
        /W4             # 경고 레벨 4
        /EHsc           # 예외 처리 모델
        /utf-8          # UTF-8 소스 인코딩
        /Zc:__cplusplus # __cplusplus 매크로 올바르게
    )

    # Release 최적화
    target_compile_options(cpyhwpx PRIVATE
        $<$<CONFIG:Release>:/O2>
        $<$<CONFIG:Release>:/DNDEBUG>
    )
endif()

#==============================================================================
# 설치 설정
#==============================================================================

# 설치 경로 (pip install 시 자동 처리됨)
install(TARGETS cpyhwpx
    LIBRARY DESTINATION .
    RUNTIME DESTINATION .
)

#==============================================================================
# 디버그 정보
#==============================================================================

message(STATUS "=== cpyhwpx Build Configuration ===")
message(STATUS "CMake version: ${CMAKE_VERSION}")
message(STATUS "Build type: ${CMAKE_BUILD_TYPE}")
message(STATUS "Platform: ${CMAKE_GENERATOR_PLATFORM}")
message(STATUS "C++ standard: ${CMAKE_CXX_STANDARD}")
message(STATUS "Python executable: ${PYTHON_EXECUTABLE}")
message(STATUS "pybind11 version: ${pybind11_VERSION}")
message(STATUS "===================================")
