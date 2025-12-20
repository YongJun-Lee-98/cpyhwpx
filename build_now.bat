@echo off
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
cd /d D:\Code\hwp_test\cpyhwpx

set PYBIND11_INC=C:\Users\leep0\AppData\Local\Programs\Python\Python312\Lib\site-packages\pybind11\include
set PYTHON_INC=C:\Users\leep0\AppData\Local\Programs\Python\Python312\include
set PYTHON_LIB=C:\Users\leep0\AppData\Local\Programs\Python\Python312\libs

set SOURCES=src\HwpWrapper.cpp src\HwpCtrl.cpp src\HwpAction.cpp src\HwpParameter.cpp src\Utils.cpp src\FontDefs.cpp src\bindings.cpp

set CL_OPTS=/EHsc /O2 /MD /std:c++17 /utf-8 /W3
set CL_OPTS=%CL_OPTS% /DUNICODE /D_UNICODE /DWIN32_LEAN_AND_MEAN /DNOMINMAX

set INCLUDES=/I"%PYBIND11_INC%" /I"%PYTHON_INC%" /I"src"

set LINK_OPTS=/LIBPATH:"%PYTHON_LIB%" python312.lib ole32.lib oleaut32.lib uuid.lib advapi32.lib shell32.lib shlwapi.lib

if not exist "build64" mkdir build64

for %%f in (%SOURCES%) do (
    echo Compiling %%f...
    cl /c %CL_OPTS% %INCLUDES% /Fo"build64\%%~nf.obj" "%%f"
    if errorlevel 1 exit /b 1
)

echo Linking...
link /DLL /OUT:cpyhwpx\_native\cpyhwpx.pyd build64\*.obj %LINK_OPTS%
if errorlevel 1 exit /b 1

echo Build successful!
dir cpyhwpx\_native\cpyhwpx.pyd
