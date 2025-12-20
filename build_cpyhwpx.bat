@echo off
setlocal EnableDelayedExpansion

echo ============================================================
echo Building cpyhwpx
echo ============================================================
echo.

REM Set up Visual Studio environment
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
if errorlevel 1 (
    echo ERROR: Failed to set up Visual Studio environment
    exit /b 1
)

cd /d D:\Code\cpyhwpx

REM Set paths
set PYBIND11_INC=C:\Users\leep0\AppData\Local\Programs\Python\Python312\Lib\site-packages\pybind11\include
set PYTHON_INC=C:\Users\leep0\AppData\Local\Programs\Python\Python312\include
set PYTHON_LIB=C:\Users\leep0\AppData\Local\Programs\Python\Python312\libs

REM Source files
set SOURCES=src\HwpWrapper.cpp src\HwpCtrl.cpp src\HwpAction.cpp src\HwpParameter.cpp src\Utils.cpp src\FontDefs.cpp src\bindings.cpp

REM Compile options
set CL_OPTS=/EHsc /O2 /MD /std:c++17 /utf-8 /W3
set CL_OPTS=%CL_OPTS% /DUNICODE /D_UNICODE /DWIN32_LEAN_AND_MEAN /DNOMINMAX

REM Include paths
set INCLUDES=/I"%PYBIND11_INC%" /I"%PYTHON_INC%" /I"src"

REM Link options
set LINK_OPTS=/LIBPATH:"%PYTHON_LIB%" python312.lib ole32.lib oleaut32.lib uuid.lib advapi32.lib shell32.lib

echo Compiling cpyhwpx...
echo.
echo Sources: %SOURCES%
echo.

REM Create object files directory
if not exist "build" mkdir build

REM Compile each source file
for %%f in (%SOURCES%) do (
    echo Compiling %%f...
    cl /c %CL_OPTS% %INCLUDES% /Fo"build\%%~nf.obj" "%%f"
    if errorlevel 1 (
        echo ERROR: Failed to compile %%f
        exit /b 1
    )
)

echo.
echo Linking cpyhwpx.pyd...
echo.

REM Link all object files
link /DLL /OUT:cpyhwpx.pyd build\*.obj %LINK_OPTS%
if errorlevel 1 (
    echo ERROR: Failed to link
    exit /b 1
)

echo.
echo ============================================================
echo Build successful!
echo ============================================================
dir cpyhwpx.pyd

echo.
echo Testing import...
py -3 -c "import cpyhwpx; print('cpyhwpx imported successfully!'); print('Version:', cpyhwpx.__doc__)"

endlocal
