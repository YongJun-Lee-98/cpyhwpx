@echo off
setlocal

REM Set VS environment for 64-bit
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
if errorlevel 1 (
    echo Failed to set up Visual Studio environment
    exit /b 1
)

cd /d D:\Code\cpyhwpx

REM Build with pip
py -3 -m pip install . --no-build-isolation -v

endlocal
