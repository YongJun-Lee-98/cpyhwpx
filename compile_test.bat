@echo off
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvars64.bat"
if errorlevel 1 goto :error

cd /d D:\Code\cpyhwpx

echo.
echo Compiling test_module.cpp...
echo.

cl /EHsc /O2 /MD /LD /Fe:test_module.pyd ^
    /I"C:\Users\leep0\AppData\Local\Programs\Python\Python312\Lib\site-packages\pybind11\include" ^
    /I"C:\Users\leep0\AppData\Local\Programs\Python\Python312\include" ^
    test_module.cpp ^
    /link /LIBPATH:"C:\Users\leep0\AppData\Local\Programs\Python\Python312\libs" python312.lib

if errorlevel 1 goto :error

echo.
echo SUCCESS!
dir *.pyd
goto :end

:error
echo.
echo FAILED with error %errorlevel%

:end
