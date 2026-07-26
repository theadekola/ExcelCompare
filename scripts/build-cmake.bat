@echo off
setlocal

cd /d "%~dp0\.."

if "%QT_ROOT%"=="" (
    echo ERROR: QT_ROOT is not set.
    echo Example:
    echo set QT_ROOT=C:\Qt\6.8.3\msvc2022_64
    exit /b 1
)

if not exist "%QT_ROOT%\lib\cmake\Qt6\Qt6Config.cmake" (
    echo ERROR: Qt was not found at:
    echo %QT_ROOT%
    exit /b 1
)

cmake --version
if errorlevel 1 exit /b 1

cmake -S . -B build ^
-G "Visual Studio 18 2026" ^
-A x64 ^
-T v143 ^
-DCMAKE_PREFIX_PATH="%QT_ROOT%"

if errorlevel 1 (
    echo.
    echo CMake configuration failed.
    exit /b 1
)

cmake --build build --config Release

if errorlevel 1 (
    echo.
    echo Build failed.
    exit /b 1
)

echo.
echo Build completed successfully.
echo Executable location:
echo build\Release\ExcelCompareProfessional.exe

endlocal