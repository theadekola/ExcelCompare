@echo off
setlocal

cd /d "%~dp0\.."

where cmake >nul 2>nul || (
    echo ERROR: CMake is not installed or not in PATH.
    exit /b 1
)

where git >nul 2>nul || (
    echo ERROR: Git is not installed or not in PATH.
    exit /b 1
)

if "%QT_ROOT%"=="" (
    echo ERROR: QT_ROOT is not set.
    echo Example:
    echo set QT_ROOT=C:\Qt\6.8.3\msvc2022_64
    exit /b 1
)

if not exist "%QT_ROOT%\bin\windeployqt.exe" (
    echo ERROR: windeployqt was not found at:
    echo %QT_ROOT%\bin\windeployqt.exe
    exit /b 1
)

if not exist "%QT_ROOT%\lib\cmake\Qt6\Qt6Config.cmake" (
    echo ERROR: Qt6Config.cmake was not found at:
    echo %QT_ROOT%\lib\cmake\Qt6\Qt6Config.cmake
    exit /b 1
)

set "PATH=%QT_ROOT%\bin;%PATH%"

echo.
echo ============================================
echo Cleaning previous package files
echo ============================================

rem Keep the build folder to avoid compiling QXlsx again.
if exist package rmdir /s /q package
if exist dist rmdir /s /q dist

echo.
echo ============================================
echo Configuring project
echo ============================================

cmake -S . -B build -G "Visual Studio 18 2026" -A x64 -T v143 -DCMAKE_PREFIX_PATH="%QT_ROOT%"

if errorlevel 1 (
    echo.
    echo ERROR: CMake configuration failed.
    exit /b 1
)

echo.
echo ============================================
echo Building Release version
echo ============================================

cmake --build build --config Release

if errorlevel 1 (
    echo.
    echo ERROR: Application build failed.
    exit /b 1
)

if not exist "build\Release\ExcelCompareProfessional.exe" (
    echo ERROR: ExcelCompareProfessional.exe was not generated.
    exit /b 1
)

echo.
echo ============================================
echo Creating deployment package
echo ============================================

mkdir package

copy /y "build\Release\ExcelCompareProfessional.exe" "package\ExcelCompareProfessional.exe"

if errorlevel 1 (
    echo ERROR: Failed to copy the application executable.
    exit /b 1
)

echo.
echo ============================================
echo Deploying Qt dependencies
echo ============================================

"%QT_ROOT%\bin\windeployqt.exe" --release --compiler-runtime --no-translations "package\ExcelCompareProfessional.exe"

if errorlevel 1 (
    echo.
    echo ERROR: Qt deployment failed.
    exit /b 1
)

echo.
echo ============================================
echo Testing deployed application files
echo ============================================

if not exist "package\Qt6Core.dll" (
    echo ERROR: Qt6Core.dll was not deployed.
    exit /b 1
)

if not exist "package\Qt6Gui.dll" (
    echo ERROR: Qt6Gui.dll was not deployed.
    exit /b 1
)

if not exist "package\Qt6Widgets.dll" (
    echo ERROR: Qt6Widgets.dll was not deployed.
    exit /b 1
)

if not exist "package\platforms\qwindows.dll" (
    echo ERROR: qwindows.dll was not deployed.
    exit /b 1
)

echo.
echo ============================================
echo Copying commercial package files
echo ============================================

if exist "docs\Excel_Compare_Professional_User_Manual.pdf" (
    mkdir "package\Documentation" 2>nul
    copy /y "docs\Excel_Compare_Professional_User_Manual.pdf" "package\Documentation\"
)

if exist "licenses\LICENCE.txt" (
    copy /y "licenses\LICENCE.txt" "package\"
)

if exist "samples" (
    xcopy /e /i /y "samples" "package\samples"
)

if exist "assets" (
    xcopy /e /i /y "assets" "package\assets"
)

if exist "tools" (
    xcopy /e /i /y "tools" "package\tools"
)

if exist "docs\FORMAT_SUPPORT.md" (
    copy /y "docs\FORMAT_SUPPORT.md" "package\Documentation\"
)

if exist "docs\NATIVE_XLS_SUPPORT.md" (
    copy /y "docs\NATIVE_XLS_SUPPORT.md" "package\Documentation\"
)

if exist "docs\THIRD_PARTY_NOTICES.md" (
    copy /y "docs\THIRD_PARTY_NOTICES.md" "package\Documentation\"
)

echo.
echo ============================================
echo Locating Inno Setup
echo ============================================

set "ISCC=%ProgramFiles(x86)%\Inno Setup 6\ISCC.exe"

if not exist "%ISCC%" (
    set "ISCC=%ProgramFiles%\Inno Setup 6\ISCC.exe"
)

if not exist "%ISCC%" (
    set "ISCC=%ProgramFiles(x86)%\Inno Setup 7\ISCC.exe"
)

if not exist "%ISCC%" (
    set "ISCC=%ProgramFiles%\Inno Setup 7\ISCC.exe"
)

if not exist "%ISCC%" (
    echo ERROR: Inno Setup 6 or 7 was not found.
    echo Install Inno Setup and run this script again.
    exit /b 1
)

echo Found Inno Setup:
echo %ISCC%

echo.
echo ============================================
echo Creating Windows installer
echo ============================================

"%ISCC%" "installer\ExcelCompareProfessional.iss"

if errorlevel 1 (
    echo.
    echo ERROR: Installer compilation failed.
    exit /b 1
)

echo.
echo ============================================
echo SUCCESS
echo ============================================
echo Installer created in the dist folder.
echo.
echo Expected file:
echo dist\ExcelCompareProfessional-Setup-v1.4.0.exe
echo.

endlocal
pause