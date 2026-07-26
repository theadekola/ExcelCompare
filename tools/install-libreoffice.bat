@echo off
setlocal

echo Excel Compare Professional - Legacy Format Support
echo AAT-Tech Ltd
echo.

if exist "C:\Program Files\LibreOffice\program\soffice.exe" (
    echo LibreOffice is already installed.
    echo Legacy spreadsheet support is ready.
    pause
    exit /b 0
)

if exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" (
    echo LibreOffice is already installed.
    echo Legacy spreadsheet support is ready.
    pause
    exit /b 0
)

where winget >nul 2>nul
if not errorlevel 1 (
    echo Installing LibreOffice using Windows Package Manager...
    winget install --id TheDocumentFoundation.LibreOffice --exact --source winget --accept-package-agreements --accept-source-agreements
    if not errorlevel 1 (
        echo.
        echo LibreOffice installation completed.
        echo Restart Excel Compare Professional before opening legacy files.
        pause
        exit /b 0
    )
)

echo Automatic installation was not available.
echo Opening the official LibreOffice download page...
start "" "https://www.libreoffice.org/download/download-libreoffice/"
echo Install LibreOffice, then restart Excel Compare Professional.
pause
endlocal
