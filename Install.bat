@echo off
REM Install.bat - Water Bill Processor Installer for Windows

echo.
echo ========================================
echo  Water Bill Processor Installer
echo ========================================
echo.

REM Get current directory
set "INSTALL_DIR=%~dp0"
set "target=%INSTALL_DIR%WaterBillProcessor.exe"

echo Installing Water Bill Processor...

REM Check if executable exists
if not exist "%target%" (
    echo ERROR: WaterBillProcessor.exe not found!
    echo Please make sure all files were extracted properly.
    pause
    exit /b 1
)

REM Test if the application can start
echo Testing application startup...
"%target%" --help >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo WARNING: Application may be missing required dependencies.
    echo.
    echo This application requires:
    echo   - Tesseract OCR
    echo   - Poppler PDF utilities
    echo   - Microsoft Visual C++ Redistributable
    echo.
    echo Please install these dependencies or the application may not work.
    echo.
    echo You can download them from:
    echo   - Tesseract: https://github.com/UB-Mannheim/tesseract/wiki
    echo   - Poppler: https://github.com/oschwartz10612/poppler-windows
    echo   - VC++ Redist: https://aka.ms/vs/17/release/vc_redist.x64.exe
    echo.
)

REM Create shortcut using VBScript (more reliable than PowerShell)
echo Creating desktop shortcut...

echo Set oWS = WScript.CreateObject("WScript.Shell") > "%temp%\CreateShortcut.vbs"
echo sLinkFile = oWS.SpecialFolders("Desktop") ^& "\Water Bill Processor.lnk" >> "%temp%\CreateShortcut.vbs"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%temp%\CreateShortcut.vbs"
echo oLink.TargetPath = "%target%" >> "%temp%\CreateShortcut.vbs"
echo oLink.WorkingDirectory = "%INSTALL_DIR%" >> "%temp%\CreateShortcut.vbs"
echo oLink.Description = "Water Bill PDF Processor" >> "%temp%\CreateShortcut.vbs"
echo oLink.Save >> "%temp%\CreateShortcut.vbs"

cscript //nologo "%temp%\CreateShortcut.vbs"
del "%temp%\CreateShortcut.vbs"

if %ERRORLEVEL% EQU 0 (
    echo Shortcut created successfully.
) else (
    echo Shortcut creation failed - you can run the .exe directly.
)

echo.
echo Installation complete!
echo.
echo To run the application:
echo   1. Double-click "Water Bill Processor" on your Desktop
echo   2. Or double-click WaterBillProcessor.exe in this folder
echo.
echo If the application doesn't start, you may need to install dependencies.
echo See the warnings above for download links.
echo.

pause