@echo off
REM Install.bat - Water Bill Processor Installer for Windows

echo.
echo ========================================
echo  Water Bill Processor Installer
echo ========================================
echo.

REM Get current directory
set "INSTALL_DIR=%~dp0"

REM Create Desktop shortcut
set "desktop=%USERPROFILE%\Desktop"
set "target=%INSTALL_DIR%WaterBillProcessor.exe"
set "shortcut_name=Water Bill Processor"

echo Installing Water Bill Processor...

REM Check if executable exists
if not exist "%target%" (
    echo ERROR: WaterBillProcessor.exe not found!
    echo Please make sure all files were extracted properly.
    pause
    exit /b 1
)

echo Creating desktop shortcut...

REM Create shortcut using PowerShell
powershell -command "& { $ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%desktop%\%shortcut_name%.lnk'); $s.TargetPath = '%target%'; $s.WorkingDirectory = '%INSTALL_DIR%'; $s.Description = 'Water Bill PDF Processor'; $s.Save() }"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo SUCCESS! Installation complete.
    echo.
    echo You can now run the application by:
    echo   * Double-clicking "%shortcut_name%" on your Desktop
    echo   * Or double-clicking WaterBillProcessor.exe in this folder
    echo.
    echo The application will create a "Bills" folder for processed files.
    echo.
) else (
    echo.
    echo Installation completed but shortcut creation may have failed.
    echo You can still run WaterBillProcessor.exe directly.
    echo.
)

echo Press any key to continue...
pause >nul