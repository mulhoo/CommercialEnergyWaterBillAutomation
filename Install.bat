@echo off
REM Install.bat - Water Bill Processor Installer for Windows

echo.
echo ========================================
echo  Water Bill Processor Installer
echo ========================================
echo.

REM Get current directory (remove trailing backslash if present)
set "INSTALL_DIR=%~dp0"
if "%INSTALL_DIR:~-1%"=="\" set "INSTALL_DIR=%INSTALL_DIR:~0,-1%"

REM Get desktop path and validate it exists
set "desktop=%USERPROFILE%\Desktop"
if not exist "%desktop%" (
    echo Warning: Desktop folder not found at %desktop%
    echo Trying alternative locations...
    set "desktop=%USERPROFILE%\OneDrive\Desktop"
    if not exist "%desktop%" (
        set "desktop=%PUBLIC%\Desktop"
        if not exist "%desktop%" (
            echo ERROR: Could not locate Desktop folder.
            echo Please create a shortcut manually.
            goto :skip_shortcut
        )
    )
)

REM Set paths
set "target=%INSTALL_DIR%\WaterBillProcessor.exe"
set "shortcut_name=Water Bill Processor"
set "shortcut_path=%desktop%\%shortcut_name%.lnk"

echo Installing Water Bill Processor...
echo Install directory: %INSTALL_DIR%
echo Target executable: %target%
echo Desktop path: %desktop%

REM Check if executable exists
if not exist "%target%" (
    echo ERROR: WaterBillProcessor.exe not found!
    echo Please make sure all files were extracted properly.
    pause
    exit /b 1
)

echo Creating desktop shortcut...

REM Create shortcut using PowerShell with proper escaping
powershell -ExecutionPolicy Bypass -Command "& { try { $ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut('%shortcut_path%'); $s.TargetPath = '%target%'; $s.WorkingDirectory = '%INSTALL_DIR%'; $s.Description = 'Water Bill PDF Processor'; $s.Save(); Write-Host 'Shortcut created successfully' } catch { Write-Host 'Shortcut creation failed:' $_.Exception.Message; exit 1 } }"

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
    goto :end
) else (
    :skip_shortcut
    echo.
    echo Installation completed but shortcut creation failed.
    echo You can still run WaterBillProcessor.exe directly.
    echo.
    echo To create a shortcut manually:
    echo 1. Right-click on WaterBillProcessor.exe
    echo 2. Select "Create shortcut"
    echo 3. Move the shortcut to your Desktop
    echo.
)

:end
echo Press any key to continue...
pause >nul