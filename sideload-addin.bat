@echo off
echo Installing Word Add-in via Registry...
echo.

REM Get the current directory
set "ADDIN_PATH=%~dp0"
set "ADDIN_PATH=%ADDIN_PATH:~0,-1%"

REM Convert backslashes to forward slashes for file:// URL
set "ADDIN_PATH=%ADDIN_PATH:\=/%"
set "ADDIN_PATH=file:///%ADDIN_PATH%/"

echo Add-in path: %ADDIN_PATH%
echo.

REM Add to registry (Office 365 / Office 2016+)
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\TrustedCatalogs\TextFormatter" /v "CatalogUrl" /t REG_SZ /d "%ADDIN_PATH%" /f

if %ERRORLEVEL% EQU 0 (
    echo.
    echo SUCCESS! Add-in registered.
    echo.
    echo Please RESTART Microsoft Word for the changes to take effect.
    echo After restarting, the add-in should appear in the Home tab ribbon.
) else (
    echo.
    echo ERROR: Could not add to registry. Please run as Administrator.
    pause
)

