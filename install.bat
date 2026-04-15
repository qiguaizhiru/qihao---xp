@echo off
title iMouseXP V2.0 - Install
cd /d "%~dp0"

echo ============================================
echo    iMouseXP V2.0 - Install
echo ============================================
echo.

echo [1/3] Checking Python...
call :find_python
if defined PYTHON_EXE goto :python_ok

echo Python not found, installing...
echo.
if not exist "%~dp0python-installer.exe" (
    echo [Error] python-installer.exe not found!
    goto :end
)
echo Installing Python 3.12 ...
start /wait "" "%~dp0python-installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=1
echo Done.
echo.

call :refresh_path
call :find_python
if not defined PYTHON_EXE (
    echo [Error] Python not found after install.
    goto :end
)

:python_ok
echo Found: %PYTHON_EXE%
"%PYTHON_EXE%" --version
echo.

echo [2/3] Installing packages...
"%PYTHON_EXE%" -m pip install --no-index --find-links="%~dp0packages" imouse-py openpyxl colorlog pydantic websocket-client requests pillow
if %errorlevel% neq 0 (
    echo.
    echo [Error] Package install failed.
    goto :end
)
echo.

echo [3/3] Fixing compatibility...
"%PYTHON_EXE%" "%~dp0fix_helper.py"
echo.

echo ============================================
echo    Done! Run start.bat to launch.
echo ============================================

goto :end

:find_python
set "PYTHON_EXE="
where python >nul 2>nul
if %errorlevel% equ 0 (
    for /f "delims=" %%P in ('where python 2^>nul') do (
        if not defined PYTHON_EXE set "PYTHON_EXE=%%P"
    )
    if defined PYTHON_EXE goto :eof
)
if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    goto :eof
)
if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
    set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python314\python.exe"
    goto :eof
)
if exist "%LOCALAPPDATA%\Programs\Python\Python313\python.exe" (
    set "PYTHON_EXE=%LOCALAPPDATA%\Programs\Python\Python313\python.exe"
    goto :eof
)
if exist "C:\Python312\python.exe" (
    set "PYTHON_EXE=C:\Python312\python.exe"
    goto :eof
)
if exist "C:\Python3\python.exe" (
    set "PYTHON_EXE=C:\Python3\python.exe"
    goto :eof
)
goto :eof

:refresh_path
for /f "tokens=2*" %%A in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "USERPATH=%%B"
if defined USERPATH set "PATH=%USERPATH%;%PATH%"
goto :eof

:end
echo.
echo Press any key to exit...
pause >nul
