@echo off
title iMouseXP V2.0
cd /d "%~dp0"

call :find_python
if not defined PYTHON_EXE (
    call :refresh_path
    call :find_python
)
if not defined PYTHON_EXE (
    echo Python not found!
    if exist "%~dp0python-installer.exe" (
        echo Installing Python 3.12...
        start /wait "" "%~dp0python-installer.exe" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_tcltk=1
        call :refresh_path
        call :find_python
    )
)
if not defined PYTHON_EXE (
    echo [Error] Python not found. Run install.bat first.
    goto :end
)

"%PYTHON_EXE%" "%~dp0fix_helper.py" >nul 2>nul

"%PYTHON_EXE%" -c "import imouse" >nul 2>nul
if %errorlevel% neq 0 (
    echo Installing packages...
    "%PYTHON_EXE%" -m pip install --no-index --find-links="%~dp0packages" imouse-py openpyxl colorlog pydantic websocket-client requests pillow >nul 2>nul
    "%PYTHON_EXE%" "%~dp0fix_helper.py" >nul 2>nul
    echo Done.
    echo.
)

rem 应用更新：如果存在 _apply_update.bat，先执行它把 .new 文件覆盖原文件
if exist "%~dp0_apply_update.bat" (
    call "%~dp0_apply_update.bat"
    del /q "%~dp0_apply_update.bat" >nul 2>nul
)

"%PYTHON_EXE%" "%~dp0transfer_gui.py"
if %errorlevel% neq 0 (
    echo.
    echo [Error] Program error. See above.
    goto :end
)
goto :eof

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
