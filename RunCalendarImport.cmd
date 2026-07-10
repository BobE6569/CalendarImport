@echo off
setlocal

set "PROJECT_ROOT=%~dp0"
set "VENV_PYTHON=%PROJECT_ROOT%.venv\Scripts\python.exe"

if not exist "%VENV_PYTHON%" (
    where py >nul 2>nul
    if %errorlevel%==0 (
        py -3 -m venv "%PROJECT_ROOT%.venv"
    ) else (
        where python >nul 2>nul
        if errorlevel 1 goto no_python
        python -m venv "%PROJECT_ROOT%.venv"
    )
    if errorlevel 1 goto fail
)

"%VENV_PYTHON%" -m pip install -r "%PROJECT_ROOT%requirements.txt"
if errorlevel 1 goto fail

"%VENV_PYTHON%" -m pip install -e "%PROJECT_ROOT%"
if errorlevel 1 goto fail

"%VENV_PYTHON%" -m calendar_import
if errorlevel 1 goto fail

exit /b 0

:no_python
echo Python 3 was not found. Install Python 3.10 or newer, then run this file again.
pause
exit /b 1

:fail
echo.
echo CalendarImport could not start. Review the error above.
pause
exit /b 1

