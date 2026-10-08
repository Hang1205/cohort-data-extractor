@echo off
rem Called by the launchers; return a verified Python executable in DEID_PYTHON.
set "DEID_PYTHON="
call :probe "%~dp0.venv\Scripts\python.exe"
if defined DEID_PYTHON exit /b 0
for /f "delims=" %%P in ('where py.exe 2^>nul') do for /f "delims=" %%Q in ('"%%P" -3.12 -c "import sys; print(sys.executable)" 2^>nul') do call :probe "%%Q"
if defined DEID_PYTHON exit /b 0
for /f "delims=" %%P in ('where python.exe 2^>nul') do call :probe "%%P"
if defined DEID_PYTHON exit /b 0
call :probe "%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
call :probe "%ProgramFiles%\Python312\python.exe"
call :probe "C:\Python312\python.exe"
if defined DEID_PYTHON exit /b 0
for /f "tokens=2,*" %%A in ('reg query "HKCU\Software\Python\PythonCore\3.12\InstallPath" /ve 2^>nul ^| find "REG_SZ"') do call :probe "%%B\python.exe"
for /f "tokens=2,*" %%A in ('reg query "HKLM\Software\Python\PythonCore\3.12\InstallPath" /ve 2^>nul ^| find "REG_SZ"') do call :probe "%%B\python.exe"
if defined DEID_PYTHON exit /b 0
exit /b 1

:probe
if defined DEID_PYTHON exit /b 0
if not exist "%~1" exit /b 1
"%~1" -c "import sys, struct, tkinter; assert sys.version_info[:2] == (3,12); assert struct.calcsize('P') == 8" >nul 2>&1
if errorlevel 1 exit /b 1
set "DEID_PYTHON=%~1"
exit /b 0
