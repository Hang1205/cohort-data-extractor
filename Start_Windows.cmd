@echo off
setlocal
pushd "%~dp0" || exit /b 1
if not exist "cohort_extract_gui.py" (
  echo The GUI file is missing. Do not run this launcher from inside the ZIP.
  echo Right-click the ZIP, choose Extract All, and run Start_Windows.cmd in the extracted folder.
  pause
  exit /b 1
)
if not exist "vendor\openpyxl\__init__.py" if not exist "vendor.zip" (
  echo Extract the ENTIRE ZIP, including the vendor folder, to enable Excel id lists.
  pause
  exit /b 1
)
if not exist "Find_Python312.cmd" (
  echo Extract the entire updated ZIP, including Find_Python312.cmd.
  pause
  popd
  exit /b 1
)
call "%~dp0Find_Python312.cmd"
if errorlevel 1 (
  echo Python 3.12 64-bit with Tkinter was not found.
  echo Use the Windows DESKTOP package to run without installing Python.
  pause
  popd
  exit /b 1
)
"%DEID_PYTHON%" cohort_extract_gui.py
if errorlevel 1 pause
popd
