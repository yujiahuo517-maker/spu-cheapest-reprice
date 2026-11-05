@echo off
setlocal
set "FILE=%~1"

rem no argument: open file picker dialog (defaults to Downloads)
if "%FILE%"=="" (
  for /f "usebackq delims=" %%f in (`powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; $d = New-Object System.Windows.Forms.OpenFileDialog; $d.Filter = 'Excel (*.xlsx)|*.xlsx'; $d.Title = 'select exported JX-SKU xlsx'; $d.InitialDirectory = [Environment]::GetFolderPath('UserProfile') + '\Downloads'; if ($d.ShowDialog() -eq 'OK') { $d.FileName }"`) do set "FILE=%%f"
)
if "%FILE%"=="" (
  echo no file selected
  pause
  exit /b 1
)

set "LOG=%TEMP%\reprice_last.log"
python "%~dp0scripts\fill_reprice.py" "%FILE%" > "%LOG%" 2>&1
type "%LOG%"
if errorlevel 1 (
  echo.
  echo FAILED: install Python and run: pip install openpyxl
  pause
  exit /b 1
)

set "OUT="
for /f "usebackq tokens=1,* delims==" %%a in (`findstr /b RESULT_FILE= "%LOG%"`) do set "OUT=%%b"
if defined OUT (
  explorer /select,"%OUT%"
) else (
  for %%F in ("%FILE%") do explorer "%%~dpF"
)
