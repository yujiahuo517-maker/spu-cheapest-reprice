@echo off
setlocal
set "DIR=%~dp0"
if "%~1"=="" (
  echo 用法: 把导出商品JX-SKU信息 xlsx 拖到本文件上，或命令行传入文件路径
  pause
  exit /b 1
)
python "%DIR%scripts\fill_reprice.py" "%~1"
if errorlevel 1 (
  echo.
  echo 运行失败：请确认已安装 Python 且执行过 pip install openpyxl
  pause
  exit /b 1
)
echo.
explorer /select,"%~dpn1_改价结果.xlsx"
