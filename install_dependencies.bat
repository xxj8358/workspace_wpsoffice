@echo off

rem 设置编码为UTF-8
chcp 65001 >nul

echo 正在安装财务数据处理所需的依赖库...
echo.

echo 安装pandas库...
python -m pip install pandas

if %errorlevel% neq 0 (
    echo 安装pandas失败！请确保Python已正确安装并添加到系统环境变量。
    pause
    exit /b 1
)

echo 安装openpyxl库...
python -m pip install openpyxl

if %errorlevel% neq 0 (
    echo 安装openpyxl失败！
    pause
    exit /b 1
)

echo 安装xlrd库...
python -m pip install xlrd

if %errorlevel% neq 0 (
    echo 安装xlrd失败！
    pause
    exit /b 1
)

echo.
echo 所有依赖库安装成功！
echo 现在可以运行run_test_copy_table.bat来处理财务数据了。
echo.
pause