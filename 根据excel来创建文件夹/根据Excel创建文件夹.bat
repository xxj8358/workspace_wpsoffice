@echo off

rem 根据Excel文件C列创建文件夹的批处理脚本

echo 正在启动Excel文件分类创建文件夹工具...
echo.  

rem 检查Python是否安装
python --version > nul 2>&1
if %errorlevel% neq 0 (
    echo 错误：未找到Python。请先安装Python并添加到系统路径。
    echo.  
    pause
    exit /b 1
)

rem 检查pandas是否安装
python -c "import pandas" > nul 2>&1
if %errorlevel% neq 0 (
    echo 正在安装必要的依赖包...
    pip install pandas openpyxl
)

rem 运行Python脚本
echo 正在运行脚本...
echo.  
python create_folders_from_excel.py

rem 等待用户按键退出
echo.  
echo 程序已完成。
pause