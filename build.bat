@echo off
chcp 65001 > nul
echo ========================================
echo 正在打包窗口置顶工具...
echo ========================================
echo.

REM 安装依赖
echo [1/2] 正在安装依赖...
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if %errorlevel% neq 0 (
    echo 依赖安装失败！
    pause
    exit /b 1
)

REM 打包程序
echo.
echo [2/2] 正在打包程序...
python -m PyInstaller --onefile ^
    --windowed ^
    --name "窗口置顶工具" ^
    --icon "C:\Users\likaiyi\CodeBuddy\20260212132732\icon\push_pin_bfny858du3i2.ico" ^
    --add-data "C:\Users\likaiyi\CodeBuddy\20260212132732\icon\push_pin_bfny858du3i2.ico;icon" ^
    main.py

if %errorlevel% neq 0 (
    echo 打包失败！
    pause
    exit /b 1
)

echo.
echo ========================================
echo 打包完成！
echo 可执行文件位于: dist\窗口置顶工具.exe
echo ========================================
pause
