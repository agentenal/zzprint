@echo off
chcp 65001
title ZZ打票 打包工具

echo ==============================================
echo       正在开始自动化打包流程 (路径适配版)
echo ==============================================

:: 切换到脚本所在目录，防止路径错误
cd /d "%~dp0"

:: 1. 清理旧缓存
echo [1/5] 清理旧缓存...
if exist "build" rd /s /q "build"
if exist "dist" rd /s /q "dist"

:: 2. 检查并激活你的虚拟环境 (适配 venv_printer)
if exist "venv_printer\Scripts\activate.bat" (
    echo [2/5] 检测到 venv_printer，正在激活...
    call "venv_printer\Scripts\activate.bat"
) else (
    echo [2/5] 未检测到环境，正在创建新环境...
    python -m venv venv_printer
    call "venv_printer\Scripts\activate.bat"
)

:: 3. 确保安装了打包工具
echo [3/5] 检查核心依赖...
pip install PyQt6 PyMuPDF pyinstaller

:: 4. 执行打包 (注意：这里给所有参数都加上了双引号)
echo [4/5] 正在打包，请稍候...
pyinstaller --noconsole ^
            --onedir ^
            --name "ZZ打票" ^
            --clean ^
            --exclude-module pandas ^
            --exclude-module numpy ^
            --exclude-module matplotlib ^
            --exclude-module scipy ^
            --exclude-module PyQt5 ^
            "zzprint.py"

:: 5. 检查结果
if exist "dist\ZZ打票" (
    echo.
    echo [5/5] ★★★ 打包成功 ★★★
    echo 成品位置: "%cd%\dist\ZZ打票"
    start "" "dist"
) else (
    echo.
    echo [错误] 打包失败，请检查上方是否有红色报错信息！
)

pause