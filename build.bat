@echo off
chcp 65001
title ZZ打票 自动化打包工具 (v3.2 全能版)

echo ==============================================
echo       正在开始自动化打包流程 (含台账引擎)
echo ==============================================

:: 切换到脚本所在目录，防止路径错误
cd /d "%~dp0"

:: 1. 清理旧缓存 (彻底清理，防止旧代码残留)
echo [1/5] 清理旧缓存...
if exist "build" rd /s /q "build"
if exist "dist" rd /s /q "dist"

:: 2. 检查并激活虚拟环境 (适配 venv_printer)
if exist "venv_printer\Scripts\activate.bat" (
    echo [2/5] 检测到 venv_printer，正在激活...
    call "venv_printer\Scripts\activate.bat"
) else (
    echo [2/5] 未检测到环境，正在创建新环境...
    python -m venv venv_printer
    call "venv_printer\Scripts\activate.bat"
)

:: 3. 安装/更新关键依赖 (增加了 pandas, pdfplumber, openpyxl)
echo [3/5] 正在安装核心依赖...
pip install PyQt6 PyMuPDF pdfplumber pandas openpyxl pyinstaller

:: 4. 执行增强型打包
:: 注意：
:: 1. 移除了对 pandas/numpy 的排除 (exclude)
:: 2. 增加了 hidden-import 防止核心库丢失
:: 3. 排除 matplotlib/scipy 以减小体积
echo [4/5] 正在执行 PyInstaller 打包 (体积较大，请耐心等待)...
pyinstaller --noconsole ^
            --onedir ^
            --name "ZZ打票" ^
            --clean ^
            --hidden-import pdfplumber ^
            --hidden-import pandas ^
            --hidden-import openpyxl ^
            --exclude-module matplotlib ^
            --exclude-module scipy ^
            --exclude-module PyQt5 ^
            --exclude-module tkinter ^
            "zzprint.py"

:: 5. 完成提示
if exist "dist\ZZ打票\ZZ打票.exe" (
    echo.
    echo ==============================================
    echo          ★★★ 打包成功 (v3.2) ★★★
    echo ==============================================
    echo.
    echo [重要提示]
    echo 1. 成品位置: "%cd%\dist\ZZ打票"
    echo 2. 请务必发送整个 "ZZ打票" 文件夹，不要只发 exe！
    echo 3. 由于包含 Excel 引擎，体积会比旧版大，属正常现象。
    echo.
    start "" "dist\ZZ打票"
) else (
    echo.
    echo [错误] 打包失败，请截图上方红色报错信息！
)

pause