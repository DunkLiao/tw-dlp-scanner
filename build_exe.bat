@echo off
chcp 65001 >nul
title DLP Scanner 打包工具

echo.
echo ==========================================
echo   DLP Scanner - Windows EXE 打包工具
echo ==========================================
echo.

REM 檢查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 Python，請先安裝 Python 並加入 PATH。
    pause
    exit /b 1
)

echo [1/5] 檢查 pip...
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo [錯誤] 找不到 pip。
    pause
    exit /b 1
)

echo.
echo [2/5] 安裝必要套件...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller

if errorlevel 1 (
    echo [錯誤] 套件安裝失敗。
    pause
    exit /b 1
)

echo.
echo [3/5] 清除舊的 build / dist 資料夾...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist dlp_scanner.spec del /q dlp_scanner.spec

echo.
echo [4/5] 開始打包 EXE...
python -m PyInstaller ^
  --onefile ^
  --windowed ^
  --exclude-module PyQt5 ^
  --exclude-module PyQt6 ^
  --exclude-module PySide2 ^
  --exclude-module PySide6 ^
  --exclude-module pandas ^
  --exclude-module numpy ^
  --exclude-module scipy ^
  --exclude-module matplotlib ^
  --exclude-module pytest ^
  --exclude-module IPython ^
  --exclude-module dask ^
  --exclude-module sphinx ^
  --exclude-module black ^
  --name DLPScanner ^
  dlp_scanner.py

if errorlevel 1 (
    echo [錯誤] EXE 打包失敗。
    pause
    exit /b 1
)

echo.
echo [5/5] 打包完成。
echo.
echo EXE 位置：
echo dist\DLPScanner.exe
echo.

if exist dist\DLPScanner.exe (
    echo 是否開啟 dist 資料夾？
    choice /c YN /m "請選擇"
    if errorlevel 2 goto end
    if errorlevel 1 explorer dist
)

:end
echo.
echo 完成。
pause
