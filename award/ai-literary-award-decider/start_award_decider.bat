@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

where py >nul 2>nul
if not errorlevel 1 (
    py -3 -c "import tkinter" >nul 2>nul
    if errorlevel 1 goto tkinter_missing
    py -3 award_decider_gui.py
    if errorlevel 1 goto launch_failed
    exit /b 0
)

where python >nul 2>nul
if not errorlevel 1 (
    python -c "import tkinter" >nul 2>nul
    if errorlevel 1 goto tkinter_missing
    python award_decider_gui.py
    if errorlevel 1 goto launch_failed
    exit /b 0
)

echo Python 3 が見つかりません。
echo https://www.python.org/downloads/windows/ からPython 3をインストールしてください。
echo インストール時に「Add Python to PATH」と「tcl/tk and IDLE」を有効にしてください。
pause
exit /b 1

:tkinter_missing
echo tkinter が利用できません。
echo Pythonを修復または再インストールし、「tcl/tk and IDLE」を有効にしてください。
pause
exit /b 1

:launch_failed
echo アプリを起動できませんでした。上に表示されたエラーを確認してください。
pause
exit /b 1
