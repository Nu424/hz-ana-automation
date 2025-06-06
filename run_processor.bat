@echo off
chcp 65001 > nul

REM 管理者権限チェック
net session >nul 2>&1
if %errorLevel% == 0 (
    echo 管理者権限で実行中...
) else (
    echo 管理者権限が必要です。管理者権限で再実行します...
    echo.
    powershell -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)
cd /d "%~dp0"

echo ===================================
echo 測定データ処理自動化ツール
echo ===================================
echo.

REM Pythonがインストールされているかチェック
python --version >nul 2>&1
if errorlevel 1 (
    echo エラー: Pythonがインストールされていません。
    echo Pythonをインストールしてから再実行してください。
    pause
    exit /b 1
)

REM 仮想環境を起動
echo 仮想環境を起動中...
call venv/Scripts/activate.bat

REM 必要なライブラリをインストール
echo 必要なライブラリをインストール中...
pip install -r requirements.txt

if errorlevel 1 (
    echo エラー: ライブラリのインストールに失敗しました。
    pause
    exit /b 1
)

echo.
echo ツールを起動します...
echo.

REM メインプログラムを実行
python measurement_data_processor.py

echo.
echo プログラムが終了しました。
deactivate
pause 