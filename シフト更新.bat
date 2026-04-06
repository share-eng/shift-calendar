@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   シフトカレンダー更新ツール
echo ========================================
echo.

REM 引数があればそのファイルを変換、なければDownloadsを自動検出
if "%~1" neq "" (
    echo 指定ファイルを変換します...
    python convert.py %*
) else (
    echo Downloadsフォルダのシフト管理Excelを自動検出します...
    python convert.py
)

if errorlevel 1 (
    echo.
    echo [エラー] 変換に失敗しました。
    pause
    exit /b 1
)

echo.
echo GitHubにアップロードしています...
git add data/
git commit -m "シフトデータ更新"
git push

if errorlevel 1 (
    echo.
    echo [エラー] GitHubへのアップロードに失敗しました。
    pause
    exit /b 1
)

echo.
echo ========================================
echo   完了！1分ほどでスマホに反映されます
echo   https://share-eng.github.io/shift-calendar/
echo ========================================
echo.
pause
