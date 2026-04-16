@echo off

cd /d "%~dp0"

echo.
echo ========================================
echo   Shift Calendar Update
echo ========================================
echo.

if "%~1" neq "" (
    echo Converting specified file...
    python convert.py %*
) else (
    echo Auto-detecting from Downloads...
    python convert.py
)

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Convert failed.
    pause
    exit /b 1
)

echo.
echo Uploading to GitHub...
git add data\
git commit -m "update shift data"
git push

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Push failed.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   Done!
echo   https://share-eng.github.io/shift-calendar/
echo ========================================
echo.
pause
