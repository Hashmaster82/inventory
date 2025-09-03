@echo off
chcp 65001 >nul

set REPO_DIR=%~dp0

where git >nul 2>nul
if errorlevel 1 (
    echo ❌ Git не найден. Установите Git: https://git-scm.com/download/win
    pause
    exit /b
)

cd /d "%REPO_DIR%"

if not exist ".git" (
    echo 📥 Клонирование репозитория...
    git clone https://github.com/Hashmaster82/inventory.git "%REPO_DIR%"
) else (
    echo 🔄 Обновление репозитория...
    git pull origin master
)

echo 🚀 Запуск программы...
if exist "%REPO_DIR%dist\app.exe" (
    start "" "%REPO_DIR%dist\app.exe"
) else (
    python "%REPO_DIR%app.py"
)