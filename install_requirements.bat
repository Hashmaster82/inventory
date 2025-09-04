@echo off
REM Активация виртуального окружения (если используется)
REM call venv\Scripts\activate.bat

REM Установка пакетов из requirements.txt
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

pause
