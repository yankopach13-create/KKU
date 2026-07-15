@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ========================================
echo   Запуск приложения анализа ККУ (справочник XLSX, операции TXT UTF-8)
echo ========================================
echo.

where python >nul 2>&1
if errorlevel 1 (
    echo [ОШИБКА] Python не найден. Установите Python и добавьте его в PATH.
    pause
    exit /b 1
)

if not exist "venv\" (
    echo Создание виртуального окружения...
    python -m venv venv
    if errorlevel 1 (
        echo [ОШИБКА] Не удалось создать виртуальное окружение.
        pause
        exit /b 1
    )
)

call venv\Scripts\activate.bat

echo Установка зависимостей...
echo (при первом запуске это может занять несколько минут)
echo.

python -m pip install --upgrade pip
if errorlevel 1 (
    echo [ОШИБКА] Не удалось обновить pip.
    pause
    exit /b 1
)

python -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить зависимости.
    pause
    exit /b 1
)

echo.
echo Зависимости установлены.
echo Приложение откроется в браузере: http://localhost:8501
echo Для остановки нажмите Ctrl+C в этом окне.
echo.

python -m streamlit run app.py

pause
