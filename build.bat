@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul

REM ============================================================
REM Настройки проекта
REM ============================================================
set APP_NAME=PIIMasker
set SCRIPT_NAME=main.py
set VENV_DIR=.venv
set ICON_FILE=icon.ico
set HTML_FILE=index.html
set REQ_FILE=requirements.txt
REM ============================================================

echo [1/8] Проверка Python...
where python >nul 2>nul
if errorlevel 1 (
    echo [ОШИБКА] Python не найден в PATH. Установите Python 3.11 и добавьте его в PATH.
    pause
    exit /b 1
)
python --version
echo.

echo [2/8] Создание виртуального окружения "%VENV_DIR%"...
if not exist "%VENV_DIR%" (
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo [ОШИБКА] Не удалось создать venv.
        pause
        exit /b 1
    )
) else (
    echo Виртуальное окружение уже существует, пропускаем.
)
echo.

echo [3/8] Активация venv...
call "%VENV_DIR%\Scripts\activate.bat"
if errorlevel 1 (
    echo [ОШИБКА] Не удалось активировать venv.
    pause
    exit /b 1
)
echo.

echo [4/8] Обновление pip / setuptools / wheel...
python -m pip install --upgrade pip setuptools wheel
if errorlevel 1 (
    echo [ОШИБКА] Не удалось обновить pip.
    pause
    exit /b 1
)
echo.

echo [5/8] Проверка исходных файлов...
if not exist "%SCRIPT_NAME%" (
    echo [ОШИБКА] Файл "%SCRIPT_NAME%" не найден рядом с build.bat.
    pause
    exit /b 1
)
if not exist "%HTML_FILE%" (
    echo [ОШИБКА] Файл "%HTML_FILE%" не найден рядом с build.bat.
    pause
    exit /b 1
)
if not exist "%REQ_FILE%" (
    echo [ОШИБКА] Файл "%REQ_FILE%" не найден рядом с build.bat.
    pause
    exit /b 1
)
echo Все исходные файлы на месте.
echo.

echo [6/8] Установка зависимостей из %REQ_FILE%...
python -m pip install -r "%REQ_FILE%"
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить зависимости.
    pause
    exit /b 1
)
echo.

echo [7/8] Установка / обновление PyInstaller...
python -m pip install --upgrade pyinstaller
if errorlevel 1 (
    echo [ОШИБКА] Не удалось установить PyInstaller.
    pause
    exit /b 1
)
echo.

echo [8/8] Сборка %APP_NAME%.exe...
set ICON_ARG=
if exist "%ICON_FILE%" (
    set ICON_ARG=--icon "%ICON_FILE%"
    echo Используется иконка: %ICON_FILE%
) else (
    echo Иконка %ICON_FILE% не найдена — сборка без иконки.
)

REM --add-data "index.html;." — кладёт index.html внутрь бандла (sys._MEIPASS)
REM --noconsole — без окна консоли (для запуска как службы)
REM --onefile — один exe
REM --uac-admin — запрос прав администратора (для установки как службы)
REM --collect-submodules — подстраховка для динамических импортов FastAPI/uvicorn

pyinstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --noconsole ^
    --uac-admin ^
    --name "%APP_NAME%" ^
    %ICON_ARG% ^
    --add-data "%HTML_FILE%;." ^
    --hidden-import uvicorn.logging ^
    --hidden-import uvicorn.loops.auto ^
    --hidden-import uvicorn.loops.asyncio ^
    --hidden-import uvicorn.protocols.http.auto ^
    --hidden-import uvicorn.protocols.http.h11_impl ^
    --hidden-import uvicorn.protocols.websockets.auto ^
    --hidden-import uvicorn.lifespan.on ^
    --hidden-import uvicorn.lifespan.off ^
    --hidden-import engineio.async_drivers.threading ^
    --hidden-import sqlalchemy.sql.default_comparator ^
    --hidden-import sqlalchemy.dialects.sqlite ^
    --hidden-import openpyxl.cell._writer ^
    --collect-submodules faker ^
    --collect-submodules uvicorn ^
    "%SCRIPT_NAME%"

if errorlevel 1 (
    echo.
    echo [ОШИБКА] Сборка не удалась. Смотрите вывод выше.
    pause
    exit /b 1
)

echo.
echo ==================================================
echo  Готово! Файл: dist\%APP_NAME%.exe
echo  Рядом с exe будет создана папка "data"
echo  (users.json, progress.db, activity.log, uploads, masked)
echo ==================================================
echo.

if exist "dist" (
    explorer "dist"
)

pause
endlocal