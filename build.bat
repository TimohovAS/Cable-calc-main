@echo off
chcp 65001 >nul
echo Сборка исполняемого файла CableCalc...
echo.

if not exist "cable_calc_gui.py" (
    echo Ошибка: файл cable_calc_gui.py не найден. Запустите build.bat из папки проекта.
    pause
    exit /b 1
)

pip show pyinstaller >nul 2>&1
if errorlevel 1 (
    echo Установка PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo Не удалось установить PyInstaller.
        pause
        exit /b 1
    )
)

pip show openpyxl >nul 2>&1
if errorlevel 1 (
    echo Установка openpyxl...
    pip install -r requirements.txt
)

echo.
echo Запуск PyInstaller...
pyinstaller --noconfirm cable_calc.spec

if errorlevel 1 (
    echo Сборка завершилась с ошибкой.
    pause
    exit /b 1
)

echo.
echo Готово. Исполняемый файл: dist\CableCalc.exe
echo Его можно копировать на любой компьютер с Windows и запускать без установки Python.
echo.
pause
