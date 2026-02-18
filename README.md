# Cable-calc

Расчёт кабелей по IEC 60364 (GUI на tkinter).

## Запуск без установки Python

Собранный исполняемый файл лежит в папке **dist**: `dist\CableCalc.exe`. Скопируйте его на любой компьютер с Windows и запускайте двойным щелчком — Python устанавливать не нужно.

## Сборка исполняемого файла

1. Установите зависимости: `pip install -r requirements.txt` и `pip install pyinstaller`
2. Запустите `build.bat` или выполните: `pyinstaller --noconfirm cable_calc.spec`
3. Готовый файл появится в `dist\CableCalc.exe`