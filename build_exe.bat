@echo off
setlocal

REM Собирает main.py в Windows .exe через PyInstaller.
REM Запускать на Windows (CMD), из папки проекта.

cd /d "%~dp0"

echo.
echo [1/3] Установка зависимостей...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo [2/3] Очистка старых сборок...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "PriceList.spec" del /q "PriceList.spec"

echo.
echo [3/3] Сборка .exe...
pyinstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "PriceList" ^
  "main.py"

echo.
echo Готово. Ищите файл: dist\PriceList.exe
pause

