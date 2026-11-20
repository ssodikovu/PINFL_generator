@echo off
chcp 65001 >nul
python -m pip install --upgrade pip
python -m pip install openpyxl pyinstaller
python -m PyInstaller --noconfirm --clean --onefile --windowed --name PINFL_Test_Generator pinfl_test_generator.py
echo.
echo Готово. EXE находится в папке dist.
pause
