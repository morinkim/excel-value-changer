@echo off
echo ========================================
echo Excel Value Changer - EXE Build
echo ========================================
echo.

echo [1/2] Installing packages...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo Package install failed!
    pause
    exit /b 1
)

echo.
echo [2/2] Building EXE...
python -m PyInstaller --onefile --windowed --name "ExcelValueChanger" excel_changer.py
if errorlevel 1 (
    echo EXE build failed!
    pause
    exit /b 1
)

echo.
echo ========================================
echo Build complete!
echo EXE location: dist\ExcelValueChanger.exe
echo ========================================
echo.
pause
