@echo off
echo Building Description Splitter executable...
echo.

REM Tjek om PyInstaller er installeret
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

REM Ryd op i tidligere builds
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del *.spec

REM Byg .exe filen
echo Building executable...
pyinstaller --onefile ^
    --noconsole ^
    --name "DescriptionSplitter" ^
    --hidden-import=openpyxl ^
    --hidden-import=tkinter ^
    --hidden-import=tkinter.ttk ^
    --hidden-import=tkinter.filedialog ^
    --hidden-import=tkinter.messagebox ^
    --hidden-import=openpyxl.styles ^
    --hidden-import=openpyxl.utils ^
    --clean ^
    description_splitter.py

if exist "dist\DescriptionSplitter.exe" (
    echo.
    echo Build successful!
    echo Executable created: dist\DescriptionSplitter.exe
    echo.
    echo To avoid Windows Defender issues:
    echo 1. Right-click the .exe file
    echo 2. Select "Properties"
    echo 3. Check "Unblock" at the bottom
    echo 4. Click "Apply" and "OK"
    echo.
    pause
) else (
    echo Build failed!
    pause
) 