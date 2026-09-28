@echo off
py -m pip install -r requirements.txt
if errorlevel 1 exit /b 1
py -m PyInstaller --noconfirm --clean --onefile --windowed --name "중앙부처동향" app.py
if errorlevel 1 exit /b 1
echo EXE: dist\중앙부처동향.exe
pause
