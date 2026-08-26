@echo off
setlocal
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m PyInstaller --noconfirm --clean --onefile --windowed --name HolyokeTransferConverter desktop_app.py

echo.
echo Build complete.
echo The EXE is in the dist folder:
echo dist\HolyokeTransferConverter.exe
pause
