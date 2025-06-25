@echo off
REM Change directory to the app folder
cd ..\app

REM Run PyInstaller to create a one-folder executable for launcher.py
pyinstaller --onefile --windowed launcher.py

REM Move the generated exe from dist to the main dist folder
move dist\launcher.exe ..\dist\launcher.exe

REM Cleanup build files
rd /s /q build
rd /s /q dist
del launcher.spec

echo Launcher.exe build complete!
pause
