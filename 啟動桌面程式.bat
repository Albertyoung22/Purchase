@echo off
cd /d "%~dp0"
if exist "C:\Users\USER\AppData\Local\Programs\Python\Python311\python.exe" (
    "C:\Users\USER\AppData\Local\Programs\Python\Python311\python.exe" app.py
) else (
    python app.py
)
pause
