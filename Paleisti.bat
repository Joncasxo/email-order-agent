@echo off
cd /d "%~dp0"
where pythonw >nul 2>nul && (start "" pythonw app.py & exit /b)
where pyw >nul 2>nul && (start "" pyw app.py & exit /b)
python app.py
