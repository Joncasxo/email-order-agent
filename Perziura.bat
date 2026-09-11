@echo off
cd /d "%~dp0"
where pyw >nul 2>nul && (start "" pyw perziura.py & exit /b)
where pythonw >nul 2>nul && (start "" pythonw perziura.py & exit /b)
where py >nul 2>nul && (py perziura.py & exit /b)
python perziura.py
