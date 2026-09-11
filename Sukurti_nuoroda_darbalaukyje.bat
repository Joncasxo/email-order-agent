@echo off
rem Sukuria "Uzsakymu perziura" nuoroda darbalaukyje (paleisti viena karta)
set "T=%~dp0Uzsakymu perziura.exe"
if not exist "%T%" set "T=%~dp0Perziura.bat"
powershell -NoProfile -Command "$ws = New-Object -ComObject WScript.Shell; $sc = $ws.CreateShortcut([Environment]::GetFolderPath('Desktop') + '\Uzsakymu perziura.lnk'); $sc.TargetPath = '%T%'; $sc.WorkingDirectory = '%~dp0'; $sc.Description = 'Uzsakymu stebejimas'; $sc.Save()"
echo Nuoroda sukurta darbalaukyje. Galite uzdaryti si langa.
pause
