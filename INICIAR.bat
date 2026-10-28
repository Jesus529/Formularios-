@echo off
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
  echo Creando entorno virtual...
  py -m venv .venv
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe app.py
pause
