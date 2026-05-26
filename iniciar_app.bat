@echo off
echo Iniciando entorno virtual...
if not exist "venv\Scripts\activate" (
    echo El entorno virtual no existe. Por favor ejecuta instalacion.bat primero.
    pause
    exit
)
call venv\Scripts\activate
echo Iniciando aplicacion web...
streamlit run app.py
pause
