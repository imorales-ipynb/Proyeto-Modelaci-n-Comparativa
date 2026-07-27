@echo off
echo Iniciando entorno virtual...
if not exist "C:\dev\venvs\proyecto-modelacion-comparativa\Scripts\activate" (
    echo El entorno virtual no existe en C:\dev\venvs\proyecto-modelacion-comparativa
    pause
    exit
)
call C:\dev\venvs\proyecto-modelacion-comparativa\Scripts\activate
echo Iniciando aplicacion web...
streamlit run app.py
pause
