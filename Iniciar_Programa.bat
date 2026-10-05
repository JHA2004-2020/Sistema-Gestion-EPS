@echo off
title Sistema de EPS - Nueva EPS y Savia Salud
echo ===================================================
echo Iniciando el Sistema de Gestion de EPS...
echo ===================================================

:: 1. Se asegura de estar en la carpeta donde esta el .bat
cd /d "%~dp0"

:: 2. Verifica si el entorno virtual tiene Python adentro. Si no, lo crea de cero.
IF NOT EXIST ".venv\Scripts\python.exe" (
    echo [Info] Entorno virtual no encontrado o incompleto. Creando...
    python -m venv .venv
    
    echo [Info] Instalando librerias necesarias...
    .venv\Scripts\python.exe -m pip install --upgrade pip
    .venv\Scripts\python.exe -m pip install streamlit pandas openpyxl
) ELSE (
    echo [Info] Entorno virtual detectado correctamente.
)

:: 3. Entra a la carpeta del proyecto
echo [Info] Entrando a la carpeta 202...
cd 202

:: 4. Ejecuta la aplicacion de forma directa e infalible
echo [Info] Ejecutando la aplicacion...
..\.venv\Scripts\python.exe -m streamlit run app.py

:: Si algo falla, el programa se pausara aqui para poder leer el error
pause