@echo off
chcp 65001 >nul
title SIPSA Insumos — App Web

cd /d "%~dp0"

:: Credenciales (modificar antes de usar en producción)
set INSUMOS_USER=sipsa
set INSUMOS_PASS=insumos2024

:: Puerto
set PORT=8080

echo ============================================================
echo   SIPSA Insumos Agropecuarios — Interfaz Web
echo   http://localhost:%PORT%
echo   Usuario: %INSUMOS_USER%
echo ============================================================
echo.

:: Activar entorno virtual si existe
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
) else if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

:: Verificar que uvicorn esté disponible
uvicorn --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] uvicorn no encontrado. Ejecute: pip install uvicorn fastapi jinja2
    pause
    exit /b 1
)

echo Iniciando servidor en http://localhost:%PORT% ...
echo Presione Ctrl+C para detener.
echo.

uvicorn app:app --host 0.0.0.0 --port %PORT% --reload

pause
