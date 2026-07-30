@echo off
chcp 65001 >nul
title SIPSA Insumos - App Web

cd /d "%~dp0"

if not exist ".env" (
    echo [ERROR] No existe .env - copie .env.example a .env y complete las credenciales.
    pause
    exit /b 1
)

:: Puerto y credenciales se leen de .env (app.py hace load_dotenv)
set PORT=8080
for /f "usebackq tokens=1,2 delims==" %%A in (".env") do (
    if "%%A"=="PORT" set PORT=%%B
)

echo ============================================================
echo   SIPSA Insumos Agropecuarios - Interfaz Web
echo   http://localhost:%PORT%
echo   Credenciales: ver .env
echo ============================================================
echo.

:: Activar entorno virtual si existe
if exist ".venv\Scripts\activate.bat" (
    call .venv\Scripts\activate.bat
) else if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
)

:: Verificar que uvicorn este disponible
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
