@echo off
title Pizarra Digital de Produccion - Servidor en Tiempo Real
cd /d "%~dp0"
echo Iniciando servidor del Tablero Kanban de Produccion...
python app.py --open
if errorlevel 1 (
    echo.
    echo [ERROR] No se pudo iniciar con "python". Intentando con "py"...
    py app.py --open
)
pause
