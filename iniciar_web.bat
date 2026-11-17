@echo off
title VeriDocs - Servidor Web MES
cd /d "%~dp0"

echo ========================================================
echo   Iniciando VeriDocs (Modo Web)
echo ========================================================
echo Abriendo servidor web y navegador...
echo Para acceder desde otros PCs de la red, usa la IP mostrada en pantalla.
echo.

python app_web.py
pause
