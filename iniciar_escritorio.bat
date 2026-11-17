@echo off
title VeriDocs - Aplicacion de Escritorio
cd /d "%~dp0"

echo Iniciando VeriDocs Desktop...
start "" pythonw app_gui.py
