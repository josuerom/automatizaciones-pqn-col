@echo off
title CONVERTIDOR DE PY A EXE
color 02

echo =====================================================
echo ===:    CONVERTIDOR DE PROGRAMAS PYTHON A EXE    :===
echo ===:               AUTOR: JOSUEROM               :===

:loop
echo =====================================================
echo.
cd /d %~dp0

set /p exeName=Nombre para el archivo .EXE: 
set /p pyFile=Nombre del programa .PY: 
echo.

python -m PyInstaller --onefile --noconsole --name="%exeName%" --icon="static\ico\icono.ico" src\main\%pyFile%.py
del "%pyFile%.spec"

echo.
echo Iniciando programa para prueba de funcionalidad...
.\dist\%exeName%.exe

echo.
pause
goto loop