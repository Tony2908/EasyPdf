@echo off
title Compilador Automático - Word a PDF Pro
color 0A
cls

echo =======================================================
echo    INICIANDO PROCESO DE COMPILACION (WORD TO PDF)
echo =======================================================
echo.

:: 1. Ir a la carpeta donde está este archivo (tu Escritorio)
cd /d "%~dp0"

:: 2. Limpiar versiones anteriores automáticamente
echo [+] Paso 1: Limpiando archivos viejos...
if exist "dist\EASYDOC SUITE.exe" del /f /q "dist\EASYDOC SUITE.exe"
if exist "build" rmdir /s /q "build"
if exist "EASYDOC SUITE.spec" del /f /q "EASYDOC SUITE.spec"

:: 3. Instalar/Actualizar PyInstaller y dependencias de UI por si acaso
echo [+] Paso 2: Verificando entorno...
pip install pyinstaller customtkinter tkinterdnd2 --upgrade --quiet

:: 4. Compilación de alta precisión
:: Usamos --collect-all para asegurar que CustomTkinter y DnD funcionen en el EXE
echo [+] Paso 3: Generando ejecutable (esto puede tardar un minuto)...
python -m PyInstaller --noconfirm --onefile --windowed ^
    --name "EASYDOC SUITE" ^
    --icon="icono.ico" ^
    --add-data "icono.ico;." ^
    --add-data "loading.gif;." ^
    --collect-all customtkinter ^
    --collect-all tkinterdnd2 ^
    "CodigoEasyPdf.py"

echo.
echo =======================================================
echo    PROCESO COMPLETADO EXITOSAMENTE
echo    Tu nuevo .exe esta en la carpeta 'dist'
echo =======================================================
echo.
pause