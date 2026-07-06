@echo off
chcp 65001 > nul
title Compilador HidraDis

echo ============================================
echo   Compilando HidraDis.exe con PyInstaller
echo ============================================
echo.

:: Ir a la carpeta del script
cd /d "%~dp0"

:: Limpiar compilaciones anteriores
echo [1/3] Limpiando compilaciones anteriores...
if exist "build" rmdir /s /q "build"
if exist "dist\HidraDis.exe" del /q "dist\HidraDis.exe"

:: Compilar
echo [2/3] Compilando HidraDis.py ...
pyinstaller HidraDis.spec

:: Verificar resultado
echo.
echo [3/3] Verificando resultado...
if exist "dist\HidraDis.exe" (
    echo.
    echo  EXITO: HidraDis.exe generado correctamente.
    echo  Ubicacion: %~dp0dist\HidraDis.exe
) else (
    echo.
    echo  ERROR: No se genero el ejecutable. Revisa los mensajes anteriores.
)

echo.
pause
