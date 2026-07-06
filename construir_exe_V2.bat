@echo off
chcp 65001 >nul
echo ============================================================
echo   HIDRA DIS V2 - Generador de ejecutable independiente
echo ============================================================
cd /d "%~dp0"

echo Carpeta de trabajo actual:
echo %cd%
echo.

if not exist "Hidra_Dis_V2.py" (
    echo ============================================================
    echo   ERROR: No se encontro Hidra_Dis_V2.py en esta carpeta:
    echo   %cd%
    echo.
    echo   Este archivo .bat debe estar guardado EN LA MISMA CARPETA
    echo   que Hidra_Dis_V2.py, Imagenes\ e icono_Hidra_Dis.ico
    echo   ^(C:\Users\User\Desktop\PYTHON\TIC^).
    echo   Si lo moviste o copiaste a otro lugar, vuelve a colocarlo
    echo   ahi y ejecutalo de nuevo con doble clic.
    echo ============================================================
    pause
    exit /b 1
)

where pyinstaller >nul 2>nul
if errorlevel 1 (
    echo PyInstaller no esta instalado. Instalando...
    pip install pyinstaller
)

echo.
echo Compilando Hidra_Dis_V2.py ...
echo.

pyinstaller --noconfirm --onefile --windowed ^
    --name "HidraDis_V2" ^
    --icon "Iconos\icono_Hidra_Dis.ico" ^
    --add-data "Imagenes;Imagenes" ^
    --add-data "Iconos\icono_Hidra_Dis.ico;." ^
    --collect-all customtkinter ^
    --noupx ^
    Hidra_Dis_V2.py

echo.
if exist "dist\HidraDis_V2.exe" (
    echo ============================================================
    echo   LISTO: dist\HidraDis_V2.exe
    echo   Ese archivo es independiente: se puede copiar y ejecutar
    echo   en otra computadora con Windows SIN tener Python instalado.
    echo ============================================================
) else (
    echo Hubo un error durante la compilacion. Revisa el mensaje anterior.
)
pause
