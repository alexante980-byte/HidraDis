@echo off
echo ============================================
echo   Subiendo Hidra_Dis a GitHub
echo ============================================
echo.

cd /d "C:\Users\User\Desktop\PYTHON\TIC"

echo [1/5] Inicializando repositorio git...
git init

echo.
echo [2/5] Conectando con GitHub...
git remote remove origin 2>nul
git remote add origin https://github.com/alexante980-byte/Hidra_Dis.git
git branch -M main

echo.
echo [3/5] Agregando archivos...
git add .
git status

echo.
echo [4/5] Creando commit...
git commit -m "Hidra_Dis v1.0 - codigo fuente"

echo.
echo [5/5] Subiendo a GitHub...
echo (Se pediran tus credenciales de GitHub)
echo Usuario: alexante980@gmail.com
echo Password: usa tu token personal de GitHub
echo.
git push -u origin main

echo.
echo ============================================
echo   LISTO - Codigo subido a GitHub
echo   https://github.com/alexante980-byte/Hidra_Dis
echo ============================================
pause
