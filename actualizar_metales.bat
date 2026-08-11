@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo ========================================================
echo   1. CONSULTANDO COTIZACIONES Y ACTUALIZANDO DATOS...
echo ========================================================
python actualizar_metales.py

echo.
echo ========================================================
echo   2. PUBLICANDO DATOS ACTUALIZADOS EN GITHUB PAGES...
echo ========================================================
git add .
git commit -m "Actualización diaria de cotizaciones - %date%"
git push

echo.
echo ========================================================
echo   ¡PROCESO COMPLETADO EXITOSAMENTE!
echo   La web interactiva ya está al día en tu enlace de GitHub.
echo ========================================================
pause
