@echo off
echo ==========================================
echo   My Microsoft 365 Dashboard
echo ==========================================
echo.
echo Iniciando servidor en puerto 8080...
echo.
echo Una vez iniciado, abre tu navegador en:
echo   http://localhost:8080
echo.
echo Presiona Ctrl+C para detener el servidor
echo.
echo ==========================================
echo.

python -m http.server 8080

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Error: Python no esta instalado o no esta en el PATH
    echo.
    echo Instala Python desde https://www.python.org/downloads/
    echo.
    pause
)
