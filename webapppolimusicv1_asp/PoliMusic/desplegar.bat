@echo off
REM ===========================================================================
REM  PoliMusic - Despliegue en IIS
REM  Haga doble clic en este archivo, o ejecutelo desde una consola.
REM  El script solicitara privilegios de administrador (aviso de UAC).
REM ===========================================================================
setlocal

set SCRIPT=%~dp0desplegar.ps1

if not exist "%SCRIPT%" (
    echo.
    echo  ERROR: no se encontro desplegar.ps1 en la misma carpeta.
    echo.
    pause
    exit /b 1
)

echo.
echo  Iniciando el despliegue de PoliMusic en IIS...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*

echo.
pause
endlocal
