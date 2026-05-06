@echo off
chcp 65001 >nul
title CMDB SENA - Lanzador de Inventario
color 0A

echo ============================================
echo   CMDB SENA CCYS - Lanzador de Inventario
echo ============================================
echo.

set "SCRIPT_DIR=%~dp0"
set "PS1_FILE=%SCRIPT_DIR%Inventario.ps1"

if not exist "%PS1_FILE%" (
    echo [ERROR] No se encontro el archivo Inventario.ps1
    echo.
    echo Asegurate de que este archivo .bat este en la misma carpeta que Inventario.ps1
    echo.
    pause
    exit /b 1
)

echo [1/3] Desbloqueando archivo si fue descargado de internet...
powershell -NoProfile -Command "Unblock-File -Path '%PS1_FILE%'" 2>nul
echo       Listo.
echo.

echo [2/3] Ejecutando inventario con permisos necesarios...
echo.
echo ============================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1_FILE%"

if %errorlevel% neq 0 (
    echo.
    echo ============================================
    echo [ERROR] El script termino con codigo %errorlevel%
    echo.
    echo Si ves un error de "ExecutionPolicy", contacta al administrador.
    echo.
    pause
    exit /b %errorlevel%
)

echo.
echo ============================================
echo [OK] Proceso completado.
echo.
pause
