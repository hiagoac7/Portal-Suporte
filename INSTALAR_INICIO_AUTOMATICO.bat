@echo off
chcp 65001 >nul
title Instalar inicio automatico do Portal de Suporte

:: Pede elevacao (administrador) automaticamente
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Solicitando permissao de administrador...
    powershell -NoProfile -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ============================================================
echo   Instalando inicio automatico do Portal de Suporte
echo ============================================================
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar_portal_startup.ps1"
