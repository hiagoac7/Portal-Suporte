@echo off
chcp 65001 >nul
title Instalar inicio automatico (Tarefa Agendada)

net session >nul 2>&1
if %errorLevel% neq 0 (
    echo Solicitando permissao de administrador...
    powershell -NoProfile -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo ============================================================
echo   Portal de Suporte - inicio automatico via Tarefa Agendada
echo ============================================================
echo.
echo Use este SO SE o atalho de Inicializacao nao funcionar no seu PC.
echo Ele tambem libera a porta 8501 no Firewall do Windows.
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0instalar_portal_startup.ps1"
