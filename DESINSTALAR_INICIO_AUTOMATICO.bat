@echo off
chcp 65001 >nul
title Remover inicio automatico do Portal de Suporte

net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -NoProfile -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo Removendo a tarefa agendada 'PortalSuporte'...
powershell -NoProfile -ExecutionPolicy Bypass -Command "Unregister-ScheduledTask -TaskName 'PortalSuporte' -Confirm:$false -ErrorAction SilentlyContinue; Write-Host '[OK] Removido (se existia).' -ForegroundColor Green"
echo.
pause
