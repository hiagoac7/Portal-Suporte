@echo off
chcp 65001 >nul
title Remover inicio automatico do Portal de Suporte

net session >nul 2>&1
if %errorLevel% neq 0 (
    powershell -NoProfile -Command "Start-Process '%~f0' -Verb RunAs"
    exit /b
)

echo Removendo os DOIS metodos de inicio automatico...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Unregister-ScheduledTask -TaskName 'PortalSuporte' -Confirm:$false -ErrorAction SilentlyContinue; ^
   $s = [Environment]::GetFolderPath('Startup'); ^
   foreach ($f in @('PortalSuporte.lnk','iniciar_suporte_invisivel.vbs','portal_startup.vbs','iniciar_suporte_invisivel.lnk')) { ^
     $p = Join-Path $s $f; if (Test-Path $p) { Remove-Item $p -Force; Write-Host ('  removido: ' + $f) } }; ^
   Write-Host '[OK] Inicio automatico removido.' -ForegroundColor Green"
echo.
echo O portal continua funcionando pelo "LIGAR PORTAL.bat" na pasta principal.
echo.
pause
