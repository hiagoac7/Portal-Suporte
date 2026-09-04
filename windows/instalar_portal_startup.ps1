$ErrorActionPreference = 'Stop'

$base     = Split-Path -Parent $PSScriptRoot
$pyw      = Join-Path $base 'venv\Scripts\pythonw.exe'
$app      = Join-Path $base 'app\app.py'
$wrapper  = Join-Path $PSScriptRoot 'iniciar_portal_oculto.pyw'
$taskName = 'PortalSuporte'

Write-Host "Verificando arquivos..." -ForegroundColor Cyan
if (-not (Test-Path $pyw))     { Write-Host "[ERRO] Nao encontrei: $pyw" -ForegroundColor Red; Read-Host "Enter para sair"; exit 1 }
if (-not (Test-Path $app))     { Write-Host "[ERRO] Nao encontrei: $app" -ForegroundColor Red; Read-Host "Enter para sair"; exit 1 }
if (-not (Test-Path $wrapper)) { Write-Host "[ERRO] Nao encontrei: $wrapper" -ForegroundColor Red; Read-Host "Enter para sair"; exit 1 }

$startup = [Environment]::GetFolderPath('Startup')
$lnkAntigo = Join-Path $startup 'PortalSuporte.lnk'
if (Test-Path $lnkAntigo) { Remove-Item $lnkAntigo -Force; Write-Host "  removido atalho de Inicializacao (metodo concorrente)" }

$arguments = '"' + $wrapper + '"'
$action  = New-ScheduledTaskAction -Execute $pyw -Argument $arguments -WorkingDirectory $base

$trigger = New-ScheduledTaskTrigger -AtLogOn
$trigger.Delay = 'PT20S'

$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
            -StartWhenAvailable -ExecutionTimeLimit ([TimeSpan]::Zero)

$principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
    -Settings $settings -Principal $principal `
    -Description 'Inicia o Portal de Suporte (Streamlit) no logon, oculto.' | Out-Null

Write-Host "Liberando a porta 8501 no Firewall do Windows..." -ForegroundColor Cyan
try {
    $fwRuleName = "Portal de Suporte (Streamlit 8501)"
    Remove-NetFirewallRule -DisplayName $fwRuleName -ErrorAction SilentlyContinue | Out-Null
    New-NetFirewallRule -DisplayName $fwRuleName -Direction Inbound -Protocol TCP `
        -LocalPort 8501 -Action Allow -Profile Private,Domain | Out-Null
    Write-Host "[OK] Regra de firewall criada (porta 8501, redes privadas)." -ForegroundColor Green
} catch {
    Write-Host "[AVISO] Nao consegui criar a regra de firewall automaticamente: $_" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "[OK] Tarefa '$taskName' criada com sucesso!" -ForegroundColor Green
Write-Host "     - Roda automaticamente ~20s depois do login, sem janela."
Write-Host "     - Portal neste PC: http://localhost:8501"

$ips = Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
    $_.IPAddress -notlike '169.254.*' -and $_.IPAddress -ne '127.0.0.1'
} | Select-Object -ExpandProperty IPAddress
foreach ($ip in $ips) { Write-Host "     - Outros PCs da rede: http://$ip`:8501" }

Write-Host ""
Write-Host "Testando agora (iniciando a tarefa)..." -ForegroundColor Cyan
Start-ScheduledTask -TaskName $taskName
Write-Host "Aguarde ~20s e abra http://localhost:8501 no navegador."
Write-Host ""
Read-Host "Enter para fechar"
