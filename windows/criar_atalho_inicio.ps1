$ErrorActionPreference = 'Stop'

$base    = Split-Path -Parent $PSScriptRoot
$pyw     = Join-Path $base 'venv\Scripts\pythonw.exe'
$app     = Join-Path $base 'app\app.py'
$wrapper = Join-Path $PSScriptRoot 'iniciar_portal_oculto.pyw'

if (-not (Test-Path $pyw))     { Write-Host "[ERRO] Nao encontrei: $pyw" -ForegroundColor Red; Read-Host "Enter para sair"; exit 1 }
if (-not (Test-Path $app))     { Write-Host "[ERRO] Nao encontrei: $app" -ForegroundColor Red; Read-Host "Enter para sair"; exit 1 }
if (-not (Test-Path $wrapper)) { Write-Host "[ERRO] Nao encontrei: $wrapper" -ForegroundColor Red; Read-Host "Enter para sair"; exit 1 }

Write-Host "Limpando metodos antigos..." -ForegroundColor Cyan
Unregister-ScheduledTask -TaskName 'PortalSuporte' -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
$startup = [Environment]::GetFolderPath('Startup')
foreach ($old in @('iniciar_suporte_invisivel.vbs','portal_startup.vbs','iniciar_suporte_invisivel.lnk')) {
    $p = Join-Path $startup $old
    if (Test-Path $p) { Remove-Item $p -Force; Write-Host "  removido: $old" }
}

$lnkPath = Join-Path $startup 'PortalSuporte.lnk'
$wsh = New-Object -ComObject WScript.Shell
$lnk = $wsh.CreateShortcut($lnkPath)
$lnk.TargetPath       = $pyw
$lnk.Arguments        = '"' + $wrapper + '"'
$lnk.WorkingDirectory = $base
$lnk.WindowStyle      = 7
$lnk.Description      = 'Portal de Suporte'
$lnk.Save()

Write-Host ""
Write-Host "[OK] Atalho criado em:" -ForegroundColor Green
Write-Host "     $lnkPath"
Write-Host "     apontando para: $wrapper"
Write-Host ""
Write-Host "Testando agora (iniciando o portal)..." -ForegroundColor Cyan
Start-Process -FilePath $pyw -ArgumentList ('"' + $wrapper + '"') -WorkingDirectory $base -WindowStyle Hidden
Write-Host "Aguarde ~20s e abra http://localhost:8501"
Write-Host ""

$fwRuleName = "Portal de Suporte (Streamlit 8501)"
$fwOk = Get-NetFirewallRule -DisplayName $fwRuleName -ErrorAction SilentlyContinue
if (-not $fwOk) {
    Write-Host "[AVISO] A porta 8501 ainda nao esta liberada no Firewall do Windows." -ForegroundColor Yellow
    Write-Host "        Sem isso, OUTROS PCs da rede nao conseguem abrir o portal (so este PC consegue)."
    Write-Host "        Rode uma vez, em PowerShell como Administrador:"
    Write-Host "        New-NetFirewallRule -DisplayName '$fwRuleName' -Direction Inbound -Protocol TCP -LocalPort 8501 -Action Allow -Profile Private"
    Write-Host ""
}

Write-Host "Pronto. Da proxima vez que ligar o notebook e fizer login, o portal sobe sozinho." -ForegroundColor Green
Write-Host ""
Read-Host "Enter para fechar"
