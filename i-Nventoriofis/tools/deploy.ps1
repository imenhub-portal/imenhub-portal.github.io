# Deploys Code.gs to Apps Script and republishes the SAME web app URL.
#
# Why this exists: GitHub Pages serves index.html straight from the repo, so
# a push is the deploy. Apps Script does not read the repo at all — it runs
# only the copy in its own editor, frozen at the last published version — so
# Code.gs had to be pasted in by hand every time. This closes that gap.
#
# It deploys by deployment ID on purpose. "New deployment" in the Apps Script
# UI mints a brand-new /exec URL and retires the old one, which silently
# breaks the published page; updating the existing deployment keeps the URL
# that index.html already points at.

$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot\..

# The deployment behind the /exec URL baked into index.html.
$DeploymentId = 'AKfycbyY2fSJbt6FMYtH8fIun8D_O4JWbhBZlN9hfPG-QVLs0T6m0OePE5_WxVr7XSRawaGE'

if (-not (Test-Path '.clasp.json')) {
  Write-Host ""
  Write-Host "  .clasp.json tiada — jalankan tools\setup-clasp.ps1 dahulu (sekali sahaja)." -ForegroundColor Yellow
  Write-Host ""
  exit 1
}

Write-Host ""
Write-Host "  Menghantar Code.gs ke Apps Script..." -ForegroundColor Cyan
npx --yes @google/clasp push --force
if ($LASTEXITCODE -ne 0) { throw "clasp push gagal" }

Write-Host ""
Write-Host "  Menerbitkan versi baharu (URL kekal sama)..." -ForegroundColor Cyan
npx --yes @google/clasp deploy --deploymentId $DeploymentId --description "auto $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
if ($LASTEXITCODE -ne 0) { throw "clasp deploy gagal" }

Write-Host ""
Write-Host "  Mengesahkan pelayan menjawab..." -ForegroundColor Cyan
$body = '{"fn":"handleAction","args":["Ping",{}]}'
try {
  $r = Invoke-RestMethod -Uri "https://script.google.com/macros/s/$DeploymentId/exec" `
                         -Method Post -ContentType 'text/plain;charset=utf-8' -Body $body -TimeoutSec 60
  if ($r.ok -and $r.result.result -eq 'pong') {
    Write-Host ""
    Write-Host "  SIAP. Pelayan menjawab 'pong' — doPost berfungsi." -ForegroundColor Green
    Write-Host ""
  } else {
    Write-Host ""
    Write-Host "  Deploy selesai, tetapi pelayan menjawab luar jangka:" -ForegroundColor Yellow
    $r | ConvertTo-Json -Depth 5
  }
} catch {
  Write-Host ""
  Write-Host "  Deploy selesai, tetapi ujian POST gagal: $_" -ForegroundColor Red
  Write-Host "  Jika ini 'Page not found', doPost tiada dalam kod yang di-deploy." -ForegroundColor Red
}
