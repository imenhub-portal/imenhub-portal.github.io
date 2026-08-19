# Deploys Code.gs to Apps Script and republishes the SAME web app URL.
#
# Why this exists: GitHub Pages serves index.html straight from the repo, so
# a push is the deploy. Apps Script does not read the repo at all - it runs
# only the copy in its own editor, frozen at the last published version - so
# Code.gs had to be pasted in by hand every time. This closes that gap.
#
# It deploys by deployment ID on purpose. "New deployment" in the Apps Script
# UI mints a brand-new /exec URL and retires the old one, which silently
# breaks the published page; updating the existing deployment keeps the URL
# that index.html already points at.

$ErrorActionPreference = 'Stop'
Set-Location (Join-Path $PSScriptRoot '..')

# The deployment behind the /exec URL baked into index.html.
$DeploymentId = 'AKfycbyY2fSJbt6FMYtH8fIun8D_O4JWbhBZlN9hfPG-QVLs0T6m0OePE5_WxVr7XSRawaGE'

if (-not (Test-Path '.clasp.json')) {
  Write-Host ""
  Write-Host "  .clasp.json tiada - jalankan 1-SETUP-clasp.cmd dahulu (sekali sahaja)." -ForegroundColor Yellow
  Write-Host ""
  exit 1
}

Write-Host ""
Write-Host "  Menghantar Code.gs ke Apps Script..." -ForegroundColor Cyan
npx --yes @google/clasp push --force
if ($LASTEXITCODE -ne 0) { throw "clasp push gagal" }

Write-Host ""
Write-Host "  Menerbitkan versi baharu (URL kekal sama)..." -ForegroundColor Cyan
$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
npx --yes @google/clasp deploy --deploymentId $DeploymentId --description "auto $stamp"
if ($LASTEXITCODE -ne 0) { throw "clasp deploy gagal" }

# Deploying is not the same as working. The failure this tool exists to
# prevent - a deployed script with no doPost - looks completely fine at the
# deploy step and only shows up when the page makes its first API call. So
# ask the server directly before claiming success.
Write-Host ""
Write-Host "  Mengesahkan pelayan menjawab..." -ForegroundColor Cyan
$uri = "https://script.google.com/macros/s/$DeploymentId/exec"
$body = '{"fn":"handleAction","args":["Ping",{}]}'
try {
  $r = Invoke-RestMethod -Uri $uri -Method Post -ContentType 'text/plain;charset=utf-8' -Body $body -TimeoutSec 90
  if ($r.ok -and $r.result.result -eq 'pong') {
    Write-Host ""
    Write-Host "  SIAP. Pelayan menjawab 'pong' - doPost berfungsi." -ForegroundColor Green
    Write-Host ""
  } else {
    Write-Host ""
    Write-Host "  Deploy selesai, TETAPI pelayan menjawab luar jangka:" -ForegroundColor Yellow
    $r | ConvertTo-Json -Depth 5
    Write-Host ""
  }
} catch {
  Write-Host ""
  Write-Host "  Deploy selesai, TETAPI ujian POST gagal:" -ForegroundColor Red
  Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
  Write-Host ""
  Write-Host "  Jika ini 'Page not found', doPost tiada dalam kod yang di-deploy." -ForegroundColor Red
  Write-Host ""
}
