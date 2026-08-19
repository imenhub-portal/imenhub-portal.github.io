# One-time clasp setup. After this, deploying is just tools\deploy.ps1.
$ErrorActionPreference = 'Stop'
Set-Location $PSScriptRoot\..

Write-Host ""
Write-Host "  LANGKAH 1 — benarkan Apps Script API (sekali sahaja)" -ForegroundColor Cyan
Write-Host "  Buka: https://script.google.com/home/usersettings"
Write-Host "  Hidupkan 'Google Apps Script API'. Tekan Enter bila siap."
Read-Host

Write-Host ""
Write-Host "  LANGKAH 2 — log masuk clasp" -ForegroundColor Cyan
Write-Host "  Pelayar akan terbuka. Pilih akaun imenmakmal@gmail.com."
npx --yes @google/clasp login
if ($LASTEXITCODE -ne 0) { throw "clasp login gagal" }

Write-Host ""
Write-Host "  LANGKAH 3 — Script ID" -ForegroundColor Cyan
Write-Host "  Dalam editor Apps Script: Project Settings (ikon gear) -> Script ID."
$id = Read-Host "  Paste Script ID di sini"
if ([string]::IsNullOrWhiteSpace($id)) { throw "Script ID kosong" }

@{ scriptId = $id.Trim(); rootDir = "." } | ConvertTo-Json | Set-Content '.clasp.json' -Encoding utf8
Write-Host ""
Write-Host "  Siap. Mulai sekarang, deploy dengan: tools\deploy.ps1" -ForegroundColor Green
Write-Host ""
