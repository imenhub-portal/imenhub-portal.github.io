# One-time clasp setup. After this, deploying is just 2-DEPLOY.cmd.
#
# Run interactively: it opens a browser for Google sign-in and asks for the
# Script ID, so it cannot be automated from here.

$ErrorActionPreference = 'Stop'
Set-Location (Join-Path $PSScriptRoot '..')

Write-Host ""
Write-Host "  LANGKAH 1 - benarkan Apps Script API (sekali sahaja)" -ForegroundColor Cyan
Write-Host "  Buka: https://script.google.com/home/usersettings"
Write-Host "  Hidupkan 'Google Apps Script API'."
Write-Host ""
Read-Host "  Tekan Enter bila siap"

Write-Host ""
Write-Host "  LANGKAH 2 - log masuk clasp" -ForegroundColor Cyan
Write-Host "  Pelayar akan terbuka. Pilih akaun imenmakmal@gmail.com."
Write-Host ""
npx --yes @google/clasp login
if ($LASTEXITCODE -ne 0) { throw "clasp login gagal" }

Write-Host ""
Write-Host "  LANGKAH 3 - Script ID" -ForegroundColor Cyan
Write-Host "  Dalam editor Apps Script: Project Settings (ikon gear), cari 'Script ID'."
Write-Host ""
$id = Read-Host "  Paste Script ID di sini"
if ([string]::IsNullOrWhiteSpace($id)) { throw "Script ID kosong" }

$cfg = [ordered]@{ scriptId = $id.Trim(); rootDir = '.' }
$cfg | ConvertTo-Json | Set-Content '.clasp.json' -Encoding utf8

Write-Host ""
Write-Host "  Siap. Mulai sekarang, deploy dengan klik dua kali: 2-DEPLOY.cmd" -ForegroundColor Green
Write-Host ""
