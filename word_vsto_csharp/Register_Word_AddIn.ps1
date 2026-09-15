# PowerShell script to Register Kruti Dev Word Add-in for Microsoft Word LTSC 2024 / 2016-365
# Uses HKCU user-level COM registration (No Administrator rights required!)

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dllPath = (Join-Path $scriptDir "KrutiDevWordAddIn.dll").Replace("\", "/")
$clsid = "{A789B1C2-3D4E-5F6A-7B8C-9D0E1F2A3B4C}"
$progId = "KrutiDevWordAddIn.Connect"
$codebaseUri = "file:///" + $dllPath

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host " Registering Kruti Dev 010 Add-In for Microsoft Word 2024" -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Cyan

# 1. Register ProgID in HKCU:\Software\Classes
Write-Host "`n[1/3] Registering ProgID in HKCU Classes..." -ForegroundColor Green
$progIdPath = "HKCU:\Software\Classes\$progId"
if (-not (Test-Path $progIdPath)) { New-Item -Path $progIdPath -Force | Out-Null }
Set-ItemProperty -Path $progIdPath -Name "(Default)" -Value "KrutiDevWordAddIn.Connect"
$clsidSubPath = Join-Path $progIdPath "CLSID"
if (-not (Test-Path $clsidSubPath)) { New-Item -Path $clsidSubPath -Force | Out-Null }
Set-ItemProperty -Path $clsidSubPath -Name "(Default)" -Value $clsid

# 2. Register CLSID in HKCU:\Software\Classes\CLSID
Write-Host "[2/3] Registering CLSID in HKCU..." -ForegroundColor Green
$clsidPath = "HKCU:\Software\Classes\CLSID\$clsid"
if (-not (Test-Path $clsidPath)) { New-Item -Path $clsidPath -Force | Out-Null }
Set-ItemProperty -Path $clsidPath -Name "(Default)" -Value "KrutiDevWordAddIn.Connect"

$inprocPath = Join-Path $clsidPath "InprocServer32"
if (-not (Test-Path $inprocPath)) { New-Item -Path $inprocPath -Force | Out-Null }
Set-ItemProperty -Path $inprocPath -Name "(Default)" -Value "mscoree.dll"
Set-ItemProperty -Path $inprocPath -Name "ThreadingModel" -Value "Both"
Set-ItemProperty -Path $inprocPath -Name "Class" -Value "KrutiDevWordAddIn.Connect"
Set-ItemProperty -Path $inprocPath -Name "Assembly" -Value "KrutiDevWordAddIn, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null"
Set-ItemProperty -Path $inprocPath -Name "RuntimeVersion" -Value "v4.0.30319"
Set-ItemProperty -Path $inprocPath -Name "CodeBase" -Value $codebaseUri

$progIdSub = Join-Path $clsidPath "ProgID"
if (-not (Test-Path $progIdSub)) { New-Item -Path $progIdSub -Force | Out-Null }
Set-ItemProperty -Path $progIdSub -Name "(Default)" -Value $progId

# 3. Register Word Add-in in HKCU:\Software\Microsoft\Office\Word\Addins
Write-Host "[3/3] Registering Add-in into Microsoft Word 2024..." -ForegroundColor Green
$wordAddinPath = "HKCU:\Software\Microsoft\Office\Word\Addins\$progId"
if (-not (Test-Path $wordAddinPath)) { New-Item -Path $wordAddinPath -Force | Out-Null }
Set-ItemProperty -Path $wordAddinPath -Name "FriendlyName" -Value "शिक्षा प्रारूपक (कृति देव 010 सहायक)"
Set-ItemProperty -Path $wordAddinPath -Name "Description" -Value "Jharkhand Education Dept Kruti Dev 010 Drafting & Spellcheck Add-in"
Set-ItemProperty -Path $wordAddinPath -Name "LoadBehavior" -Value 3 -Type DWord
Set-ItemProperty -Path $wordAddinPath -Name "CommandLineSafe" -Value 0 -Type DWord

Write-Host "`n==========================================================" -ForegroundColor Cyan
Write-Host " SUCCESS! Kruti Dev Word Add-in is now registered!" -ForegroundColor Green
Write-Host " When you open Microsoft Word, you will see the" -ForegroundColor Yellow
Write-Host " '[शिक्षा प्रारूपक (कृति देव)]' tab on top in the Word Ribbon." -ForegroundColor Yellow
Write-Host "==========================================================" -ForegroundColor Cyan
