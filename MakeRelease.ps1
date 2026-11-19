# MakeRelease.ps1 - สร้าง update.json (SHA-256) และ ZIP ของเวอร์ชันที่เพิ่ง build
# ใช้: powershell -File MakeRelease.ps1 [-BaseUrl "https://raw.githubusercontent.com/<user>/<repo>/main/"]
# ไม่ใส่ BaseUrl = ใช้ลิงก์ GitHub Raw ของ Minos117/sd8 (installer/SD8System_Setup.exe)
param([string]$BaseUrl = "")
if (-not $BaseUrl) { $BaseUrl = "RELEASE_AUTO" }
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

$iss = Get-Content (Join-Path $root "installer.iss") -Raw -Encoding UTF8
if ($iss -notmatch '#define\s+MyAppVersion\s+"([^"]+)"') { throw "อ่าน MyAppVersion จาก installer.iss ไม่ได้" }
$ver = $Matches[1]

$setup = Join-Path $root "Output\SD8System_Setup.exe"
if (-not (Test-Path $setup)) { throw "ไม่พบ $setup - ต้องสร้างตัวติดตั้งด้วย Inno Setup ก่อน" }

$hash = (Get-FileHash $setup -Algorithm SHA256).Hash.ToLower()
$size = (Get-Item $setup).Length
$name = "SD8System_Beta_$ver"
$dir  = Join-Path $root "release\$name"
if (Test-Path $dir) { Remove-Item $dir -Recurse -Force }
New-Item -ItemType Directory -Path $dir | Out-Null
Copy-Item $setup $dir

$json = [ordered]@{
    setup_rev = 1
    version  = $ver
    channel  = "beta"
    file     = "SD8System_Setup.exe"
    url      = if ($BaseUrl -eq "RELEASE_AUTO") { "https://raw.githubusercontent.com/Minos117/sd8/main/installer/SD8System_Setup.exe" } else { $BaseUrl.TrimEnd('/') + "/SD8System_Setup.exe" }
    sha256   = $hash
    size     = $size
    released = (Get-Date -Format "yyyy-MM-dd")
} | ConvertTo-Json
[IO.File]::WriteAllText((Join-Path $dir "update.json"), $json, (New-Object Text.UTF8Encoding($false)))
Copy-Item (Join-Path $dir "update.json") (Join-Path $root "release\update.json") -Force

$zip = Join-Path $root "release\$name.zip"
if (Test-Path $zip) { Remove-Item $zip -Force }
Compress-Archive -Path (Join-Path $dir "*") -DestinationPath $zip
Write-Host "OK  version=$ver  sha256=$hash"
Write-Host "ZIP: $zip"
