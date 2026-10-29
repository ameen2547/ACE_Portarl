[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$DeploymentId,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^\d+$')]
    [string]$VersionNumber
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$configPath = Join-Path $projectRoot '.clasp.json'

if (-not (Get-Command clasp -ErrorAction SilentlyContinue)) {
    throw 'ไม่พบ clasp กรุณาติดตั้งด้วย: npm install --global @google/clasp'
}
if (-not (Test-Path -LiteralPath $configPath)) {
    throw 'ยังไม่ได้ตั้งค่า Script ID กรุณารัน scripts/setup-gas.ps1 ก่อน'
}

$config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
if ([string]::IsNullOrWhiteSpace([string]$config.scriptId)) {
    throw '.clasp.json ไม่มีค่า scriptId กรุณาตรวจสอบการตั้งค่า'
}

Push-Location $projectRoot
try {
    & clasp show-authorized-user *> $null
    if ($LASTEXITCODE -ne 0) {
        throw 'ยังไม่ได้เข้าสู่ระบบ clasp กรุณารัน clasp login ด้วยบัญชีที่มีสิทธิ์แก้ไข Apps Script ปลายทาง'
    }

    Write-Host ('Script ID: ' + $config.scriptId) -ForegroundColor Cyan
    Write-Host ('Deployment ID ที่จะอัปเดต: ' + $DeploymentId) -ForegroundColor Cyan
    Write-Host ('เวอร์ชัน: ' + $VersionNumber) -ForegroundColor Cyan
    Write-Host 'ตรวจสอบว่าได้ push และทดสอบ URL /dev ของเวอร์ชันนี้แล้ว' -ForegroundColor Yellow
    $answer = Read-Host 'การดำเนินการนี้จะเปลี่ยน URL /exec ของ Deployment ที่ระบุ พิมพ์ DEPLOY เพื่อยืนยัน'
    if ($answer -cne 'DEPLOY') {
        Write-Host 'ยกเลิกแล้ว ไม่มีการเปลี่ยน Deployment' -ForegroundColor Yellow
        exit 0
    }

    & clasp redeploy $DeploymentId -V $VersionNumber
    if ($LASTEXITCODE -ne 0) {
        throw 'อัปเดต Deployment ไม่สำเร็จ ตรวจสอบ Deployment ID, Version Number และสิทธิ์ Apps Script'
    }

    Write-Host 'อัปเดต Deployment สำเร็จ' -ForegroundColor Green
} finally {
    Pop-Location
}
