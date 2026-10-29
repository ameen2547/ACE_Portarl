[CmdletBinding()]
param(
    [Parameter()]
    [switch]$Yes
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
    throw '.clasp.json ไม่มีค่า scriptId กรุณารัน scripts/setup-gas.ps1 ใหม่'
}

Push-Location $projectRoot
try {
    & clasp show-authorized-user *> $null
    if ($LASTEXITCODE -ne 0) {
        throw 'ยังไม่ได้เข้าสู่ระบบ clasp กรุณารัน clasp login ด้วยบัญชีที่มีสิทธิ์แก้ไข Apps Script ปลายทาง'
    }

    Write-Host ('ปลายทาง Script ID: ' + $config.scriptId) -ForegroundColor Cyan
    Write-Host 'รายการไฟล์ที่จะส่งขึ้น Apps Script:' -ForegroundColor Cyan
    & clasp status
    if ($LASTEXITCODE -ne 0) {
        throw 'clasp ตรวจสอบรายการไฟล์ไม่สำเร็จ จึงหยุดก่อนอัปโหลด'
    }

    if (-not $Yes) {
        $answer = Read-Host 'การอัปโหลดจะเขียนทับไฟล์ต้นฉบับใน Apps Script และอัปเดต URL /dev พิมพ์ PUSH เพื่อยืนยัน'
        if ($answer -cne 'PUSH') {
            Write-Host 'ยกเลิกแล้ว ไม่มีการอัปโหลด' -ForegroundColor Yellow
            exit 0
        }
    }

    & clasp push
    if ($LASTEXITCODE -ne 0) {
        throw 'อัปโหลดไม่สำเร็จ ตรวจสอบข้อความจาก clasp และสิทธิ์ของบัญชี แล้วลองใหม่'
    }

    Write-Host 'อัปโหลดสำเร็จ URL /dev ใช้โค้ดล่าสุดแล้ว ส่วน URL /exec จะไม่เปลี่ยนจนกว่าจะปล่อยเวอร์ชันใหม่' -ForegroundColor Green
} finally {
    Pop-Location
}
