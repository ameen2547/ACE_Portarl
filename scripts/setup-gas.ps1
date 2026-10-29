[CmdletBinding()]
param(
    [Parameter()]
    [string]$ScriptId
)

$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$configPath = Join-Path $projectRoot '.clasp.json'

if ([string]::IsNullOrWhiteSpace($ScriptId)) {
    $ScriptId = Read-Host 'กรอก Script ID จากหน้า Project Settings ของ Google Apps Script'
}

$ScriptId = $ScriptId.Trim()
if ([string]::IsNullOrWhiteSpace($ScriptId) -or $ScriptId -match '\s') {
    throw 'Script ID ว่างหรือมีช่องว่าง กรุณาตรวจสอบค่าแล้วลองใหม่'
}

if (Test-Path -LiteralPath $configPath) {
    $current = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
    if ($current.scriptId -ne $ScriptId) {
        Write-Host ('โปรเจกต์นี้ผูกกับ Script ID อื่นอยู่: ' + $current.scriptId) -ForegroundColor Yellow
        $answer = Read-Host 'พิมพ์ REPLACE เพื่อเปลี่ยนปลายทาง'
        if ($answer -cne 'REPLACE') {
            Write-Host 'ยกเลิกการตั้งค่า ไม่มีการเปลี่ยนไฟล์' -ForegroundColor Yellow
            exit 0
        }
    } else {
        Write-Host 'Script ID นี้ถูกตั้งค่าไว้แล้ว' -ForegroundColor Green
        exit 0
    }
}

$config = [ordered]@{ scriptId = $ScriptId } | ConvertTo-Json
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($configPath, $config + [Environment]::NewLine, $utf8NoBom)

Write-Host 'ตั้งค่าปลายทาง Apps Script แล้ว (.clasp.json ถูกยกเว้นจาก Git)' -ForegroundColor Green
Write-Host 'ขั้นตอนถัดไป: clasp login แล้วเรียก scripts/upload-dev.ps1'
