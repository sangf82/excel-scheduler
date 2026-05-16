param(
    [switch]$WhatIf,
    [switch]$RestoreLegacy
)

$ErrorActionPreference = 'Stop'
chcp 65001 > $null
$OutputEncoding = [System.Text.Encoding]::UTF8

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$codexHome   = Join-Path $env:USERPROFILE '.codex'
$agentsHome  = Join-Path $env:USERPROFILE '.agents'

$beginMarker = '# BEGIN MEDMATE'
$endMarker   = '# END MEDMATE'

function Write-Plan($message) {
    if ($WhatIf) {
        Write-Host "[WhatIf] $message" -ForegroundColor Yellow
    } else {
        Write-Host $message -ForegroundColor Cyan
    }
}

function Write-TextNoBom($path, $text) {
    if (-not $WhatIf) {
        [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding $false))
    }
}

if (-not (Test-Path $codexHome)) {
    Write-Error "Không tìm thấy Codex home tại $codexHome."
    exit 1
}

Write-Host "Trình gỡ cài đặt MedMate Scheduler" -ForegroundColor Green

# 1. Remove plugin link/copy
$pluginLink = Join-Path $codexHome 'plugins\medmate-scheduler'
if (Test-Path $pluginLink) {
    Write-Plan "Xóa liên kết plugin: $pluginLink"
    if (-not $WhatIf) {
        $item = Get-Item $pluginLink -Force
        if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
            cmd /c rmdir "`"$pluginLink`"" | Out-Null
        } else {
            Remove-Item -Path $pluginLink -Recurse -Force
        }
    }
} else {
    Write-Plan "Liên kết plugin không tồn tại tại $pluginLink."
}

# 2. Strip MEDMATE block + restore plugin states from config.toml
$configPath = Join-Path $codexHome 'config.toml'
if (Test-Path $configPath) {
    $existingConfig = Get-Content $configPath -Raw
    if ($null -eq $existingConfig) { $existingConfig = '' }

    # Remove MEDMATE block
    $pattern = '\r?\n?' + [regex]::Escape($beginMarker) + '[\s\S]*?' + [regex]::Escape($endMarker) + '\r?\n?'
    $rx = New-Object System.Text.RegularExpressions.Regex($pattern)
    $newConfig = $existingConfig
    if ($rx.IsMatch($existingConfig)) {
        Write-Plan "Xóa khối MEDMATE khỏi config.toml"
        $newConfig = $rx.Replace($existingConfig, "`n")
    } else {
        Write-Plan "Không có khối MEDMATE trong config.toml."
    }

    # Restore plugin states from snapshot
    $snapshotPath = Join-Path $codexHome '.tmp\medmate-config-snapshot.json'
    if (Test-Path $snapshotPath) {
        try {
            $snapshot = Get-Content $snapshotPath -Raw | ConvertFrom-Json
            foreach ($name in $snapshot.PSObject.Properties.Name) {
                $oldValue = $snapshot.$name
                $pluginPattern = '(\[plugins\.' + '"' + [regex]::Escape($name) + '"' + '\]\s*\r?\n\s*enabled\s*=\s*)(true|false)'
                if ([regex]::IsMatch($newConfig, $pluginPattern)) {
                    $newConfig = [regex]::Replace($newConfig, $pluginPattern, "`${1}$oldValue")
                    Write-Plan "Khôi phục $name -> enabled = $oldValue"
                }
            }
            if (-not $WhatIf) {
                Remove-Item $snapshotPath -Force
            }
        } catch {
            Write-Warning "Không đọc được snapshot: $_"
        }
    }

    if (-not $WhatIf) {
        Write-TextNoBom $configPath $newConfig
    }
} else {
    Write-Plan "config.toml không tồn tại."
}

# 2b. Remove AGENTS.md block from ~/.codex/AGENTS.md
$agentsTarget = Join-Path $codexHome 'AGENTS.md'
$agentsBegin = '# BEGIN MEDMATE AGENTS'
$agentsEnd   = '# END MEDMATE AGENTS'
if (Test-Path $agentsTarget) {
    $agentsExisting = Get-Content $agentsTarget -Raw -Encoding UTF8
    if ($null -eq $agentsExisting) { $agentsExisting = '' }
    $agentsPattern = '\r?\n?' + [regex]::Escape($agentsBegin) + '[\s\S]*?' + [regex]::Escape($agentsEnd) + '\r?\n?'
    $agentsRx = New-Object System.Text.RegularExpressions.Regex($agentsPattern)
    if ($agentsRx.IsMatch($agentsExisting)) {
        Write-Plan "Xóa MedMate rules khỏi ~/.codex/AGENTS.md"
        if (-not $WhatIf) {
            $newAgents = $agentsRx.Replace($agentsExisting, "`n").Trim()
            Write-TextNoBom $agentsTarget $newAgents
        }
    } else {
        Write-Plan "Không có MedMate rules trong ~/.codex/AGENTS.md."
    }
} else {
    Write-Plan "~/.codex/AGENTS.md không tồn tại."
}

# 3. Remove marketplace entry
$marketplacePath = Join-Path $agentsHome 'plugins\marketplace.json'
if (Test-Path $marketplacePath) {
    try {
        $marketplace = Get-Content $marketplacePath -Raw | ConvertFrom-Json
    } catch {
        Write-Warning "marketplace.json không phải JSON hợp lệ; bỏ qua."
        $marketplace = $null
    }
    if ($marketplace -and $marketplace.plugins) {
        $remaining = @()
        foreach ($p in $marketplace.plugins) {
            if ($p.name -ne 'medmate-scheduler') { $remaining += $p }
        }
        if ($remaining.Count -ne $marketplace.plugins.Count) {
            Write-Plan "Xóa medmate-scheduler khỏi marketplace.json"
            if (-not $WhatIf) {
                $marketplace.plugins = $remaining
                ($marketplace | ConvertTo-Json -Depth 10) | Set-Content -Path $marketplacePath -Encoding UTF8
            }
        } else {
            Write-Plan "marketplace.json không có mục medmate-scheduler."
        }
    }
} else {
    Write-Plan "marketplace.json không tồn tại."
}

# 4. Remove memory file
$memoryTarget = Join-Path $codexHome 'memories\medmate-scheduler.md'
if (Test-Path $memoryTarget) {
    Write-Plan "Xóa file bộ nhớ: $memoryTarget"
    if (-not $WhatIf) {
        Remove-Item -Path $memoryTarget -Force
    }
} else {
    Write-Plan "File bộ nhớ không tồn tại."
}

# 5. Optionally restore legacy excel skill
if ($RestoreLegacy) {
    $tmpDir = Join-Path $codexHome '.tmp'
    if (Test-Path $tmpDir) {
        $latest = Get-ChildItem -Path $tmpDir -Directory -Filter 'excel.bak.*' |
                  Sort-Object LastWriteTime -Descending |
                  Select-Object -First 1
        if ($latest) {
            $restoreTarget = Join-Path $codexHome 'skills\excel'
            if (Test-Path $restoreTarget) {
                Write-Warning "skills\excel đã tồn tại; không ghi đè. Bản sao lưu mới nhất: $($latest.FullName)"
            } else {
                Write-Plan "Khôi phục skill excel cũ từ $($latest.FullName)"
                if (-not $WhatIf) {
                    Move-Item -Path $latest.FullName -Destination $restoreTarget -Force
                }
            }
        } else {
            Write-Plan "Không tìm thấy bản sao lưu excel.bak.* trong $tmpDir."
        }
    } else {
        Write-Plan "$tmpDir không tồn tại; không có gì để khôi phục."
    }
}

Write-Host "Đã gỡ cài MedMate Scheduler hoàn tất." -ForegroundColor Green
