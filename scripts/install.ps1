param(
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'
chcp 65001 > $null
$OutputEncoding = [System.Text.Encoding]::UTF8

# Handle remote execution (irm ... | iex) — $PSScriptRoot is empty
if (-not $PSScriptRoot) {
    Write-Host "Detected remote execution. Cloning repository first..." -ForegroundColor Cyan
    $repoUrl = 'https://github.com/sangf82/excel-scheduler.git'
    $cloneDir = Join-Path $env:TEMP "medmate-scheduler-$(Get-Random)"

    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        Write-Error "Git is required for remote install. Please install git or clone the repo manually:`n  git clone $repoUrl`nThen run .\scripts\install.ps1 from the cloned folder."
        exit 1
    }

    git clone $repoUrl $cloneDir 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Error "Failed to clone repository. If this is a private repo, ensure your git credentials are configured.`nAlternatively, clone manually:`n  git clone $repoUrl"
        exit 1
    }

    Write-Host "Repository cloned to $cloneDir" -ForegroundColor Green
    $invokeArgs = @{}
    if ($WhatIf) { $invokeArgs['WhatIf'] = $true }
    & (Join-Path $cloneDir 'scripts\install.ps1') @invokeArgs
    exit $LASTEXITCODE
}

$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path

# Validate we are in the correct repository
$repoMarker = Join-Path $projectRoot 'AGENTS.md'
if (-not (Test-Path $repoMarker)) {
    Write-Error "This script must be run from the MedMate Scheduler repository.`nPlease clone the repo and run .\scripts\install.ps1 from the project root.`n  git clone https://github.com/sangf82/excel-scheduler.git"
    exit 1
}

$codexHome   = Join-Path $env:USERPROFILE '.codex'
$agentsHome  = Join-Path $env:USERPROFILE '.agents'

$beginMarker = '# BEGIN MEDMATE'
$endMarker   = '# END MEDMATE'

$medmateBlock = @"
$beginMarker
[projects."c:\\projects\\medmate\\excel-scheduler"]
trust_level = "trusted"

[plugins."browser@openai-bundled"]
enabled = false

[plugins."presentations@openai-primary-runtime"]
enabled = false

[plugins."documents@openai-primary-runtime"]
enabled = false

[plugins."spreadsheets@openai-primary-runtime"]
enabled = true
$endMarker
"@

function Write-Plan($message) {
    if ($WhatIf) {
        Write-Host "[WhatIf] $message" -ForegroundColor Yellow
    } else {
        Write-Host $message -ForegroundColor Cyan
    }
}

function Ensure-Directory($path) {
    if (-not (Test-Path $path)) {
        Write-Plan "Tạo thư mục: $path"
        if (-not $WhatIf) {
            New-Item -ItemType Directory -Path $path -Force | Out-Null
        }
    }
}

function Write-TextNoBom($path, $text) {
    if (-not $WhatIf) {
        [System.IO.File]::WriteAllText($path, $text, (New-Object System.Text.UTF8Encoding $false))
    }
}

function Test-Dependencies {
    Write-Host "--- Kiểm tra phụ thuộc / Dependency check ---" -ForegroundColor Cyan

    # 1. Python (3.8+)
    $script:pythonExe = $null
    foreach ($candidate in @('python', 'py', 'python3')) {
        $cmd = Get-Command $candidate -ErrorAction SilentlyContinue
        if ($cmd) {
            try {
                $verStr = & $cmd.Source --version 2>&1
                if ($verStr -match '(\d+)\.(\d+)') {
                    $major = [int]$Matches[1]
                    $minor = [int]$Matches[2]
                    if ($major -gt 3 -or ($major -eq 3 -and $minor -ge 8)) {
                        $script:pythonExe = $cmd.Source
                        Write-Host "  Python OK: $verStr ($($cmd.Source))" -ForegroundColor Green
                        break
                    }
                }
            } catch {
                continue
            }
        }
    }

    if (-not $script:pythonExe) {
        Write-Host "" -ForegroundColor Red
        Write-Host "=== LỖI: Không tìm thấy Python 3.8+ ===" -ForegroundColor Red
        Write-Host "Python chua duoc cai dat hoac phien ban qua cu." -ForegroundColor Red
        Write-Host "" -ForegroundColor Red
        Write-Host "Huong dan cai dat:" -ForegroundColor Yellow
        Write-Host "  1. Tai Python tu: https://www.python.org/downloads/" -ForegroundColor Yellow
        Write-Host "  2. Khi cai dat, tich chon 'Add Python to PATH'" -ForegroundColor Yellow
        Write-Host "  3. Mo lai PowerShell va chay lai install.ps1" -ForegroundColor Yellow
        Write-Host "" -ForegroundColor Red
        Write-Host "=== ERROR: Python 3.8+ not found ===" -ForegroundColor Red
        Write-Host "Please install Python from: https://www.python.org/downloads/" -ForegroundColor Yellow
        Write-Host "Make sure to check 'Add Python to PATH' during installation." -ForegroundColor Yellow
        exit 1
    }

    # 2. openpyxl (auto-install)
    try {
        & $script:pythonExe -c "import openpyxl" 2>$null
        if ($LASTEXITCODE -eq 0) {
            $opVer = & $script:pythonExe -c "import openpyxl; print(openpyxl.__version__)" 2>$null
            Write-Host "  openpyxl OK: $opVer" -ForegroundColor Green
        } else {
            throw "missing"
        }
    } catch {
        Write-Host "  openpyxl chua co. Dang cai dat..." -ForegroundColor Yellow
        & $script:pythonExe -m pip install --user openpyxl
        if ($LASTEXITCODE -ne 0) {
            Write-Host "" -ForegroundColor Red
            Write-Host "=== LỖI: Cai dat openpyxl that bai ===" -ForegroundColor Red
            Write-Host "Chay tay: pip install openpyxl" -ForegroundColor Yellow
            Write-Host "=== ERROR: openpyxl installation failed ===" -ForegroundColor Red
            Write-Host "Try manually: pip install openpyxl" -ForegroundColor Yellow
            exit 1
        }
        Write-Host "  openpyxl da cai dat xong" -ForegroundColor Green
    }

    # 3. Node.js (16+)
    $nodeExe = $null
    $nodeCmd = Get-Command 'node' -ErrorAction SilentlyContinue
    if ($nodeCmd) {
        try {
            $verStr = & $nodeCmd.Source --version 2>&1
            if ($verStr -match 'v?(\d+)') {
                $major = [int]$Matches[1]
                if ($major -ge 16) {
                    $nodeExe = $nodeCmd.Source
                    Write-Host "  Node.js OK: $verStr ($($nodeCmd.Source))" -ForegroundColor Green
                }
            }
        } catch {}
    }

    if (-not $nodeExe) {
        Write-Host "" -ForegroundColor Red
        Write-Host "=== LỖI: Khong tim thay Node.js 16+ ===" -ForegroundColor Red
        Write-Host "Node.js can thiet de chay excel-mcp-server qua npx." -ForegroundColor Red
        Write-Host "" -ForegroundColor Yellow
        Write-Host "Huong dan cai dat:" -ForegroundColor Yellow
        Write-Host "  1. Tai Node.js tu: https://nodejs.org/ (chon LTS)" -ForegroundColor Yellow
        Write-Host "  2. Cai dat voi mac dinh (se tu them PATH)" -ForegroundColor Yellow
        Write-Host "  3. Mo lai PowerShell va chay lai install.ps1" -ForegroundColor Yellow
        Write-Host "" -ForegroundColor Red
        Write-Host "=== ERROR: Node.js 16+ not found ===" -ForegroundColor Red
        Write-Host "Node.js is required for the excel MCP server (npx)." -ForegroundColor Red
        Write-Host "Download from: https://nodejs.org/ (choose LTS)" -ForegroundColor Yellow
        exit 1
    }

    # 4. npx
    $npxCmd = Get-Command 'npx' -ErrorAction SilentlyContinue
    if (-not $npxCmd) {
        # Try npx.cmd on Windows
        $npxCmd = Get-Command 'npx.cmd' -ErrorAction SilentlyContinue
    }
    if ($npxCmd) {
        try {
            $npxVer = & $npxCmd.Source --version 2>&1
            Write-Host "  npx OK: $npxVer" -ForegroundColor Green
        } catch {
            Write-Host "  npx: co nhung khong chay duoc. Kiem tra lai Node.js." -ForegroundColor Yellow
        }
    } else {
        Write-Host "" -ForegroundColor Red
        Write-Host "=== CANH BAO: Khong tim thay npx ===" -ForegroundColor Yellow
        Write-Host "npx thuong di kem voi Node.js. Hay kiem tra lai cai dat Node.js." -ForegroundColor Yellow
        Write-Host "=== WARNING: npx not found ===" -ForegroundColor Yellow
        Write-Host "npx usually comes with Node.js. Please verify your Node.js installation." -ForegroundColor Yellow
        # Not fatal — user may still have it via npm prefix or global path issues
    }

    Write-Host "--- Phu thuoc da san sang / Dependencies OK ---" -ForegroundColor Green
    Write-Host ""
}

if (-not (Test-Path $codexHome)) {
    Write-Error "Không tìm thấy Codex home tại $codexHome. Vui lòng cài Codex Desktop trước."
    exit 1
}

Write-Host "Trình cài đặt MedMate Scheduler" -ForegroundColor Green
Write-Host "Thư mục dự án: $projectRoot"
Write-Host "Codex home:    $codexHome"
if ($WhatIf) {
    Write-Host "Đang chạy ở chế độ -WhatIf. Sẽ không thay đổi gì trên ổ đĩa." -ForegroundColor Yellow
}

# --- Dependency check ---
if (-not $WhatIf) {
    Test-Dependencies
} else {
    Write-Host "[WhatIf] Bo qua kiem tra phu thuoc." -ForegroundColor Yellow
}

# 1. Plugin junction
$pluginLink = Join-Path $codexHome 'plugins\medmate-scheduler'
Ensure-Directory (Split-Path $pluginLink -Parent)

if (Test-Path $pluginLink) {
    Write-Plan "Liên kết plugin đã tồn tại tại $pluginLink; giữ nguyên."
} else {
    Write-Plan "Tạo junction: $pluginLink -> $projectRoot"
    if (-not $WhatIf) {
        try {
            New-Item -ItemType Junction -Path $pluginLink -Target $projectRoot | Out-Null
        } catch {
            Write-Warning "Junction thất bại ($($_.Exception.Message)). Chuyển sang dùng robocopy /MIR."
            $robocopyExclude = @('legacy', '.git', 'spreadsheets')
            $excludeArgs = @()
            foreach ($d in $robocopyExclude) {
                $excludeArgs += '/XD'
                $excludeArgs += (Join-Path $projectRoot $d)
            }
            & robocopy $projectRoot $pluginLink /MIR @excludeArgs | Out-Null
            if ($LASTEXITCODE -ge 8) {
                throw "robocopy thất bại với mã lỗi $LASTEXITCODE"
            }
        }
    }
}

# 2. Patch ~/.codex/config.toml
$configPath = Join-Path $codexHome 'config.toml'
if (-not (Test-Path $configPath)) {
    Write-Plan "Tạo mới config.toml tại $configPath"
    if (-not $WhatIf) {
        New-Item -ItemType File -Path $configPath -Force | Out-Null
    }
}

$existingConfig = ''
if (Test-Path $configPath) {
    $existingConfig = Get-Content $configPath -Raw
    if ($null -eq $existingConfig) { $existingConfig = '' }
}

# Snapshot current plugin states (only once — first install)
$snapshotPath = Join-Path $codexHome '.tmp\medmate-config-snapshot.json'
$pluginNames = @(
    'browser@openai-bundled',
    'presentations@openai-primary-runtime',
    'documents@openai-primary-runtime',
    'spreadsheets@openai-primary-runtime'
)

if (-not (Test-Path $snapshotPath)) {
    $snapshot = @{}
    foreach ($name in $pluginNames) {
        $pattern = '\[plugins\.' + '"' + [regex]::Escape($name) + '"' + '\]\s*\r?\n\s*enabled\s*=\s*(true|false)'
        $m = [regex]::Match($existingConfig, $pattern)
        if ($m.Success) {
            $snapshot[$name] = $m.Groups[1].Value
        }
    }
    if ($snapshot.Count -gt 0) {
        Ensure-Directory (Split-Path $snapshotPath -Parent)
        if (-not $WhatIf) {
            ($snapshot | ConvertTo-Json) | Set-Content -Path $snapshotPath -Encoding UTF8
        }
        Write-Plan "Đã lưu snapshot trạng thái plugin trước khi thay đổi."
    }
}

# Remove any old MEDMATE block first
$medmatePattern = [regex]::Escape($beginMarker) + '[\s\S]*?' + [regex]::Escape($endMarker)
$newConfig = [regex]::Replace($existingConfig, $medmatePattern, '').Trim()

# Update existing plugin entries in-place; add missing ones at the end
$pluginReplacements = @{
    'browser@openai-bundled'               = 'false'
    'presentations@openai-primary-runtime' = 'false'
    'documents@openai-primary-runtime'     = 'false'
    'spreadsheets@openai-primary-runtime'  = 'true'
}

$missingPlugins = @()
foreach ($name in $pluginReplacements.Keys) {
    $pattern = '(\[plugins\.' + '"' + [regex]::Escape($name) + '"' + '\]\s*\r?\n\s*enabled\s*=\s*)(true|false)'
    if ([regex]::IsMatch($newConfig, $pattern)) {
        $newConfig = [regex]::Replace($newConfig, $pattern, "`${1}$($pluginReplacements[$name])")
        Write-Plan "Cập nhật $name -> enabled = $($pluginReplacements[$name])"
    } else {
        $missingPlugins += "[plugins.`"$name`"]`nenabled = $($pluginReplacements[$name])`n"
    }
}

# Append missing plugin tables + project trust block at the end
$appendBlock = ""
if ($missingPlugins.Count -gt 0) {
    $appendBlock += ($missingPlugins -join "`n") + "`n"
}

$projectSection = "[projects.`"c:\\projects\\medmate\\excel-scheduler`"]"
$projectTrustBlock = ""
if (-not $newConfig.ToLower().Contains($projectSection.ToLower())) {
    $projectTrustBlock = "$projectSection`ntrust_level = `"trusted`"`n"
}

if (-not ($newConfig -match '(?m)^mode\s*=')) {
    $newConfig = "mode = `"planning`"`n" + $newConfig
} else {
    $newConfig = $newConfig -replace '(?m)^mode\s*=.*', 'mode = "planning"'
}

if ($appendBlock.Length -gt 0 -or $projectTrustBlock.Length -gt 0) {
    $separator = if ($newConfig.Length -eq 0 -or $newConfig.EndsWith("`n")) { '' } else { "`n" }
    $newConfig = $newConfig + $separator + "`n$beginMarker`n" + $appendBlock + $projectTrustBlock + "$endMarker`n"
}

Write-Plan "Ghi config.toml (UTF-8, no BOM)"
Write-TextNoBom $configPath $newConfig

# 2b. Install AGENTS.md into ~/.codex/AGENTS.md (global rule enforcement)
$agentsSource = Join-Path $projectRoot 'AGENTS.md'
$agentsTarget = Join-Path $codexHome 'AGENTS.md'
$agentsBegin = '# BEGIN MEDMATE AGENTS'
$agentsEnd   = '# END MEDMATE AGENTS'

$agentsExisting = ''
if (Test-Path $agentsTarget) {
    $agentsExisting = Get-Content $agentsTarget -Raw -Encoding UTF8
    if ($null -eq $agentsExisting) { $agentsExisting = '' }
}

$agentsPattern = [regex]::Escape($agentsBegin) + '[\s\S]*?' + [regex]::Escape($agentsEnd)
$agentsExisting = [regex]::Replace($agentsExisting, $agentsPattern, '').Trim()

if (Test-Path $agentsSource) {
    $sourceAgents = Get-Content $agentsSource -Raw -Encoding UTF8
    $agentsBlock = "`n$agentsBegin`n$sourceAgents`n$agentsEnd`n"
    $agentsSep = if ($agentsExisting.Length -eq 0 -or $agentsExisting.EndsWith("`n")) { '' } else { "`n" }
    $newAgents = $agentsExisting + $agentsSep + $agentsBlock
    Write-Plan "Cập nhật ~/.codex/AGENTS.md với MedMate Scheduler rules"
    Write-TextNoBom $agentsTarget $newAgents
} else {
    Write-Plan "Không tìm thấy AGENTS.md trong thư mục dự án; bỏ qua."
}

# 3. Patch ~/.agents/plugins/marketplace.json
$marketplacePath = Join-Path $agentsHome 'plugins\marketplace.json'
Ensure-Directory (Split-Path $marketplacePath -Parent)

$marketplaceEntry = [ordered]@{
    name   = 'medmate-scheduler'
    source = [ordered]@{
        source = 'local'
        path   = $projectRoot
    }
    category = 'Productivity'
    policy = [ordered]@{
        installation   = 'AVAILABLE'
        authentication = 'ON_INSTALL'
    }
}

if (Test-Path $marketplacePath) {
    try {
        $marketplace = Get-Content $marketplacePath -Raw | ConvertFrom-Json
    } catch {
        Write-Warning "marketplace.json hiện tại không phải JSON hợp lệ. Sao lưu và tạo lại."
        if (-not $WhatIf) {
            Copy-Item $marketplacePath "$marketplacePath.bak" -Force
        }
        $marketplace = $null
    }
} else {
    $marketplace = $null
}

if (-not $marketplace) {
    $marketplace = [pscustomobject]@{
        name      = 'medmate'
        interface = [pscustomobject]@{ displayName = 'MedMate Marketplace' }
        plugins   = @()
    }
}

if (-not ($marketplace.PSObject.Properties.Name -contains 'plugins')) {
    $marketplace | Add-Member -NotePropertyName 'plugins' -NotePropertyValue @() -Force
}

$plugins = @()
if ($marketplace.plugins) {
    foreach ($p in $marketplace.plugins) {
        if ($p.name -ne 'medmate-scheduler') { $plugins += $p }
    }
}
$plugins += [pscustomobject]$marketplaceEntry
$marketplace.plugins = $plugins

Write-Plan "Ghi marketplace.json với $($plugins.Count) plugin"
if (-not $WhatIf) {
    ($marketplace | ConvertTo-Json -Depth 10) | Set-Content -Path $marketplacePath -Encoding UTF8
}

# 3b. Setup ~/.codex/mcp.json (inject excel MCP server)
$mcpJsonPath = Join-Path $codexHome 'mcp.json'
if (-not (Test-Path $mcpJsonPath)) {
    Write-Plan "Tạo mới mcp.json tại $mcpJsonPath"
    $mcpData = [ordered]@{ mcpServers = [ordered]@{} }
} else {
    try {
        $mcpData = Get-Content $mcpJsonPath -Raw | ConvertFrom-Json
    } catch {
        Write-Warning "mcp.json hiện tại không phải JSON hợp lệ. Tạo lại."
        if (-not $WhatIf) {
            Copy-Item $mcpJsonPath "$mcpJsonPath.bak" -Force
        }
        $mcpData = [ordered]@{ mcpServers = [ordered]@{} }
    }
}

if (-not $mcpData.mcpServers) {
    $mcpData | Add-Member -NotePropertyName 'mcpServers' -NotePropertyValue [ordered]@{} -Force
}

$mcpData.mcpServers | Add-Member -NotePropertyName 'excel' -NotePropertyValue ([ordered]@{
    command = "npx"
    args = @("-y", "@negokaz/excel-mcp-server")
}) -Force

Write-Plan "Cập nhật mcp.json với excel MCP server"
if (-not $WhatIf) {
    ($mcpData | ConvertTo-Json -Depth 10) | Set-Content -Path $mcpJsonPath -Encoding UTF8
}

# 4. Seed memory file
$memorySource = Join-Path $projectRoot 'memories\medmate-scheduler.seed.md'
$memoryTarget = Join-Path $codexHome  'memories\medmate-scheduler.md'
Ensure-Directory (Split-Path $memoryTarget -Parent)
if (Test-Path $memoryTarget) {
    Write-Plan "File bộ nhớ đã có tại $memoryTarget; giữ nguyên."
} else {
    Write-Plan "Sao chép bộ nhớ mẫu sang $memoryTarget"
    if (-not $WhatIf) {
        Copy-Item $memorySource $memoryTarget -Force
    }
}

# 5. Backup legacy ~/.codex/skills/excel
$legacySkill = Join-Path $codexHome 'skills\excel'
if (Test-Path $legacySkill) {
    $timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
    $backupRoot = Join-Path $codexHome ".tmp\excel.bak.$timestamp"
    Write-Plan "Sao lưu skill cũ: $legacySkill -> $backupRoot"
    if (-not $WhatIf) {
        Ensure-Directory (Split-Path $backupRoot -Parent)
        Move-Item -Path $legacySkill -Destination $backupRoot -Force
    }
} else {
    Write-Plan "Không có ~/.codex/skills/excel cũ để sao lưu."
}

# 6. Move project-local excel/ -> legacy/excel/
$projectExcel = Join-Path $projectRoot 'excel'
$legacyDir    = Join-Path $projectRoot 'legacy'
if (Test-Path $projectExcel) {
    Ensure-Directory $legacyDir
    $target = Join-Path $legacyDir 'excel'
    if (Test-Path $target) {
        Write-Plan "legacy\excel đã tồn tại; để nguyên project excel/."
    } else {
        Write-Plan "Di chuyển $projectExcel -> $target"
        if (-not $WhatIf) {
            Move-Item -Path $projectExcel -Destination $target -Force
        }
    }
} else {
    Write-Plan "Không có thư mục excel/ cục bộ để di chuyển."
}

# 7. Build the workbook
$buildScript = Join-Path $projectRoot 'scripts\build_template.py'
Write-Plan "Chạy build_template.py để tạo lại scheduling-template.xlsx"
if (-not $WhatIf) {
    # pythonExe already discovered & validated in Test-Dependencies
    $env:PYTHONIOENCODING = 'utf-8'
    & $script:pythonExe $buildScript
    if ($LASTEXITCODE -ne 0) {
        Write-Error "build_template.py thất bại."
        exit 1
    }
}
Write-Plan "Kiểm tra tính hợp lệ của config.toml (Healthcheck)"
if (-not $WhatIf) {
    try {
        python -c "import sys; import tomli; tomli.load(open(r'$configPath', 'rb'))"
        if ($LASTEXITCODE -ne 0) {
            Write-Error "Codex config.toml bị lỗi cấu trúc! Vui lòng khôi phục từ thư mục backup."
            exit 1
        }
        Write-Plan "Healthcheck OK: config.toml hợp lệ."
    } catch {
        Write-Plan "Bỏ qua healthcheck do python không khả dụng hoặc thiếu module tomli."
    }
}

Write-Host ""
Write-Host "Đã cài đặt MedMate Scheduler. Hãy khởi động lại Codex Desktop, mở scheduling-template.xlsx (sheet INPUT), rồi gõ: Thêm bệnh nhân BN036 tên Nguyễn Văn X loại Ghép gan" -ForegroundColor Green
