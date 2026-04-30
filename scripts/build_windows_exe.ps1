param(
    [string]$AppName = "Excel AI Assistant",
    [switch]$NoClean,
    [switch]$NoZip
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Resolve-Path (Join-Path $PSScriptRoot "..")
Set-Location $ProjectRoot

$python = if ($env:PYTHON) { $env:PYTHON } else { "python" }

& $python -m PyInstaller --version | Out-Null

$sourceFiles = @(
    "agent.py",
    "backup.py",
    "compress.py",
    "config.py",
    "constants.py",
    "excel_event_watcher.py",
    "excel_query.py",
    "excel_tools.py",
    "exceptions.py",
    "formula_validator.py",
    "logger.py",
    "macro.py",
    "main.py",
    "session.py",
    "telemetry.py",
    "utils.py"
)

$addData = @()
foreach ($file in $sourceFiles) {
    $addData += @("--add-data", "$file;.")
}
foreach ($dir in @("excel", "providers", "tools", "ui")) {
    $addData += @("--add-data", "$dir;$dir")
}
foreach ($file in @("requirements.txt", ".env.example", "README.md")) {
    if (Test-Path $file) {
        $addData += @("--add-data", "$file;.")
    }
}

$cleanArgs = if ($NoClean) { @() } else { @("--clean") }

$pyInstallerArgs = @(
    "-m", "PyInstaller"
) + $cleanArgs + @(
    "--noconfirm",
    "--onedir",
    "--console",
    "--name", $AppName,
    "--collect-all", "streamlit",
    "--collect-all", "altair",
    "--collect-all", "pydeck",
    "--collect-all", "pyarrow",
    "--collect-all", "pandas",
    "--collect-all", "openpyxl",
    "--collect-all", "dotenv",
    "--collect-all", "openai",
    "--collect-all", "httpx",
    "--collect-submodules", "win32com",
    "--hidden-import", "pythoncom",
    "--hidden-import", "pywintypes",
    "--hidden-import", "win32timezone",
    "--hidden-import", "win32com.client"
) + $addData + @(
    "packaging\streamlit_launcher.py"
)

& $python @pyInstallerArgs

$distDir = Join-Path $ProjectRoot "dist\$AppName"
if (-not (Test-Path $distDir)) {
    throw "Expected output folder was not created: $distDir"
}

$notes = @(
    "Excel AI Assistant - Windows distribution",
    "",
    "How to run:",
    "1. Extract the whole folder. Do not copy only the exe.",
    "2. Make sure Microsoft Excel is installed.",
    "3. Run $AppName.exe.",
    "4. The browser will open http://localhost:8501/.",
    "5. If your model server is different, update Server URL and Model ID in the sidebar.",
    "",
    "Notes:",
    "- This package includes the Python runtime. Colleagues do not need to install Python.",
    "- Deliver the whole folder. Do not delete the _internal folder.",
    "- This package does not include .env or test workbooks.",
    "- Excel automation requires Microsoft Excel and the current Windows user permissions."
) -join [Environment]::NewLine
Set-Content -Path (Join-Path $distDir "README_FOR_COLLEAGUES.txt") -Value $notes -Encoding UTF8

if (-not $NoZip) {
    $zipPath = Join-Path $ProjectRoot ("dist\{0}-windows.zip" -f ($AppName -replace "\s+", "-"))
    if (Test-Path $zipPath) {
        Remove-Item -LiteralPath $zipPath -Force
    }
    $compressed = $false
    for ($attempt = 1; $attempt -le 5; $attempt++) {
        try {
            Compress-Archive -Path (Join-Path $distDir "*") -DestinationPath $zipPath -Force
            $compressed = $true
            break
        } catch {
            if (Test-Path $zipPath) {
                Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
            }
            Start-Sleep -Seconds (2 * $attempt)
        }
    }
    if (-not $compressed) {
        $tar = Get-Command tar.exe -ErrorAction SilentlyContinue
        if (-not $tar) {
            throw "Compress-Archive failed and tar.exe was not found."
        }
        & $tar.Source -a -cf $zipPath -C (Join-Path $ProjectRoot "dist") $AppName
        if ($LASTEXITCODE -ne 0) {
            throw "tar.exe zip fallback failed with exit code $LASTEXITCODE"
        }
        $compressed = $true
    }
    Write-Host "ZIP: $zipPath"
}

$exePath = Join-Path $distDir "$AppName.exe"
Write-Host "EXE: $exePath"
