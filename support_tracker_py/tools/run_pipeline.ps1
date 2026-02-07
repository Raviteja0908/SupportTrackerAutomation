param(
    [string]$FolderPaths = "",
    [string]$StartDate = "",
    [string]$EndDate = "",
    [string]$OutputPstPath = ""
)

$ErrorActionPreference = "Stop"

Write-Host "=== Support Tracker Pipeline ==="
Write-Host "Step 1: Export from Outlook to PST (daily chunks)"

$exportScript = Join-Path $PSScriptRoot "export_outlook.ps1"
Write-Host ("Using export script: {0}" -f $exportScript)
try {
    $verLine = Select-String -Path $exportScript -Pattern "Export script version" | Select-Object -First 1
    if ($verLine) {
        Write-Host $verLine.Line
    } else {
        Write-Host "WARNING: export script version not found (script may be outdated)."
    }
} catch {
    Write-Host "WARNING: unable to read export script for version."
}
try {
    & $exportScript -FolderPathsInput $FolderPaths -StartDate $StartDate -EndDate $EndDate -OutputPstPath $OutputPstPath
    if (-not $?) {
        throw "Export script returned failure."
    }
} catch {
    throw "Export script failed: $($_.Exception.Message)"
}

if ([string]::IsNullOrWhiteSpace($OutputPstPath)) {
    $OutputPstPath = Read-Host "Enter the same PST path you used for export"
}

if (-not (Test-Path $OutputPstPath)) {
    throw "PST not found at: $OutputPstPath"
}

$pstDir = "D:\Support_Tracker\PstFiles"
if ((Split-Path -Parent $OutputPstPath) -ne $pstDir) {
    Write-Host ("WARNING: PST path is not in {0}. Docker will mount {0}." -f $pstDir)
    Write-Host "For consistency, store the PST in: D:\\Support_Tracker\\PstFiles"
}

$maxWaitSeconds = 600
$waitIntervalSeconds = 5
$startWait = Get-Date

function Test-FileUnlocked {
    param([string]$path)
    try {
        $fs = [System.IO.File]::Open($path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $fs.Close()
        return $true
    } catch {
        return $false
    }
}

Write-Host "Step 1b: Ensure Outlook export is finished before Docker"
if (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) {
    Write-Host "Outlook is still running. Please close Outlook to avoid PST file locks."
    Read-Host "Press Enter once Outlook is closed"
}

while (-not (Test-FileUnlocked $OutputPstPath)) {
    $elapsed = (Get-Date) - $startWait
    if ($elapsed.TotalSeconds -gt $maxWaitSeconds) {
        throw "PST file is still locked after $maxWaitSeconds seconds. Please close Outlook and retry."
    }
    Write-Host "Waiting for PST file to be released..."
    Start-Sleep -Seconds $waitIntervalSeconds
}

$outputDir = $null
while ($true) {
    $outputDir = Read-Host "Docker output folder (default: D:\\Support_Tracker\\DockerOutput)"
    if ([string]::IsNullOrWhiteSpace($outputDir)) {
        $outputDir = "D:\\Support_Tracker\\DockerOutput"
    }

    if ($outputDir.ToLower().EndsWith(".pst")) {
        Write-Host "That looks like a PST file path. Please enter a folder path for Docker output."
        continue
    }

    if (Test-Path $outputDir) {
        $item = Get-Item $outputDir
        if ($item -and -not $item.PSIsContainer) {
            Write-Host "That path is a file. Please enter a folder path for Docker output."
            continue
        }
    } else {
        try {
            New-Item -ItemType Directory -Path $outputDir | Out-Null
        } catch {
            Write-Host "Unable to create output folder. Please enter a valid folder path."
            continue
        }
    }
    break
}

Write-Host "Step 2: Run Docker pipeline"
Write-Host ("Using PST folder: {0}" -f $pstDir)
Write-Host ("Using output folder: {0}" -f $outputDir)

Write-Host "Running Docker..."
& docker run --rm -v "${pstDir}:/app/input" -v "${outputDir}:/app/output" support-tracker

Write-Host "=== Done ==="
