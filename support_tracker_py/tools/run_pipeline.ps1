param(
    [string]$FolderPaths = "",
    [string]$StartDate = "",
    [string]$EndDate = "",
    [string]$OutputPstPath = "",
    [switch]$NoVolume,
    [switch]$KeepVolumeData,
    [switch]$CleanVolume
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
$useVolume = -not $NoVolume
if ($useVolume) {
    $volumeName = "support_tracker_output"
    $template = Get-ChildItem $outputDir -Filter "*.xlsx" |
        Where-Object { $_.Name -notmatch "filled|done|automation_output" -and $_.Length -gt 0 } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
    if (-not $template) {
        $zeroes = Get-ChildItem $outputDir -Filter "*.xlsx" |
            Where-Object { $_.Length -eq 0 } |
            Select-Object -ExpandProperty Name
        if ($zeroes) {
            Write-Host ("Found zero-byte .xlsx files: {0}" -f ($zeroes -join ", "))
        }
        throw "No valid (non-empty) template .xlsx found in output folder. Please place a real template in: $outputDir"
    }
    try {
        $header = Get-Content -Path $template.FullName -Encoding Byte -TotalCount 2
        if ($header.Length -lt 2 -or $header[0] -ne 0x50 -or $header[1] -ne 0x4B) {
            throw "Template does not look like a valid .xlsx (missing PK header). Close Excel or choose the correct file."
        }
    } catch {
        throw "Unable to read template file. Close Excel if it is open, then retry. Details: $($_.Exception.Message)"
    }
    Write-Host ("Using Docker volume for output: {0}" -f $volumeName)
    & docker volume create $volumeName | Out-Null
    Write-Host ("Copying template into volume: {0}" -f $template.Name)
    & docker run --rm --entrypoint sh -v "${volumeName}:/app/output" -v "$($template.FullName):/host/template.xlsx:ro" support-tracker -c "cp /host/template.xlsx /app/output/"
    try {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        & docker run --rm -v "${pstDir}:/app/input" -v "${volumeName}:/app/output" support-tracker
        $sw.Stop()
        Write-Host ("Sheet fill time: {0}" -f $sw.Elapsed)
        Write-Host "Copying output files to host folder..."
        & docker run --rm --entrypoint sh -v "${volumeName}:/app/output" -v "${outputDir}:/host" support-tracker -c "cp /app/output/*_filled*.xlsx /host/ 2>/dev/null || true; cp /app/output/automation_output.csv /host/ 2>/dev/null || true; cp /app/output/debug_subjects.csv /host/ 2>/dev/null || true; cp /app/output/processing.log /host/ 2>/dev/null || true"
    } finally {
        if (-not $KeepVolumeData) {
            Write-Host "Clearing volume data..."
            & docker run --rm --entrypoint sh -v "${volumeName}:/app/output" support-tracker -c "find /app/output -mindepth 1 -maxdepth 1 -exec rm -rf {} +"
            Write-Host "Cleanup complete (volume cleared)."
        } else {
            Write-Host "Keeping volume data (EMLs retained)."
        }
        if ($CleanVolume) {
            if ($KeepVolumeData) {
                Write-Host "NOTE: -CleanVolume will remove the volume (data will be lost)."
            }
            Write-Host ("Removing Docker volume: {0}" -f $volumeName)
            try {
                & docker volume rm $volumeName | Out-Null
            } catch {
                Write-Host "WARNING: unable to remove Docker volume."
            }
        }
    }
} else {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    & docker run --rm -v "${pstDir}:/app/input" -v "${outputDir}:/app/output" support-tracker
    $sw.Stop()
    Write-Host ("Sheet fill time: {0}" -f $sw.Elapsed)
}

Write-Host "=== Done ==="
