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

$Script:DefaultScriptsDir = $PSScriptRoot
$Script:DefaultTrackerRoot = Split-Path -Parent $Script:DefaultScriptsDir
$Script:DefaultPstDir = Join-Path $Script:DefaultTrackerRoot "PstFiles"
$Script:DefaultOutputDir = Join-Path $Script:DefaultTrackerRoot "DockerOutput"

foreach ($dir in @($Script:DefaultTrackerRoot, $Script:DefaultScriptsDir, $Script:DefaultPstDir, $Script:DefaultOutputDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

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
    $exportResult = & $exportScript -FolderPathsInput $FolderPaths -StartDate $StartDate -EndDate $EndDate -OutputPstPath $OutputPstPath
    if (-not $?) {
        throw "Export script returned failure."
    }
    if ($exportResult) {
        $resultObj = @($exportResult) | Where-Object { $_ -and $_.PSObject.Properties['OutputPstPath'] } | Select-Object -Last 1
        if ($resultObj -and -not [string]::IsNullOrWhiteSpace($resultObj.OutputPstPath)) {
            $OutputPstPath = [string]$resultObj.OutputPstPath
        }
    }
} catch {
    Write-Host ""
    Write-Host "Export failed."
    Write-Host $_.Exception.Message
    exit 1
}

if (-not (Test-Path $OutputPstPath)) {
    throw "Export did not create/open the PST as expected: $OutputPstPath"
}

$pstDir = Split-Path -Parent $OutputPstPath
if ([string]::IsNullOrWhiteSpace($pstDir)) {
    throw "Unable to determine PST folder from path: $OutputPstPath"
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

function Close-OutlookGracefully {
    $outlookProcs = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    if (-not $outlookProcs) {
        return $true
    }

    Write-Host "Outlook is still running. Trying to close it safely..."

    try {
        $outlookApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        if ($outlookApp) {
            $outlookApp.Quit()
        }
    } catch {
    }

    foreach ($proc in @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) {
        try {
            if ($proc.MainWindowHandle -ne 0) {
                $null = $proc.CloseMainWindow()
            }
        } catch {
        }
    }

    $deadline = (Get-Date).AddSeconds(60)
    while ((Get-Date) -lt $deadline) {
        if (-not (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) {
            return $true
        }
        Start-Sleep -Seconds 2
    }

    return (-not (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue))
}

function Get-WorkbookKind {
    param([string]$FileName)
    $name = [regex]::Replace(($FileName | ForEach-Object { $_.ToLower() }), '[^a-z0-9]+', ' ')
    if ($name -match 'incident' -and $name -match 'self' -and $name -match 'service') {
        return 'incident_self_service'
    }
    if ($name -match 'task' -and $name -match 'business') {
        return 'task_business'
    }
    if ($name -match 'incident' -and $name -match 'business' -and $name -notmatch 'self') {
        return 'incident_business'
    }
    return ''
}

function Get-WorkbookLabel {
    param([string]$Kind)
    switch ($Kind) {
        'incident_business' { return 'Incident Business' }
        'task_business' { return 'Task Business' }
        'incident_self_service' { return 'Incident Self Service' }
        default { return $Kind }
    }
}

Write-Host "Step 1b: Ensure Outlook export is finished before Docker"
if (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) {
    $closed = Close-OutlookGracefully
    if (-not $closed) {
        Write-Host "Outlook did not close automatically. Please close Outlook to avoid PST file locks."
        Read-Host "Press Enter once Outlook is closed"
    }
}

while (-not (Test-FileUnlocked $OutputPstPath)) {
    $elapsed = (Get-Date) - $startWait
    if ($elapsed.TotalSeconds -gt $maxWaitSeconds) {
        throw "PST file is still locked after $maxWaitSeconds seconds. Please close Outlook and run again."
    }
    Write-Host "Waiting for PST file to be released..."
    Start-Sleep -Seconds $waitIntervalSeconds
}

$outputDir = $Script:DefaultOutputDir
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
Write-Host ("Using default Docker output folder: {0}" -f $outputDir)

Write-Host "Step 2: Run Docker pipeline"
Write-Host ("Using PST folder: {0}" -f $pstDir)
Write-Host ("Using output folder: {0}" -f $outputDir)

Write-Host "Running Docker..."
$useVolume = -not $NoVolume
if ($useVolume) {
    $volumeName = "support_tracker_output"
    $templates = @()
    while ($templates.Count -eq 0) {
        $candidates = Get-ChildItem $outputDir -Filter "*.xlsx" |
            Where-Object { $_.Name -notmatch "filled|done|automation_output" -and $_.Length -gt 0 } |
            Sort-Object LastWriteTime -Descending
        $recognized = @()
        $byKind = @{}
        foreach ($candidate in $candidates) {
            $kind = Get-WorkbookKind -FileName $candidate.Name
            if ([string]::IsNullOrWhiteSpace($kind)) {
                continue
            }
            if ($byKind.ContainsKey($kind)) {
                throw ("Multiple workbook files detected for {0}: {1}, {2}. Keep only one per workbook type." -f (Get-WorkbookLabel $kind), $byKind[$kind].Name, $candidate.Name)
            }
            $byKind[$kind] = $candidate
            $recognized += [PSCustomObject]@{
                Kind  = $kind
                Label = Get-WorkbookLabel -Kind $kind
                File  = $candidate
            }
        }
        $templates = @($recognized | Sort-Object @{Expression = {
            switch ($_.Kind) {
                'incident_business' { 1 }
                'incident_self_service' { 2 }
                'task_business' { 3 }
                default { 99 }
            }
        }}, @{Expression = { $_.File.Name }})
        if ($templates.Count -gt 0) {
            break
        }
        $zeroes = Get-ChildItem $outputDir -Filter "*.xlsx" -ErrorAction SilentlyContinue |
            Where-Object { $_.Length -eq 0 } |
            Select-Object -ExpandProperty Name
        Write-Host ("No recognized workbook .xlsx found in output folder: {0}" -f $outputDir)
        if ($zeroes) {
            Write-Host ("Found zero-byte .xlsx files: {0}" -f ($zeroes -join ", "))
        }
        $resume = Read-Host "Place the sheet in this folder, then press Enter to continue. Type quit to stop"
        if ($resume -match '^(?i)\s*(quit|exit)\s*$') {
            throw "Workbook template not provided."
        }
    }
    foreach ($templateInfo in $templates) {
        $template = $templateInfo.File
        try {
            $header = Get-Content -Path $template.FullName -Encoding Byte -TotalCount 2
            if ($header.Length -lt 2 -or $header[0] -ne 0x50 -or $header[1] -ne 0x4B) {
                throw "Template does not look like a valid .xlsx (missing PK header). Close Excel or choose the correct file."
            }
        } catch {
            throw "Unable to read template file. Close Excel if it is open, then run again. Details: $($_.Exception.Message)"
        }
    }
    Write-Host ("Using Docker volume for output: {0}" -f $volumeName)
    & docker volume create $volumeName | Out-Null
    foreach ($templateInfo in $templates) {
        $template = $templateInfo.File
        Write-Host ("Copying workbook into volume: {0}" -f $template.Name)
        & docker run --rm --entrypoint sh -v "${volumeName}:/app/output" -v "$($template.FullName):/host/template.xlsx:ro" support-tracker -c "cp /host/template.xlsx '/app/output/$($template.Name)'"
    }
    try {
        $totalSw = [Diagnostics.Stopwatch]::StartNew()
        foreach ($templateInfo in $templates) {
            $template = $templateInfo.File
            Write-Host ("Starting workbook fill: {0} ({1})" -f $templateInfo.Label, $template.Name)
            $sw = [Diagnostics.Stopwatch]::StartNew()
            & docker run --rm -e "TEMPLATE_PATH=/app/output/$($template.Name)" -v "${pstDir}:/app/input" -v "${volumeName}:/app/output" support-tracker
            $sw.Stop()
            Write-Host ("Finished workbook fill: {0} | Time: {1}" -f $templateInfo.Label, $sw.Elapsed)
            Write-Host ("Copying workbook back to host file: {0}" -f $template.Name)
            & docker run --rm --entrypoint sh -v "${volumeName}:/app/output" -v "${outputDir}:/host" support-tracker -c "cp '/app/output/$($template.Name)' '/host/$($template.Name)'"
        }
        $totalSw.Stop()
        Write-Host ("Total workbook fill time: {0}" -f $totalSw.Elapsed)
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
    $templates = @()
    $seenKinds = @{}
    Get-ChildItem $outputDir -Filter "*.xlsx" |
        Where-Object { $_.Name -notmatch "filled|done|automation_output" -and $_.Length -gt 0 } |
        Sort-Object LastWriteTime -Descending |
        ForEach-Object {
            $kind = Get-WorkbookKind -FileName $_.Name
            if ([string]::IsNullOrWhiteSpace($kind)) { return $null }
            if ($seenKinds.ContainsKey($kind)) {
                throw ("Multiple workbook files detected for {0}: {1}, {2}. Keep only one per workbook type." -f (Get-WorkbookLabel $kind), $seenKinds[$kind].Name, $_.Name)
            }
            $seenKinds[$kind] = $_
            $templates += [PSCustomObject]@{
                Kind  = $kind
                Label = Get-WorkbookLabel -Kind $kind
                File  = $_
            }
        }
    if (-not $templates -or $templates.Count -eq 0) {
        throw "No recognized workbook .xlsx found in output folder."
    }
    $totalSw = [Diagnostics.Stopwatch]::StartNew()
    foreach ($templateInfo in $templates) {
        $sw = [Diagnostics.Stopwatch]::StartNew()
        Write-Host ("Starting workbook fill: {0} ({1})" -f $templateInfo.Label, $templateInfo.File.Name)
        & docker run --rm -e "TEMPLATE_PATH=/app/output/$($templateInfo.File.Name)" -v "${pstDir}:/app/input" -v "${outputDir}:/app/output" support-tracker
        $sw.Stop()
        Write-Host ("Finished workbook fill: {0} | Time: {1}" -f $templateInfo.Label, $sw.Elapsed)
    }
    $totalSw.Stop()
    Write-Host ("Total workbook fill time: {0}" -f $totalSw.Elapsed)
}

Write-Host "=== Done ==="
