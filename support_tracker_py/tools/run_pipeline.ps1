param(
    [string]$FolderPaths = "",
    [string]$StartDate = "",
    [string]$EndDate = "",
    [string]$OutputPstPath = "",
    [switch]$NoVolume,
    [switch]$UseVolume,
    [switch]$KeepVolumeData,
    [switch]$CleanVolume
)

$ErrorActionPreference = "Stop"

function Resolve-TrackerRoot {
    if (-not [string]::IsNullOrWhiteSpace($env:SUPPORT_TRACKER_ROOT)) {
        return [System.IO.Path]::GetFullPath($env:SUPPORT_TRACKER_ROOT)
    }
    $dir = $PSScriptRoot
    for ($i = 0; $i -lt 8 -and -not [string]::IsNullOrWhiteSpace($dir); $i++) {
        $hasRuntimeDirs = (Test-Path (Join-Path $dir "PstFiles")) -or (Test-Path (Join-Path $dir "DockerOutput"))
        $hasRepoDir = Test-Path (Join-Path $dir "SupportTrackerAutomation")
        if ($hasRuntimeDirs -or $hasRepoDir) {
            return $dir
        }
        $parent = Split-Path -Parent $dir
        if ($parent -eq $dir) { break }
        $dir = $parent
    }
    return (Split-Path -Parent $PSScriptRoot)
}

$Script:DefaultScriptsDir = $PSScriptRoot
$Script:DefaultTrackerRoot = Resolve-TrackerRoot
$Script:DefaultPstDir = Join-Path $Script:DefaultTrackerRoot "PstFiles"
$Script:DefaultOutputDir = Join-Path $Script:DefaultTrackerRoot "DockerOutput"

function Write-Log {
    param(
        [string]$Level,
        [string]$Message
    )
    $label = switch ($Level.ToLower()) {
        "ok" { "ok" }
        "warn" { "warn" }
        "fail" { "fail" }
        "progress" { "progress" }
        "hint" { "hint" }
        default { "info" }
    }
    Write-Host ("[{0}] {1}" -f $label, $Message)
}

function Write-Section {
    param([string]$Title)
    Write-Host ""
    Write-Host ("== {0} ==" -f $Title)
}

foreach ($dir in @($Script:DefaultTrackerRoot, $Script:DefaultScriptsDir, $Script:DefaultPstDir, $Script:DefaultOutputDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

Write-Section "Support Tracker Pipeline"
Write-Log "info" "Step 1: export Outlook folders to PST"

$exportScript = Join-Path $PSScriptRoot "export_outlook.ps1"
Write-Log "info" ("Export script: {0}" -f $exportScript)
try {
    $verLine = Select-String -Path $exportScript -Pattern "Export script version" | Select-Object -First 1
    if ($verLine) {
        Write-Log "info" $verLine.Line
    } else {
        Write-Log "warn" "Export script version not found; script may be outdated."
    }
} catch {
    Write-Log "warn" "Unable to read export script version."
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
    Write-Section "Export Failed"
    Write-Log "fail" ($_.Exception.Message)
    exit 1
}

if (-not (Test-Path $OutputPstPath)) {
    throw "Export did not create/open the PST as expected: $OutputPstPath"
}

$pstDir = Split-Path -Parent $OutputPstPath
if ([string]::IsNullOrWhiteSpace($pstDir)) {
    throw "Unable to determine PST folder from path: $OutputPstPath"
}

$maxWaitSeconds = 120
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
    param([int]$WaitSeconds = 20)

    $outlookProcs = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)
    if (-not $outlookProcs) {
        return $true
    }

    Write-Log "info" "Outlook is still running; trying a bounded close."

    try {
        $outlookApp = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application")
        if ($outlookApp) {
            $outlookApp.Quit()
        }
    } catch {
    } finally {
        try {
            if ($outlookApp) {
                [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlookApp)
            }
        } catch {
        }
        $outlookApp = $null
    }

    foreach ($proc in @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) {
        try {
            if ($proc.MainWindowHandle -ne 0) {
                $null = $proc.CloseMainWindow()
            }
        } catch {
        }
    }

    $deadline = (Get-Date).AddSeconds($WaitSeconds)
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

function Get-TemplateOutputStem {
    param([string]$FileName)
    $stem = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $stem = $stem.ToLowerInvariant()
    $stem = [regex]::Replace($stem, '[^a-z0-9]+', '_').Trim('_')
    return $stem
}

function Assert-LastExitCode {
    param([string]$Context)
    if ($LASTEXITCODE -ne 0) {
        throw ("{0} failed with exit code {1}." -f $Context, $LASTEXITCODE)
    }
}

function Clear-TemplateArtifactsInVolume {
    param(
        [string]$VolumeName,
        [string]$TemplateName
    )
    $stem = Get-TemplateOutputStem -FileName $TemplateName
    & docker run --rm --entrypoint sh -v "${VolumeName}:/app/output" support-tracker -c "rm -f '/app/output/automation_output_${stem}.csv' '/app/output/debug_subjects_${stem}.csv'"
    Assert-LastExitCode ("Clearing prior CSV artifacts for {0}" -f $TemplateName)
    return $stem
}

function Get-TemplateArtifactsFromVolume {
    param(
        [string]$VolumeName,
        [string]$TemplateName
    )
    $stem = Get-TemplateOutputStem -FileName $TemplateName
    $cmd = 'a=''/app/output/automation_output_{0}.csv''; d=''/app/output/debug_subjects_{0}.csv''; test -s "$a" && test -s "$d" || exit 11; a_lines=$(wc -l < "$a"); d_lines=$(wc -l < "$d"); [ "$a_lines" -gt 1 ] && [ "$d_lines" -gt 1 ] || exit 12; printf ''%s|%s'' "$a_lines" "$d_lines"' -f $stem
    $summary = & docker run --rm --entrypoint sh -v "${VolumeName}:/app/output" support-tracker -c $cmd
    Assert-LastExitCode ("Verifying fresh CSV outputs for {0}" -f $TemplateName)
    $parts = (($summary | Out-String).Trim()) -split '\|'
    if ($parts.Count -ne 2) {
        throw ("Unexpected CSV verification output for {0}: {1}" -f $TemplateName, $summary)
    }
    return [PSCustomObject]@{
        AutomationLines = [int]$parts[0]
        DebugLines = [int]$parts[1]
    }
}

function Clear-TemplateArtifactsOnHost {
    param(
        [string]$OutputDir,
        [string]$TemplateName
    )
    $stem = Get-TemplateOutputStem -FileName $TemplateName
    foreach ($name in @("automation_output_${stem}.csv", "debug_subjects_${stem}.csv")) {
        $path = Join-Path $OutputDir $name
        if (Test-Path $path) {
            Remove-Item -LiteralPath $path -Force
        }
    }
    return $stem
}

function Get-TemplateArtifactsOnHost {
    param(
        [string]$OutputDir,
        [string]$TemplateName
    )
    $stem = Get-TemplateOutputStem -FileName $TemplateName
    $automationPath = Join-Path $OutputDir "automation_output_${stem}.csv"
    $debugPath = Join-Path $OutputDir "debug_subjects_${stem}.csv"
    foreach ($path in @($automationPath, $debugPath)) {
        if (-not (Test-Path $path)) {
            throw ("Expected output file was not created: {0}" -f $path)
        }
    }
    $automationLines = (Get-Content -LiteralPath $automationPath | Measure-Object -Line).Lines
    $debugLines = (Get-Content -LiteralPath $debugPath | Measure-Object -Line).Lines
    if ($automationLines -le 1 -or $debugLines -le 1) {
        throw ("Output CSVs look empty for {0}: automation={1}, debug={2}" -f $TemplateName, $automationLines, $debugLines)
    }
    return [PSCustomObject]@{
        AutomationLines = [int]$automationLines
        DebugLines = [int]$debugLines
    }
}

function Assert-WorkbookCopiedBack {
    param([string]$WorkbookPath)
    if (-not (Test-Path $WorkbookPath)) {
        throw ("Workbook was not copied back: {0}" -f $WorkbookPath)
    }
    $item = Get-Item -LiteralPath $WorkbookPath
    if ($item.Length -le 0) {
        throw ("Workbook copy looks empty: {0}" -f $WorkbookPath)
    }
}

function Reset-VolumeRunArtifacts {
    param([string]$VolumeName)
    $cmd = @'
find /app/output -mindepth 1 -maxdepth 1 -exec rm -rf {} +
'@
    & docker run --rm --entrypoint sh -v "${VolumeName}:/app/output" support-tracker -c $cmd
    Assert-LastExitCode ("Resetting run artifacts in Docker volume {0}" -f $VolumeName)
}

Write-Section "PST Lock Check"
if (-not (Test-FileUnlocked $OutputPstPath)) {
    Write-Log "warn" "PST is still locked; trying bounded Outlook close."
    if (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue) {
        $closed = Close-OutlookGracefully -WaitSeconds 20
        if (-not $closed) {
            Write-Log "warn" "Outlook still running after 20 seconds; waiting only for the PST lock now."
        }
    }
} else {
    Write-Log "ok" "PST file is already released; skipping Outlook close."
}

while (-not (Test-FileUnlocked $OutputPstPath)) {
    $elapsed = (Get-Date) - $startWait
    if ($elapsed.TotalSeconds -gt $maxWaitSeconds) {
        throw "PST file is still locked after $maxWaitSeconds seconds. Please close Outlook and run again."
    }
    Write-Log "progress" "Waiting for PST file to be released..."
    Start-Sleep -Seconds $waitIntervalSeconds
}

$outputDir = $Script:DefaultOutputDir
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
Write-Log "info" ("Docker output folder: {0}" -f $outputDir)

Write-Section "Docker Pipeline"
Write-Log "info" ("PST folder: {0}" -f $pstDir)
Write-Log "info" ("Output folder: {0}" -f $outputDir)

Write-Log "info" "Starting Docker..."
$useVolume = $UseVolume -and -not $NoVolume
if ($useVolume) {
    Write-Log "info" "Docker output mode: named volume"
} else {
    Write-Log "info" "Docker output mode: direct bind mount"
}
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
        Write-Log "warn" ("No recognized workbook .xlsx found in output folder: {0}" -f $outputDir)
        if ($zeroes) {
            Write-Log "warn" ("Zero-byte .xlsx files: {0}" -f ($zeroes -join ", "))
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
    Write-Log "info" ("Docker volume: {0}" -f $volumeName)
    & docker volume create $volumeName | Out-Null
    Assert-LastExitCode ("Creating Docker volume {0}" -f $volumeName)
    if ($KeepVolumeData) {
        Write-Log "info" "Keeping Docker volume after run; clearing prior artifacts before starting."
    } else {
        Write-Log "info" "Clearing prior Docker volume artifacts before starting."
    }
    Reset-VolumeRunArtifacts -VolumeName $volumeName
    foreach ($templateInfo in $templates) {
        $template = $templateInfo.File
        Write-Log "info" ("Copy workbook into volume: {0}" -f $template.Name)
        & docker run --rm --entrypoint sh -v "${volumeName}:/app/output" -v "$($template.FullName):/host/template.xlsx:ro" support-tracker -c "cp /host/template.xlsx '/app/output/$($template.Name)'"
        Assert-LastExitCode ("Copying workbook {0} into Docker volume" -f $template.Name)
    }
    try {
        $totalSw = [Diagnostics.Stopwatch]::StartNew()
        foreach ($templateInfo in $templates) {
            $template = $templateInfo.File
            $null = Clear-TemplateArtifactsInVolume -VolumeName $volumeName -TemplateName $template.Name
            Write-Section ("Workbook: {0}" -f $templateInfo.Label)
            Write-Log "info" ("Template: {0}" -f $template.Name)
            $sw = [Diagnostics.Stopwatch]::StartNew()
            & docker run --rm -e "TEMPLATE_PATH=/app/output/$($template.Name)" -v "${pstDir}:/app/input" -v "${volumeName}:/app/output" support-tracker
            Assert-LastExitCode ("Docker workbook fill for {0}" -f $template.Name)
            $sw.Stop()
            $csvSummary = Get-TemplateArtifactsFromVolume -VolumeName $volumeName -TemplateName $template.Name
            Write-Log "ok" ("Workbook fill complete; time={0}" -f $sw.Elapsed)
            Write-Log "ok" ("Fresh outputs: dataRows={0}; debugRows={1}" -f ($csvSummary.AutomationLines - 1), ($csvSummary.DebugLines - 1))
            Write-Log "info" ("Copy workbook back: {0}" -f $template.Name)
            & docker run --rm --entrypoint sh -v "${volumeName}:/app/output" -v "${outputDir}:/host" support-tracker -c "cp '/app/output/$($template.Name)' '/host/$($template.Name)'"
            Assert-LastExitCode ("Copying workbook {0} back to host" -f $template.Name)
            Assert-WorkbookCopiedBack -WorkbookPath (Join-Path $outputDir $template.Name)
        }
        $totalSw.Stop()
        Write-Log "ok" ("Total workbook fill time: {0}" -f $totalSw.Elapsed)
    } finally {
        if (-not $KeepVolumeData) {
            Write-Log "info" "Clearing volume data..."
            & docker run --rm --entrypoint sh -v "${volumeName}:/app/output" support-tracker -c "find /app/output -mindepth 1 -maxdepth 1 -exec rm -rf {} +"
            Assert-LastExitCode ("Clearing Docker volume {0}" -f $volumeName)
            Write-Log "ok" "Volume cleanup complete."
        } else {
            Write-Log "info" "Keeping volume data."
        }
        if ($CleanVolume) {
            if ($KeepVolumeData) {
                Write-Log "warn" "-CleanVolume will remove the kept volume."
            }
            Write-Log "info" ("Removing Docker volume: {0}" -f $volumeName)
            try {
                & docker volume rm $volumeName | Out-Null
            } catch {
                Write-Log "warn" "Unable to remove Docker volume."
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
        $null = Clear-TemplateArtifactsOnHost -OutputDir $outputDir -TemplateName $templateInfo.File.Name
        Write-Section ("Workbook: {0}" -f $templateInfo.Label)
        Write-Log "info" ("Template: {0}" -f $templateInfo.File.Name)
        & docker run --rm -e "TEMPLATE_PATH=/app/output/$($templateInfo.File.Name)" -v "${pstDir}:/app/input" -v "${outputDir}:/app/output" support-tracker
        Assert-LastExitCode ("Docker workbook fill for {0}" -f $templateInfo.File.Name)
        $sw.Stop()
        $csvSummary = Get-TemplateArtifactsOnHost -OutputDir $outputDir -TemplateName $templateInfo.File.Name
        Write-Log "ok" ("Workbook fill complete; time={0}" -f $sw.Elapsed)
        Write-Log "ok" ("Fresh outputs: dataRows={0}; debugRows={1}" -f ($csvSummary.AutomationLines - 1), ($csvSummary.DebugLines - 1))
        Assert-WorkbookCopiedBack -WorkbookPath $templateInfo.File.FullName
    }
    $totalSw.Stop()
    Write-Log "ok" ("Total workbook fill time: {0}" -f $totalSw.Elapsed)
}

Write-Section "Done"
