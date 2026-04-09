param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ForwardArgs
)

$ErrorActionPreference = "Stop"

function Resolve-SupportTrackerRoot {
    $dRoot = "D:\Support_Tracker"
    $cRoot = "C:\Support_Tracker"

    if (Test-Path $dRoot) { return $dRoot }
    if (Test-Path $cRoot) { return $cRoot }
    if (Test-Path "D:\") { return $dRoot }
    return $cRoot
}

$root = Resolve-SupportTrackerRoot
$scriptsDir = Join-Path $root "Scripts"
$pstDir = Join-Path $root "PstFiles"
$outputDir = Join-Path $root "DockerOutput"
$selfPath = $MyInvocation.MyCommand.Path

foreach ($dir in @($root, $scriptsDir, $pstDir, $outputDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

$image = if ([string]::IsNullOrWhiteSpace($env:SUPPORT_TRACKER_IMAGE)) { "support-tracker" } else { $env:SUPPORT_TRACKER_IMAGE }
$runPipelinePath = Join-Path $scriptsDir "run_pipeline.ps1"
$exportPath = Join-Path $scriptsDir "export_outlook.ps1"

Write-Host ("Refreshing scripts into: {0}" -f $root)
try {
    & docker run --rm -e EMIT_SCRIPT=1 -v "${root}:/hostroot" $image | Out-Null
} catch {
    if ((Test-Path $runPipelinePath) -and (Test-Path $exportPath)) {
        Write-Host "WARNING: Unable to refresh scripts from Docker image. Using existing local scripts."
    } else {
        throw "Unable to refresh scripts from Docker image '$image'. $($_.Exception.Message)"
    }
}

if (-not (Test-Path $runPipelinePath)) {
    throw "Pipeline script not found: $runPipelinePath"
}

Write-Host ("Using root: {0}" -f $root)
Write-Host ("Using scripts: {0}" -f $scriptsDir)
Write-Host ("Using PST folder: {0}" -f $pstDir)
Write-Host ("Using output folder: {0}" -f $outputDir)

$installedScriptsDir = [System.IO.Path]::GetFullPath($scriptsDir)
$selfDir = if ($selfPath) { [System.IO.Path]::GetFullPath((Split-Path -Parent $selfPath)) } else { "" }
if ($selfPath -and ($selfDir -ne $installedScriptsDir)) {
    Write-Host ("Setup complete. Run next: powershell -ExecutionPolicy Bypass -File ""{0}""" -f $runPipelinePath)
    try {
        Remove-Item -LiteralPath $selfPath -Force -ErrorAction Stop
        Write-Host ("Removed bootstrap script: {0}" -f $selfPath)
    } catch {
        Write-Host ("WARNING: Unable to remove bootstrap script: {0}" -f $selfPath)
    }
    exit 0
}

Write-Host ("Next time, run: powershell -ExecutionPolicy Bypass -File ""{0}""" -f $runPipelinePath)

& powershell -ExecutionPolicy Bypass -File $runPipelinePath @ForwardArgs
exit $LASTEXITCODE
