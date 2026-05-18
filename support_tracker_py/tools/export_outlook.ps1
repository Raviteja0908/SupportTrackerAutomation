param(
    [string]$FolderPathsInput,
    [string]$StartDate,
    [string]$EndDate,
    [string]$OutputPstPath
)

$ErrorActionPreference = "Stop"

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

function Release-ComObjectQuietly {
    param([object]$ComObject)
    try {
        if ($ComObject) {
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ComObject)
        }
    } catch {
    }
}

$includeCreationTime = $true
if ($env:EXCLUDE_CREATION_TIME) {
    $val = $env:EXCLUDE_CREATION_TIME.ToString().ToLower()
    if ($val -in @("1", "true", "yes", "y", "on")) { $includeCreationTime = $false }
}

Write-Section "Outlook Export"
Write-Log "info" "Mode: PST append, copy only"
Write-Log "info" ("Start: {0}" -f (Get-Date))
Write-Log "info" "Script version: 2026-04-06-1"
Write-Log "hint" "Prompt commands: back, quit"

$Script:BackToken = "__SCRIPT_BACK__"
$Script:DefaultScriptsDir = $PSScriptRoot
$Script:DefaultTrackerRoot = Resolve-TrackerRoot
$Script:DefaultPstDir = Join-Path $Script:DefaultTrackerRoot "PstFiles"
$Script:DefaultOutputDir = Join-Path $Script:DefaultTrackerRoot "DockerOutput"
$Script:DefaultPstPath = Join-Path $Script:DefaultPstDir "export_filtered.pst"

foreach ($dir in @($Script:DefaultTrackerRoot, $Script:DefaultScriptsDir, $Script:DefaultPstDir, $Script:DefaultOutputDir)) {
    if (-not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
}

Write-Section "Outlook"
Write-Log "info" "Connecting..."
$outlookWasRunning = @(Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0
$startupPstHint = if ([string]::IsNullOrWhiteSpace($OutputPstPath)) { $Script:DefaultPstPath } else { $OutputPstPath }
if ($startupPstHint -match '^[\\/]' -and $startupPstHint -notmatch '^[A-Za-z]:') {
    $defaultDrive = [System.IO.Path]::GetPathRoot($Script:DefaultTrackerRoot).TrimEnd('\')
    $startupPstHint = "$defaultDrive$startupPstHint"
}
$outlookReadyHintTimer = $null
if (-not (Test-Path $startupPstHint)) {
    $timerCallback = [System.Threading.TimerCallback]{
        param($state)
        Write-Host "Outlook seems to be waiting on a missing export PST popup. Click OK once, then the script will continue."
    }
    $outlookReadyHintTimer = New-Object System.Threading.Timer($timerCallback, $null, 8000, [System.Threading.Timeout]::Infinite)
}
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
if ($outlookReadyHintTimer) {
    $outlookReadyHintTimer.Dispose()
}

$fastExport = $true
if ($env:FAST_EXPORT) {
    $val = $env:FAST_EXPORT.ToString().ToLower()
    if ($val -in @("0", "false", "no", "n", "off")) {
        $fastExport = $false
    } elseif ($val -in @("1", "true", "yes", "y", "on")) {
        $fastExport = $true
    }
}
Write-Log "info" ("Export mode: {0}" -f ($(if ($fastExport) { "FAST (MailItem + ReceivedTime)" } else { "COMPLETE (All items + SentOn fallback)" })))

$closeOutlookAfterExport = $false
if ($env:CLOSE_OUTLOOK_AFTER_EXPORT) {
    $val = $env:CLOSE_OUTLOOK_AFTER_EXPORT.ToString().ToLower()
    if ($val -in @("1", "true", "yes", "y", "on")) {
        $closeOutlookAfterExport = $true
    }
}

function Split-InboxCombinedPath {
    param([string]$path)
    if ([string]::IsNullOrWhiteSpace($path)) { return @() }
    $p = $path.Trim()
    $lower = $p.ToLower()
    $needle = "inbox\\"
    $indices = @()
    $idx = $lower.IndexOf($needle)
    while ($idx -ge 0) {
        $indices += $idx
        $idx = $lower.IndexOf($needle, $idx + 1)
    }
    if ($indices.Count -le 1) { return @($p) }
    $parts = @()
    for ($i = 0; $i -lt $indices.Count; $i++) {
        $start = $indices[$i]
        $end = if ($i -lt $indices.Count - 1) { $indices[$i + 1] } else { $p.Length }
        $segment = $p.Substring($start, $end - $start).Trim()
        if (-not [string]::IsNullOrWhiteSpace($segment)) {
            $parts += $segment
        }
    }
    return $parts
}

function Read-ScriptInput {
    param(
        [string]$Prompt,
        [switch]$AllowBlank,
        [switch]$AllowBack
    )
    while ($true) {
        $value = Read-Host $Prompt
        if ($null -eq $value) { $value = "" }
        $trimmed = $value.Trim()
        if ($trimmed -ieq ":quit" -or $trimmed -ieq ":exit" -or $trimmed -ieq "quit" -or $trimmed -ieq "exit") {
            throw "User cancelled export."
        }
        if ($trimmed -ieq ":back" -or $trimmed -ieq "back") {
            if ($AllowBack) {
                return $Script:BackToken
            }
            Write-Host "Back is not available for this prompt. Type quit to exit."
            continue
        }
        if (-not $AllowBlank -and [string]::IsNullOrWhiteSpace($value)) {
            Write-Host "Input cannot be empty. Type back to return when supported, or quit to exit."
            continue
        }
        return $value
    }
}

function Read-DateValue {
    param(
        [string]$Label,
        [string]$Initial,
        [switch]$AllowBack
    )
    $current = $Initial
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($current)) {
            $current = Read-ScriptInput -Prompt "$Label (DD-MM-YYYY)" -AllowBack:$AllowBack
        }
        if ($current -eq $Script:BackToken) {
            return @($Script:BackToken, $null)
        }
        $dt = [DateTime]::MinValue
        $ok = [DateTime]::TryParseExact($current, "dd-MM-yyyy", $null, [Globalization.DateTimeStyles]::None, [ref]$dt)
        if ($ok) {
            return @($current, $dt)
        }
        Write-Host "Invalid date format. Use DD-MM-YYYY."
        $current = $null
    }
}

function Test-ProtectedStoreName {
    param([string]$text)
    if ([string]::IsNullOrWhiteSpace($text)) { return $false }
    return $text.ToLower().Contains("@invenio-solutions.com")
}

function Get-SafeStores {
    $stores = @()
    $count = 0
    try {
        $count = [int]$namespace.Stores.Count
    } catch {
        return @()
    }
    for ($i = 1; $i -le $count; $i++) {
        try {
            $store = $namespace.Stores.Item($i)
            if ($store) {
                $stores += $store
            }
        } catch {
            # Skip broken/stale entries without failing the whole export.
            continue
        }
    }
    return $stores
}

function Remove-StaleManagedPstStores {
    param(
        [string]$ManagedDir,
        [string]$TargetPstPath = ""
    )
    $managedDirLower = ""
    if (-not [string]::IsNullOrWhiteSpace($ManagedDir)) {
        $managedDirLower = [System.IO.Path]::GetFullPath($ManagedDir).TrimEnd('\').ToLower()
    }
    $targetLower = ""
    if (-not [string]::IsNullOrWhiteSpace($TargetPstPath)) {
        try {
            $targetLower = [System.IO.Path]::GetFullPath($TargetPstPath).ToLower()
        } catch {
            $targetLower = $TargetPstPath.ToLower()
        }
    }

    foreach ($store in (Get-SafeStores)) {
        $displayName = $null
        $root = $null
        $rootName = $null
        $storePath = $null
        try { $displayName = [string]$store.DisplayName } catch { $displayName = $null }
        try {
            $root = $store.GetRootFolder()
            if ($root) { $rootName = [string]$root.Name }
        } catch {
            $root = $null
            $rootName = $null
        }
        try { $storePath = [string]$store.FilePath } catch { $storePath = $null }

        if (Test-ProtectedStoreName $displayName) { continue }
        if (Test-ProtectedStoreName $rootName) { continue }
        if (Test-ProtectedStoreName $storePath) { continue }
        if ([string]::IsNullOrWhiteSpace($storePath)) { continue }
        if (-not $storePath.ToLower().EndsWith(".pst")) { continue }

        $storePathLower = ""
        $storeDirLower = ""
        try {
            $storePathLower = [System.IO.Path]::GetFullPath($storePath).ToLower()
            $storeDirLower = [System.IO.Path]::GetDirectoryName($storePathLower)
        } catch {
            $storePathLower = $storePath.ToLower()
            $storeDirLower = (Split-Path -Parent $storePath).ToLower()
        }

        $matchesTarget = (-not [string]::IsNullOrWhiteSpace($targetLower)) -and ($storePathLower -eq $targetLower)
        $inManagedDir = (-not [string]::IsNullOrWhiteSpace($managedDirLower)) -and ($storeDirLower -eq $managedDirLower)
        if (-not $matchesTarget -and -not $inManagedDir) { continue }
        if (Test-Path $storePath) { continue }
        if (-not $root) { continue }

        try {
            $namespace.RemoveStore($root)
            Write-Log "ok" ("Removed stale PST store: {0}" -f $storePath)
        } catch {
            Write-Log "warn" ("Unable to remove stale PST store: {0}" -f $storePath)
        }
    }
}

function Remove-ManagedPstStoreByPath {
    param(
        [string]$TargetPstPath,
        [int]$Retries = 3,
        [int]$RetryDelaySeconds = 2
    )
    if ([string]::IsNullOrWhiteSpace($TargetPstPath)) { return $false }

    $targetLower = ""
    try {
        $targetLower = [System.IO.Path]::GetFullPath($TargetPstPath).ToLower()
    } catch {
        $targetLower = $TargetPstPath.ToLower()
    }

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        $matchingStore = $null
        $root = $null
        try {
            foreach ($store in (Get-SafeStores)) {
                $displayName = $null
                $rootName = $null
                $storePath = $null
                try { $displayName = [string]$store.DisplayName } catch { $displayName = $null }
                try {
                    $root = $store.GetRootFolder()
                    if ($root) { $rootName = [string]$root.Name }
                } catch {
                    $root = $null
                    $rootName = $null
                }
                try { $storePath = [string]$store.FilePath } catch { $storePath = $null }

                if (Test-ProtectedStoreName $displayName) { continue }
                if (Test-ProtectedStoreName $rootName) { continue }
                if (Test-ProtectedStoreName $storePath) { continue }
                if ([string]::IsNullOrWhiteSpace($storePath)) { continue }
                if (-not $storePath.ToLower().EndsWith(".pst")) { continue }

                $storePathLower = ""
                try {
                    $storePathLower = [System.IO.Path]::GetFullPath($storePath).ToLower()
                } catch {
                    $storePathLower = $storePath.ToLower()
                }

                if ($storePathLower -ne $targetLower) { continue }
                if (-not $root) { continue }

                $matchingStore = $store
                break
            }

            if (-not $matchingStore -or -not $root) {
                return $true
            }

            $namespace.RemoveStore($root)
            Write-Log "ok" ("Detached export PST: {0}" -f $TargetPstPath)
            Start-Sleep -Seconds 1
            return $true
        } catch {
            if ($attempt -ge $Retries) {
                Write-Log "warn" ("Unable to detach export PST after {0} attempt(s): {1}" -f $Retries, $TargetPstPath)
                return $false
            }
            Write-Log "warn" ("Detach attempt {0}/{1} failed; retrying: {2}" -f $attempt, $Retries, $TargetPstPath)
            Start-Sleep -Seconds $RetryDelaySeconds
        } finally {
            foreach ($comObj in @($root, $matchingStore)) {
                try {
                    if ($comObj) {
                        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObj)
                    }
                } catch {
                }
            }
            $root = $null
            $matchingStore = $null
        }
    }

    return (-not (Find-StoreByPath -targetPath $TargetPstPath))
}

Remove-StaleManagedPstStores -ManagedDir $Script:DefaultPstDir

# Build store lookup (display name + root name) for shared mailboxes
$storeInfos = @()
foreach ($s in (Get-SafeStores)) {
    $root = $null
    $displayName = $null
    try { $root = $s.GetRootFolder() } catch { $root = $null }
    try { $displayName = $s.DisplayName } catch { $displayName = $null }
    if (-not $displayName -and -not $root) {
        # Ignore stale/broken Outlook store entries safely.
        continue
    }
    $storeInfos += [PSCustomObject]@{
        Store       = $s
        DisplayName = $displayName
        RootFolder  = $root
        RootName    = if ($root) { $root.Name } else { $null }
    }
}
Write-Host "Available mailboxes/stores:"
foreach ($info in $storeInfos) {
    if ($info.DisplayName) {
        Write-Host (" - {0}" -f $info.DisplayName)
    } elseif ($info.RootName) {
        Write-Host (" - {0}" -f $info.RootName)
    }
}

$defaultMailboxName = $null
$defaultInboxFolder = $null

function Normalize-UserPath {
    param([string]$raw)
    $clean = ($raw -replace "/", "\\").Replace('"', '').Replace("'", '').Trim()
    if ([string]::IsNullOrWhiteSpace($clean)) { return $null }
    $clean = $clean.TrimStart("\")
    $storePrefixMatch = $false
    foreach ($info in $storeInfos) {
        if ($info.DisplayName -and $clean -like "$($info.DisplayName)\\*") {
            $storePrefixMatch = $true
            break
        }
        if ($info.RootName -and $clean -like "$($info.RootName)\\*") {
            $storePrefixMatch = $true
            break
        }
    }
    if (-not ($clean -match '^(?i)inbox\\') -and -not $storePrefixMatch) {
        $clean = "Inbox\$clean"
    }
    return $clean
}

function Get-StoreInfo {
    param([string]$storeName)
    if ([string]::IsNullOrWhiteSpace($storeName)) { return $null }
    foreach ($info in $storeInfos) {
        if ($info.DisplayName -and $info.DisplayName -eq $storeName) { return $info }
        if ($info.RootName -and $info.RootName -eq $storeName) { return $info }
    }
    foreach ($info in $storeInfos) {
        if ($info.DisplayName -and $info.DisplayName.ToLower() -eq $storeName.ToLower()) { return $info }
        if ($info.RootName -and $info.RootName.ToLower() -eq $storeName.ToLower()) { return $info }
    }
    return $null
}

function Get-StoreRootFolder {
    param([string]$storeName)
    $info = Get-StoreInfo -storeName $storeName
    if ($info) { return $info.RootFolder }
    return $null
}

function Get-StoreInboxFolder {
    param([string]$storeName)
    $info = Get-StoreInfo -storeName $storeName
    if (-not $info) { return $null }
    $root = $info.RootFolder
    if (-not $root) { return $null }
    try {
        $inbox = $info.Store.GetDefaultFolder(6)
        if ($inbox) { return $inbox }
    } catch {
        # ignore and fallback to root child lookup
    }
    foreach ($f in @($root.Folders)) {
        if ($f.Name -eq "Inbox") {
            return $f
        }
    }
    return $null
}

function Find-StoreByPath {
    param([string]$targetPath)
    if ([string]::IsNullOrWhiteSpace($targetPath)) { return $null }
    $targetLower = $targetPath.ToLower()
    foreach ($store in (Get-SafeStores)) {
        $storePath = $null
        try {
            $storePath = $store.FilePath
        } catch {
            # Ignore stale/broken Outlook store entries safely.
            continue
        }
        if ([string]::IsNullOrWhiteSpace($storePath)) {
            continue
        }
        if ($storePath.ToLower() -eq $targetLower) {
            return $store
        }
    }
    return $null
}

function Wait-ForStoreByPath {
    param(
        [string]$TargetPath,
        [int]$Retries = 8,
        [int]$RetryDelaySeconds = 2
    )
    if ([string]::IsNullOrWhiteSpace($TargetPath)) { return $null }
    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        $store = Find-StoreByPath -targetPath $TargetPath
        if ($store) {
            if ($attempt -gt 1) {
                Write-Host ("PST store became available after retry {0}/{1}." -f $attempt, $Retries)
            }
            return $store
        }
        if ($attempt -lt $Retries) {
            Start-Sleep -Seconds $RetryDelaySeconds
        }
    }
    return $null
}

function Resolve-ChildPath {
    param(
        [object]$root,
        [string[]]$parts
    )
    $current = $root
    foreach ($seg in $parts) {
        $next = $null
        foreach ($f in @($current.Folders)) {
            if ($f.Name -eq $seg) {
                $next = $f
                break
            }
        }
        if (-not $next) {
            return $null
        }
        $current = $next
    }
    return $current
}

function Resolve-RelativeInboxAcrossStores {
    param(
        [string[]]$parts
    )
    foreach ($info in $storeInfos) {
        $inboxFolder = $null
        if ($info.DisplayName) {
            $inboxFolder = Get-StoreInboxFolder -storeName $info.DisplayName
        }
        if (-not $inboxFolder -and $info.RootName) {
            $inboxFolder = Get-StoreInboxFolder -storeName $info.RootName
        }
        if (-not $inboxFolder) { continue }
        $resolved = Resolve-ChildPath -root $inboxFolder -parts $parts
        if ($resolved) {
            $storeName = if ($info.DisplayName) { $info.DisplayName } else { $info.RootName }
            Write-Host ("FOUND under store: {0}" -f $storeName)
            return $resolved
        }
    }
    return $null
}

function Debug-Resolve-FolderNotFound {
    param(
        [string]$path,
        [string[]]$relativeParts
    )
    return
}

function Enumerate-Folders {
    param(
        [object]$root,
        [string]$prefix
    )
    $stack = New-Object System.Collections.Stack
    $stack.Push([PSCustomObject]@{ Folder = $root; Prefix = $prefix })
    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        foreach ($f in @($current.Folder.Folders)) {
            $path = if ([string]::IsNullOrWhiteSpace($current.Prefix)) { $f.Name } else { "$($current.Prefix)\\$($f.Name)" }
            [PSCustomObject]@{ Folder = $f; Path = $path }
            $stack.Push([PSCustomObject]@{ Folder = $f; Prefix = $path })
        }
    }
}

function Find-FolderByLeafNameAcrossStores {
    param(
        [string]$leafName
    )
    $found = @()
    foreach ($info in $storeInfos) {
        $root = $info.RootFolder
        if (-not $root) { continue }
        $storeName = if ($info.DisplayName) { $info.DisplayName } else { $info.RootName }
        foreach ($item in (Enumerate-Folders -root $root -prefix $storeName)) {
            if ($item.Path -match "\\$([regex]::Escape($leafName))$") {
                $found += $item
            }
        }
    }
    return $found
}

function Print-FolderTree {
    param(
        [object]$root,
        [string]$prefix,
        [int]$depth,
        [int]$maxDepth
    )
    $stack = New-Object System.Collections.Stack
    $stack.Push([PSCustomObject]@{ Folder = $root; Prefix = $prefix; Depth = $depth })
    while ($stack.Count -gt 0) {
        $current = $stack.Pop()
        if ($current.Depth -ge $maxDepth) { continue }
        $children = @($current.Folder.Folders)
        for ($i = $children.Count - 1; $i -ge 0; $i--) {
            $f = $children[$i]
            $path = if ([string]::IsNullOrWhiteSpace($current.Prefix)) { $f.Name } else { "$($current.Prefix)\\$($f.Name)" }
            Write-Host ("{0}{1}" -f (" " * ($current.Depth * 2)), $path)
            $stack.Push([PSCustomObject]@{ Folder = $f; Prefix = $path; Depth = ($current.Depth + 1) })
        }
    }
}

function Prompt-InteractiveExportDetails {
    param(
        [string]$InitialStartDate,
        [string]$InitialEndDate,
        [string]$InitialOutputPstPath
    )

    $listChoice = ""
    $mb = ""
    $inboxOnly = ""
    $usedListing = $false
    $folderCount = 0
    $folderPathsLocal = @()
    $folderIndex = 1
    $startDateLocal = $InitialStartDate
    $startDateObjLocal = $null
    $endDateLocal = $InitialEndDate
    $endDateObjLocal = $null
    $outputPstLocal = $InitialOutputPstPath
    $state = "listChoice"
    $maxDepth = 6

    while ($true) {
        switch ($state) {
            "listChoice" {
                $listChoice = Read-ScriptInput -Prompt "List folder paths for a mailbox? (y/N)" -AllowBlank -AllowBack
                if ($listChoice -eq $Script:BackToken) {
                    Write-Host "Back is not available before the first prompt."
                    continue
                }
                if ($listChoice -match '^(y|yes)$') {
                    $usedListing = $true
                    $state = "mailbox"
                } else {
                    $usedListing = $false
                    $state = "count"
                }
                continue
            }
            "mailbox" {
                $mb = Read-ScriptInput -Prompt "Mailbox/store name (exact as listed above)" -AllowBlank -AllowBack
                if ($mb -eq $Script:BackToken) {
                    $state = "listChoice"
                    continue
                }
                $storeInfo = if (-not [string]::IsNullOrWhiteSpace($mb)) { Get-StoreInfo -storeName $mb } else { $null }
                if (-not ($storeInfo -and $storeInfo.RootFolder)) {
                    Write-Host "Store not found. Type back to return, or quit to exit."
                    continue
                }
                $defaultMailboxName = $mb
                $state = "inboxOnly"
                continue
            }
            "inboxOnly" {
                $inboxOnly = Read-ScriptInput -Prompt "Start at Inbox only? (y/N)" -AllowBlank -AllowBack
                if ($inboxOnly -eq $Script:BackToken) {
                    $state = "mailbox"
                    continue
                }
                if ($inboxOnly -match '^(y|yes)$') {
                    $root = Get-StoreInboxFolder -storeName $mb
                    $defaultInboxFolder = $root
                    Write-Host ("Folder paths under: {0}\\Inbox" -f $mb)
                    if ($root) {
                        Print-FolderTree -root $root -prefix ("$mb\\Inbox") -depth 0 -maxDepth $maxDepth
                    } else {
                        Write-Host "Inbox not found for that mailbox."
                    }
                } else {
                    $storeInfo = Get-StoreInfo -storeName $mb
                    Write-Host ("Folder paths under: {0}" -f $mb)
                    Print-FolderTree -root $storeInfo.RootFolder -prefix $mb -depth 0 -maxDepth $maxDepth
                }
                $state = "count"
                continue
            }
            "count" {
                $countStr = Read-ScriptInput -Prompt 'How many folders to export?' -AllowBack
                if ($countStr -eq $Script:BackToken) {
                    if ($usedListing) {
                        $state = "inboxOnly"
                    } else {
                        $state = "listChoice"
                    }
                    continue
                }
                if (-not ($countStr -match '^\d+$') -or [int]$countStr -le 0) {
                    Write-Host 'Please enter a valid positive number.'
                    continue
                }
                $folderCount = [int]$countStr
                $folderPathsLocal = New-Object string[] $folderCount
                $folderIndex = 1
                $state = "folderPath"
                continue
            }
            "folderPath" {
                $pRaw = Read-ScriptInput -Prompt ("Folder $folderIndex path (under Inbox)") -AllowBack
                if ($pRaw -eq $Script:BackToken) {
                    if ($folderIndex -gt 1) {
                        $folderIndex -= 1
                        $folderPathsLocal[$folderIndex - 1] = $null
                        Write-Host ("Back to Folder {0} path." -f $folderIndex)
                    } else {
                        $state = "count"
                    }
                    continue
                }
                $p = Normalize-UserPath -raw $pRaw
                if ([string]::IsNullOrWhiteSpace($p)) {
                    Write-Host 'Please enter a valid path like Inbox\\My Team.'
                    continue
                }
                $parts = Split-InboxCombinedPath $p
                if ($parts.Count -gt 1) {
                    Write-Host 'It looks like multiple paths were pasted. Please enter only one path for this item.'
                    continue
                }
                $folderPathsLocal[$folderIndex - 1] = $p
                if ($folderIndex -lt $folderCount) {
                    $folderIndex += 1
                    continue
                }
                $folderPathsLocal = $folderPathsLocal | Where-Object { $_ -and $_.Trim() -ne '' }
                $state = "startDate"
                continue
            }
            "startDate" {
                $result = Read-DateValue -Label "Start date" -Initial $startDateLocal -AllowBack
                if ($result[0] -eq $Script:BackToken) {
                    if ($folderCount -gt 0) {
                        $folderIndex = $folderCount
                        $state = "folderPath"
                    } else {
                        $state = "count"
                    }
                    $startDateLocal = ""
                    continue
                }
                $startDateLocal = $result[0]
                $startDateObjLocal = $result[1]
                $state = "endDate"
                continue
            }
            "endDate" {
                $result = Read-DateValue -Label "End date" -Initial $endDateLocal -AllowBack
                if ($result[0] -eq $Script:BackToken) {
                    $startDateLocal = ""
                    $endDateLocal = ""
                    $state = "startDate"
                    continue
                }
                $endDateLocal = $result[0]
                $endDateObjLocal = $result[1]
                $state = "outputPst"
                continue
            }
            "outputPst" {
                if ([string]::IsNullOrWhiteSpace($outputPstLocal)) {
                    $outputPstLocal = $Script:DefaultPstPath
                    Write-Host ("Using default PST path: {0}" -f $outputPstLocal)
                }
                $state = "confirmDetails"
                continue
            }
            "confirmDetails" {
                Write-Host ("Selected start date: {0}" -f $startDateLocal)
                Write-Host ("Selected end date: {0}" -f $endDateLocal)
                Write-Host ("Selected PST path: {0}" -f $outputPstLocal)
                $confirm = Read-ScriptInput -Prompt "Press Enter to continue, or type back to change the end date" -AllowBlank -AllowBack
                if ($confirm -eq $Script:BackToken) {
                    $endDateLocal = ""
                    $state = "endDate"
                    continue
                }
                return [PSCustomObject]@{
                    FolderPaths       = @($folderPathsLocal)
                    StartDate         = $startDateLocal
                    StartDateObj      = $startDateObjLocal
                    EndDate           = $endDateLocal
                    EndDateObj        = $endDateObjLocal
                    OutputPstPath     = $outputPstLocal
                    DefaultMailbox    = $defaultMailboxName
                    DefaultInbox      = $defaultInboxFolder
                }
            }
        }
    }
}

$folderPaths = @()
$datesCollected = $false
if (-not [string]::IsNullOrWhiteSpace($FolderPathsInput)) {
    # Split on comma/semicolon/newline, trim whitespace
    $rawPaths = ($FolderPathsInput -split '[,;\r\n]+' ) | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
    foreach ($raw in $rawPaths) {
        $clean = Normalize-UserPath -raw $raw
        if ([string]::IsNullOrWhiteSpace($clean)) { continue }
        $parts = Split-InboxCombinedPath $clean
        foreach ($part in $parts) {
            if (-not [string]::IsNullOrWhiteSpace($part)) {
                $folderPaths += ,$part
            }
        }
    }
} else {
    Write-Host 'Enter Outlook folder paths under Inbox.'
    $interactive = Prompt-InteractiveExportDetails -InitialStartDate $StartDate -InitialEndDate $EndDate -InitialOutputPstPath $OutputPstPath
    $folderPaths = @($interactive.FolderPaths)
    $StartDate = $interactive.StartDate
    $startDateObj = $interactive.StartDateObj
    $EndDate = $interactive.EndDate
    $endDateObj = $interactive.EndDateObj
    $OutputPstPath = $interactive.OutputPstPath
    if ($interactive.DefaultMailbox) { $defaultMailboxName = $interactive.DefaultMailbox }
    if ($interactive.DefaultInbox) { $defaultInboxFolder = $interactive.DefaultInbox }
    $datesCollected = $true
}

# Final normalize: split any combined Inbox paths
$finalList = New-Object System.Collections.Generic.List[string]
foreach ($fp in $folderPaths) {
    foreach ($part in (Split-InboxCombinedPath $fp)) {
        if (-not [string]::IsNullOrWhiteSpace($part)) {
            $finalList.Add($part)
        }
    }
}
$folderPaths = $finalList.ToArray()

if ($folderPaths.Count -eq 0) {
    $folderPaths = @("Inbox")
}

Write-Host "Parsed folder paths:"
foreach ($p in $folderPaths) {
    Write-Host (" - {0}" -f $p)
}
Write-Host ("Folder path count: {0}" -f $folderPaths.Count)

# Hard stop if any combined paths remain
$badPaths = @()
foreach ($p in $folderPaths) {
    $count = ([regex]::Matches($p, "Inbox\\", "IgnoreCase")).Count
    if ($count -gt 1) {
        $badPaths += $p
    }
}
if ($badPaths.Count -gt 0) {
    Write-Host "ERROR: Combined folder paths detected. Please re-enter them one by one."
    foreach ($bp in $badPaths) {
        Write-Host (" - {0}" -f $bp)
    }
    throw "Combined folder paths detected."
}

Write-Host "Folder list (indexed):"
for ($i = 0; $i -lt $folderPaths.Count; $i++) {
    Write-Host (" {0}. {1}" -f ($i + 1), $folderPaths[$i])
}

if (-not $datesCollected) {
    while ($true) {
        $result = Read-DateValue -Label "Start date" -Initial $StartDate
        if ($result[0] -eq $Script:BackToken) {
            Write-Host "Back is not available before Start date."
            $StartDate = ""
            continue
        }
        $StartDate = $result[0]
        $startDateObj = $result[1]

        $result = Read-DateValue -Label "End date" -Initial $EndDate -AllowBack
        if ($result[0] -eq $Script:BackToken) {
            $StartDate = ""
            $EndDate = ""
            continue
        }
        $EndDate = $result[0]
        $endDateObj = $result[1]
        break
    }
    if ([string]::IsNullOrWhiteSpace($OutputPstPath)) {
        $OutputPstPath = $Script:DefaultPstPath
        Write-Host ("Using default PST path: {0}" -f $OutputPstPath)
    }
}

if ([string]::IsNullOrWhiteSpace($OutputPstPath)) {
    throw "Output PST path is required."
}

# Auto-fix missing drive (e.g. \Support_Tracker\...)
if ($OutputPstPath -match '^[\\/]' -and $OutputPstPath -notmatch '^[A-Za-z]:') {
    $defaultDrive = [System.IO.Path]::GetPathRoot($Script:DefaultTrackerRoot).TrimEnd('\')
    $OutputPstPath = "$defaultDrive$OutputPstPath"
    Write-Log "warn" ("Drive letter missing; using {0}" -f $OutputPstPath)
}
$startDateObj = $startDateObj.Date
$endDateObj = $endDateObj.Date
if ($endDateObj -lt $startDateObj) {
    Write-Log "warn" ("End date {0} is earlier than start date {1}; swapping." -f $endDateObj.ToString("dd-MM-yyyy"), $startDateObj.ToString("dd-MM-yyyy"))
    $tmp = $startDateObj
    $startDateObj = $endDateObj
    $endDateObj = $tmp
}

$outDir = Split-Path -Parent $OutputPstPath
if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}
if ($outDir -ne $Script:DefaultPstDir) {
    Write-Log "warn" ("PST path is outside the team default: {0}" -f $outDir)
    Write-Log "hint" ("Team default PST folder: {0}" -f $Script:DefaultPstDir)
}

Remove-StaleManagedPstStores -ManagedDir $outDir -TargetPstPath $OutputPstPath
Remove-ManagedPstStoreByPath -TargetPstPath $OutputPstPath | Out-Null

Write-Section "PST"
Write-Log "info" ("Path: {0}" -f $OutputPstPath)
Write-Log "info" "Checking Outlook attachment..."
$destStore = Wait-ForStoreByPath -TargetPath $OutputPstPath -Retries 2 -RetryDelaySeconds 1
if ($destStore) {
    Write-Log "ok" "PST already attached; skipping AddStore."
} else {
    $addStoreAction = {
        if ($namespace.PSObject.Methods.Name -contains "AddStoreEx") {
            Write-Log "info" "Creating/attaching Unicode PST with AddStoreEx."
            $namespace.AddStoreEx($OutputPstPath, 2) | Out-Null
        } else {
            Write-Log "info" "Attaching PST with AddStore."
            $namespace.AddStore($OutputPstPath) | Out-Null
        }
    }
    try {
        & $addStoreAction
    } catch {
        $firstError = $_.Exception.Message
        throw ("AddStore failed: {0}`nManual action needed: detach and remove the stale export PST from Outlook Data Files, then run again.`nSafe target only: {1}`nNever remove anything containing @invenio-solutions.com." -f $firstError, $OutputPstPath)
    }
    $destStore = Wait-ForStoreByPath -TargetPath $OutputPstPath -Retries 10 -RetryDelaySeconds 2
}

if (-not $destStore) {
    throw "Unable to open PST store: $OutputPstPath"
}

$destRoot = $destStore.GetRootFolder()
$inbox = $namespace.GetDefaultFolder(6) # olFolderInbox

Write-Section "Export Plan"
Write-Log "info" ("Date range: {0} to {1}" -f $startDateObj.ToString('dd-MM-yyyy'), $endDateObj.ToString('dd-MM-yyyy'))
Write-Log "info" "Folders:"
if ($folderPaths.Count -eq 0) {
    Write-Host "  - (none)"
} else {
    $idx = 1
    foreach ($p in $folderPaths) {
        Write-Host ("  {0}. {1}" -f $idx, $p)
        $idx++
    }
}

function Get-OrCreate-Folder {
    param(
        [object]$root,
        [string]$relativePath
    )
    $current = $root
    $parts = $relativePath.Split("\") | Where-Object { $_ -ne "" }
    foreach ($part in $parts) {
        $next = $null
        foreach ($f in @($current.Folders)) {
            if ($f.Name -eq $part) {
                $next = $f
                break
            }
        }
        if (-not $next) {
            $next = $current.Folders.Add($part)
        }
        $current = $next
    }
    return $current
}

function Resolve-SourceFolder {
    param(
        [string]$path
    )
    $parts = @($path -split "\\" | Where-Object { $_ -ne "" })
    if ($parts.Length -eq 0) {
        return $null
    }

    $current = $null
    $startIdx = 0

    if ($parts[0].ToLower() -eq "inbox") {
        $relativeParts = @()
        if ($parts.Length -gt 1) {
            $relativeParts = $parts[1..($parts.Length - 1)]
        }
        if ($defaultInboxFolder) {
            $current = Resolve-ChildPath -root $defaultInboxFolder -parts $relativeParts
        }
        if (-not $current) {
            $current = Resolve-ChildPath -root $inbox -parts $relativeParts
        }
        if (-not $current) {
            $current = Resolve-RelativeInboxAcrossStores -parts $relativeParts
        }
        if (-not $current) {
            $leaf = if ($relativeParts.Length -gt 0) { $relativeParts[$relativeParts.Length - 1] } else { $null }
            if ($leaf) {
                $candidates = Find-FolderByLeafNameAcrossStores -leafName $leaf
                if ($candidates.Count -eq 1) {
                    return $candidates[0].Folder
                }
            }
            Debug-Resolve-FolderNotFound -path $path -relativeParts $relativeParts
            return $null
        }
        return $current
    } else {
        $storeRoot = Get-StoreRootFolder -storeName $parts[0]
        if (-not $storeRoot) {
            Write-Host "Unknown mailbox/store or path must start with Inbox. Skipping: $path"
            return $null
        }
        if ($parts.Length -gt 1 -and $parts[1].ToLower() -eq "inbox") {
            $current = Get-StoreInboxFolder -storeName $parts[0]
            $startIdx = 2
        } else {
            $current = $storeRoot
            $startIdx = 1
        }
        if (-not $current) {
            Write-Host "Folder not found: $path"
            return $null
        }
        $remaining = @()
        if ($parts.Length -gt $startIdx) {
            $remaining = $parts[$startIdx..($parts.Length - 1)]
        }
        $resolved = Resolve-ChildPath -root $current -parts $remaining
        if (-not $resolved) {
            $rel = @()
            if ($parts.Length -gt $startIdx) { $rel = $parts[$startIdx..($parts.Length - 1)] }
            $leaf = if ($rel.Length -gt 0) { $rel[$rel.Length - 1] } else { $null }
            if ($leaf) {
                $candidates = Find-FolderByLeafNameAcrossStores -leafName $leaf
                if ($candidates.Count -eq 1) {
                    return $candidates[0].Folder
                }
            }
            Debug-Resolve-FolderNotFound -path $path -relativeParts $rel
            return $null
        }
        return $resolved
    }
}

# Validate folders before export (FOUND / NOT FOUND)
Write-Section "Folder Validation"
$validated = @()
foreach ($path in $folderPaths) {
    $resolved = Resolve-SourceFolder -path $path
    if ($resolved) {
        Write-Log "ok" $path
        $validated += $path
    } else {
        Write-Log "fail" $path
    }
}
$folderPaths = $validated

function Get-Item-Date {
    param([object]$item)
    $rt = $null
    $st = $null
    $ct = $null
    try { $rt = $item.ReceivedTime } catch { $rt = $null }
    try { $st = $item.SentOn } catch { $st = $null }
    if ($rt -and $rt -ne [DateTime]::MinValue) { return @($rt, $true) }
    if ($st -and $st -ne [DateTime]::MinValue) { return @($st, $false) }
    if ($includeCreationTime) {
        try { $ct = $item.CreationTime } catch { $ct = $null }
        if ($ct -and $ct -ne [DateTime]::MinValue) { return @($ct, $false) }
    }
    return @($null, $false)
}

function Test-DateInWindow {
    param(
        [object]$Value,
        [DateTime]$WindowStart,
        [DateTime]$WindowEnd
    )
    if (-not $Value) { return $false }
    try {
        $dt = [DateTime]$Value
    } catch {
        return $false
    }
    if ($dt -eq [DateTime]::MinValue) { return $false }
    return ($dt -ge $WindowStart -and $dt -le $WindowEnd)
}

function Get-Item-DateInWindow {
    param(
        [object]$item,
        [DateTime]$WindowStart,
        [DateTime]$WindowEnd
    )

    $rt = $null
    $st = $null
    $ct = $null
    try { $rt = $item.ReceivedTime } catch { $rt = $null }
    if (Test-DateInWindow -Value $rt -WindowStart $WindowStart -WindowEnd $WindowEnd) { return @($rt, "ReceivedTime") }
    try { $st = $item.SentOn } catch { $st = $null }
    if (Test-DateInWindow -Value $st -WindowStart $WindowStart -WindowEnd $WindowEnd) { return @($st, "SentOn") }
    if ($includeCreationTime) {
        try { $ct = $item.CreationTime } catch { $ct = $null }
        if (Test-DateInWindow -Value $ct -WindowStart $WindowStart -WindowEnd $WindowEnd) { return @($ct, "CreationTime") }
    }
    return @($null, "")
}

function Export-FolderWindow {
    param(
        [object]$sourceFolder,
        [string]$relativePath,
        [DateTime]$windowStart,
        [DateTime]$windowEnd,
        [string]$windowLabel = ""
    )

    $destFolder = Get-OrCreate-Folder -root $destRoot -relativePath $relativePath
    $count = 0
    $processed = 0
    $copyErrors = 0
    $sampleErrors = New-Object System.Collections.Generic.List[string]
    $progressEvery = 100
    $folderStart = Get-Date

    $items = $sourceFolder.Items
    $startSql = $windowStart.ToString("yyyy-MM-dd HH:mm")
    $endSql = $windowEnd.ToString("yyyy-MM-dd HH:mm")
    if ($fastExport) {
        $sql = "@SQL=""urn:schemas:httpmail:datereceived"" >= '$startSql' AND ""urn:schemas:httpmail:datereceived"" <= '$endSql'"
    } else {
        $conds = @(
            """urn:schemas:httpmail:datereceived"" >= '$startSql' AND ""urn:schemas:httpmail:datereceived"" <= '$endSql'",
            """urn:schemas:httpmail:datesent"" >= '$startSql' AND ""urn:schemas:httpmail:datesent"" <= '$endSql'"
        )
        if ($includeCreationTime) {
            $conds += """urn:schemas:httpmail:datecreated"" >= '$startSql' AND ""urn:schemas:httpmail:datecreated"" <= '$endSql'"
        }
        $sql = "@SQL=(" + ($conds -join " OR ") + ")"
    }

    $filtered = $null
    $usedFilterMode = "manual"
    try {
        $filtered = $items.Restrict($sql)
    } catch {
        $filtered = $null
    }

    $forceManual = $false
    if ($filtered -ne $null) {
        $usedFilterMode = "sql"
        Write-Log "ok" "SQL Restrict filter ready."
        $filteredCount = 0
        try {
            $filteredCount = $filtered.Count
            Write-Log "info" ("Filtered items: {0}" -f $filteredCount)
            if ($filteredCount -eq 0) {
                Write-Log "info" "Filtered items: 0, no items in range."
                if ($includeCreationTime) {
                    $forceManual = $true
                    Write-Log "warn" "Falling back to manual scan for CreationTime items."
                }
            }
        } catch {
            Write-Log "info" "Filtered items: unknown"
        }
        if (-not $forceManual) {
            try {
                $filtered.Sort("[ReceivedTime]", $true)
                Write-Log "ok" "Sorted by ReceivedTime descending."
            } catch {
                Write-Log "warn" "Unable to sort filtered items; strict date window still applies."
            }
            for ($itemIndex = 1; $itemIndex -le $filteredCount; $itemIndex++) {
                $item = $null
                $copy = $null
                $moved = $null
                try {
                    $item = $filtered.Item($itemIndex)
                    if (-not $item) { continue }
                    if ($fastExport) {
                        if ($item.Class -ne 43) { continue }
                        $rt = $item.ReceivedTime
                        if (-not (Test-DateInWindow -Value $rt -WindowStart $windowStart -WindowEnd $windowEnd)) { continue }
                    } else {
                        $dtInfo = Get-Item-DateInWindow -item $item -WindowStart $windowStart -WindowEnd $windowEnd
                        $rt = $dtInfo[0]
                        if (-not $rt) { continue }
                    }
                    $copy = $item.Copy()
                    $moved = $copy.Move($destFolder)
                    $count++
                    $processed++
                    if (($processed % $progressEvery) -eq 0) {
                        Write-Log "progress" ("processed={0}; exported={1}; current={2}" -f $processed, $count, $rt)
                    }
                } catch {
                    $copyErrors++
                    if ($sampleErrors.Count -lt 5) {
                        $sampleErrors.Add($_.Exception.Message) | Out-Null
                    }
                    continue
                } finally {
                    Release-ComObjectQuietly $moved
                    Release-ComObjectQuietly $copy
                    Release-ComObjectQuietly $item
                }
            }
        }
    } else {
        # Sort descending and break once received-time scans pass the start date.
        $manualSorted = $false
        try {
            $items.Sort("[ReceivedTime]", $true)
            $manualSorted = $true
        } catch {
            Write-Log "warn" "Unable to sort source items; manual scan continues without early date ordering."
        }
        Write-Log "warn" "SQL Restrict failed; using manual scan."
        Write-Log "info" "Manual scan: descending ReceivedTime with early break."
        $itemCount = 0
        try {
            $itemCount = $items.Count
        } catch {
            $itemCount = 0
        }
        for ($itemIndex = 1; $itemIndex -le $itemCount; $itemIndex++) {
            $item = $null
            $copy = $null
            $moved = $null
            try {
                $item = $items.Item($itemIndex)
                if (-not $item) { continue }
                if ($fastExport) {
                    if ($item.Class -ne 43) { continue }
                    $rt = $item.ReceivedTime
                    if (-not $rt) { continue }
                    if ($rt -lt $windowStart) {
                        if ($manualSorted) { break }
                        continue
                    }
                    if ($rt -gt $windowEnd) { continue }
                } else {
                    $dtInfo = Get-Item-DateInWindow -item $item -WindowStart $windowStart -WindowEnd $windowEnd
                    $rt = $dtInfo[0]
                    if (-not $rt) { continue }
                }
                $copy = $item.Copy()
                $moved = $copy.Move($destFolder)
                $count++
                $processed++
                if (($processed % $progressEvery) -eq 0) {
                    Write-Log "progress" ("processed={0}; exported={1}; current={2}" -f $processed, $count, $rt)
                }
            } catch {
                $copyErrors++
                if ($sampleErrors.Count -lt 5) {
                    $sampleErrors.Add($_.Exception.Message) | Out-Null
                }
                continue
            } finally {
                Release-ComObjectQuietly $moved
                Release-ComObjectQuietly $copy
                Release-ComObjectQuietly $item
            }
        }
    }

    $elapsed = (Get-Date) - $folderStart
    if ([string]::IsNullOrWhiteSpace($windowLabel)) {
        $windowLabel = "{0} -> {1}" -f $windowStart.ToString("dd-MM-yyyy HH:mm"), $windowEnd.ToString("dd-MM-yyyy HH:mm")
    }
    Write-Log "ok" ("{0} | exported={1}; window={2}; time={3}" -f $relativePath, $count, $windowLabel, $elapsed.ToString())
    if ($copyErrors -gt 0) {
        Write-Log "warn" ("{0} copy error(s) in {1} for {2}." -f $copyErrors, $relativePath, $windowLabel)
        foreach ($sample in $sampleErrors) {
            Write-Host ("  - {0}" -f $sample)
        }
    }

    # No subfolder export (explicit folders only)
    return [PSCustomObject]@{
        RelativePath = $relativePath
        WindowStart  = $windowStart
        WindowEnd    = $windowEnd
        WindowLabel  = $windowLabel
        Exported     = $count
        Processed    = $processed
        CopyErrors   = $copyErrors
        FilterMode   = $usedFilterMode
    }
}

function Invoke-FolderExportWithRetry {
    param(
        [object]$sourceFolder,
        [string]$relativePath,
        [DateTime]$rangeStart,
        [DateTime]$rangeEnd
    )

    $folderResults = New-Object System.Collections.Generic.List[object]
    $failedDays = New-Object System.Collections.Generic.List[DateTime]
    $rangeLabel = "{0} to {1}" -f $rangeStart.ToString("dd-MM-yyyy"), $rangeEnd.ToString("dd-MM-yyyy")
    $fullRangeStart = $rangeStart.Date
    $fullRangeEnd = $rangeEnd.Date.AddDays(1).AddSeconds(-1)

    Write-Section ("Folder: {0}" -f $relativePath)
    Write-Log "info" ("Window: {0}" -f $rangeLabel)
    try {
        $rangeResult = Export-FolderWindow -sourceFolder $sourceFolder -relativePath $relativePath -windowStart $fullRangeStart -windowEnd $fullRangeEnd -windowLabel $rangeLabel
        $folderResults.Add([PSCustomObject]@{
            Folder         = $relativePath
            Scope          = "range"
            WindowLabel    = $rangeLabel
            Success        = $true
            RetryAttempt   = 0
            Exported       = $rangeResult.Exported
            CopyErrors     = $rangeResult.CopyErrors
            FilterMode     = $rangeResult.FilterMode
            ErrorMessage   = $null
        }) | Out-Null
        Write-Log "ok" "Full-range export complete."
        return $folderResults
    } catch {
        $rangeError = $_.Exception.Message
        Write-Log "warn" ("Full-range export failed: {0}" -f $rangeError)
        Write-Log "info" "Falling back to day-by-day export for exact retry windows."
        $folderResults.Add([PSCustomObject]@{
            Folder         = $relativePath
            Scope          = "range"
            WindowLabel    = $rangeLabel
            Success        = $false
            RetryAttempt   = 0
            Exported       = 0
            CopyErrors     = 0
            FilterMode     = $null
            ErrorMessage   = $rangeError
        }) | Out-Null
    }

    for ($day = $rangeStart.Date; $day -le $rangeEnd.Date; $day = $day.AddDays(1)) {
        $dayStart = $day
        $dayEnd = $day.AddDays(1).AddSeconds(-1)
        $dayLabel = $day.ToString("dd-MM-yyyy")
        try {
            Write-Log "info" ("Trying day: {0}" -f $dayLabel)
            $dayResult = Export-FolderWindow -sourceFolder $sourceFolder -relativePath $relativePath -windowStart $dayStart -windowEnd $dayEnd -windowLabel $dayLabel
            $folderResults.Add([PSCustomObject]@{
                Folder         = $relativePath
                Scope          = "day"
                WindowLabel    = $dayLabel
                Success        = $true
                RetryAttempt   = 0
                Exported       = $dayResult.Exported
                CopyErrors     = $dayResult.CopyErrors
                FilterMode     = $dayResult.FilterMode
                ErrorMessage   = $null
            }) | Out-Null
            Write-Log "ok" ("Day complete: {0}" -f $dayLabel)
        } catch {
            $dayError = $_.Exception.Message
            $failedDays.Add($day) | Out-Null
            $folderResults.Add([PSCustomObject]@{
                Folder         = $relativePath
                Scope          = "day"
                WindowLabel    = $dayLabel
                Success        = $false
                RetryAttempt   = 0
                Exported       = 0
                CopyErrors     = 0
                FilterMode     = $null
                ErrorMessage   = $dayError
            }) | Out-Null
            Write-Log "fail" ("Day failed: {0}; {1}" -f $dayLabel, $dayError)
        }
    }

    if ($failedDays.Count -gt 0) {
        Write-Log "warn" ("Retrying failed day(s): {0}" -f (($failedDays | ForEach-Object { $_.ToString('dd-MM-yyyy') }) -join ", "))
    }
    foreach ($failedDay in $failedDays) {
        $retryStart = $failedDay.Date
        $retryEnd = $retryStart.AddDays(1).AddSeconds(-1)
        $retryLabel = $retryStart.ToString("dd-MM-yyyy")
        try {
            $retryResult = Export-FolderWindow -sourceFolder $sourceFolder -relativePath $relativePath -windowStart $retryStart -windowEnd $retryEnd -windowLabel ("Retry " + $retryLabel)
            $folderResults.Add([PSCustomObject]@{
                Folder         = $relativePath
                Scope          = "retry-day"
                WindowLabel    = $retryLabel
                Success        = $true
                RetryAttempt   = 1
                Exported       = $retryResult.Exported
                CopyErrors     = $retryResult.CopyErrors
                FilterMode     = $retryResult.FilterMode
                ErrorMessage   = $null
            }) | Out-Null
            Write-Log "ok" ("Retry complete: {0}" -f $retryLabel)
        } catch {
            $retryError = $_.Exception.Message
            $folderResults.Add([PSCustomObject]@{
                Folder         = $relativePath
                Scope          = "retry-day"
                WindowLabel    = $retryLabel
                Success        = $false
                RetryAttempt   = 1
                Exported       = 0
                CopyErrors     = 0
                FilterMode     = $null
                ErrorMessage   = $retryError
            }) | Out-Null
            Write-Log "fail" ("Retry failed: {0}; {1}" -f $retryLabel, $retryError)
        }
    }

    return $folderResults
}

$exportResult = $null
try {
    $allExportLogs = New-Object System.Collections.Generic.List[object]
    foreach ($path in $folderPaths) {
        $sourceFolder = Resolve-SourceFolder -path $path
        if (-not $sourceFolder) {
            Write-Log "fail" ("Folder could not be resolved: {0}" -f $path)
            $allExportLogs.Add([PSCustomObject]@{
                Folder         = $path
                Scope          = "folder"
                WindowLabel    = "n/a"
                Success        = $false
                RetryAttempt   = 0
                Exported       = 0
                CopyErrors     = 0
                FilterMode     = $null
                ErrorMessage   = "Folder could not be resolved."
            }) | Out-Null
            continue
        }

        foreach ($log in (Invoke-FolderExportWithRetry -sourceFolder $sourceFolder -relativePath $path -rangeStart $startDateObj -rangeEnd $endDateObj)) {
            $allExportLogs.Add($log) | Out-Null
        }
    }

    Write-Section "Export Summary"
    Write-Log "ok" "Export pass complete."
    Write-Log "info" ("PST: {0}" -f $OutputPstPath)
    foreach ($log in $allExportLogs) {
        $status = if ($log.Success -and $log.CopyErrors -gt 0) { "INCOMPLETE" } elseif ($log.Success) { "SUCCESS" } else { "FAILED" }
        $retryText = if ($log.RetryAttempt -gt 0) { "retry $($log.RetryAttempt)" } else { "initial" }
        $extra = if ($log.Success) {
            "exported=$($log.Exported); copyErrors=$($log.CopyErrors); mode=$($log.FilterMode)"
        } else {
            "error=$($log.ErrorMessage)"
        }
        $level = if ($status -eq "SUCCESS") { "ok" } elseif ($status -eq "INCOMPLETE") { "warn" } else { "fail" }
        Write-Log $level ("{0} | {1} | {2} | {3} | {4}" -f $log.Folder, $log.Scope, $log.WindowLabel, $retryText, $extra)
    }

    $totalCopyErrors = 0
    foreach ($log in $allExportLogs) {
        $totalCopyErrors += [int]$log.CopyErrors
    }
    if ($totalCopyErrors -gt 0) {
        throw ("Export incomplete: {0} item(s) could not be copied. The output PST is likely full or unable to accept more items. Use a fresh PST path, then rerun." -f $totalCopyErrors)
    }

    $exportResult = [PSCustomObject]@{
        OutputPstPath = $OutputPstPath
        StartDate     = $StartDate
        EndDate       = $EndDate
        FolderPaths   = @($folderPaths)
        Logs          = @($allExportLogs)
    }
} finally {
    try {
        $null = Remove-ManagedPstStoreByPath -TargetPstPath $OutputPstPath -Retries 4 -RetryDelaySeconds 3
        Release-ComObjectQuietly $destRoot
        Release-ComObjectQuietly $destStore
        $destRoot = $null
        $destStore = $null
    } catch {
        Write-Log "warn" ("Cleanup detach failed: {0}" -f ($_.Exception.Message))
    }
    try {
        if (-not $outlookWasRunning -and $closeOutlookAfterExport) {
            try {
                if ($outlook) {
                    $outlook.Quit()
                    Write-Log "ok" "Closed Outlook instance started by export script."
                }
            } catch {
            }
        } elseif (-not $outlookWasRunning) {
            Write-Log "info" "Leaving Outlook open after export cleanup to avoid shutdown hangs."
        }
    } finally {
        foreach ($comObj in @($inbox, $namespace, $outlook)) {
            try {
                if ($comObj) {
                    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObj)
                }
            } catch {
            }
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

return $exportResult
