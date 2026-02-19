param(
    [string]$FolderPathsInput,
    [string]$StartDate,
    [string]$EndDate,
    [string]$OutputPstPath
)

$ErrorActionPreference = "Stop"

$includeCreationTime = $true
if ($env:EXCLUDE_CREATION_TIME) {
    $val = $env:EXCLUDE_CREATION_TIME.ToString().ToLower()
    if ($val -in @("1", "true", "yes", "y", "on")) { $includeCreationTime = $false }
}

Write-Host "Outlook export (PST append) - copy only, no move"
Write-Host ("Start time: {0}" -f (Get-Date))
Write-Host "Export script version: 2026-02-05-1"
Write-Host ""

Write-Host "Connecting to Outlook..."
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

$fastExport = $true
if ($env:FAST_EXPORT) {
    $val = $env:FAST_EXPORT.ToString().ToLower()
    if ($val -in @("0", "false", "no", "n", "off")) {
        $fastExport = $false
    } elseif ($val -in @("1", "true", "yes", "y", "on")) {
        $fastExport = $true
    }
}
Write-Host ("Export mode: {0}" -f ($(if ($fastExport) { "FAST (MailItem + ReceivedTime)" } else { "COMPLETE (All items + SentOn fallback)" })))

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

function Read-DateValue {
    param(
        [string]$Label,
        [string]$Initial
    )
    $current = $Initial
    while ($true) {
        if ([string]::IsNullOrWhiteSpace($current)) {
            $current = Read-Host "$Label (DD-MM-YYYY)"
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

# Build store lookup (display name + root name) for shared mailboxes
$storeInfos = @()
foreach ($s in $namespace.Stores) {
    $root = $null
    try { $root = $s.GetRootFolder() } catch { $root = $null }
    $storeInfos += [PSCustomObject]@{
        Store       = $s
        DisplayName = $s.DisplayName
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

function Debug-Print-InboxChildren {
    param(
        [string]$storeName,
        [object]$inboxFolder
    )
    if (-not $inboxFolder) { return }
    $limit = 50
    $names = @()
    foreach ($f in @($inboxFolder.Folders)) {
        if ($names.Count -ge $limit) { break }
        $names += $f.Name
    }
    Write-Host ("[DEBUG] Inbox children for {0} (showing up to {1}): {2}" -f $storeName, $limit, ($names -join ", "))
}

function Debug-Resolve-FolderNotFound {
    param(
        [string]$path,
        [string[]]$relativeParts
    )
    Write-Host ("[DEBUG] Folder NOT FOUND: {0}" -f $path)
    Write-Host ("[DEBUG] Relative parts: {0}" -f ($relativeParts -join " | "))
    Write-Host "[DEBUG] Stores available:"
    foreach ($info in $storeInfos) {
        $name = if ($info.DisplayName) { $info.DisplayName } else { $info.RootName }
        Write-Host ("[DEBUG]  - {0}" -f $name)
        $inbox = $null
        if ($info.DisplayName) { $inbox = Get-StoreInboxFolder -storeName $info.DisplayName }
        if (-not $inbox -and $info.RootName) { $inbox = Get-StoreInboxFolder -storeName $info.RootName }
        Debug-Print-InboxChildren -storeName $name -inboxFolder $inbox
    }
}

function Enumerate-Folders {
    param(
        [object]$root,
        [string]$prefix
    )
    foreach ($f in @($root.Folders)) {
        $path = if ([string]::IsNullOrWhiteSpace($prefix)) { $f.Name } else { "$prefix\\$($f.Name)" }
        [PSCustomObject]@{ Folder = $f; Path = $path }
        foreach ($child in (Enumerate-Folders -root $f -prefix $path)) {
            $child
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
    if ($depth -ge $maxDepth) { return }
    foreach ($f in @($root.Folders)) {
        $path = if ([string]::IsNullOrWhiteSpace($prefix)) { $f.Name } else { "$prefix\\$($f.Name)" }
        Write-Host ("{0}{1}" -f (" " * ($depth * 2)), $path)
        Print-FolderTree -root $f -prefix $path -depth ($depth + 1) -maxDepth $maxDepth
    }
}

# Optional: list folder paths to help users pick correct paths (before asking for inputs)
$listChoice = Read-Host "List folder paths for a mailbox? (y/N)"
if ($listChoice -match '^(y|yes)$') {
    $mb = Read-Host "Mailbox/store name (exact as listed above)"
    $storeInfo = if (-not [string]::IsNullOrWhiteSpace($mb)) { Get-StoreInfo -storeName $mb } else { $null }
    if ($storeInfo -and $storeInfo.RootFolder) {
        $defaultMailboxName = $mb
        $depthStr = Read-Host "Max depth to list (default 4)"
        if (-not ($depthStr -match '^\d+$')) { $depthStr = '4' }
        $maxDepth = [int]$depthStr
        $inboxOnly = Read-Host "Start at Inbox only? (y/N)"
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
            Write-Host ("Folder paths under: {0}" -f $mb)
            Print-FolderTree -root $storeInfo.RootFolder -prefix $mb -depth 0 -maxDepth $maxDepth
        }
    } else {
        Write-Host "Store not found. Skipping folder listing."
    }
}

$folderPaths = @()
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
    $countStr = Read-Host 'How many folders to export?'
    while (-not ($countStr -match '^\d+$') -or [int]$countStr -le 0) {
        Write-Host 'Please enter a valid positive number.'
        $countStr = Read-Host 'How many folders to export?'
    }
    $count = [int]$countStr
    $folderPaths = New-Object string[] $count
    for ($i = 1; $i -le $count; $i++) {
        while ($true) {
            $pRaw = Read-Host ("Folder $i path (under Inbox)")
            if ([string]::IsNullOrWhiteSpace($pRaw)) {
                Write-Host 'Path cannot be empty.'
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
            $folderPaths[$i-1] = $p
            Write-Host ("[DEBUG] Added path: {0}" -f $p)
            break
        }
    }
    # Remove any empty slots just in case
    $folderPaths = $folderPaths | Where-Object { $_ -and $_.Trim() -ne '' }
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

$result = Read-DateValue -Label "Start date" -Initial $StartDate
$StartDate = $result[0]
$startDateObj = $result[1]

$result = Read-DateValue -Label "End date" -Initial $EndDate
$EndDate = $result[0]
$endDateObj = $result[1]

if ([string]::IsNullOrWhiteSpace($OutputPstPath)) {
    $defaultPst = "D:\\Support_Tracker\\PstFiles\\export_filtered.pst"
    $OutputPstPath = Read-Host ("Output PST path (default: {0})" -f $defaultPst)
    if ([string]::IsNullOrWhiteSpace($OutputPstPath)) {
        $OutputPstPath = $defaultPst
    }
}

if ([string]::IsNullOrWhiteSpace($OutputPstPath)) {
    throw "Output PST path is required."
}

# Auto-fix missing drive (e.g. \Support_Tracker\...)
if ($OutputPstPath -match '^[\\/]' -and $OutputPstPath -notmatch '^[A-Za-z]:') {
    $OutputPstPath = "D:$OutputPstPath"
    Write-Host ("WARNING: Drive letter missing; using {0}" -f $OutputPstPath)
}
$startDateObj = $startDateObj.Date
$endDateObj = $endDateObj.Date
if ($endDateObj -lt $startDateObj) {
    Write-Host ("WARNING: End date {0} is earlier than start date {1}. Swapping." -f $endDateObj.ToString("dd-MM-yyyy"), $startDateObj.ToString("dd-MM-yyyy"))
    $tmp = $startDateObj
    $startDateObj = $endDateObj
    $endDateObj = $tmp
}

$outDir = Split-Path -Parent $OutputPstPath
if ($outDir -and -not (Test-Path $outDir)) {
    New-Item -ItemType Directory -Path $outDir | Out-Null
}
if ($outDir -ne "D:\Support_Tracker\PstFiles") {
    Write-Host ("WARNING: PST path is not in D:\Support_Tracker\PstFiles -> {0}" -f $outDir)
    Write-Host "For team consistency, please use: D:\\Support_Tracker\\PstFiles"
}

Write-Host "Opening/creating PST at: $OutputPstPath"
Write-Host "Checking if PST is already attached..."
$destStore = $namespace.Stores | Where-Object { $_.FilePath -and $_.FilePath.ToLower() -eq $OutputPstPath.ToLower() }
if ($destStore) {
    Write-Host "PST already attached. Skipping AddStore."
} else {
    try {
        if ($namespace.PSObject.Methods.Name -contains "AddStoreEx") {
            Write-Host "Using AddStoreEx (Unicode PST)..."
            $namespace.AddStoreEx($OutputPstPath, 3) | Out-Null
        } else {
            Write-Host "Using AddStore..."
            $namespace.AddStore($OutputPstPath) | Out-Null
        }
    } catch {
        throw "AddStore failed: $($_.Exception.Message)"
    }
}

$destStore = $namespace.Stores | Where-Object { $_.FilePath -eq $OutputPstPath }
if (-not $destStore) {
    $destStore = $namespace.Stores | Where-Object { $_.FilePath.ToLower() -eq $OutputPstPath.ToLower() }
}
if (-not $destStore) {
    throw "Unable to open PST store: $OutputPstPath"
}

$destRoot = $destStore.GetRootFolder()
$inbox = $namespace.GetDefaultFolder(6) # olFolderInbox

Write-Host "Date filter (daily chunks): $($startDateObj.ToString('dd-MM-yyyy')) to $($endDateObj.ToString('dd-MM-yyyy'))"
Write-Host "Folders to export:"
if ($folderPaths.Count -eq 0) {
    Write-Host " - (none)"
} else {
    $idx = 1
    foreach ($p in $folderPaths) {
        Write-Host (" {0}. {1}" -f $idx, $p)
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
                    Write-Host ("[DEBUG] Fallback matched leaf '{0}' at: {1}" -f $leaf, $candidates[0].Path)
                    return $candidates[0].Folder
                } elseif ($candidates.Count -gt 1) {
                    Write-Host ("[DEBUG] Multiple fallback matches for leaf '{0}':" -f $leaf)
                    foreach ($c in $candidates) {
                        Write-Host ("[DEBUG]  - {0}" -f $c.Path)
                    }
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
                    Write-Host ("[DEBUG] Fallback matched leaf '{0}' at: {1}" -f $leaf, $candidates[0].Path)
                    return $candidates[0].Folder
                } elseif ($candidates.Count -gt 1) {
                    Write-Host ("[DEBUG] Multiple fallback matches for leaf '{0}':" -f $leaf)
                    foreach ($c in $candidates) {
                        Write-Host ("[DEBUG]  - {0}" -f $c.Path)
                    }
                }
            }
            Debug-Resolve-FolderNotFound -path $path -relativeParts $rel
            return $null
        }
        return $resolved
    }
}

# Validate folders before export (FOUND / NOT FOUND)
Write-Host "Validating folder paths..."
$validated = @()
foreach ($path in $folderPaths) {
    $resolved = Resolve-SourceFolder -path $path
    if ($resolved) {
        Write-Host ("FOUND: {0}" -f $path)
        $validated += $path
    } else {
        Write-Host ("NOT FOUND: {0}" -f $path)
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

function Export-Folder {
    param(
        [object]$sourceFolder,
        [string]$relativePath,
        [DateTime]$dayStart,
        [DateTime]$dayEnd
    )

    $destFolder = Get-OrCreate-Folder -root $destRoot -relativePath $relativePath
    $count = 0
    $processed = 0
    $progressEvery = 100
    $folderStart = Get-Date

    $items = $sourceFolder.Items
    $startSql = $dayStart.ToString("yyyy-MM-dd HH:mm")
    $endSql = $dayEnd.ToString("yyyy-MM-dd HH:mm")
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
    try {
        $filtered = $items.Restrict($sql)
    } catch {
        $filtered = $null
    }

    $forceManual = $false
    if ($filtered -ne $null) {
        Write-Host "SQL Restrict succeeded."
        Write-Host "Using SQL Restrict filter..."
        try {
            $filteredCount = $filtered.Count
            Write-Host ("  Filtered items count: {0}" -f $filteredCount)
            if ($filteredCount -eq 0) {
                Write-Host "  Filtered count is 0 (no items in range)."
                if ($includeCreationTime) {
                    $forceManual = $true
                    Write-Host "  Falling back to manual scan for CreationTime items..."
                }
            }
        } catch {
            Write-Host "  Filtered items count: unknown"
        }
        if (-not $forceManual) {
            foreach ($item in @($filtered)) {
                try {
                    if (-not $item) { continue }
                    if ($fastExport) {
                        if ($item.Class -ne 43) { continue }
                        $rt = $item.ReceivedTime
                        if (-not $rt) { continue }
                    } else {
                        $dtInfo = Get-Item-Date $item
                        $rt = $dtInfo[0]
                        if (-not $rt) { continue }
                    }
                    $copy = $item.Copy()
                    $copy.Move($destFolder) | Out-Null
                    $count++
                    $processed++
                    if (($processed % $progressEvery) -eq 0) {
                        Write-Host ("  Processed {0} items | Exported {1} | Current Time: {2}" -f $processed, $count, $rt)
                    }
                } catch {
                    continue
                }
            }
        }
    } else {
        # Fast mode: sort descending and break once we pass the start date
        $items.Sort("[ReceivedTime]", $true)
        Write-Host "SQL Restrict failed (falling back to fast manual scan)..."
        Write-Host "Using fast manual scan (descending, early break)..."
        foreach ($item in @($items)) {
            try {
                if (-not $item) { continue }
                if ($fastExport) {
                    if ($item.Class -ne 43) { continue }
                    $rt = $item.ReceivedTime
                    if (-not $rt) { continue }
                    if ($rt -lt $dayStart) { break }
                    if ($rt -gt $dayEnd) { continue }
                } else {
                    $dtInfo = Get-Item-Date $item
                    $rt = $dtInfo[0]
                    $usedReceived = $dtInfo[1]
                    if (-not $rt) { continue }
                    if ($usedReceived -and $rt -lt $dayStart) { break }
                    if ($rt -gt $dayEnd) { continue }
                }
                $copy = $item.Copy()
                $copy.Move($destFolder) | Out-Null
                $count++
                $processed++
                if (($processed % $progressEvery) -eq 0) {
                    Write-Host ("  Processed {0} items | Exported {1} | Current Time: {2}" -f $processed, $count, $rt)
                }
            } catch {
                continue
            }
        }
    }

    $elapsed = (Get-Date) - $folderStart
    Write-Host ("Exported {0} items from {1} | Time: {2}" -f $count, $relativePath, $elapsed.ToString())

    # No subfolder export (explicit folders only)
}

for ($day = $startDateObj; $day -le $endDateObj; $day = $day.AddDays(1)) {
    $dayStart = $day
    $dayEnd = $day.AddDays(1).AddSeconds(-1)
    Write-Host ("--- Exporting day: {0} ---" -f $day.ToString("dd-MM-yyyy"))
    foreach ($path in $folderPaths) {
        $sourceFolder = Resolve-SourceFolder -path $path
        if (-not $sourceFolder) {
            continue
        }
        Write-Host "Exporting folder: $path"
        Export-Folder -sourceFolder $sourceFolder -relativePath $path -dayStart $dayStart -dayEnd $dayEnd
    }
}

Write-Host "Export completed."
Write-Host "PST written to: $OutputPstPath"

try {
    $namespace.RemoveStore($destRoot)
} catch {
    # ignore
}
