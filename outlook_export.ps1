[CmdletBinding()]
param(
    [string]$OutputDir = 'C:\Users\wuttke\Documents\outlook-export',
    [int]$MaxItemsPerFolder = 0   # 0 = no limit; useful for first smoke test
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# ---- constants ----
$olMailItem = 0
$PR_HEADERS = 'http://schemas.microsoft.com/mapi/proptag/0x007D001E'
$SkipRoles = @{ 3='DeletedItems'; 9='Calendar'; 10='Contacts'; 11='Journal';
                12='Notes'; 13='Tasks'; 19='Conflicts'; 20='SyncIssues';
                21='LocalFailures'; 22='ServerFailures'; 23='Junk'; 28='ToDo' }
$SkipNames = @('Files','Yammer-Stamm','Verlauf der Unterhaltung','Conversation History',
    'Social Activity Notifications','Conversation Action Settings',
    'ExternalContacts','PersonMetadata','EventCheckPoints',
    'Einstellungen f' + [char]0xFC + 'r QuickSteps','Quick Step Settings')

if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir | Out-Null }
$StatePath = Join-Path $OutputDir '_sync_state.json'

# ---- load state ----
$state = @{}
if (Test-Path $StatePath) {
    $raw = Get-Content $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($p in $raw.folders.PSObject.Properties) {
        $seen = New-Object 'System.Collections.Generic.HashSet[string]'
        foreach ($k in $p.Value.seen_keys) { $null = $seen.Add($k) }
        $state[$p.Name] = @{
            mbox = $p.Value.mbox
            watermark = if ($p.Value.watermark) { [DateTime]$p.Value.watermark } else { [DateTime]'1900-01-01' }
            seen = $seen
            count = [int]$p.Value.count
        }
    }
}

# ---- connect to Outlook ----
$ol = New-Object -ComObject Outlook.Application
$ns = $ol.GetNamespace('MAPI')
$store = $ns.DefaultStore
$root = $store.GetRootFolder()
$skipIds = @{}
foreach ($r in $SkipRoles.Keys) {
    try { $skipIds[$ns.GetDefaultFolder($r).EntryID] = $SkipRoles[$r] } catch {}
}

# ---- helpers ----
function SanitizeFilename([string]$s) {
    $bad = [System.IO.Path]::GetInvalidFileNameChars() + [char[]]'/\:'
    foreach ($c in $bad) { $s = $s.Replace($c, '_') }
    $s
}
function MboxFromLine([DateTime]$d, [string]$from) {
    if (-not $from) { $from = 'MAILER-DAEMON' }
    $stamp = $d.ToUniversalTime().ToString('ddd MMM dd HH:mm:ss yyyy', [Globalization.CultureInfo]::InvariantCulture)
    "From $from $stamp"
}
function EscapeBody([string]$body) {
    if (-not $body) { return '' }
    # mboxrd: escape lines starting with ">*From " by prefixing with '>'
    ($body -split "`r?`n" | ForEach-Object {
        if ($_ -match '^>*From ') { '>' + $_ } else { $_ }
    }) -join "`n"
}
function BuildHeaders($item, $folderPath) {
    $sb = New-Object System.Text.StringBuilder
    $raw = $null
    try { $raw = $item.PropertyAccessor.GetProperty($PR_HEADERS) } catch {}
    if ($raw -and $raw.Trim()) {
        # Approach A: use original RFC822 headers, normalize line endings, strip trailing blank lines
        $raw = ($raw -replace "`r`n","`n").TrimEnd("`n")
        [void]$sb.Append($raw).Append("`n")
    } else {
        # Approach B: synthesize minimal headers
        $msgid = ''
        try { $msgid = [string]$item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F') } catch {}
        if (-not $msgid) { $msgid = '<' + $item.EntryID + '@outlook-export.local>' }
        $date = $null
        try { $date = $item.ReceivedTime } catch {}
        if (-not $date -or $date.Year -lt 1990) { try { $date = $item.SentOn } catch {} }
        if (-not $date -or $date.Year -lt 1990) { $date = $item.CreationTime }
        $dstr = $date.ToString('ddd, dd MMM yyyy HH:mm:ss zzz', [Globalization.CultureInfo]::InvariantCulture)
        [void]$sb.AppendLine("Date: $dstr")
        [void]$sb.AppendLine("From: $(($item.SenderName)) <$(($item.SenderEmailAddress))>")
        [void]$sb.AppendLine("To: $(($item.To))")
        if ($item.CC) { [void]$sb.AppendLine("Cc: $(($item.CC))") }
        [void]$sb.AppendLine("Subject: $(($item.Subject))")
        [void]$sb.AppendLine("Message-ID: $msgid")
        [void]$sb.AppendLine("MIME-Version: 1.0")
        [void]$sb.AppendLine("Content-Type: text/plain; charset=utf-8")
        [void]$sb.AppendLine("Content-Transfer-Encoding: 8bit")
    }
    # X-Outlook-* extras (always appended)
    [void]$sb.AppendLine("X-Outlook-EntryID: $($item.EntryID)")
    [void]$sb.AppendLine("X-Outlook-Folder: $folderPath")
    try { if ($item.UnRead) { [void]$sb.AppendLine("X-Outlook-Unread: true") } } catch {}
    try { if ($item.Categories) { [void]$sb.AppendLine("X-Outlook-Categories: $($item.Categories)") } } catch {}
    if ($item.Attachments.Count -gt 0) {
        $names = @()
        foreach ($a in $item.Attachments) { $names += $a.FileName }
        [void]$sb.AppendLine("X-Outlook-Attachments: " + ($names -join '; '))
    }
    $sb.ToString()
}

# ---- walk and export ----
$totalNew = 0
function Process($folder, $path, $parentSkip) {
    foreach ($sub in $folder.Folders) {
        $name = $sub.Name
        $sp = if ($path) { "$path/$name" } else { $name }
        $skip = $parentSkip -or $skipIds.ContainsKey($sub.EntryID) -or ($SkipNames -contains $name) -or ($sub.DefaultItemType -ne $olMailItem)
        Process $sub $sp $skip
        if ($skip) { continue }
        if ($sub.Items.Count -eq 0) { continue }
        ExportFolder $sub $sp
    }
}
function ExportFolder($folder, $path) {
    $key = $path
    if (-not $state.ContainsKey($key)) {
        $state[$key] = @{
            mbox = (SanitizeFilename ($path -replace '/','__')) + '.mbox'
            watermark = [DateTime]'1900-01-01'
            seen = New-Object 'System.Collections.Generic.HashSet[string]'
            count = 0
        }
    }
    $entry = $state[$key]
    $mboxFile = Join-Path $OutputDir $entry.mbox

    $items = $folder.Items
    $items.Sort('[ReceivedTime]', $false)
    $wm = $entry.watermark
    if ($wm -gt [DateTime]'1900-01-01') {
        $restrict = "[ReceivedTime] > '" + $wm.ToString("MM/dd/yyyy HH:mm tt", [Globalization.CultureInfo]::InvariantCulture) + "'"
        try { $items = $items.Restrict($restrict) } catch {}
    }
    $total = $items.Count
    if ($total -eq 0) { Write-Host ("  [{0,-45}] up-to-date" -f $path); return }
    Write-Host ("  [{0,-45}] {1} candidate(s) since {2:yyyy-MM-dd HH:mm}" -f $path, $total, $wm)

    $sw = New-Object System.IO.StreamWriter($mboxFile, $true, [System.Text.UTF8Encoding]::new($false))
    $sw.NewLine = "`n"
    $written = 0; $skippedDup = 0; $errors = 0
    for ($i = 1; $i -le $total; $i++) {
        try {
            $m = $items.Item($i)
        } catch { $errors++; continue }
        if ($MaxItemsPerFolder -gt 0 -and $written -ge $MaxItemsPerFolder) { break }
        # dedup key
        $msgid = $null
        try { $msgid = [string]$m.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F') } catch {}
        $dkey = if ($msgid) { $msgid } else { $m.EntryID }
        if ($entry.seen.Contains($dkey)) { $skippedDup++; continue }
        try {
            $rt = $null
            try { $rt = $m.ReceivedTime } catch {}
            if (-not $rt -or $rt.Year -lt 1990) { try { $rt = $m.SentOn } catch {} }
            if (-not $rt -or $rt.Year -lt 1990) { $rt = $m.CreationTime }
            $senderEmail = ''
            try { $senderEmail = [string]$m.SenderEmailAddress } catch {}
            $sw.WriteLine((MboxFromLine $rt $senderEmail))
            $sw.WriteLine((BuildHeaders $m $path).TrimEnd("`n"))
            $sw.WriteLine('')
            $body = ''
            try { $body = [string]$m.Body } catch {}
            $sw.WriteLine((EscapeBody $body))
            $sw.WriteLine('')
            [void]$entry.seen.Add($dkey)
            if ($rt -gt $entry.watermark) { $entry.watermark = $rt }
            $entry.count++
            $written++
            $script:totalNew++
        } catch { $errors++ }
        if ($written % 100 -eq 0 -and $written -gt 0) {
            Write-Host ("    ... {0}/{1}" -f $written, $total)
            $sw.Flush()
        }
    }
    $sw.Close()
    Write-Host ("    -> wrote {0}, dup {1}, err {2}, watermark {3:yyyy-MM-dd HH:mm}" -f $written, $skippedDup, $errors, $entry.watermark) -ForegroundColor Green
    SaveState
}
function SaveState {
    $out = @{ version = 1; exported_at = (Get-Date).ToString('o'); folders = @{} }
    foreach ($k in $state.Keys) {
        $out.folders[$k] = @{
            mbox = $state[$k].mbox
            watermark = $state[$k].watermark.ToString('o')
            count = $state[$k].count
            seen_keys = @($state[$k].seen)
        }
    }
    $out | ConvertTo-Json -Depth 6 | Set-Content -Path $StatePath -Encoding UTF8
}

Write-Host "Store : $($store.DisplayName)" -ForegroundColor Cyan
Write-Host "Out   : $OutputDir" -ForegroundColor Cyan
Write-Host ""
Process $root '' $false
SaveState
Write-Host ""
Write-Host ("Done. New items this run: {0}" -f $totalNew) -ForegroundColor Green
