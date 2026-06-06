# Read attachments from a live Outlook mail by EntryID.
#
# Usage:
#   list:  outlook_attachment.ps1 -Action list -EntryID <id>
#          -> stdout: JSON { entry_id, subject, attachments: [{filename, size, type}] }
#   read:  outlook_attachment.ps1 -Action read -EntryID <id> -Filename <name> -OutFile <path>
#          -> writes binary to <path>, stdout: JSON { filename, size, type, out_file }
#
# Output is always a single JSON object on stdout. Errors -> JSON { error: "..." }
# and a non-zero exit code.

[CmdletBinding()]
param(
    [ValidateSet('list','read')] [string]$Action,
    [string]$EntryID,
    [string]$Filename,
    [string]$OutFile,
    [string]$ArgsFile   # internal: used by the IL-drop re-exec to ferry args safely
)

if ($ArgsFile -and (Test-Path $ArgsFile)) {
    $j = Get-Content -Raw -Encoding UTF8 $ArgsFile | ConvertFrom-Json
    if ($j.Action)   { $Action   = [string]$j.Action }
    if ($j.EntryID)  { $EntryID  = [string]$j.EntryID }
    if ($j.Filename) { $Filename = [string]$j.Filename }
    if ($j.OutFile)  { $OutFile  = [string]$j.OutFile }
    Remove-Item -Force $ArgsFile
}

$ErrorActionPreference = 'Stop'
# Force UTF-8 so non-ASCII filenames and subjects survive the WSL hop.
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ---- Drop integrity level from High to Medium if needed ----
# Outlook runs at Medium IL; an elevated parent (e.g. WSL launched at High IL)
# would hit CO_E_SERVER_EXEC_FAILURE (0x80080005) on COM activation. Mirrors
# the trick used in outlook_export.ps1.
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public class IlDrop {
    [DllImport("advapi32.dll", SetLastError=true)] public static extern bool OpenProcessToken(IntPtr h, uint da, out IntPtr tok);
    [DllImport("advapi32.dll", SetLastError=true)] public static extern bool GetTokenInformation(IntPtr tok, int tic, IntPtr ti, int tis, out int rl);
    [DllImport("advapi32.dll", SetLastError=true)] public static extern bool DuplicateTokenEx(IntPtr h, uint da, IntPtr sa, int il, int tt, out IntPtr nt);
    [DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr GetCurrentProcess();
    [DllImport("kernel32.dll", SetLastError=true)] public static extern IntPtr OpenProcess(uint da, bool inh, int pid);
    [DllImport("kernel32.dll", SetLastError=true)] public static extern bool CloseHandle(IntPtr h);
    [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool CreateProcessWithTokenW(IntPtr hToken, uint flags, string app, string cmd, uint cflags, IntPtr env, string cwd, ref STARTUPINFO si, out PROCESS_INFORMATION pi);
    [DllImport("kernel32.dll", SetLastError=true)] public static extern uint WaitForSingleObject(IntPtr h, int ms);
    [DllImport("kernel32.dll", SetLastError=true)] public static extern bool GetExitCodeProcess(IntPtr h, out uint c);
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    public struct STARTUPINFO {
        public int cb; public string lpReserved, lpDesktop, lpTitle;
        public uint dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
        public ushort wShowWindow, cbReserved2; public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION { public IntPtr hProcess, hThread; public uint dwProcessId, dwThreadId; }
}
"@

function Get-CurrentIL {
    $tok = [IntPtr]::Zero
    if (-not [IlDrop]::OpenProcessToken([IlDrop]::GetCurrentProcess(), 0x0008, [ref]$tok)) { return 'Unknown' }
    $rl = 0; [IlDrop]::GetTokenInformation($tok, 25, [IntPtr]::Zero, 0, [ref]$rl) | Out-Null
    $buf = [Runtime.InteropServices.Marshal]::AllocHGlobal($rl)
    [IlDrop]::GetTokenInformation($tok, 25, $buf, $rl, [ref]$rl) | Out-Null
    $sidPtr = [Runtime.InteropServices.Marshal]::ReadIntPtr($buf)
    $sc = [Runtime.InteropServices.Marshal]::ReadByte($sidPtr, 1)
    $rid = [Runtime.InteropServices.Marshal]::ReadInt32($sidPtr, 8 + 4*($sc - 1))
    [Runtime.InteropServices.Marshal]::FreeHGlobal($buf)
    [IlDrop]::CloseHandle($tok) | Out-Null
    switch ($rid) { 4096{'Low'} 8192{'Medium'} 8448{'MediumPlus'} 12288{'High'} 16384{'System'} default{"RID=$rid"} }
}

$currentIL = Get-CurrentIL
if ($currentIL -eq 'High' -or $currentIL -eq 'System') {
    $sid = (Get-Process -Id $PID).SessionId
    $explorer = Get-Process -Name explorer -ErrorAction SilentlyContinue |
        Where-Object { $_.SessionId -eq $sid } | Select-Object -First 1
    if (-not $explorer) {
        @{ error = "No explorer.exe in session $sid; cannot drop integrity." } | ConvertTo-Json -Compress
        exit 1
    }
    $h = [IlDrop]::OpenProcess(0x0400, $false, $explorer.Id)
    $srcTok = [IntPtr]::Zero
    [IlDrop]::OpenProcessToken($h, 0x000B, [ref]$srcTok) | Out-Null
    [IlDrop]::CloseHandle($h) | Out-Null
    $newTok = [IntPtr]::Zero
    [IlDrop]::DuplicateTokenEx($srcTok, 0x02000000, [IntPtr]::Zero, 2, 1, [ref]$newTok) | Out-Null
    [IlDrop]::CloseHandle($srcTok) | Out-Null
    $exe = (Get-Process -Id $PID).Path
    $scriptPath = $MyInvocation.MyCommand.Path
    # Serialize all bound params into a temp JSON file so the re-execed child
    # gets exact arg values regardless of quoting (filenames may contain
    # spaces, parens, single quotes, etc.).
    $argsFile = [IO.Path]::Combine($env:TEMP, "outlook-att-args-$([Guid]::NewGuid().ToString('N')).json")
    $payload = @{
        Action = $Action; EntryID = $EntryID
        Filename = $Filename; OutFile = $OutFile
    }
    ($payload | ConvertTo-Json -Compress) | Set-Content -Path $argsFile -Encoding UTF8
    $tmpOut = [IO.Path]::Combine($env:TEMP, "outlook-attach-$([Guid]::NewGuid().ToString('N')).out")
    $cmd = "`"$exe`" -NoProfile -ExecutionPolicy Bypass -Command `"& '$scriptPath' -ArgsFile '$argsFile' *>&1 | Out-File -Encoding utf8 -FilePath '$tmpOut'`""
    $si = New-Object IlDrop+STARTUPINFO
    $si.cb = [Runtime.InteropServices.Marshal]::SizeOf($si)
    $pi = New-Object IlDrop+PROCESS_INFORMATION
    $ok = [IlDrop]::CreateProcessWithTokenW($newTok, 0, $exe, $cmd, 0x08000000, [IntPtr]::Zero, 'C:\Windows\System32', [ref]$si, [ref]$pi)
    [IlDrop]::CloseHandle($newTok) | Out-Null
    if (-not $ok) {
        @{ error = "CreateProcessWithTokenW failed: $([Runtime.InteropServices.Marshal]::GetLastWin32Error())" } | ConvertTo-Json -Compress
        exit 1
    }
    [IlDrop]::WaitForSingleObject($pi.hProcess, -1) | Out-Null
    $code = 0
    [IlDrop]::GetExitCodeProcess($pi.hProcess, [ref]$code) | Out-Null
    [IlDrop]::CloseHandle($pi.hProcess) | Out-Null
    [IlDrop]::CloseHandle($pi.hThread) | Out-Null
    if (Test-Path $tmpOut) {
        Get-Content -Raw -Encoding UTF8 $tmpOut | Write-Host -NoNewline
        Remove-Item -Force $tmpOut
    }
    exit $code
}

function Emit($obj) {
    $obj | ConvertTo-Json -Depth 6 -Compress
}
function Die($msg) {
    Emit @{ error = $msg }
    exit 1
}

try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace('MAPI')
} catch {
    Die "Cannot connect to Outlook COM: $($_.Exception.Message)"
}

try {
    $item = $ns.GetItemFromID($EntryID)
} catch {
    Die "GetItemFromID failed for entry: $($_.Exception.Message)"
}
if (-not $item) { Die "Item not found" }

# Outlook attachment types: 1=ByValue (normal), 5=Embedded, 4=OLE, 6=Reference.
function TypeName($t) {
    switch ([int]$t) {
        1 { 'byvalue' }
        4 { 'ole' }
        5 { 'embedded' }
        6 { 'reference' }
        default { "type$t" }
    }
}

if ($Action -eq 'list') {
    $items = @()
    foreach ($a in $item.Attachments) {
        $size = $null
        try { $size = $a.Size } catch {}
        $items += @{
            filename = $a.FileName
            display_name = $a.DisplayName
            size = $size
            type = TypeName $a.Type
            index = $a.Index
        }
    }
    Emit @{
        entry_id = $EntryID
        subject = $item.Subject
        attachments = $items
    }
    exit 0
}

if ($Action -eq 'read') {
    if (-not $Filename) { Die "-Filename required for 'read'" }
    if (-not $OutFile) { Die "-OutFile required for 'read'" }
    $match = $null
    foreach ($a in $item.Attachments) {
        if ($a.FileName -eq $Filename) { $match = $a; break }
    }
    if (-not $match) {
        $avail = @()
        foreach ($a in $item.Attachments) { $avail += $a.FileName }
        Emit @{ error = "no attachment named '$Filename'"; available = $avail }
        exit 1
    }
    # Outlook can only save attachments to local paths; ensure parent exists.
    $dir = Split-Path -Parent $OutFile
    if ($dir -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    try {
        $match.SaveAsFile($OutFile)
    } catch {
        Die "SaveAsFile failed: $($_.Exception.Message)"
    }
    $size = $null
    try { $size = (Get-Item $OutFile).Length } catch {}
    Emit @{
        filename = $match.FileName
        display_name = $match.DisplayName
        size = $size
        type = TypeName $match.Type
        out_file = $OutFile
    }
    exit 0
}
