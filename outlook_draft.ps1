# Create or send a draft mail in the live Outlook session.
#
# Always invoked with a single -ArgsFile <path> argument that points at a
# UTF-8 JSON file with shape:
#   {
#     "To": "a@x.com; b@y.com",   (optional)
#     "Cc": "...",                 (optional)
#     "Bcc": "...",                (optional)
#     "Subject": "...",            (optional)
#     "Body": "..." | null,        (plain text body)
#     "HtmlBody": "..." | null,    (HTML body; takes precedence over Body)
#     "InReplyToEntryID": "...",   (optional: create as reply to this mail)
#     "ReplyAll": true | false,    (when InReplyToEntryID set; default false)
#     "SendEntryID": "..."         (optional: send this existing draft;
#                                   all create-fields above are ignored)
#   }
#
# stdout: single-line JSON { entry_id, subject, draft_folder, ... } on create,
#         or { entry_id, sent: true, subject, to, ... } on send.
# stderr/exit-1 on error: JSON { error: "..." }

[CmdletBinding()]
param(
    [string]$ArgsFile
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ---- IL-drop (see outlook_export.ps1 for full explanation) ----
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
    if (-not $explorer) { @{ error = "No explorer.exe in session $sid" } | ConvertTo-Json -Compress; exit 1 }
    $h = [IlDrop]::OpenProcess(0x0400, $false, $explorer.Id)
    $srcTok = [IntPtr]::Zero
    [IlDrop]::OpenProcessToken($h, 0x000B, [ref]$srcTok) | Out-Null
    [IlDrop]::CloseHandle($h) | Out-Null
    $newTok = [IntPtr]::Zero
    [IlDrop]::DuplicateTokenEx($srcTok, 0x02000000, [IntPtr]::Zero, 2, 1, [ref]$newTok) | Out-Null
    [IlDrop]::CloseHandle($srcTok) | Out-Null
    $exe = (Get-Process -Id $PID).Path
    $scriptPath = $MyInvocation.MyCommand.Path
    $tmpOut = [IO.Path]::Combine($env:TEMP, "outlook-draft-$([Guid]::NewGuid().ToString('N')).out")
    $cmd = "`"$exe`" -NoProfile -ExecutionPolicy Bypass -Command `"& '$scriptPath' -ArgsFile '$ArgsFile' *>&1 | Out-File -Encoding utf8 -FilePath '$tmpOut'`""
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

function Emit($obj) { $obj | ConvertTo-Json -Depth 6 -Compress }
function Die($msg)  { Emit @{ error = $msg }; exit 1 }

if (-not $ArgsFile -or -not (Test-Path $ArgsFile)) { Die "missing or unreadable -ArgsFile" }
try {
    $a = Get-Content -Raw -Encoding UTF8 $ArgsFile | ConvertFrom-Json
} catch { Die "invalid ArgsFile JSON: $($_.Exception.Message)" }
# Caller is responsible for unlinking the args file (it lives in TEMP).

try {
    $outlook = New-Object -ComObject Outlook.Application
    $ns = $outlook.GetNamespace('MAPI')
} catch { Die "Cannot connect to Outlook COM: $($_.Exception.Message)" }

# --- Send path: an existing draft EntryID was supplied --------------------
$sendId = [string]$a.SendEntryID
if ($sendId) {
    try { $item = $ns.GetItemFromID($sendId) } catch { Die "GetItemFromID failed: $($_.Exception.Message)" }
    if (-not $item) { Die "draft not found for send" }
    # Sanity: must be a MailItem (olMail = 43)
    if ($item.Class -ne 43) { Die "item is not a MailItem (class=$($item.Class))" }
    $snapshotSubject = [string]$item.Subject
    $snapshotTo      = [string]$item.To
    $snapshotCc      = [string]$item.CC
    $snapshotBcc     = [string]$item.BCC
    if (-not $snapshotTo -and -not $snapshotCc -and -not $snapshotBcc) {
        Die "refusing to send: draft has no recipients"
    }
    try {
        $item.Send()
    } catch { Die "Send() failed: $($_.Exception.Message)" }
    Emit @{
        entry_id = $sendId
        sent     = $true
        subject  = $snapshotSubject
        to       = $snapshotTo
        cc       = $snapshotCc
        bcc      = $snapshotBcc
    }
    exit 0
}

# --- Create path ----------------------------------------------------------
$inReply = [string]$a.InReplyToEntryID
if ($inReply) {
    try { $orig = $ns.GetItemFromID($inReply) } catch { Die "GetItemFromID failed: $($_.Exception.Message)" }
    if (-not $orig) { Die "original mail not found for reply" }
    if ($a.ReplyAll) { $draft = $orig.ReplyAll() } else { $draft = $orig.Reply() }
} else {
    # olMailItem = 0
    $draft = $outlook.CreateItem(0)
}

if ($a.To)      { $draft.To      = [string]$a.To }
if ($a.Cc)      { $draft.CC      = [string]$a.Cc }
if ($a.Bcc)     { $draft.BCC     = [string]$a.Bcc }
if ($a.Subject) { $draft.Subject = [string]$a.Subject }

# HtmlBody wins over Body when both are present. For replies, prepending text
# to the quoted history is the user's job; we just set the body verbatim.
if ($a.HtmlBody) {
    $draft.BodyFormat = 2  # olFormatHTML
    $draft.HTMLBody   = [string]$a.HtmlBody
} elseif ($a.Body) {
    $draft.BodyFormat = 1  # olFormatPlain
    $draft.Body       = [string]$a.Body
}

try {
    $draft.Save()
} catch { Die "Save() failed: $($_.Exception.Message)" }

# Move into the default Drafts folder of the originating account when it
# isn't there already (replies normally start there; new items always do).
try {
    # olFolderDrafts = 16
    $drafts = $ns.GetDefaultFolder(16)
    $folderPath = $drafts.FolderPath
} catch { $folderPath = $null }

Emit @{
    entry_id     = $draft.EntryID
    subject      = $draft.Subject
    to           = $draft.To
    cc           = $draft.CC
    body_format  = $draft.BodyFormat
    draft_folder = $folderPath
    size_bytes   = $draft.Size
    in_reply_to  = $inReply
}
exit 0
