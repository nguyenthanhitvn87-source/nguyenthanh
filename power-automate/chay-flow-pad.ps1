# ==========================================================
# Launcher: chay mot flow Power Automate Desktop
# ==========================================================
# File nay do Scheduled Task goi moi ngay (chay tay cung duoc de test).
#
# Vi sao tach ra file rieng thay vi tro thang vao PAD.Console.Host.exe:
#   Ban PAD cai tu Microsoft Store nam trong C:\Program Files\WindowsApps\
#   Microsoft.PowerAutomateDesktop_<VERSION>_x64__8wekyb3d8bbwe\
#   -> moi lan PAD tu cap nhat thi <VERSION> doi, path cung se chet lang le.
#   File nay giai quyet lai duong dan MOI LAN CHAY.
# ==========================================================

param(
    [string]$TenFlow = "run job",

    # Tham so dau vao cho flow, dang JSON. Vi du:
    #   -InputArguments '{"NgayChay":"2026-08-11","SoLan":3}'
    [string]$InputArguments = "",

    # Chi kiem tra moi truong, KHONG chay flow
    [switch]$ChiKiemTra
)

$ErrorActionPreference = "Stop"

# ---------------- GHI LOG ----------------
$LogDir = Join-Path $env:LOCALAPPDATA "PAD-Scheduler"
if (-not (Test-Path -LiteralPath $LogDir)) {
    New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
}
$LogFile = Join-Path $LogDir "chay-flow.log"

function Write-Log {
    param([string]$Message, [string]$Muc = "INFO")
    $dong = "{0}  [{1}]  {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Muc, $Message
    Write-Host $dong
    try { Add-Content -LiteralPath $LogFile -Value $dong -Encoding UTF8 } catch { }
}

# ---------------- DO DUONG DAN PAD (du phong) ----------------
function Get-PADConsolePath {
    # Cach 1: ban cai tu Microsoft Store (MSIX)
    try {
        $pkg = Get-AppxPackage -Name "Microsoft.PowerAutomateDesktop*" -ErrorAction SilentlyContinue |
               Sort-Object -Property Version -Descending |
               Select-Object -First 1
        if ($pkg -and $pkg.InstallLocation) {
            $p = Join-Path $pkg.InstallLocation "PAD.Console.Host.exe"
            if (Test-Path -LiteralPath $p) { return $p }
        }
    } catch { }

    # Cach 2: ban cai bang bo cai MSI truyen thong
    foreach ($c in @(
        "C:\Program Files (x86)\Power Automate Desktop\PAD.Console.Host.exe",
        "C:\Program Files\Power Automate Desktop\PAD.Console.Host.exe"
    )) {
        if (Test-Path -LiteralPath $c) { return $c }
    }

    # Cach 3: quet thang WindowsApps (phong khi Get-AppxPackage bi chan)
    try {
        $thuMuc = Get-ChildItem -LiteralPath "C:\Program Files\WindowsApps" `
                    -Filter "Microsoft.PowerAutomateDesktop_*" -Directory -ErrorAction SilentlyContinue |
                  Sort-Object -Property Name -Descending
        foreach ($d in $thuMuc) {
            $p = Join-Path $d.FullName "PAD.Console.Host.exe"
            if (Test-Path -LiteralPath $p) { return $p }
        }
    } catch { }

    return $null
}

# Protocol handler 'ms-powerautomate:' co duoc dang ky khong?
function Test-PADProtocol {
    return (Test-Path -LiteralPath "Registry::HKEY_CLASSES_ROOT\ms-powerautomate")
}

# ---------------- KIEM TRA MOI TRUONG ----------------
$CoProtocol = Test-PADProtocol
$PADPath    = Get-PADConsolePath

Write-Log ("Protocol 'ms-powerautomate:' : {0}" -f $(if ($CoProtocol) { "co" } else { "KHONG co" }))
Write-Log ("PAD.Console.Host.exe        : {0}" -f $(if ($PADPath) { $PADPath } else { "KHONG tim thay" }))

if (-not $CoProtocol -and -not $PADPath) {
    Write-Log "Khong tim thay Power Automate Desktop tren may nay. Da cai chua?" "LOI"
    exit 1
}

if ($ChiKiemTra) {
    Write-Log "Che do kiem tra - khong chay flow."
    exit 0
}

# ---------------- CHAY FLOW ----------------
# Ten flow phai duoc URL-encode: 'run job' -> 'run%20job'
$uri = "ms-powerautomate:/console/flow/run?workflowName=" + [uri]::EscapeDataString($TenFlow)
if ($InputArguments) {
    $uri += "&inputArguments=" + [uri]::EscapeDataString($InputArguments)
}

Write-Log "Goi flow '$TenFlow'"
Write-Log "URI: $uri"

# Cach A - kich hoat qua protocol handler.
# Uu tien cach nay: voi ban Store (MSIX), Windows se kich hoat dung goi ung dung.
# Chay thang file .exe trong WindowsApps hay bi chan ACL.
if ($CoProtocol) {
    try {
        Start-Process -FilePath $uri
        Write-Log "Da gui lenh chay flow '$TenFlow' (qua protocol handler)."
        exit 0
    } catch {
        Write-Log ("Protocol handler that bai: " + $_.Exception.Message + " - thu cach du phong.") "CANH BAO"
    }
}

# Cach B - goi thang PAD.Console.Host.exe
if ($PADPath) {
    try {
        Start-Process -FilePath $PADPath -ArgumentList "`"$uri`""
        Write-Log "Da gui lenh chay flow '$TenFlow' (goi thang PAD.Console.Host.exe)."
        exit 0
    } catch {
        Write-Log ("Goi thang .exe that bai: " + $_.Exception.Message) "LOI"
    }
}

Write-Log "Khong the khoi chay flow '$TenFlow' bang bat ky cach nao." "LOI"
exit 1
