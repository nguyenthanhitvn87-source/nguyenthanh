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

# ---------------- DO DUONG DAN PAD ----------------
# Tra ve DANH SACH cac duong dan ung vien, xep theo do tin cay giam dan.
function Get-PADConsolePaths {
    $ketQua = New-Object System.Collections.Generic.List[string]
    function Them([string]$p) {
        if ($p -and (Test-Path -LiteralPath $p) -and -not $ketQua.Contains($p)) { $ketQua.Add($p) }
    }

    # Uu tien 1: App Execution Alias cua ban Store.
    # Goi qua alias nay thi Windows kich hoat dung goi MSIX (co package identity),
    # khac han voi viec goi thang file .exe nam trong WindowsApps (hay bi chan ACL).
    Them (Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\PAD.Console.Host.exe")

    # Uu tien 2: ban cai bang bo cai MSI truyen thong
    Them "C:\Program Files (x86)\Power Automate Desktop\PAD.Console.Host.exe"
    Them "C:\Program Files\Power Automate Desktop\PAD.Console.Host.exe"

    # Uu tien 3: thu muc cai dat that cua goi MSIX
    try {
        $pkg = Get-AppxPackage -Name "Microsoft.PowerAutomateDesktop*" -ErrorAction SilentlyContinue |
               Sort-Object -Property Version -Descending |
               Select-Object -First 1
        if ($pkg -and $pkg.InstallLocation) {
            Them (Join-Path $pkg.InstallLocation "PAD.Console.Host.exe")
        }
    } catch { }

    # Uu tien 4: quet thang WindowsApps (phong khi Get-AppxPackage bi chan)
    try {
        $thuMuc = Get-ChildItem -LiteralPath "C:\Program Files\WindowsApps" `
                    -Filter "Microsoft.PowerAutomateDesktop_*" -Directory -ErrorAction SilentlyContinue |
                  Sort-Object -Property Name -Descending
        foreach ($d in $thuMuc) { Them (Join-Path $d.FullName "PAD.Console.Host.exe") }
    } catch { }

    return $ketQua
}

# Protocol handler 'ms-powerautomate:' co xuat hien trong registry khong?
# CHI DE THAM KHAO khi doc log - KHONG duoc dung lam dieu kien chan.
# Voi ung dung MSIX, protocol dang ky theo app model chu khong phai lúc nao cung
# tao khoa o HKEY_CLASSES_ROOT, nen ket qua $false o day rat hay la bao dong gia.
function Test-PADProtocol {
    foreach ($k in @(
        "Registry::HKEY_CLASSES_ROOT\ms-powerautomate",
        "HKCU:\Software\Classes\ms-powerautomate"
    )) {
        if (Test-Path -LiteralPath $k) { return $true }
    }
    return $false
}

# ---------------- KIEM TRA MOI TRUONG ----------------
$CoProtocol = Test-PADProtocol
$PADPaths   = Get-PADConsolePaths

Write-Log ("Protocol 'ms-powerautomate:' trong registry: {0} (chi tham khao)" -f $(if ($CoProtocol) { "co" } else { "khong thay" }))
if ($PADPaths.Count -eq 0) {
    Write-Log "PAD.Console.Host.exe: KHONG tim thay duong dan nao."
} else {
    for ($i = 0; $i -lt $PADPaths.Count; $i++) {
        Write-Log ("PAD.Console.Host.exe [{0}]: {1}" -f ($i + 1), $PADPaths[$i])
    }
}

if ($ChiKiemTra) {
    if ($PADPaths.Count -eq 0 -and -not $CoProtocol) {
        Write-Log "Khong thay dau vet nao cua Power Automate Desktop tren may nay. Da cai chua?" "LOI"
        exit 1
    }
    Write-Log "Che do kiem tra - khong chay flow. Moi truong dat yeu cau."
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

# Cach A - kich hoat qua protocol handler, de Windows tu tim ung dung.
# LUON thu cach nay truoc, khong phu thuoc ket qua kiem tra registry o tren:
# neu protocol chua dang ky that thi Start-Process se nem loi va roi xuong Cach B.
try {
    Start-Process -FilePath $uri -ErrorAction Stop
    Write-Log "Da gui lenh chay flow '$TenFlow' (qua protocol handler)."
    exit 0
} catch {
    Write-Log ("Protocol handler khong dung duoc: " + $_.Exception.Message) "CANH BAO"
}

# Cach B - goi thang PAD.Console.Host.exe, thu lan luot cac duong dan tim duoc
foreach ($p in $PADPaths) {
    try {
        Start-Process -FilePath $p -ArgumentList "`"$uri`"" -ErrorAction Stop
        Write-Log "Da gui lenh chay flow '$TenFlow' qua: $p"
        exit 0
    } catch {
        Write-Log ("Goi that bai [{0}]: {1}" -f $p, $_.Exception.Message) "CANH BAO"
    }
}

Write-Log "Khong the khoi chay flow '$TenFlow' bang bat ky cach nao." "LOI"
exit 1
