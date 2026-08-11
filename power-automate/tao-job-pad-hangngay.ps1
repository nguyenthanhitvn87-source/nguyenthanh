# ==========================================================
# Tao Scheduled Task chay flow Power Automate Desktop moi ngay
# ==========================================================
# CACH DUNG:
# 1. De file nay CUNG THU MUC voi file "chay-flow-pad.ps1"
# 2. Sua phan "CAU HINH" ben duoi
# 3. Mo PowerShell quyen Administrator (chuot phai > Run as Administrator)
# 4. Chay:  powershell -ExecutionPolicy Bypass -File ".\tao-job-pad-hangngay.ps1"
#
# GO TASK RA:
#   Unregister-ScheduledTask -TaskName "ChayPowerAutomateFlowHangNgay" -Confirm:$false
# ==========================================================

#Requires -RunAsAdministrator

# ---------------- CAU HINH ----------------
$TenTask = "ChayPowerAutomateFlowHangNgay"
$GioChay = "15:55"       # Gio chay moi ngay, dinh dang HH:mm (24h)

# CACH CHI DINH FLOW - chon MOT trong hai:
#
# Cach 1 (NEN DUNG): dan nguyen URI lay tu Power Automate Desktop.
#   Lay o dau: chuot phai flow trong PAD > Create desktop shortcut,
#   roi chuot phai shortcut vua tao > Properties > copy o "Target".
#   URI nay dung workflowid (GUID) nen doi ten flow ve sau khong lam hong lich chay,
#   va co san environmentid nen khong nham moi truong.
$UriFlow = "ms-powerautomate:/console/flow/run?environmentid=Default-fe9c4641-3b53-43d0-af72-8f3e64d3aa05&workflowid=a27713c9-83fd-44d9-be54-c1d1760e88bd&source=Other"
#
# Cach 2 (du phong): de $UriFlow = "" o tren, roi dien ten flow vao day.
#   Ten phai dung y het, ke ca hoa/thuong va dau cach.
$TenFlow = "run job"

# Chay PAD voi quyen Administrator?
#   $false (mac dinh) - giong nhu ban tu mo PAD binh thuong. Nen dung, vi ban PAD
#                       cai tu Microsoft Store (MSIX) thuong KHONG khoi chay duoc
#                       khi bi elevate.
#   $true             - chi bat khi flow can thao tac len ung dung chay quyen admin.
$QuyenCao = $false
# -------------------------------------------

$ErrorActionPreference = "Stop"

# --- Tim file launcher ---
$Launcher = Join-Path $PSScriptRoot "chay-flow-pad.ps1"
if (-not (Test-Path -LiteralPath $Launcher)) {
    Write-Host "KHONG tim thay 'chay-flow-pad.ps1' trong thu muc: $PSScriptRoot" -ForegroundColor Red
    Write-Host "Hai file phai nam cung mot thu muc." -ForegroundColor Red
    exit 1
}

# --- Xac dinh se goi flow theo cach nao ---
if ($UriFlow) {
    $ThamSoFlow = '-UriFlow "{0}"' -f $UriFlow.Trim()
    $MoTaFlow   = "URI: " + $UriFlow.Trim()
} elseif ($TenFlow) {
    $ThamSoFlow = '-TenFlow "{0}"' -f $TenFlow
    $MoTaFlow   = "Ten flow: " + $TenFlow
} else {
    Write-Host "Phai dien MOT trong hai: `$UriFlow hoac `$TenFlow o phan CAU HINH." -ForegroundColor Red
    exit 1
}

# --- Kiem tra truoc: may nay co Power Automate Desktop khong? ---
Write-Host "Dang kiem tra Power Automate Desktop..." -ForegroundColor Cyan
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Launcher -ChiKiemTra
if ($LASTEXITCODE -ne 0) {
    Write-Host "Kiem tra that bai - dung tao task. Cai/mo Power Automate Desktop roi chay lai." -ForegroundColor Red
    exit 1
}

# --- Action: goi launcher qua PowerShell, chay an cua so ---
# Luu y: KHONG dat ten bien la $Args - do la bien tu dong cua PowerShell.
# URI chua ky tu '&' nhung khong sao: Task Scheduler dua chuoi nay thang cho
# CreateProcess, khong di qua cmd.exe, nen '&' khong bi dien giai.
$PSExe   = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$ThamSo  = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden ' +
           ('-File "{0}" {1}' -f $Launcher, $ThamSoFlow)

$Action = New-ScheduledTaskAction -Execute $PSExe -Argument $ThamSo -WorkingDirectory $PSScriptRoot

# --- Trigger: moi ngay dung gio ---
# Parse bang InvariantCulture de khong phu thuoc dinh dang gio cua Windows
$MocGio  = [datetime]::ParseExact($GioChay, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
$Trigger = New-ScheduledTaskTrigger -Daily -At $MocGio

# --- Principal ---
# QUAN TRONG: phai la Interactive, KHONG duoc dung S4U.
# S4U ("Run whether user is logged on or not") chay o session khong co desktop
# tuong tac -> moi buoc UI automation cua PAD (click chuot, go phim, mo app)
# se treo hoac loi. Ban PAD attended (mien phi) bat buoc dung Interactive.
$UserId    = "$env:USERDOMAIN\$env:USERNAME"
$RunLevel  = if ($QuyenCao) { "Highest" } else { "Limited" }
$Principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel $RunLevel

# --- Settings ---
$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -WakeToRun `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 5)

# --- Dang ky ---
Register-ScheduledTask -TaskName $TenTask `
    -Action $Action `
    -Trigger $Trigger `
    -Principal $Principal `
    -Settings $Settings `
    -Description "Tu dong chay flow trong Power Automate Desktop luc $GioChay moi ngay. $MoTaFlow" `
    -Force | Out-Null

$task = Get-ScheduledTask -TaskName $TenTask
$info = Get-ScheduledTaskInfo -TaskName $TenTask

Write-Host ""
Write-Host "=== DA TAO TASK THANH CONG ===" -ForegroundColor Green
Write-Host ("  Ten task   : {0}" -f $TenTask)
Write-Host ("  Flow       : {0}" -f $MoTaFlow)
Write-Host ("  Chay luc   : {0} moi ngay" -f $GioChay)
Write-Host ("  Chay duoi  : {0} (Interactive, RunLevel={1})" -f $UserId, $RunLevel)
Write-Host ("  Lan ke tiep: {0}" -f $info.NextRunTime)
Write-Host ("  Trang thai : {0}" -f $task.State)
Write-Host ""
Write-Host "CHAY THU NGAY (khong can doi toi gio):" -ForegroundColor Yellow
Write-Host ("  Start-ScheduledTask -TaskName `"{0}`"" -f $TenTask)
Write-Host ""
Write-Host "XEM LOG:" -ForegroundColor Yellow
Write-Host ("  notepad `"{0}\PAD-Scheduler\chay-flow.log`"" -f $env:LOCALAPPDATA)
Write-Host ""
Write-Host "LUU Y: man hinh phai dang DANG NHAP va KHONG KHOA thi flow moi thao tac" -ForegroundColor Yellow
Write-Host "       duoc giao dien. Xem README.md de biet cach xu ly." -ForegroundColor Yellow
