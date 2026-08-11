# ==========================================================
# Tao Scheduled Task chay flow Power Automate Desktop moi ngay
# Ho tro NHIEU JOB, moi job mot URI va mot gio chay rieng.
# ==========================================================
# CACH DUNG:
# 1. De file nay CUNG THU MUC voi file "chay-flow-pad.ps1"
# 2. Sua phan "CAU HINH" ben duoi
# 3. Mo PowerShell quyen Administrator (chuot phai > Run as Administrator)
# 4. Chay:  powershell -ExecutionPolicy Bypass -File ".\tao-job-pad-hangngay.ps1"
#
# XEM CAC TASK DA TAO:
#   Get-ScheduledTask -TaskName "PAD - *" | Format-Table TaskName, State
# GO HET RA:
#   Get-ScheduledTask -TaskName "PAD - *" | Unregister-ScheduledTask -Confirm:$false
# ==========================================================

#Requires -RunAsAdministrator

# ---------------- CAU HINH ----------------
#
# Moi job la mot dong @{ ... } trong danh sach duoi day.
#   Ten = ten goi nho, se thanh ten task: "PAD - <Ten>"
#   Gio = gio chay moi ngay, dinh dang HH:mm (24h)
#   Uri = URI lay tu Power Automate Desktop
#
# LAY Uri O DAU:
#   Trong PAD, chuot phai vao flow > Create desktop shortcut.
#   Chuot phai shortcut vua tao > Properties > copy o "Target".
#   Dang: ms-powerautomate:/console/flow/run?environmentid=<GUID>&workflowid=<GUID>&source=Other
#
# THEM JOB THU 3, 4...: copy them mot khoi @{ ... }, nho dau phay ngan cach.

$DanhSachJob = @(
    @{
        Ten = "run job"
        Gio = "15:55"
        Uri = "ms-powerautomate:/console/flow/run?environmentid=Default-fe9c4641-3b53-43d0-af72-8f3e64d3aa05&workflowid=a27713c9-83fd-44d9-be54-c1d1760e88bd&source=Other"
    },
    @{
        Ten = "job 2"
        Gio = "16:10"
        Uri = "DAN_URI_CUA_JOB_THU_HAI_VAO_DAY"
    }
)

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

# --- Kiem tra cau hinh truoc khi dung toi Task Scheduler ---
$loi = @()
if ($DanhSachJob.Count -eq 0) { $loi += "Danh sach job dang trong." }

$daThayTen = @{}
foreach ($job in $DanhSachJob) {
    $ten = [string]$job.Ten
    if ([string]::IsNullOrWhiteSpace($ten)) { $loi += "Co job thieu 'Ten'."; continue }

    if ($daThayTen.ContainsKey($ten)) { $loi += "Job '$ten' bi trung ten - moi job phai co ten khac nhau." }
    $daThayTen[$ten] = $true

    if ([string]::IsNullOrWhiteSpace($job.Uri)) {
        $loi += "Job '$ten' thieu 'Uri'."
    } elseif ($job.Uri -notmatch '^ms-powerautomate:') {
        $loi += "Job '$ten': 'Uri' phai bat dau bang 'ms-powerautomate:'. Dang co: $($job.Uri)"
    }

    if ([string]::IsNullOrWhiteSpace($job.Gio)) {
        $loi += "Job '$ten' thieu 'Gio'."
    } else {
        try {
            [void][datetime]::ParseExact($job.Gio, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
        } catch {
            $loi += "Job '$ten': 'Gio' phai dinh dang HH:mm (vi du 09:00 hoac 15:55). Dang co: $($job.Gio)"
        }
    }
}

if ($loi.Count -gt 0) {
    Write-Host "CAU HINH CHUA DUNG - chua tao task nao:" -ForegroundColor Red
    foreach ($l in $loi) { Write-Host "  - $l" -ForegroundColor Red }
    exit 1
}

# --- Kiem tra may co Power Automate Desktop khong ---
Write-Host "Dang kiem tra Power Automate Desktop..." -ForegroundColor Cyan
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $Launcher -ChiKiemTra
if ($LASTEXITCODE -ne 0) {
    Write-Host "Kiem tra that bai - dung tao task. Cai/mo Power Automate Desktop roi chay lai." -ForegroundColor Red
    exit 1
}

# --- Phan dung chung cho moi task ---
$PSExe = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"

# QUAN TRONG: LogonType phai la Interactive, KHONG duoc dung S4U.
# S4U ("Run whether user is logged on or not") chay o session khong co desktop
# tuong tac -> moi buoc UI automation cua PAD (click chuot, go phim, mo app)
# se treo hoac loi. Ban PAD attended (mien phi) bat buoc dung Interactive.
$UserId    = "$env:USERDOMAIN\$env:USERNAME"
$RunLevel  = if ($QuyenCao) { "Highest" } else { "Limited" }
$Principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel $RunLevel

$Settings = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -WakeToRun `
    -MultipleInstances IgnoreNew `
    -ExecutionTimeLimit (New-TimeSpan -Hours 2) `
    -RestartCount 2 `
    -RestartInterval (New-TimeSpan -Minutes 5)

# --- Tao tung task ---
$daTao = @()

foreach ($job in $DanhSachJob) {
    $ten     = [string]$job.Ten
    $gio     = [string]$job.Gio
    $uri     = ([string]$job.Uri).Trim()
    $tenTask = "PAD - $ten"

    # Luu y: KHONG dat ten bien la $Args - do la bien tu dong cua PowerShell.
    # URI chua ky tu '&' nhung khong sao: Task Scheduler dua chuoi nay thang cho
    # CreateProcess, khong di qua cmd.exe, nen '&' khong bi dien giai.
    $thamSo = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden ' +
              ('-File "{0}" -TenJob "{1}" -UriFlow "{2}"' -f $Launcher, $ten, $uri)

    $Action  = New-ScheduledTaskAction -Execute $PSExe -Argument $thamSo -WorkingDirectory $PSScriptRoot
    $MocGio  = [datetime]::ParseExact($gio, "HH:mm", [System.Globalization.CultureInfo]::InvariantCulture)
    $Trigger = New-ScheduledTaskTrigger -Daily -At $MocGio

    Register-ScheduledTask -TaskName $tenTask `
        -Action $Action `
        -Trigger $Trigger `
        -Principal $Principal `
        -Settings $Settings `
        -Description "Tu dong chay flow '$ten' trong Power Automate Desktop luc $gio moi ngay. $uri" `
        -Force | Out-Null

    $info = Get-ScheduledTaskInfo -TaskName $tenTask
    $daTao += [pscustomobject]@{
        Task      = $tenTask
        Gio       = $gio
        LanKeTiep = $info.NextRunTime
    }
}

# --- Ket qua ---
Write-Host ""
Write-Host ("=== DA TAO {0} TASK ===" -f $daTao.Count) -ForegroundColor Green
$daTao | Format-Table -AutoSize
Write-Host ("Chay duoi: {0} (Interactive, RunLevel={1})" -f $UserId, $RunLevel)
Write-Host ""
Write-Host "CHAY THU NGAY (khong can doi toi gio):" -ForegroundColor Yellow
foreach ($t in $daTao) {
    Write-Host ("  Start-ScheduledTask -TaskName `"{0}`"" -f $t.Task)
}
Write-Host ""
Write-Host "XEM LOG:" -ForegroundColor Yellow
Write-Host ("  notepad `"{0}\PAD-Scheduler\chay-flow.log`"" -f $env:LOCALAPPDATA)
Write-Host ""
Write-Host "LUU Y 1: man hinh phai dang DANG NHAP va KHONG KHOA thi flow moi thao tac" -ForegroundColor Yellow
Write-Host "         duoc giao dien."
Write-Host "LUU Y 2: dat gio cac job CACH XA NHAU (it nhat 10-15 phut). Power Automate" -ForegroundColor Yellow
Write-Host "         Desktop chay lan luot tung flow, hai flow trung gio se gianh nhau."
