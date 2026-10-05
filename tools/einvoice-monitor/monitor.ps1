# ============================================================
# Billing Pipeline Monitor (PowerShell - khong can Python)
# SAP -> PP -> VNPT -> Thue
# ============================================================
# Chay: mo run.bat (bam dup) hoac:
#   powershell -ExecutionPolicy Bypass -File monitor.ps1
#   powershell -ExecutionPolicy Bypass -File monitor.ps1 -CheckDate 20260820   (kiem tra 1 ngay qua khu)
#
# Yeu cau: WinSCP da cai tren may (dung de kiem tra SFTP).
# Neu WinSCP cai o duong dan khac, sua "winscp_path" trong config.json.
# ============================================================
param(
    [string]$CheckDate = (Get-Date -Format "yyyyMMdd")
)

$ErrorActionPreference = "Stop"
$ScriptDir = $PSScriptRoot
Set-Location $ScriptDir

$Config = Get-Content -Path "$ScriptDir\config.json" -Raw -Encoding UTF8 | ConvertFrom-Json

$Results = @()

# Bo dem so luong billing theo tung buoc (dung de hien thi tong quan tren dashboard)
$SapBillingCount = 0
# Cac loai chung tu KHONG tinh vao so lieu Server/BU (khong xac nhan duoc trang thai da ra hoa don)
$ExcludedFromBuServerCount = @("PXK", "Billing Return", "Billing Dieu Chinh", "Cancel Bill", "Z3F2 (Khong xuat HD)")
$BillingToDocType = @{}
$ServerExcludedDocType = @{}
$BuExcludedDocType = @{}
$SeenExclBnPerServer = @{}
$SeenExclBnPerBu = @{}
$SeenBnGlobal = @{}
$SeenBnPerServer = @{}
$SeenBnPerBu = @{}
$SeenBnPerDocType = @{}
$StagingBillingCount = 0
$VnptSuccessCount = 0
$CallbackCount = 0

# Danh sach chi tiet file, dung de embed vao dashboard cho o tim kiem truc tiep tren trinh duyet
$AllSapFiles = @()
$AllStagingFiles = @()
$AllBakupFiles = @()
# Danh sach cac thu muc da quet de tim billing (hien tren dashboard khi khong tim thay)
$ScannedFolders = @()

# Bang anh xa ma BU (VNxx trong ten file) sang ten Business Unit
$BuMap = @{
    "VN54" = "HEC"; "VN55" = "HEC"; "VN56" = "HEC"; "VN57" = "HEC"; "VN58" = "HEC"; "VN60" = "HEC"
    "VN03" = "CG"
    "VN78" = "PM"; "VN79" = "PM"
    "VN49" = "CDV & TEC"; "VN40" = "CDV & TEC"; "VN43" = "CDV & TEC"
    "VN20" = "ECOM"
}
$BuCounts = @{}
$UnmappedBuCodes = @{}
$BuIssuedCounts = @{}
$BillingToBu = @{}
$BillingToServer = @{}
$ServerUntrackableFiles = @{}
$BuUntrackableFiles = @{}
$BillingToSapTime = @{}
$ProcessingMinutes = @()
$BakupBillingSet = @{}
$DocTypeCounts = @{}
$UnmappedDocCodes = @{}
$ServerCounts = @{}
$ServerIssuedCounts = @{}

function Get-BillingNumber {
    param($FileName)
    if ($FileName -match '(\d{10})') { return $Matches[1] }
    return $null
}

function Get-DocType {
    param($FileName)
    $script:LastDocCode = $null
    if ([string]::IsNullOrWhiteSpace($FileName)) { return "Khac" }
    $tokens = $FileName -split '_'
    if (-not $tokens -or $tokens.Count -eq 0) { return "Khac" }
    $isBvk = ($tokens[0] -eq "BVK")
    # Uu tien xac dinh theo ma Zxx (ZF2/ZG2/ZRE/...) TRUOC -- vi 1 file BVK van co the la
    # dieu chinh (ZG2), return (ZRE), v.v. khong phai luon la "Smallunit".
    $docCode = $tokens | Where-Object { $_ -match '^Z[A-Z0-9]{2,3}$' } | Select-Object -First 1
    $script:LastDocCode = $docCode
    $result = switch ($docCode) {
        "ZF2" { if ($isBvk) { "Billing Smallunit" } else { "Billing Ban" } }
        "ZG2" { "Billing Dieu Chinh" }
        "ZL2" { "Billing Dieu Chinh" }
        "ZRE" { "Billing Return" }
        "ZVRF" { "PXK" }
        "Z3F2" { "Z3F2 (Khong xuat HD)" }
        "ZS1" { "Cancel Bill" }
        "ZS2" { "Cancel Bill" }
        default { if ($isBvk) { "Billing Smallunit" } else { "Khac" } }
    }
    if ([string]::IsNullOrWhiteSpace($result)) { return "Khac" }
    return $result
}

function Get-EmbeddedSapTime {
    # Gio SAP that su. Co 2 dinh dang gap:
    #  1) ..._250826_114205_20260825114347.txt -> 250826=ddMMyy, 114205=HHmmss
    #  2) ..._20260825-145100-021_20260825135108.txt (PXK/ZVRF) -> YYYYMMDD-HHMMSS-mmm
    param($FileName)
    if ($FileName -match '(\d{6})_(\d{6})_\d{14}\.txt$') {
        try {
            $d = $Matches[1]; $t = $Matches[2]
            $dd = $d.Substring(0,2); $mm = $d.Substring(2,2); $yy = $d.Substring(4,2)
            return [datetime]::ParseExact("20$yy$mm$dd$t", "yyyyMMddHHmmss", $null)
        } catch { return $null }
    }
    if ($FileName -match '(\d{8})-(\d{6})-\d+_\d{14}\.txt$') {
        try { return [datetime]::ParseExact("$($Matches[1])$($Matches[2])", "yyyyMMddHHmmss", $null) } catch { return $null }
    }
    return $null
}

function Get-EmbeddedPpTime {
    # Ten file luon ket thuc bang 1 chuoi 14 so (YYYYMMDDHHMMSS) truoc ".txt"
    # -- day la gio PP xu ly/nhan file (khop voi Date Modified), KHONG phai gio SAP.
    param($FileName)
    if ($FileName -match '(\d{14})\.txt$') {
        try { return [datetime]::ParseExact($Matches[1], "yyyyMMddHHmmss", $null) } catch { return $null }
    }
    return $null
}

function Get-EmbeddedBakupTime {
    # Ten file Bakup bat dau bang YYYYMMDD-HHMMSS-mmm
    param($FileName)
    if ($FileName -match '^(\d{8})-(\d{6})-\d+') {
        try { return [datetime]::ParseExact("$($Matches[1])$($Matches[2])", "yyyyMMddHHmmss", $null) } catch { return $null }
    }
    return $null
}

function Add-StagingFile {
    # Ghi lai 1 file dang nam o staging PP de o tim kiem tren dashboard biet billing dang o DAU:
    # loai (INV/DO), hang doi (thu muc con 1/2/3...), duong dan day du, da cho bao lau, co bi ket khong.
    # Doc them noi dung file de tim so billing 10 chu so (phong khi ten file khong chua so billing).
    param($File, $Kind, $RootFolder, $MaxStuck, $Now, [bool]$IsError = $false)
    $queue = ""
    if ($File.DirectoryName.TrimEnd('\') -ne $RootFolder.TrimEnd('\')) { $queue = $File.Directory.Name }
    $bnList = @()
    $bnFromName = Get-BillingNumber -FileName $File.Name
    if ($bnFromName) { $bnList += $bnFromName }
    if ($File.Length -gt 0 -and $File.Length -lt 2MB) {
        try {
            $content = [System.IO.File]::ReadAllText($File.FullName)
            $bnList += [regex]::Matches($content, '(?<!\d)\d{10}(?!\d)') | ForEach-Object { $_.Value } | Select-Object -Unique -First 200
        } catch {}
    }
    $ageMin = [math]::Round(($Now - $File.LastWriteTime).TotalMinutes)
    $script:AllStagingFiles += [PSCustomObject]@{
        n = $File.Name
        t = $File.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss")
        k = $Kind
        q = $queue
        p = $File.FullName
        a = $ageMin
        x = [bool]($ageMin -gt $MaxStuck)
        c = (@($bnList | Select-Object -Unique) -join ",")
        e = $IsError
    }
}

# ------------------------------------------------------------
# Du phong qua FTP (WinSCP) khi khong vao duoc thu muc bang UNC (\\server\...)
# ------------------------------------------------------------
function Get-FtpServerHost {
    # Tim host FTP cua server theo ten (VD: "06P" -> muc VNSGNEIVAP06P trong sap_to_pp.sftp.servers)
    param($ServerKey)
    $entry = $Config.sap_to_pp.sftp.servers | Where-Object { $_.name -match [regex]::Escape($ServerKey) -or $_.host -match [regex]::Escape($ServerKey) } | Select-Object -First 1
    if ($entry) { return $entry.host }
    return $null
}

function Invoke-FtpList {
    # Liet ke 1 thu muc qua WinSCP. Tra ve @{ Ok; Error; Items = [ {Name; IsDir; Time; Size} ] }
    param($ServerHost, $RemotePath)
    $res = [PSCustomObject]@{ Ok = $false; Error = ""; Items = @() }
    if (-not (Test-Path $WinSCPPathResolved -ErrorAction SilentlyContinue)) { $res.Error = "Khong tim thay WinSCP"; return $res }
    $siteName = $Config.sap_to_pp.sftp.winscp_site_name
    $sftpUser = $Config.sap_to_pp.sftp.username
    $sftpPass = $env:SFTP_PASSWORD
    try {
        $tmpScript = [System.IO.Path]::GetTempFileName()
        if ($siteName) {
            $openCmd = "open `"$siteName`""
        } else {
            $encodedPass = [uri]::EscapeDataString($sftpPass)
            $openCmd = "open ftp://${sftpUser}:${encodedPass}@${ServerHost}"
        }
        @("option batch abort", "option confirm off", $openCmd, "ls `"$RemotePath`"", "exit") | Out-File -FilePath $tmpScript -Encoding ASCII
        $output = & $WinSCPPathResolved /script="$tmpScript" /nointeractiveinput 2>&1
        Remove-Item $tmpScript -ErrorAction SilentlyContinue
        $outputText = $output -join "`n"
        if ($outputText -match "Authentication failed|Access denied") { $res.Error = "Sai username/password FTP"; return $res }
        if ($outputText -match "No such file|not found|Error listing directory|Could not retrieve directory|Cannot open") { $res.Error = "Khong co thu muc '$RemotePath' tren FTP"; return $res }
        $months = @{ Jan=1; Feb=2; Mar=3; Apr=4; May=5; Jun=6; Jul=7; Aug=8; Sep=9; Oct=10; Nov=11; Dec=12 }
        foreach ($line in $output) {
            $l = "$line"
            # VD: -rw-rw-rw-   1 ftp  ftp   1234 Sep 25 19:12:20 2026 ten file.txt
            if ($l -match '^(?<type>[dl-])\S{9}\s+\S+\s+\S+\s+\S+\s+(?<size>\d+)\s+(?<mon>[A-Z][a-z]{2})\s+(?<day>\d{1,2})\s+(?:(?<time>\d{1,2}:\d{2}(?::\d{2})?)\s+(?<year>\d{4})|(?<time2>\d{1,2}:\d{2})|(?<year2>\d{4}))\s+(?<name>.+?)\s*$') {
                $name = $Matches['name']
                if ($name -eq "." -or $name -eq "..") { continue }
                $t = $null
                try {
                    $yr = if ($Matches['year']) { [int]$Matches['year'] } elseif ($Matches['year2']) { [int]$Matches['year2'] } else { (Get-Date).Year }
                    $tm = if ($Matches['time']) { $Matches['time'] } elseif ($Matches['time2']) { $Matches['time2'] } else { "00:00" }
                    if ($tm.Split(':').Count -eq 2) { $tm = "$tm`:00" }
                    $hms = $tm.Split(':')
                    $t = Get-Date -Year $yr -Month $months[$Matches['mon']] -Day ([int]$Matches['day']) -Hour ([int]$hms[0]) -Minute ([int]$hms[1]) -Second ([int]$hms[2]) -Millisecond 0
                    if ($t -gt (Get-Date).AddDays(1)) { $t = $t.AddYears(-1) }   # dang "Sep 25 19:12" khong co nam
                } catch { $t = $null }
                $res.Items += [PSCustomObject]@{ Name = $name; IsDir = ($Matches['type'] -eq 'd'); Time = $t; Size = [long]$Matches['size'] }
            }
        }
        $res.Ok = $true
        return $res
    } catch {
        $res.Error = $_.Exception.Message
        return $res
    }
}

function Get-FtpFilesRecursive {
    # Lay danh sach file trong thu muc FTP, di xuong thu muc con toi da $MaxDepth cap.
    # Bo qua thu muc con cu hon $MinTime (neu co) de khong liet ke ca nam du lieu.
    param($ServerHost, $RemotePath, [int]$MaxDepth = 2, $MinTime = $null, [switch]$SkipBak)
    $out = [PSCustomObject]@{ Ok = $false; Error = ""; Files = @() }
    $root = Invoke-FtpList -ServerHost $ServerHost -RemotePath $RemotePath
    if (-not $root.Ok) { $out.Error = $root.Error; return $out }
    $out.Ok = $true
    $queue = @(@{ Path = $RemotePath.TrimEnd('/'); Items = $root.Items; Depth = 0; Sub = "" })
    while ($queue.Count -gt 0) {
        $cur = $queue[0]; $queue = @($queue | Select-Object -Skip 1)
        foreach ($it in $cur.Items) {
            $full = "$($cur.Path)/$($it.Name)"
            if ($it.IsDir) {
                if ($cur.Depth -ge $MaxDepth) { continue }
                if ($SkipBak -and $it.Name -eq "BAK") { continue }
                if ($MinTime -and $it.Time -and $it.Time -lt $MinTime) { continue }
                $sub = Invoke-FtpList -ServerHost $ServerHost -RemotePath $full
                if ($sub.Ok) { $queue += @{ Path = $full; Items = $sub.Items; Depth = $cur.Depth + 1; Sub = $it.Name } }
            } else {
                $out.Files += [PSCustomObject]@{ Name = $it.Name; FullPath = $full; Time = $it.Time; Sub = $cur.Sub; Size = $it.Size }
            }
        }
    }
    return $out
}

function ConvertTo-FakeFile {
    # Bien 1 file FTP thanh doi tuong giong FileInfo de dung chung Add-StagingFile
    param($FtpFile, $RootLabel, $ClockOffsetHours = 0)
    $t = if ($FtpFile.Time) { $FtpFile.Time.AddHours($ClockOffsetHours) } else { Get-Date }
    $dirName = if ($FtpFile.Sub) { "$RootLabel/$($FtpFile.Sub)" } else { $RootLabel }
    return [PSCustomObject]@{
        Name = $FtpFile.Name; FullName = "$RootLabel/" + $(if ($FtpFile.Sub) { "$($FtpFile.Sub)/" } else { "" }) + $FtpFile.Name
        DirectoryName = $dirName; Directory = [PSCustomObject]@{ Name = $FtpFile.Sub }
        LastWriteTime = $t; Length = 0   # Length = 0 -> khong doc noi dung (FTP chi co ten file)
    }
}

function Resolve-WinSCPPath {
    param($ConfigPath, $BaseDir)
    if ($ConfigPath -match '^[A-Za-z]:\\') {
        return $ConfigPath
    }
    return (Join-Path $BaseDir $ConfigPath)
}

$WinSCPPathResolved = Resolve-WinSCPPath -ConfigPath $Config.sap_to_pp.winscp_path -BaseDir $ScriptDir

function Add-Result {
    param(
        [string]$Stage,
        [string]$Name,
        [ValidateSet("OK", "WARNING", "ERROR")]
        [string]$Severity,
        [string]$Message
    )
    $script:Results += [PSCustomObject]@{
        Stage    = $Stage
        Name     = $Name
        Severity = $Severity
        Message  = $Message
        Time     = Get-Date
    }
    $icon = switch ($Severity) { "OK" {"[OK]"} "WARNING" {"[CANH BAO]"} "ERROR" {"[LOI]"} }
    Write-Host "$icon [$Stage] $Name -- $Message"
}

Write-Host "=== Billing Monitor chay luc $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ==="
Write-Host ""

# ------------------------------------------------------------
# BUOC A: SAP -> PP  (kiem tra qua SFTP/FTP bang WinSCP, lap qua tung server)
# ------------------------------------------------------------
$WinSCPPath = $WinSCPPathResolved
if (-not (Test-Path $WinSCPPath)) {
    Add-Result -Stage "SAP->PP" -Name "WinSCP" -Severity "ERROR" `
        -Message "Khong tim thay WinSCP tai '$WinSCPPath'. Mo WinSCP, vao Help > About de xem duong dan that, roi sua trong config.json (muc winscp_path)."
} else {
    $dateStr = $CheckDate
    $remoteFolder = $Config.sap_to_pp.sftp.remote_folder_template -replace "\{date\}", $dateStr
    $sftpUser = $Config.sap_to_pp.sftp.username
    $sftpPass = $env:SFTP_PASSWORD
    $siteName = $Config.sap_to_pp.sftp.winscp_site_name

    if (-not $sftpPass -and -not $siteName) {
        Add-Result -Stage "SAP->PP" -Name "Mat khau SFTP" -Severity "ERROR" `
            -Message "Chua co bien moi truong SFTP_PASSWORD. Kiem tra lai run.bat."
    } else {
        foreach ($server in $Config.sap_to_pp.sftp.servers) {
            $sftpHost = $server.host
            $serverLabel = $server.name
            try {
                $tmpScript = [System.IO.Path]::GetTempFileName()
                if ($siteName) {
                    @("option batch abort", "option confirm off", "open `"$siteName`"", "ls `"$remoteFolder`"", "exit") |
                        Out-File -FilePath $tmpScript -Encoding ASCII
                } else {
                    $encodedPass = [uri]::EscapeDataString($sftpPass)
                    @("option batch abort", "option confirm off", "open ftp://${sftpUser}:${encodedPass}@${sftpHost}", "ls `"$remoteFolder`"", "exit") |
                        Out-File -FilePath $tmpScript -Encoding ASCII
                }

                $output = & $WinSCPPath /script="$tmpScript" /nointeractiveinput 2>&1
                Remove-Item $tmpScript -ErrorAction SilentlyContinue
                $outputText = $output -join "`n"

                if ($outputText -match "Authentication failed" -or $outputText -match "Access denied") {
                    Add-Result -Stage "SAP->PP" -Name "Dang nhap ($serverLabel)" -Severity "ERROR" `
                        -Message "Sai username/password khi ket noi toi $sftpHost."
                    Write-Host "  --- DEBUG ($serverLabel) ---"
                    $output | ForEach-Object { Write-Host "  | $_" }
                    Write-Host "  --- HET DEBUG ---"
                } elseif ($outputText -match "No such file" -or $outputText -match "not found") {
                    Add-Result -Stage "SAP->PP" -Name "Thu muc ($serverLabel)" -Severity "ERROR" `
                        -Message "Khong tim thay thu muc '$remoteFolder' tren $serverLabel."
                } else {
                    $fileLines = $output | Where-Object { $_ -match '\.txt\s*$' }
                    if (-not $fileLines -or $fileLines.Count -eq 0) {
                        Add-Result -Stage "SAP->PP" -Name "File tu SAP ($serverLabel)" -Severity "WARNING" `
                            -Message "Khong co file .txt nao trong '$remoteFolder' tren $serverLabel."
                    } else {
                        Add-Result -Stage "SAP->PP" -Name "File tu SAP ($serverLabel)" -Severity "OK" `
                            -Message "$($fileLines.Count) file .txt tim thay tren $serverLabel."
                        if (-not $ServerCounts.ContainsKey($serverLabel)) { $ServerCounts[$serverLabel] = 0 }
                        if (-not $SeenBnPerServer.ContainsKey($serverLabel)) { $SeenBnPerServer[$serverLabel] = @{} }

                        # Thong ke theo BU dua vao ma VNxx trong ten file (VD: VN_VN57H2_ZF2_...)
                        # Dem theo SO BILLING DUY NHAT (1 billing co the co nhieu file trigger), khong dem theo file tho.
                        $skippedFileCount = 0
                        foreach ($line in $fileLines) {
                            try {
                                $parts = $line -split '\s+' | Where-Object { $_ -ne "" }
                                if (-not $parts -or $parts.Count -eq 0) { $skippedFileCount++; continue }
                                $fname = $parts[-1]
                                if ([string]::IsNullOrWhiteSpace($fname)) { $skippedFileCount++; continue }

                                $billingNum = Get-BillingNumber -FileName $fname

                                $buCode = $null
                                if ($fname -match 'VN(\d{2})') {
                                    $buCode = "VN$($Matches[1])"
                                }
                                $buLabel = "Khac"
                                if ($buCode -and $BuMap.ContainsKey($buCode)) { $buLabel = $BuMap[$buCode] }

                                $docType = Get-DocType -FileName $fname
                                if ([string]::IsNullOrWhiteSpace($docType)) { $docType = "Khac" }
                                if ($billingNum -and -not $BillingToDocType.ContainsKey($billingNum)) { $BillingToDocType[$billingNum] = $docType }

                                # Server/BU chi tinh cho cac loai chung tu co the xac nhan duoc trang thai ra hoa don
                                $countsForBuServer = ($docType -notin $ExcludedFromBuServerCount)

                                if (-not $countsForBuServer) {
                                    $isNewExclServer = (-not $billingNum) -or (-not $SeenExclBnPerServer.ContainsKey("$serverLabel|$billingNum"))
                                    if ($isNewExclServer) {
                                        if ($billingNum) { $SeenExclBnPerServer["$serverLabel|$billingNum"] = $true }
                                        if (-not $ServerExcludedDocType.ContainsKey($serverLabel)) { $ServerExcludedDocType[$serverLabel] = @{} }
                                        if ($ServerExcludedDocType[$serverLabel].ContainsKey($docType)) { $ServerExcludedDocType[$serverLabel][$docType]++ } else { $ServerExcludedDocType[$serverLabel][$docType] = 1 }
                                    }
                                    $isNewExclBu = (-not $billingNum) -or (-not $SeenExclBnPerBu.ContainsKey("$buLabel|$billingNum"))
                                    if ($isNewExclBu) {
                                        if ($billingNum) { $SeenExclBnPerBu["$buLabel|$billingNum"] = $true }
                                        if (-not $BuExcludedDocType.ContainsKey($buLabel)) { $BuExcludedDocType[$buLabel] = @{} }
                                        if ($BuExcludedDocType[$buLabel].ContainsKey($docType)) { $BuExcludedDocType[$buLabel][$docType]++ } else { $BuExcludedDocType[$buLabel][$docType] = 1 }
                                    }
                                }

                                # Neu khong co so billing, coi moi file la 1 don vi rieng (khong the dedupe)
                                $isNewBilling = (-not $billingNum) -or (-not $SeenBnGlobal.ContainsKey($billingNum))
                                $isNewForServer = $countsForBuServer -and ((-not $billingNum) -or (-not $SeenBnPerServer[$serverLabel].ContainsKey($billingNum)))

                                if (-not $SeenBnPerBu.ContainsKey($buLabel)) { $SeenBnPerBu[$buLabel] = @{} }
                                $isNewForBu = $countsForBuServer -and ((-not $billingNum) -or (-not $SeenBnPerBu[$buLabel].ContainsKey($billingNum)))

                                if (-not $SeenBnPerDocType.ContainsKey($docType)) { $SeenBnPerDocType[$docType] = @{} }
                                $isNewForDocType = (-not $billingNum) -or (-not $SeenBnPerDocType[$docType].ContainsKey($billingNum))

                                if ($isNewBilling) {
                                    $SapBillingCount++
                                    if ($billingNum) { $SeenBnGlobal[$billingNum] = $true }
                                }
                                if ($isNewForServer) {
                                    $ServerCounts[$serverLabel]++
                                    if ($billingNum) {
                                        $SeenBnPerServer[$serverLabel][$billingNum] = $true
                                    } else {
                                        if (-not $ServerUntrackableFiles.ContainsKey($serverLabel)) { $ServerUntrackableFiles[$serverLabel] = @() }
                                        $ServerUntrackableFiles[$serverLabel] += [PSCustomObject]@{ n = $fname; d = $docType }
                                    }
                                }
                                if ($isNewForBu) {
                                    if ($BuCounts.ContainsKey($buLabel)) { $BuCounts[$buLabel]++ } else { $BuCounts[$buLabel] = 1 }
                                    if ($billingNum) {
                                        $SeenBnPerBu[$buLabel][$billingNum] = $true
                                    } else {
                                        if (-not $BuUntrackableFiles.ContainsKey($buLabel)) { $BuUntrackableFiles[$buLabel] = @() }
                                        $BuUntrackableFiles[$buLabel] += [PSCustomObject]@{ n = $fname; d = $docType }
                                    }
                                    if ($buLabel -eq "Khac") {
                                        $unmappedKey = if ($buCode) { $buCode } else { "(khong co ma VN)" }
                                        if ($UnmappedBuCodes.ContainsKey($unmappedKey)) { $UnmappedBuCodes[$unmappedKey]++ } else { $UnmappedBuCodes[$unmappedKey] = 1 }
                                    }
                                }
                                if ($isNewForDocType) {
                                    if ($DocTypeCounts.ContainsKey($docType)) { $DocTypeCounts[$docType]++ } else { $DocTypeCounts[$docType] = 1 }
                                    if ($billingNum) { $SeenBnPerDocType[$docType][$billingNum] = $true }
                                    if ($docType -eq "Khac") {
                                        $unmappedDocKey = if ($script:LastDocCode) { $script:LastDocCode } else { "(khong co ma Zxx)" }
                                        if ($UnmappedDocCodes.ContainsKey($unmappedDocKey)) { $UnmappedDocCodes[$unmappedDocKey]++ } else { $UnmappedDocCodes[$unmappedDocKey] = 1 }
                                    }
                                }

                                if ($billingNum -and -not $BillingToBu.ContainsKey($billingNum)) { $BillingToBu[$billingNum] = $buLabel }
                                if ($billingNum -and -not $BillingToServer.ContainsKey($billingNum)) { $BillingToServer[$billingNum] = $serverLabel }

                                $sapTime = Get-EmbeddedSapTime -FileName $fname
                                $ppTime = Get-EmbeddedPpTime -FileName $fname
                                if ($billingNum -and $sapTime -and -not $BillingToSapTime.ContainsKey($billingNum)) {
                                    $BillingToSapTime[$billingNum] = $sapTime
                                }
                                $AllSapFiles += [PSCustomObject]@{
                                    n = $fname; s = $serverLabel; b = $buLabel; d = $docType; bn = $billingNum
                                    t = if ($sapTime) { $sapTime.ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
                                    tpp = if ($ppTime) { $ppTime.ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
                                }
                            } catch {
                                $skippedFileCount++
                            }
                        }
                        if ($skippedFileCount -gt 0) {
                            Add-Result -Stage "SAP->PP" -Name "Ten file bat thuong ($serverLabel)" -Severity "WARNING" `
                                -Message "$skippedFileCount file co ten khong dung dinh dang, da bo qua khi thong ke (khong anh huong tong so file)."
                        }
                    }
                }
            } catch {
                Add-Result -Stage "SAP->PP" -Name "Ket noi ($serverLabel)" -Severity "ERROR" `
                    -Message "Loi khi ket noi $serverLabel : $($_.Exception.Message)"
            }
        }
    }
}

# ------------------------------------------------------------
# BUOC B: PP Processing Folder (UNC path - kiem tra truc tiep)
# ------------------------------------------------------------
$StagingFolder = $Config.pp_staging.staging_folder
$MaxStuckMinutes = $Config.pp_staging.max_stuck_minutes
if ($Config.pp_staging.use_date_subfolder) {
    $StagingFolder = Join-Path $StagingFolder $CheckDate
}

if (-not (Test-Path $StagingFolder -ErrorAction SilentlyContinue)) {
    Add-Result -Stage "PP Processing" -Name "Truy cap staging folder" -Severity "ERROR" `
        -Message "Khong truy cap duoc thu muc '$StagingFolder'. Kiem tra ket noi mang / quyen truy cap."
} else {
    $now = Get-Date
    if ($Config.pp_staging.use_numbered_subfolders) {
        # Staging chia thanh nhieu thu muc con (1, 2, 3... hang doi xu ly song song) -- quet tat ca,
        # tru thu muc BAK (luu tru, khong phai hang doi thuc su)
        $files = Get-ChildItem -Path $StagingFolder -Directory -ErrorAction SilentlyContinue |
                 Where-Object { $_.Name -ne "BAK" } |
                 ForEach-Object { Get-ChildItem -Path $_.FullName -File -ErrorAction SilentlyContinue }
    } else {
        $files = Get-ChildItem -Path $StagingFolder -File -ErrorAction SilentlyContinue
    }
    $StagingBillingCount = $files.Count
    foreach ($f in $files) { Add-StagingFile -File $f -Kind "INV" -RootFolder $StagingFolder -MaxStuck $MaxStuckMinutes -Now $now }
    $ScannedFolders += [PSCustomObject]@{ k = "Staging INV"; p = $StagingFolder; c = @($files).Count; ok = $true }
    $stuckFiles = $files | Where-Object { ($now - $_.LastWriteTime).TotalMinutes -gt $MaxStuckMinutes }

    if ($stuckFiles.Count -gt 0) {
        $names = ($stuckFiles | Select-Object -First 5 | ForEach-Object { $_.Name }) -join ", "
        Add-Result -Stage "PP Processing" -Name "File bi ket tai staging" -Severity "WARNING" `
            -Message "$($stuckFiles.Count) file nam qua $MaxStuckMinutes phut chua duoc xu ly (trong so $($files.Count) file): $names"
    } else {
        Add-Result -Stage "PP Processing" -Name "File bi ket tai staging" -Severity "OK" `
            -Message "Khong co file nao bi ket qua $MaxStuckMinutes phut ($($files.Count) file dang cho xu ly)."
    }
}

# ------------------------------------------------------------
# BUOC B2: PP Processing rieng cho PXK (thu muc DO, khac voi INV)
# ------------------------------------------------------------
if ($Config.pp_staging_pxk -and $Config.pp_staging_pxk.staging_folder) {
    $PxkStagingFolder = $Config.pp_staging_pxk.staging_folder
    $PxkMaxStuck = $Config.pp_staging_pxk.max_stuck_minutes

    if (-not (Test-Path $PxkStagingFolder -ErrorAction SilentlyContinue)) {
        Add-Result -Stage "PP Processing (PXK)" -Name "Truy cap staging folder" -Severity "ERROR" `
            -Message "Khong truy cap duoc thu muc '$PxkStagingFolder'. Kiem tra ket noi mang / quyen truy cap."
    } else {
        $now = Get-Date
        if ($Config.pp_staging_pxk.use_numbered_subfolders) {
            $pxkFiles = Get-ChildItem -Path $PxkStagingFolder -Directory -ErrorAction SilentlyContinue |
                        Where-Object { $_.Name -ne "BAK" } |
                        ForEach-Object { Get-ChildItem -Path $_.FullName -File -ErrorAction SilentlyContinue }
        } else {
            $pxkFiles = Get-ChildItem -Path $PxkStagingFolder -File -ErrorAction SilentlyContinue
        }
        foreach ($f in $pxkFiles) { Add-StagingFile -File $f -Kind "DO (PXK)" -RootFolder $PxkStagingFolder -MaxStuck $PxkMaxStuck -Now $now }
        $ScannedFolders += [PSCustomObject]@{ k = "Staging DO (PXK)"; p = $PxkStagingFolder; c = @($pxkFiles).Count; ok = $true }
        $pxkStuck = $pxkFiles | Where-Object { ($now - $_.LastWriteTime).TotalMinutes -gt $PxkMaxStuck }

        if ($pxkStuck.Count -gt 0) {
            $pxkNames = ($pxkStuck | Select-Object -First 5 | ForEach-Object { $_.Name }) -join ", "
            Add-Result -Stage "PP Processing (PXK)" -Name "File bi ket tai staging (DO)" -Severity "WARNING" `
                -Message "$($pxkStuck.Count) file nam qua $PxkMaxStuck phut chua duoc xu ly (trong so $($pxkFiles.Count) file): $pxkNames"
        } else {
            Add-Result -Stage "PP Processing (PXK)" -Name "File bi ket tai staging (DO)" -Severity "OK" `
                -Message "Khong co file nao bi ket qua $PxkMaxStuck phut ($($pxkFiles.Count) file dang cho xu ly trong thu muc DO)."
        }
    }
}

# ------------------------------------------------------------
# BUOC B3: Kiem tra Service co dang chay hay khong
# (Neu con file .txt ton dong trong \\vnsgneivap05p\PRODATA (khong phai thu muc theo ngay)
#  nghia la service da dung, chua co ai lay file di xu ly)
# ------------------------------------------------------------
$ServiceStatusResults = @{}
$ServiceHistoryPath = if ($Config.service_check -and $Config.service_check.history_path) {
    $Config.service_check.history_path
} else {
    Join-Path $ScriptDir "service_history.csv"
}
$ServiceHistoryDir = Split-Path $ServiceHistoryPath -Parent
if ($ServiceHistoryDir -and -not (Test-Path $ServiceHistoryDir -ErrorAction SilentlyContinue)) {
    try { New-Item -ItemType Directory -Path $ServiceHistoryDir -Force -ErrorAction Stop | Out-Null } catch {}
}

if ($Config.service_check -and $Config.service_check.enabled) {
    $svcPattern = $Config.service_check.file_pattern
    $svcMaxAge = $Config.service_check.max_file_age_minutes
    if (-not $svcMaxAge) { $svcMaxAge = 5 }

    # Tao file lich su neu chua co (header CSV)
    if (-not (Test-Path $ServiceHistoryPath -ErrorAction SilentlyContinue)) {
        "CheckTime,Server,Status,StuckFileCount" | Out-File -FilePath $ServiceHistoryPath -Encoding UTF8
    }

    foreach ($svc in $Config.service_check.servers) {
        $svcFolder = $svc.folder
        $svcLabel = $svc.name

        if (-not (Test-Path $svcFolder -ErrorAction SilentlyContinue)) {
            Add-Result -Stage "Service Status" -Name "Truy cap thu muc ($svcLabel)" -Severity "ERROR" `
                -Message "Khong truy cap duoc thu muc '$svcFolder' de kiem tra trang thai service."
            $ServiceStatusResults[$svcLabel] = @{ Status = "UNKNOWN"; Count = 0 }
            "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$svcLabel,UNKNOWN,0" | Out-File -FilePath $ServiceHistoryPath -Append -Encoding UTF8
        } else {
            $svcFiles = Get-ChildItem -Path $svcFolder -Filter $svcPattern -File -ErrorAction SilentlyContinue
            $now = Get-Date
            $oldFiles = $svcFiles | Where-Object { ($now - $_.LastWriteTime).TotalMinutes -gt $svcMaxAge }

            if ($oldFiles.Count -gt 0) {
                $svcNames = ($oldFiles | Select-Object -First 5 | ForEach-Object { $_.Name }) -join ", "
                # Thoi diem service that su bat dau dung = LastWriteTime cua file cu nhat + nguong,
                # chinh xac hon la dung gio kiem tra (vi file co the da nam do tu truoc do lau).
                $oldestFileTime = ($oldFiles | Sort-Object LastWriteTime | Select-Object -First 1).LastWriteTime
                $inferredStopTime = $oldestFileTime.AddMinutes($svcMaxAge)
                Add-Result -Stage "Service Status" -Name "Trang thai service ($svcLabel)" -Severity "ERROR" `
                    -Message "Service co ve DA DUNG tu khoang $($inferredStopTime.ToString('HH:mm:ss')) -- con $($oldFiles.Count) file .txt nam qua $svcMaxAge phut chua duoc lay di xu ly trong '$svcFolder': $svcNames"
                $ServiceStatusResults[$svcLabel] = @{ Status = "STOPPED"; Count = $oldFiles.Count }
                "$($inferredStopTime.ToString('yyyy-MM-dd HH:mm:ss')),$svcLabel,STOPPED,$($oldFiles.Count)" | Out-File -FilePath $ServiceHistoryPath -Append -Encoding UTF8
            } else {
                Add-Result -Stage "Service Status" -Name "Trang thai service ($svcLabel)" -Severity "OK" `
                    -Message "Service dang chay binh thuong -- khong co file nao nam qua $svcMaxAge phut trong '$svcFolder' ($($svcFiles.Count) file dang co, deu con moi)."
                $ServiceStatusResults[$svcLabel] = @{ Status = "RUNNING"; Count = 0 }
                "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'),$svcLabel,RUNNING,0" | Out-File -FilePath $ServiceHistoryPath -Append -Encoding UTF8
            }
        }
    }

    # Don bot file lich su neu qua dai (chi giu 30 ngay gan nhat) de tranh phinh to theo thoi gian
    try {
        $histLines = Get-Content -Path $ServiceHistoryPath -Encoding UTF8
        if ($histLines.Count -gt 50000) {
            $cutoffHist = (Get-Date).AddDays(-30)
            $header = $histLines[0]
            $kept = $histLines[1..($histLines.Count - 1)] | Where-Object {
                $parts = $_ -split ','
                if ($parts.Count -ge 1) {
                    try { return ([datetime]$parts[0]) -gt $cutoffHist } catch { return $true }
                }
                return $true
            }
            @($header) + $kept | Out-File -FilePath $ServiceHistoryPath -Encoding UTF8
        }
    } catch {}
}

# ------------------------------------------------------------
# BUOC B4: Quet them cac thu muc khac de o tim kiem biet billing dang nam o DAU
#  - Thu muc service (PRODATA): file SAP gui sang, dang cho service lay vao xu ly
#  - Cac thu muc khai bao them trong config.json, muc "search_folders" (VD: staging server 06P,
#    thu muc Error/Loi, thu muc gui VNPT...). Moi muc: { "name": "...", "path": "...",
#    "pattern": "*.txt" (tuy chon), "recurse": false (tuy chon), "max_stuck_minutes": 30 (tuy chon),
#    "is_error": true (tuy chon -- danh dau day la thu muc LOI, billing nam o day se bao do),
#    "max_age_days": 30 (tuy chon -- chi lay file sua trong N ngay gan nhat) }
# ------------------------------------------------------------
$ExtraSearchFolders = @()
if ($Config.service_check -and $Config.service_check.servers) {
    foreach ($svc in $Config.service_check.servers) {
        $ExtraSearchFolders += [PSCustomObject]@{
            name = "Cho service lay ($($svc.name))"; path = $svc.folder
            pattern = $(if ($Config.service_check.file_pattern) { $Config.service_check.file_pattern } else { "*.txt" })
            recurse = $false; max_stuck_minutes = $(if ($Config.service_check.max_file_age_minutes) { $Config.service_check.max_file_age_minutes } else { 5 })
        }
    }
}
if ($Config.search_folders) { $ExtraSearchFolders += $Config.search_folders }

# Thu muc LOI cua PP -- luon quet (khong can khai bao trong config.json).
# Chi lay file sua trong $ErrorFolderMaxAgeDays ngay gan nhat de khong bi cham khi thu muc loi tich tu lau.
$ErrorFolderMaxAgeDays = 30
$DefaultErrorFolders = @()
foreach ($errSrv in @("vnsgneivap05p", "vnsgneivap06p")) {
    $srvShort = $errSrv.Substring($errSrv.Length - 3).ToUpper()   # 05P / 06P
    foreach ($errKind in @("INV", "DO")) {
        $DefaultErrorFolders += [PSCustomObject]@{
            name = "Loi $errKind ($srvShort)"; path = "\\$errSrv\DKSH(VN)\ERROR\$errKind"
            pattern = "*"; recurse = $true; is_error = $true; max_age_days = $ErrorFolderMaxAgeDays
            # Du phong khi UNC bi tu choi: doc qua FTP (WinSCP). Doi duong dan FTP trong config.json:
            #   "error_folders_ftp_template": "/DKSH(VN)/ERROR/{kind}"
            ftp_server = $srvShort
            ftp_path = $(if ($Config.error_folders_ftp_template) { $Config.error_folders_ftp_template -replace '\{kind\}', $errKind } else { "/DKSH(VN)/ERROR/$errKind" })
        }
    }
}
foreach ($d in $DefaultErrorFolders) {
    # Neu config.json da khai bao cung duong dan thi khong quet 2 lan
    if (-not ($ExtraSearchFolders | Where-Object { $_.path -and $_.path.TrimEnd('\') -eq $d.path })) { $ExtraSearchFolders += $d }
}

foreach ($sf in $ExtraSearchFolders) {
    if (-not $sf.path) { continue }
    $sfPattern = if ($sf.pattern) { $sf.pattern } else { "*.txt" }
    $sfStuck = if ($sf.max_stuck_minutes) { $sf.max_stuck_minutes } else { $MaxStuckMinutes }
    if (-not (Test-Path $sf.path -ErrorAction SilentlyContinue)) {
        # UNC khong vao duoc -> thu qua FTP neu co cau hinh (ftp_server + ftp_path)
        $ftpHost = if ($sf.ftp_server) { Get-FtpServerHost -ServerKey $sf.ftp_server } else { $null }
        if ($ftpHost -and $sf.ftp_path) {
            $sfNow = Get-Date
            $minTime = if ($sf.max_age_days) { $sfNow.AddDays(-[double]$sf.max_age_days) } else { $null }
            $ftpRes = Get-FtpFilesRecursive -ServerHost $ftpHost -RemotePath $sf.ftp_path -MaxDepth $(if ($sf.recurse) { 2 } else { 0 }) -MinTime $minTime -SkipBak
            if ($ftpRes.Ok) {
                # Server 06P dong ho lech +1 tieng (giong phan Bakup)
                $offset = if ("$($sf.ftp_server)" -match '06P') { -1 } else { 0 }
                $rootLabel = "ftp://$ftpHost$($sf.ftp_path)"
                $ftpFiles = @($ftpRes.Files | Where-Object { $_.Name -like $sfPattern -and (-not $minTime -or -not $_.Time -or $_.Time -ge $minTime) })
                foreach ($ff in $ftpFiles) {
                    $fake = ConvertTo-FakeFile -FtpFile $ff -RootLabel $rootLabel -ClockOffsetHours $offset
                    Add-StagingFile -File $fake -Kind $sf.name -RootFolder $rootLabel -MaxStuck $sfStuck -Now $sfNow -IsError ([bool]$sf.is_error)
                }
                $ScannedFolders += [PSCustomObject]@{ k = "$($sf.name) - qua FTP"; p = $rootLabel; c = $ftpFiles.Count; ok = $true }
                Write-Host "[OK] Quet $($sf.name) qua FTP ($rootLabel): $($ftpFiles.Count) file"
            } else {
                $ScannedFolders += [PSCustomObject]@{ k = "$($sf.name) (UNC va FTP deu loi: $($ftpRes.Error))"; p = "$($sf.path) | ftp://$ftpHost$($sf.ftp_path)"; c = 0; ok = $false }
                Write-Host "[CANH BAO] Khong vao duoc $($sf.name): UNC bi tu choi, FTP loi: $($ftpRes.Error)"
            }
            continue
        }
        $ScannedFolders += [PSCustomObject]@{ k = $sf.name; p = $sf.path; c = 0; ok = $false }
        continue
    }
    $sfNow = Get-Date
    $sfFiles = @(Get-ChildItem -Path $sf.path -Filter $sfPattern -File -Recurse:([bool]$sf.recurse) -ErrorAction SilentlyContinue |
                 Where-Object { $_.FullName -notmatch '\\BAK\\' } |
                 Where-Object { -not $sf.max_age_days -or $_.LastWriteTime -ge $sfNow.AddDays(-[double]$sf.max_age_days) })
    foreach ($f in $sfFiles) { Add-StagingFile -File $f -Kind $sf.name -RootFolder $sf.path -MaxStuck $sfStuck -Now $sfNow -IsError ([bool]$sf.is_error) }
    $ScannedFolders += [PSCustomObject]@{ k = $sf.name; p = $sf.path; c = $sfFiles.Count; ok = $true }
}

# ------------------------------------------------------------
# BUOC D: Bao cao tong billing 5 ngay gan nhat (dem trong PRODATA theo ngay, ca 2 server)
#  Mac dinh: \\vnsgneivap05p\PRODATA\yyyyMMdd va \\vnsgneivap06p\PRODATA\yyyyMMdd (quet ca thu muc con, ca BAK).
#  Doi duong dan / so ngay trong config.json neu can:
#   "daily_report": { "days": 5, "folders": [ { "name": "05P", "path_template": "\\\\vnsgneivap05p\\PRODATA\\{date}",
#                     "ftp_server": "05P", "ftp_path_template": "/PRODATA/{date}" }, ... ] }
#  Neu UNC bi tu choi, tu dong doc qua FTP (WinSCP); mac dinh dung duong dan remote_folder_template cua sap_to_pp.
#  Dem theo SO BILLING DUY NHAT (1 billing trigger nhieu lan chi tinh 1).
# ------------------------------------------------------------
$DailyReportDays = if ($Config.daily_report -and $Config.daily_report.days) { [int]$Config.daily_report.days } else { 5 }
$DailyReportFolders = if ($Config.daily_report -and $Config.daily_report.folders) { @($Config.daily_report.folders) } else {
    @(
        [PSCustomObject]@{ name = "05P"; path_template = '\\vnsgneivap05p\PRODATA\{date}'; ftp_server = "05P"; ftp_path_template = $Config.sap_to_pp.sftp.remote_folder_template }
        [PSCustomObject]@{ name = "06P"; path_template = '\\vnsgneivap06p\PRODATA\{date}'; ftp_server = "06P"; ftp_path_template = $Config.sap_to_pp.sftp.remote_folder_template }
    )
}
$DailyReport = @()
$baseDay = [datetime]::ParseExact($CheckDate, "yyyyMMdd", $null)
for ($i = 0; $i -lt $DailyReportDays; $i++) {
    $day = $baseDay.AddDays(-$i)
    $dayStr = $day.ToString("yyyyMMdd")
    $dayBn = @{}          # billing -> loai chung tu (dedupe ca 2 server)
    $dayNoBn = 0          # file khong doc duoc so billing (moi file tinh 1)
    $dayFiles = 0
    $perServer = [ordered]@{}
    foreach ($df in $DailyReportFolders) {
        $dPath = $df.path_template -replace '\{date\}', $dayStr
        $srvBn = @{}; $srvNoBn = 0
        if (Test-Path $dPath -ErrorAction SilentlyContinue) {
            $dfFiles = @(Get-ChildItem -Path $dPath -Filter "*.txt" -File -Recurse -ErrorAction SilentlyContinue)
        } else {
            # UNC khong vao duoc -> thu qua FTP (WinSCP), mac dinh dung thu muc SAP->PP theo ngay (remote_folder_template)
            $ftpHost = if ($df.ftp_server) { Get-FtpServerHost -ServerKey $df.ftp_server } else { $null }
            if (-not $ftpHost -or -not $df.ftp_path_template) { $perServer[$df.name] = $null; continue }
            $ftpRes = Get-FtpFilesRecursive -ServerHost $ftpHost -RemotePath ($df.ftp_path_template -replace '\{date\}', $dayStr) -MaxDepth 2
            if (-not $ftpRes.Ok) { $perServer[$df.name] = $null; continue }
            $dfFiles = @($ftpRes.Files | Where-Object { $_.Name -like "*.txt" })
        }
        foreach ($f in $dfFiles) {
            $dayFiles++
            $bn = Get-BillingNumber -FileName $f.Name
            if ($bn) {
                $srvBn[$bn] = $true
                if (-not $dayBn.ContainsKey($bn)) { $dayBn[$bn] = Get-DocType -FileName $f.Name }
            } else { $srvNoBn++; $dayNoBn++ }
        }
        $perServer[$df.name] = $srvBn.Count + $srvNoBn
    }
    $typeCounts = @{}
    foreach ($t in $dayBn.Values) { if ($typeCounts.ContainsKey($t)) { $typeCounts[$t]++ } else { $typeCounts[$t] = 1 } }
    $DailyReport += [PSCustomObject]@{
        Date = $day; Total = $dayBn.Count + $dayNoBn; Files = $dayFiles; PerServer = $perServer; Types = $typeCounts
        Accessible = (@($perServer.Values | Where-Object { $null -ne $_ }).Count -gt 0)
    }
}

# Ve bang HTML cho bao cao 5 ngay
$drMax = ($DailyReport | Measure-Object -Property Total -Maximum).Maximum
if (-not $drMax) { $drMax = 1 }
$drSum = ($DailyReport | Measure-Object -Property Total -Sum).Sum
$drAccessibleDays = @($DailyReport | Where-Object { $_.Accessible })
$drAvg = if ($drAccessibleDays.Count) { [math]::Round((($drAccessibleDays | Measure-Object -Property Total -Sum).Sum) / $drAccessibleDays.Count) } else { 0 }
$drHeadSrv = ($DailyReportFolders | ForEach-Object { "<th style='text-align:right'>$($_.name)</th>" }) -join ""
$drRows = ""
$drWeekday = @("CN", "T2", "T3", "T4", "T5", "T6", "T7")
foreach ($d in $DailyReport) {
    $isToday = ($d.Date.ToString("yyyyMMdd") -eq $CheckDate)
    $srvCells = ""
    foreach ($k in $d.PerServer.Keys) {
        $v = $d.PerServer[$k]
        $srvCells += if ($null -eq $v) { "<td style='text-align:right;color:#cf222e' title='Khong truy cap duoc / khong co thu muc ngay nay'>--</td>" } else { "<td style='text-align:right'>$v</td>" }
    }
    $ban = 0 + $d.Types["Billing Ban"]; $small = 0 + $d.Types["Billing Smallunit"]
    $other = $d.Total - $ban - $small
    $pct = [math]::Round(100 * $d.Total / $drMax)
    $label = "$($drWeekday[[int]$d.Date.DayOfWeek]) $($d.Date.ToString('dd/MM/yyyy'))" + $(if ($isToday) { " <span class='dr-today'>hom nay</span>" } else { "" })
    $bar = "<div class='dr-bar'><div class='dr-bar-fill' style='width:$pct%'></div></div>"
    $drRows += "<tr$(if ($isToday) {" class='dr-row-today'"})><td style='white-space:nowrap'>$label</td>$srvCells<td style='text-align:right'>$(if ($d.Accessible) { $ban } else { '--' })</td><td style='text-align:right'>$(if ($d.Accessible) { $small } else { '--' })</td><td style='text-align:right'>$(if ($d.Accessible) { $other } else { '--' })</td><td style='text-align:right;font-weight:800;color:#0969da'>$(if ($d.Accessible) { $d.Total } else { '--' })</td><td style='width:30%'>$bar</td></tr>`n"
}
$drFootSrv = ""
foreach ($df in $DailyReportFolders) {
    $sum = 0; foreach ($d in $DailyReport) { $v = $d.PerServer[$df.name]; if ($null -ne $v) { $sum += $v } }
    $drFootSrv += "<td style='text-align:right'>$sum</td>"
}
$drBanSum = 0; $drSmallSum = 0
foreach ($d in $DailyReport) { $drBanSum += 0 + $d.Types["Billing Ban"]; $drSmallSum += 0 + $d.Types["Billing Smallunit"] }
$DailyReportHtml = @"
<table class="dr-table">
  <thead><tr><th>Ngay</th>$drHeadSrv<th style='text-align:right'>Billing Ban</th><th style='text-align:right'>Smallunit</th><th style='text-align:right'>Loai khac</th><th style='text-align:right'>Tong billing</th><th></th></tr></thead>
  <tbody>
$drRows
  </tbody>
  <tfoot><tr><td>Tong $DailyReportDays ngay</td>$drFootSrv<td style='text-align:right'>$drBanSum</td><td style='text-align:right'>$drSmallSum</td><td style='text-align:right'>$($drSum - $drBanSum - $drSmallSum)</td><td style='text-align:right;color:#0969da'>$drSum</td><td style='color:#57606a;font-weight:600'>TB: $drAvg / ngay</td></tr></tfoot>
</table>
"@
$DailyReportPaths = ($DailyReportFolders | ForEach-Object { $_.path_template }) -join " &nbsp;|&nbsp; "

# ------------------------------------------------------------
# BUOC C: VNPT response + callback ve SAP (chi chay neu da dien duong dan)
# ------------------------------------------------------------
$VnptLogFolder = $Config.vnpt.response_log_folder
if ([string]::IsNullOrWhiteSpace($VnptLogFolder)) {
    Add-Result -Stage "VNPT Response" -Name "Cau hinh" -Severity "WARNING" `
        -Message "Chua dien 'response_log_folder' trong config.json -- bo qua check nay."
} elseif (-not (Test-Path $VnptLogFolder)) {
    Add-Result -Stage "VNPT Response" -Name "Truy cap thu muc" -Severity "ERROR" `
        -Message "Khong truy cap duoc thu muc '$VnptLogFolder'."
} else {
    $since = (Get-Date).AddHours(-24)
    $jsonFiles = Get-ChildItem -Path $VnptLogFolder -Filter "*.json" -File -ErrorAction SilentlyContinue |
                 Where-Object { $_.LastWriteTime -ge $since }
    if ($jsonFiles.Count -eq 0) {
        Add-Result -Stage "VNPT Response" -Name "Response trong 24h" -Severity "WARNING" `
            -Message "Khong co file response nao tu VNPT trong 24h qua."
    } else {
        $errorCount = 0
        foreach ($f in $jsonFiles) {
            try {
                $data = Get-Content $f.FullName -Raw | ConvertFrom-Json
                $status = "$($data.status)$($data.errorCode)$($data.code)"
                if ($status -match "01|02|99") { $errorCount++ }
            } catch { $errorCount++ }
        }
        $VnptSuccessCount = $jsonFiles.Count - $errorCount
        if ($errorCount -gt 0) {
            Add-Result -Stage "VNPT Response" -Name "Loi tu VNPT" -Severity "ERROR" `
                -Message "$errorCount/$($jsonFiles.Count) response bi loi trong 24h qua."
        } else {
            Add-Result -Stage "VNPT Response" -Name "Loi tu VNPT" -Severity "OK" `
                -Message "Tat ca $($jsonFiles.Count) response trong 24h deu thanh cong."
        }
    }
}

function Get-BakupFilesViaFtp {
    param($ServerHost, $ServerLabel, $RemoteFolder, $DateStr)
    $result = [PSCustomObject]@{ Success = $false; Files = @(); ErrorMessage = "" }
    $siteName = $Config.sap_to_pp.sftp.winscp_site_name
    $sftpUser = $Config.sap_to_pp.sftp.username
    $sftpPass = $env:SFTP_PASSWORD
    $mask = "$RemoteFolder/$DateStr*.csv"
    try {
        $tmpScript = [System.IO.Path]::GetTempFileName()
        if ($siteName) {
            @("option batch abort", "option confirm off", "open `"$siteName`"", "ls `"$mask`"", "exit") |
                Out-File -FilePath $tmpScript -Encoding ASCII
        } else {
            $encodedPass = [uri]::EscapeDataString($sftpPass)
            @("option batch abort", "option confirm off", "open ftp://${sftpUser}:${encodedPass}@${ServerHost}", "ls `"$mask`"", "exit") |
                Out-File -FilePath $tmpScript -Encoding ASCII
        }
        $output = & $WinSCPPathResolved /script="$tmpScript" /nointeractiveinput 2>&1
        Remove-Item $tmpScript -ErrorAction SilentlyContinue
        $outputText = $output -join "`n"
        if ($outputText -match "Authentication failed" -or $outputText -match "Access denied") {
            $result.ErrorMessage = "Sai username/password khi ket noi FTP toi $ServerHost."
            return $result
        }
        $fileLines = $output | Where-Object { $_ -match '\.csv\s*$' }
        $names = @()
        foreach ($line in $fileLines) {
            $parts = $line -split '\s+' | Where-Object { $_ -ne "" }
            if ($parts -and $parts.Count -gt 0) { $names += $parts[-1] }
        }
        $result.Success = $true
        $result.Files = $names
        return $result
    } catch {
        $result.ErrorMessage = $_.Exception.Message
        return $result
    }
}

$SeenIssuedBnGlobal = @{}
$SeenIssuedBnPerBu = @{}
$SeenIssuedBnPerServer = @{}
$MatchedTimeBn = @{}
$UnmatchedTimeReasons = @{}

function Add-BuIssuedCount {
    param($FileName, $BakupTime, $ServerLabel)
    $bn = Get-BillingNumber -FileName $FileName
    if ($bn) { $BakupBillingSet[$bn] = $true }

    $isNewGlobal = (-not $bn) -or (-not $SeenIssuedBnGlobal.ContainsKey($bn))
    if ($isNewGlobal) {
        if ($bn) { $SeenIssuedBnGlobal[$bn] = $true }
        $script:CallbackCount++
    }

    $dtForBn = if ($bn -and $BillingToDocType.ContainsKey($bn)) { $BillingToDocType[$bn] } else { $null }
    $countsForBuServer = (-not $dtForBn) -or ($dtForBn -notin $ExcludedFromBuServerCount)

    if ($countsForBuServer -and $bn -and $BillingToBu.ContainsKey($bn)) {
        $bu = $BillingToBu[$bn]
        if (-not $SeenIssuedBnPerBu.ContainsKey($bu)) { $SeenIssuedBnPerBu[$bu] = @{} }
        if (-not $SeenIssuedBnPerBu[$bu].ContainsKey($bn)) {
            $SeenIssuedBnPerBu[$bu][$bn] = $true
            if ($BuIssuedCounts.ContainsKey($bu)) { $BuIssuedCounts[$bu]++ } else { $BuIssuedCounts[$bu] = 1 }
        }
    }
    if ($countsForBuServer -and $ServerLabel) {
        # Dung server GOC (noi SAP xuat file that su) neu tim thay, thay vi server noi tim thay
        # file xac nhan Bakup/Arcdata -- vi 2 cai co the khac nhau (du lieu lan giua 2 server).
        $attribServer = if ($bn -and $BillingToServer.ContainsKey($bn)) { $BillingToServer[$bn] } else { $ServerLabel }
        if (-not $SeenIssuedBnPerServer.ContainsKey($attribServer)) { $SeenIssuedBnPerServer[$attribServer] = @{} }
        $isNewForServer = (-not $bn) -or (-not $SeenIssuedBnPerServer[$attribServer].ContainsKey($bn))
        if ($isNewForServer) {
            if ($bn) { $SeenIssuedBnPerServer[$attribServer][$bn] = $true }
            if ($ServerIssuedCounts.ContainsKey($attribServer)) { $ServerIssuedCounts[$attribServer]++ } else { $ServerIssuedCounts[$attribServer] = 1 }
        }
    }
    if ($bn) {
        $unmatchReason = $null
        if (-not $BakupTime) {
            $unmatchReason = "Khong doc duoc gio tu ten file Bakup"
        } elseif (-not $BillingToSapTime.ContainsKey($bn)) {
            $unmatchReason = "Khong doc duoc gio tu ten file SAP (hoac khong tim thay billing ben SAP)"
        } else {
            $diffMin = ($BakupTime - $BillingToSapTime[$bn]).TotalMinutes
            if ($diffMin -ge 0 -and $diffMin -lt 1440) {
                $script:ProcessingMinutes += $diffMin
                $script:MatchedTimeBn[$bn] = $true
            } elseif ($diffMin -lt 0) {
                $unmatchReason = "Gio Bakup som hon gio SAP (co the khac cap trigger)"
            } else {
                $unmatchReason = "Khoang cach thoi gian qua lon (hon 24 tieng)"
            }
        }
        if ($unmatchReason -and -not $script:MatchedTimeBn.ContainsKey($bn)) {
            $script:UnmatchedTimeReasons[$bn] = @{ Reason = $unmatchReason; FileName = $FileName; Server = $ServerLabel }
        }
    }
}

$SapCallbackFolder = $Config.vnpt.sap_callback_folder
$SapCallbackFtpFolder = $Config.vnpt.sap_callback_ftp_folder

if ($Config.pp_to_sap_unc.servers -and $Config.vnpt.sap_callback_ftp_enabled) {
    # Kiem tra qua o dia mang (UNC) -- nhanh va dang tin cay hon FTP cho thu muc lon.
    $dateStr = $CheckDate
    $CallbackCount = 0

    foreach ($server in $Config.pp_to_sap_unc.servers) {
        $serverLabel = $server.name
        $bakupPath = $server.bakup
        try {
            $uncOk = Test-Path $bakupPath -ErrorAction SilentlyContinue
            if (-not $uncOk) {
                # UNC bi tu choi hoac khong ton tai -- tu dong chuyen sang FTP cho rieng server nay
                $ftpServerEntry = $Config.sap_to_pp.sftp.servers | Where-Object { $_.name -eq $serverLabel } | Select-Object -First 1
                if (-not $ftpServerEntry -or -not $SapCallbackFtpFolder) {
                    Add-Result -Stage "SAP Callback" -Name "Truy cap Bakup ($serverLabel)" -Severity "ERROR" `
                        -Message "Khong truy cap duoc UNC '$bakupPath', va khong co cau hinh FTP du phong."
                    continue
                }
                $ftpResult = Get-BakupFilesViaFtp -ServerHost $ftpServerEntry.host -ServerLabel $serverLabel -RemoteFolder $SapCallbackFtpFolder -DateStr $dateStr
                if (-not $ftpResult.Success) {
                    Add-Result -Stage "SAP Callback" -Name "Truy cap Bakup ($serverLabel)" -Severity "ERROR" `
                        -Message "UNC bi tu choi quyen, thu qua FTP cung loi: $($ftpResult.ErrorMessage)"
                    continue
                }
                foreach ($fname in $ftpResult.Files) {
                    $bakupTime = Get-EmbeddedBakupTime -FileName $fname
                    if ($bakupTime -and $serverLabel -eq "VNSGNEIVAP06P") { $bakupTime = $bakupTime.AddHours(-1) }
                    $tVal = if ($bakupTime) { $bakupTime.ToString("yyyy-MM-ddTHH:mm:ss") } else { $null }
                    $AllBakupFiles += [PSCustomObject]@{ n = $fname; s = $serverLabel; t = $tVal }
                    Add-BuIssuedCount -FileName $fname -BakupTime $bakupTime -ServerLabel $serverLabel
                }
                if ($ftpResult.Files.Count -eq 0) {
                    Add-Result -Stage "SAP Callback" -Name "Xac nhan PP -> SAP ($serverLabel, qua FTP)" -Severity "WARNING" `
                        -Message "Chua thay file xac nhan nao cho ngay $dateStr tren $serverLabel (UNC khong co quyen, da thu qua FTP)."
                } else {
                    Add-Result -Stage "SAP Callback" -Name "Xac nhan PP -> SAP ($serverLabel, qua FTP)" -Severity "OK" `
                        -Message "$($ftpResult.Files.Count) file xac nhan cho ngay $dateStr tren $serverLabel (UNC khong co quyen, da tu dong dung FTP)."
                }
                continue
            }
            $files = @()
            $bakupFiles = Get-ChildItem -Path $bakupPath -Filter "$dateStr*.csv" -File -ErrorAction SilentlyContinue
            $files += $bakupFiles

            # Quet them Arcdata -- vi PP ghi xac nhan TRUC TIEP vao Arcdata truoc,
            # chi khi chay move_arcdata_to_bakup.bat thu cong thi moi sang Bakup.
            # Neu khong quet Arcdata, se bo sot rat nhieu xac nhan moi/chua duoc di chuyen.
            $arcdataPath = $server.arcdata
            if ($arcdataPath -and (Test-Path $arcdataPath -ErrorAction SilentlyContinue)) {
                $arcdataFiles = Get-ChildItem -Path $arcdataPath -Filter "$dateStr*.csv" -File -ErrorAction SilentlyContinue
                # Loai trung theo ten file (phong khi file da vua duoc move nhung van con thay o ca 2 noi do do tre UNC)
                $existingNames = @{}
                foreach ($f in $files) { $existingNames[$f.Name] = $true }
                foreach ($f in $arcdataFiles) {
                    if (-not $existingNames.ContainsKey($f.Name)) { $files += $f }
                }
            }

            foreach ($f in $files) {
                $bakupTime = Get-EmbeddedBakupTime -FileName $f.Name
                if ($bakupTime -and $serverLabel -eq "VNSGNEIVAP06P") { $bakupTime = $bakupTime.AddHours(-1) }
                $tVal = if ($bakupTime) { $bakupTime.ToString("yyyy-MM-ddTHH:mm:ss") } else { $f.LastWriteTime.ToString("yyyy-MM-ddTHH:mm:ss") }
                $AllBakupFiles += [PSCustomObject]@{ n = $f.Name; s = $serverLabel; t = $tVal }
                Add-BuIssuedCount -FileName $f.Name -BakupTime $bakupTime -ServerLabel $serverLabel
            }
            if ($files.Count -eq 0) {
                Add-Result -Stage "SAP Callback" -Name "Xac nhan PP -> SAP ($serverLabel)" -Severity "WARNING" `
                    -Message "Chua thay file xac nhan nao cho ngay $dateStr tren $serverLabel (da quet ca Arcdata va Bakup)."
            } else {
                Add-Result -Stage "SAP Callback" -Name "Xac nhan PP -> SAP ($serverLabel)" -Severity "OK" `
                    -Message "$($files.Count) file xac nhan (da ra hoa don) cho ngay $dateStr tren $serverLabel (gom ca Arcdata va Bakup)."
            }
        } catch {
            Add-Result -Stage "SAP Callback" -Name "Loi Bakup ($serverLabel)" -Severity "ERROR" `
                -Message "Loi khi kiem tra thu muc Bakup tren $serverLabel : $($_.Exception.Message)"
        }
    }
} elseif ([string]::IsNullOrWhiteSpace($SapCallbackFolder)) {
    if (-not [string]::IsNullOrWhiteSpace($SapCallbackFtpFolder)) {
        Add-Result -Stage "SAP Callback" -Name "Cau hinh" -Severity "WARNING" `
            -Message "Kiem tra Arcdata dang TAT (cham vi thu muc rat lon) -- chay run_check_issued.bat de kiem tra rieng."
    } else {
        Add-Result -Stage "SAP Callback" -Name "Cau hinh" -Severity "WARNING" `
            -Message "Chua dien 'sap_callback_folder' hoac 'sap_callback_ftp_folder' trong config.json -- bo qua check nay."
    }
} elseif (-not (Test-Path $SapCallbackFolder)) {
    Add-Result -Stage "SAP Callback" -Name "Truy cap thu muc" -Severity "ERROR" `
        -Message "Khong truy cap duoc thu muc '$SapCallbackFolder'."
} else {
    $files = Get-ChildItem -Path $SapCallbackFolder -File -ErrorAction SilentlyContinue
    $CallbackCount = $files.Count
    if ($files.Count -eq 0) {
        Add-Result -Stage "SAP Callback" -Name "Callback eSeri-eForm-eRunning" -Severity "WARNING" `
            -Message "Chua thay file callback nao tu PP ve SAP."
    } else {
        $latest = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
        $ageMinutes = ((Get-Date) - $latest.LastWriteTime).TotalMinutes
        $maxWait = $Config.vnpt.max_callback_wait_minutes
        if ($ageMinutes -gt $maxWait) {
            Add-Result -Stage "SAP Callback" -Name "Callback eSeri-eForm-eRunning" -Severity "ERROR" `
                -Message "Callback moi nhat '$($latest.Name)' da $([math]::Round($ageMinutes)) phut (nguong: $maxWait phut)."
        } else {
            Add-Result -Stage "SAP Callback" -Name "Callback eSeri-eForm-eRunning" -Severity "OK" `
                -Message "Callback moi nhat cach day $([math]::Round($ageMinutes)) phut."
        }
    }
}

# ------------------------------------------------------------
# XUAT DASHBOARD HTML
# ------------------------------------------------------------
$okCount = ($Results | Where-Object { $_.Severity -eq "OK" }).Count
$warnCount = ($Results | Where-Object { $_.Severity -eq "WARNING" }).Count
$errCount = ($Results | Where-Object { $_.Severity -eq "ERROR" }).Count

$rowsHtml = ""
foreach ($r in $Results) {
    $color = switch ($r.Severity) { "OK" {"#1a7f37"} "WARNING" {"#9a6700"} "ERROR" {"#cf222e"} }
    $bg = switch ($r.Severity) { "OK" {"#dafbe1"} "WARNING" {"#fff8c5"} "ERROR" {"#ffebe9"} }
    $rowsHtml += "<tr><td>$($r.Stage)</td><td>$($r.Name)</td><td><span style='background:$bg;color:$color;padding:2px 10px;border-radius:12px;font-weight:600;font-size:13px;'>$($r.Severity)</span></td><td>$($r.Message)</td><td style='color:#666;font-size:13px;'>$($r.Time.ToString('HH:mm:ss'))</td></tr>`n"
}

function Get-StageColors {
    param($StageName)
    $stageResults = $Results | Where-Object { $_.Stage -eq $StageName }
    if (-not $stageResults -or $stageResults.Count -eq 0) {
        return @{ fill = "#f6f8fa"; stroke = "#8c959f"; text = "#57606a" }
    }
    if ($stageResults | Where-Object { $_.Severity -eq "ERROR" }) {
        return @{ fill = "#ffebe9"; stroke = "#cf222e"; text = "#82231f" }
    }
    if ($stageResults | Where-Object { $_.Severity -eq "WARNING" }) {
        return @{ fill = "#fff8c5"; stroke = "#9a6700"; text = "#7d5700" }
    }
    return @{ fill = "#dafbe1"; stroke = "#1a7f37"; text = "#116329" }
}

$cSapPP = Get-StageColors -StageName "SAP->PP"
$cStaging = Get-StageColors -StageName "PP Processing"
$cVnpt = Get-StageColors -StageName "VNPT Response"
$cCallback = Get-StageColors -StageName "SAP Callback"

# Xac dinh so "da ra hoa don" tot nhat co the, uu tien VNPT response, du phong bang callback
$IssuedCount = $null
$IssuedSource = ""
if ($VnptSuccessCount -gt 0) {
    $IssuedCount = $VnptSuccessCount
    $IssuedSource = "theo VNPT response"
} elseif ($CallbackCount -gt 0) {
    $IssuedCount = $CallbackCount
    $IssuedSource = "theo callback ve SAP"
}
$IssuedDisplay = if ($null -ne $IssuedCount) { "$IssuedCount" } else { "--" }
$IssuedNote = if ($null -ne $IssuedCount) { "($IssuedSource, khong tinh Z3F2)" } else { "(chua co du lieu -- can dien response_log_folder hoac sap_callback_folder)" }

$IssuedPct = if ($null -ne $IssuedCount -and $IssuableCount -gt 0) { [math]::Round(($IssuedCount / $IssuableCount) * 100) } else { $null }

# Xay khoi o vuong trang thai service
$serviceTilesHtml = ""
foreach ($svcName in $ServiceStatusResults.Keys) {
    $st = $ServiceStatusResults[$svcName]
    $tileClass = switch ($st.Status) {
        "RUNNING" { "svc-tile-ok" }
        "STOPPED" { "svc-tile-error" }
        default { "svc-tile-unknown" }
    }
    $tileIcon = switch ($st.Status) {
        "RUNNING" { "&#9989;" }
        "STOPPED" { "&#128308;" }
        default { "&#10067;" }
    }
    $tileLabel = switch ($st.Status) {
        "RUNNING" { "DANG CHAY" }
        "STOPPED" { "DA DUNG" }
        default { "KHONG RO" }
    }
    $tileSub = if ($st.Status -eq "STOPPED") { "$($st.Count) file ton dong" } else { "Binh thuong" }
    $serviceTilesHtml += @"
    <div class="svc-tile $tileClass">
      <div class="svc-tile-icon">$tileIcon</div>
      <div class="svc-tile-name">$svcName</div>
      <div class="svc-tile-status">$tileLabel</div>
      <div class="svc-tile-sub">$tileSub</div>
    </div>
"@
}

# ------------------------------------------------------------
# Phan tich lich su service tu file service_history.csv:
# gop cac lan STOPPED lien tiep thanh tung "dot dung" (tu luc bat dau -> luc het,
# hoac "van dang dung" neu chua thay RUNNING lai sau do).
# ------------------------------------------------------------
$ServiceHistoryRowsHtml = ""
$ServiceStopIncidentCount = 0
if (Test-Path $ServiceHistoryPath -ErrorAction SilentlyContinue) {
    try {
        $histData = Import-Csv -Path $ServiceHistoryPath -Encoding UTF8
        $histByServer = $histData | Group-Object Server

        $allIncidents = @()
        foreach ($grp in $histByServer) {
            $rows = $grp.Group | Sort-Object { [datetime]$_.CheckTime }
            $curStart = $null
            $curMaxCount = 0
            for ($i = 0; $i -lt $rows.Count; $i++) {
                $r = $rows[$i]
                if ($r.Status -eq "STOPPED") {
                    if (-not $curStart) { $curStart = [datetime]$r.CheckTime }
                    if ([int]$r.StuckFileCount -gt $curMaxCount) { $curMaxCount = [int]$r.StuckFileCount }
                    $curEnd = [datetime]$r.CheckTime
                } else {
                    if ($curStart) {
                        $allIncidents += [PSCustomObject]@{ Server = $grp.Name; Start = $curStart; End = $curEnd; Ongoing = $false; MaxCount = $curMaxCount }
                        $curStart = $null; $curMaxCount = 0
                    }
                }
            }
            if ($curStart) {
                $allIncidents += [PSCustomObject]@{ Server = $grp.Name; Start = $curStart; End = $curEnd; Ongoing = $true; MaxCount = $curMaxCount }
            }
        }

        $allIncidents = $allIncidents | Sort-Object Start -Descending
        $ServiceStopIncidentCount = $allIncidents.Count

        foreach ($inc in ($allIncidents | Select-Object -First 30)) {
            $durMin = [math]::Round(($inc.End - $inc.Start).TotalMinutes, 1)
            $statusBadge = if ($inc.Ongoing) { "<span style='background:#ffebe9;color:#cf222e;font-size:10px;font-weight:700;padding:1px 8px;border-radius:8px;'>VAN DANG DUNG</span>" } else { "<span style='background:#f6f8fa;color:#57606a;font-size:10px;font-weight:700;padding:1px 8px;border-radius:8px;'>DA HET</span>" }
            $ServiceHistoryRowsHtml += @"
        <div style='padding:8px 0;border-bottom:1px solid #f1f3f5;display:flex;align-items:center;gap:10px;flex-wrap:wrap;'>
          <span style='font-weight:700;color:#24292f;min-width:130px;'>$($inc.Server)</span>
          <span style='font-size:12.5px;color:#57606a;'>$($inc.Start.ToString("dd/MM HH:mm")) &rarr; $($inc.End.ToString("dd/MM HH:mm"))</span>
          <span style='font-size:12px;color:#8250df;font-weight:600;'>($durMin phut)</span>
          <span style='font-size:11.5px;color:#8c959f;'>Toi da $($inc.MaxCount) file ton dong</span>
          $statusBadge
        </div>
"@
        }
    } catch {}
}
if ([string]::IsNullOrWhiteSpace($ServiceHistoryRowsHtml)) {
    $ServiceHistoryRowsHtml = "<div style='color:#8c959f;font-size:13px;'>Chua co du lieu lich su (hoac chua bao gio bi dung).</div>"
}


# Thoi gian xu ly trung binh (SAP -> Da ra hoa don), doi chieu qua so billing
$AvgProcessingMinutes = $null
if ($ProcessingMinutes.Count -gt 0) {
    $AvgProcessingMinutes = ($ProcessingMinutes | Measure-Object -Average).Average
}
$AvgProcessingDisplay = if ($null -ne $AvgProcessingMinutes) {
    $avgM = [math]::Floor($AvgProcessingMinutes)
    $avgS = [math]::Round(($AvgProcessingMinutes - $avgM) * 60)
    "$avgM ph $avgS s"
} else { "--" }
$AvgProcessingNote = if ($ProcessingMinutes.Count -gt 0) { "(tren $($ProcessingMinutes.Count) billing)" } else { "(chua co du lieu doi chieu)" }

$TotalProcessingDisplay = "--"
if ($ProcessingMinutes.Count -gt 0) {
    $totalMin = ($ProcessingMinutes | Measure-Object -Sum).Sum
    $th = [math]::Floor($totalMin / 60)
    $tm = [math]::Floor($totalMin % 60)
    $TotalProcessingDisplay = if ($th -gt 0) { "$th gio $tm ph" } else { "$tm ph" }
}

$UnmatchedIssuedCount = if ($null -ne $IssuedCount) { $IssuedCount - $ProcessingMinutes.Count } else { 0 }
if ($UnmatchedIssuedCount -lt 0) { $UnmatchedIssuedCount = 0 }

# Danh sach chi tiet cac billing DA RA HOA DON nhung KHONG tinh duoc thoi gian xu ly, kem ly do
$unmatchedTimeRowsHtml = ""
foreach ($bn in $UnmatchedTimeReasons.Keys) {
    $item = $UnmatchedTimeReasons[$bn]
    $unmatchedTimeRowsHtml += "<div style='padding:4px 0;border-bottom:1px solid #f1f3f5;'><span style='color:#bc4c00;font-weight:600;'>$($item.Reason)</span> &middot; $($item.Server) &middot; $($item.FileName)</div>`n"
}
if ([string]::IsNullOrWhiteSpace($unmatchedTimeRowsHtml)) {
    $unmatchedTimeRowsHtml = "<div style='color:#8c959f;font-size:13px;'>Khong co du lieu.</div>"
}

# Danh sach billing chua co xac nhan da ra hoa don (doi chieu qua so billing)
$PendingBillings = @()
if ($BakupBillingSet.Count -gt 0 -or $AllBakupFiles.Count -ge 0) {
    $seenBn = @{}
    foreach ($f in $AllSapFiles) {
        if ($f.bn -and -not $BakupBillingSet.ContainsKey($f.bn) -and -not $seenBn.ContainsKey($f.bn) -and $f.d -ne "Z3F2 (Khong xuat HD)") {
            $seenBn[$f.bn] = $true
            $PendingBillings += $f
        }
    }
}
$PendingTotal = $PendingBillings.Count

# So "con lai" chinh xac theo tung BU va Server -- tinh TRUC TIEP tu chinh tap billing
# da dem vao cnt (SeenBnPerBu / SeenBnPerServer), tru di nhung billing da co trong Bakup.
# Cach nay dam bao "Con lai" LUON khop voi "Tong - Da ra" hien tren dashboard.
function Get-PendingFilesForBillingSet {
    param($BillingNumberSet, $FilterField, $FilterValue)
    $pendingBns = @{}
    foreach ($bn in $BillingNumberSet.Keys) {
        if (-not $BakupBillingSet.ContainsKey($bn)) { $pendingBns[$bn] = $true }
    }
    $result = @()
    $seenForList = @{}
    # Vi $BillingNumberSet da chac chan thuoc dung server/BU nay roi (do la ly do no nam trong set),
    # chi can tim BAT KY file nao trong $AllSapFiles co cung so billing -- khong loc lai theo f.s/f.b nua,
    # tranh truong hop file dai dien dau tien lai bi gan nham server/BU khac do du lieu lan.
    foreach ($f in $AllSapFiles) {
        if (-not $f.bn) { continue }
        if (-not $pendingBns.ContainsKey($f.bn)) { continue }
        if ($seenForList.ContainsKey($f.bn)) { continue }
        $seenForList[$f.bn] = $true
        $result += $f
    }
    return $result
}

$BuPendingCount = @{}
$BuPendingDocType = @{}
$BuPendingFiles = @{}
foreach ($buKey in $SeenBnPerBu.Keys) {
    $files = Get-PendingFilesForBillingSet -BillingNumberSet $SeenBnPerBu[$buKey] -FilterField "b" -FilterValue $buKey
    if ($files.Count -gt 0) {
        $BuPendingCount[$buKey] = $files.Count
        $BuPendingFiles[$buKey] = $files
        $BuPendingDocType[$buKey] = @{}
        foreach ($f in $files) {
            $dt = if ($f.d) { $f.d } else { "Khac" }
            if ($BuPendingDocType[$buKey].ContainsKey($dt)) { $BuPendingDocType[$buKey][$dt]++ } else { $BuPendingDocType[$buKey][$dt] = 1 }
        }
    }
}

# Loc rieng billing ZF2 (Billing Ban) tu output ZVRD nhung chua ra hoa don -- can theo doi sat nhat
$PendingZF2 = $PendingBillings | Where-Object { $_.d -eq "Billing Ban" }

$ServerPendingDocType = @{}
$ServerPendingCount = @{}
$ServerPendingFiles = @{}
foreach ($srvKey in $SeenBnPerServer.Keys) {
    $files = Get-PendingFilesForBillingSet -BillingNumberSet $SeenBnPerServer[$srvKey] -FilterField "s" -FilterValue $srvKey
    if ($files.Count -gt 0) {
        $ServerPendingCount[$srvKey] = $files.Count
        $ServerPendingFiles[$srvKey] = $files
        $ServerPendingDocType[$srvKey] = @{}
        foreach ($f in $files) {
            $dt = if ($f.d) { $f.d } else { "Khac" }
            if ($ServerPendingDocType[$srvKey].ContainsKey($dt)) { $ServerPendingDocType[$srvKey][$dt]++ } else { $ServerPendingDocType[$srvKey][$dt] = 1 }
        }
    }
}
function Build-PendingFileListHtml {
    param($Files, $PendingBnSet)
    $rows = ""
    foreach ($f in $Files) {
        $isPending = $PendingBnSet -and $f.bn -and $PendingBnSet.ContainsKey($f.bn)
        if ($isPending) {
            $badge = " <span style='background:#ffebe9;color:#cf222e;font-size:9.5px;font-weight:700;padding:1px 6px;border-radius:8px;margin-left:4px;'>CHUA RA HD</span>"
            $rowStyle = "padding:3px 0;border-bottom:1px solid #f1f3f5;background:#fff8f6;"
        } elseif ($f.d -eq "Z3F2 (Khong xuat HD)") {
            $badge = " <span style='background:#f6f8fa;color:#57606a;font-size:9.5px;font-weight:700;padding:1px 6px;border-radius:8px;margin-left:4px;'>KHONG PHAT SINH HD</span>"
            $rowStyle = "padding:3px 0;border-bottom:1px solid #f1f3f5;"
        } elseif ($PendingBnSet) {
            # Neu co truyen PendingBnSet vao (nghia la danh sach nay co the phan biet da/chua ra),
            # nhung dong khong nam trong tap chua ra -- gan nhan da ra hoa don ro rang.
            $badge = " <span style='background:#dafbe1;color:#1a7f37;font-size:9.5px;font-weight:700;padding:1px 6px;border-radius:8px;margin-left:4px;'>DA RA HOA DON</span>"
            $rowStyle = "padding:3px 0;border-bottom:1px solid #f1f3f5;"
        } else {
            $badge = ""
            $rowStyle = "padding:3px 0;border-bottom:1px solid #f1f3f5;"
        }
        $rows += "<div style='$rowStyle'><span style='color:#0969da;font-weight:600;'>$($f.d)</span> &middot; $($f.n)$badge</div>`n"
    }
    return $rows
}

# Danh sach chi tiet cac billing thuoc loai KHONG duoc tinh (PXK/Return/D.Chinh/Cancel/Z3F2)
$ServerExcludedFiles = @{}
$BuExcludedFiles = @{}
foreach ($f in $AllSapFiles) {
    if (-not $f.d -or ($f.d -notin $ExcludedFromBuServerCount)) { continue }
    if (-not $f.bn) { continue }
    if ($f.s) {
        if (-not $ServerExcludedFiles.ContainsKey($f.s)) { $ServerExcludedFiles[$f.s] = @{} }
        if (-not $ServerExcludedFiles[$f.s].ContainsKey($f.bn)) { $ServerExcludedFiles[$f.s][$f.bn] = $f }
    }
    $bl = if ($f.b) { $f.b } else { "Khac" }
    if (-not $BuExcludedFiles.ContainsKey($bl)) { $BuExcludedFiles[$bl] = @{} }
    if (-not $BuExcludedFiles[$bl].ContainsKey($f.bn)) { $BuExcludedFiles[$bl][$f.bn] = $f }
}

# So "con lai" rieng cho cac loai KHONG duoc tinh (PXK/Return/D.Chinh/Cancel) -- de tham khao them
$ServerExcludedPending = @{}
$BuExcludedPending = @{}
$ServerExcludedPendingFiles = @{}
$BuExcludedPendingFiles = @{}
foreach ($f in $PendingBillings) {
    if (-not $f.d -or ($f.d -notin $ExcludedFromBuServerCount)) { continue }
    if ($f.s) {
        if (-not $ServerExcludedPending.ContainsKey($f.s)) { $ServerExcludedPending[$f.s] = @{} }
        if ($ServerExcludedPending[$f.s].ContainsKey($f.d)) { $ServerExcludedPending[$f.s][$f.d]++ } else { $ServerExcludedPending[$f.s][$f.d] = 1 }
        if (-not $ServerExcludedPendingFiles.ContainsKey($f.s)) { $ServerExcludedPendingFiles[$f.s] = @() }
        $ServerExcludedPendingFiles[$f.s] += $f
    }
    $bl = if ($f.b) { $f.b } else { "Khac" }
    if (-not $BuExcludedPending.ContainsKey($bl)) { $BuExcludedPending[$bl] = @{} }
    if ($BuExcludedPending[$bl].ContainsKey($f.d)) { $BuExcludedPending[$bl][$f.d]++ } else { $BuExcludedPending[$bl][$f.d] = 1 }
    if (-not $BuExcludedPendingFiles.ContainsKey($bl)) { $BuExcludedPendingFiles[$bl] = @() }
    $BuExcludedPendingFiles[$bl] += $f
}
# Xay tap hop so billing (bn) chua ra hoa don, de danh dau trong danh sach gop chung
$ServerExcludedPendingBnSet = @{}
foreach ($srvKey in $ServerExcludedPendingFiles.Keys) {
    $ServerExcludedPendingBnSet[$srvKey] = @{}
    foreach ($f in $ServerExcludedPendingFiles[$srvKey]) { if ($f.bn) { $ServerExcludedPendingBnSet[$srvKey][$f.bn] = $true } }
}
$BuExcludedPendingBnSet = @{}
foreach ($buKey in $BuExcludedPendingFiles.Keys) {
    $BuExcludedPendingBnSet[$buKey] = @{}
    foreach ($f in $BuExcludedPendingFiles[$buKey]) { if ($f.bn) { $BuExcludedPendingBnSet[$buKey][$f.bn] = $true } }
}
function Build-ExcludedSummaryWithPending {
    param($TotalCounts, $PendingCounts)
    $parts = @()
    foreach ($key in @("PXK", "Billing Return", "Billing Dieu Chinh", "Cancel Bill", "Z3F2 (Khong xuat HD)")) {
        if ($TotalCounts.ContainsKey($key) -and $TotalCounts[$key] -gt 0) {
            $shortLabel = @{ "PXK"="PXK"; "Billing Return"="Return"; "Billing Dieu Chinh"="D.Chinh"; "Cancel Bill"="Cancel"; "Z3F2 (Khong xuat HD)"="Z3F2" }
            $total = $TotalCounts[$key]
            if ($key -eq "Z3F2 (Khong xuat HD)") {
                $pendText = " (khong phat sinh HD)"
            } else {
                $pend = if ($PendingCounts -and $PendingCounts.ContainsKey($key)) { $PendingCounts[$key] } else { 0 }
                $pendText = if ($pend -gt 0) { " (con lai: $pend)" } else { " (da du lieu Bakup)" }
            }
            $parts += "$($shortLabel[$key]): $total$pendText"
        }
    }
    return ($parts -join " &middot; ")
}
$PendingZF2Total = $PendingZF2.Count
$pendingZF2RowsHtml = ""
foreach ($f in $PendingZF2) {
    $pendingZF2RowsHtml += "<div class='ls-row'><span class='ls-badge ls-found'>$($f.s)$(if($f.b){' - '+$f.b})</span><span class='ls-name'>$($f.n)</span></div>`n"
}
if ($PendingZF2Total -eq 0) {
    $pendingZF2RowsHtml = "<div style='color:#8c959f;font-size:13px;'>Khong co billing ZF2 nao dang cho -- tat ca da ra hoa don.</div>"
}

# Chi tiet theo loai chung tu cho tung the (SAP tong va Con lai chua ra hoa don)
$PendingDocTypeCounts = @{}
foreach ($f in $PendingBillings) {
    $dt = if ($f.d) { $f.d } else { "Khac" }
    if ($PendingDocTypeCounts.ContainsKey($dt)) { $PendingDocTypeCounts[$dt]++ } else { $PendingDocTypeCounts[$dt] = 1 }
}
function Build-DocTypeSummary {
    param($Counts)
    $shortLabel = @{
        "Billing Ban" = "Ban"; "Billing Dieu Chinh" = "D.Chinh"; "Billing Return" = "Return"
        "Billing Smallunit" = "SmallU"; "PXK" = "PXK"; "Z3F2 (Khong xuat HD)" = "Z3F2"; "Cancel Bill" = "Cancel"; "Khac" = "Khac"
    }
    $parts = @()
    foreach ($key in @("Billing Ban","Billing Dieu Chinh","Billing Return","Billing Smallunit","PXK","Z3F2 (Khong xuat HD)","Cancel Bill","Khac")) {
        if ($Counts.ContainsKey($key) -and $Counts[$key] -gt 0) {
            $parts += "$($shortLabel[$key]): $($Counts[$key])"
        }
    }
    return ($parts -join " &middot; ")
}
$SapDocTypeSummary = Build-DocTypeSummary -Counts $DocTypeCounts
$PendingDocTypeSummary = Build-DocTypeSummary -Counts $PendingDocTypeCounts

$pendingRowsHtml = ""
foreach ($f in $PendingBillings) {
    $pendingRowsHtml += "<div class='ls-row'><span class='ls-badge ls-found'>$($f.s)$(if($f.b){' - '+$f.b})</span>$(if($f.d){"<span class='ls-badge ls-doctype'>$($f.d)</span>"})<span class='ls-name'>$($f.n)</span></div>`n"
}
if ($PendingTotal -eq 0) {
    $pendingRowsHtml = "<div style='color:#8c959f;font-size:13px;'>Khong co billing nao dang cho -- tat ca da co xac nhan (hoac chua co du lieu Bakup).</div>"
}

$ringCirc = 163.4
$ringDash = if ($null -ne $IssuedPct) { [math]::Round($ringCirc * $IssuedPct / 100, 1) } else { 0 }
$ringSvg = if ($null -ne $IssuedPct) { @"
<svg width="34" height="34" viewBox="0 0 60 60" style="position:absolute;right:10px;top:10px;">
<circle cx="30" cy="30" r="26" fill="none" stroke="#eef1f5" stroke-width="6"/>
<circle cx="30" cy="30" r="26" fill="none" stroke="$IssuedColor" stroke-width="6" stroke-linecap="round" stroke-dasharray="$ringDash $ringCirc" transform="rotate(-90 30 30)"/>
<text x="30" y="35" text-anchor="middle" font-size="15" font-weight="700" fill="$IssuedColor">$IssuedPct%</text>
</svg>
"@ } else { "" }

$BillingBanCount = if ($DocTypeCounts.ContainsKey("Billing Ban")) { $DocTypeCounts["Billing Ban"] } else { 0 }
$Z3F2Count = if ($DocTypeCounts.ContainsKey("Z3F2 (Khong xuat HD)")) { $DocTypeCounts["Z3F2 (Khong xuat HD)"] } else { 0 }

# Mau cho o "Da ra hoa don" dua theo ty le so voi tong billing hom nay
# (loai Z3F2 ra khoi mau so vi loai nay khong phat sinh hoa don)
$IssuableCount = $SapBillingCount - $Z3F2Count
$IssuedColor = "#8c959f"   # xam mac dinh khi chua co du lieu
$IssuedBg = "#f6f8fa"
if ($null -ne $IssuedCount -and $IssuableCount -gt 0) {
    $issuedRatio = $IssuedCount / $IssuableCount
    if ($issuedRatio -ge 0.9) {
        $IssuedColor = "#1a7f37"; $IssuedBg = "#dafbe1"   # xanh -- gan du
    } elseif ($issuedRatio -ge 0.5) {
        $IssuedColor = "#9a6700"; $IssuedBg = "#fff8c5"   # vang -- thieu vua
    } else {
        $IssuedColor = "#cf222e"; $IssuedBg = "#ffebe9"   # do -- thieu nhieu
    }
}

# Xay bang thong ke theo BU
$buRowsHtml = ""
$buTotal = ($BuCounts.Values | Measure-Object -Sum).Sum
$buOrder = @("HEC", "CG", "PM", "CDV & TEC", "ECOM", "Khac")
foreach ($buName in $buOrder) {
    if ($BuCounts.ContainsKey($buName)) {
        $cnt = $BuCounts[$buName]
        $pct = if ($buTotal -gt 0) { [math]::Round(($cnt / $buTotal) * 100, 1) } else { 0 }
        $buIssuedCnt = if ($BuIssuedCounts.ContainsKey($buName)) { $BuIssuedCounts[$buName] } else { 0 }
        $buIssuedPctOfBar = if ($cnt -gt 0) { if ($buIssuedCnt -ge $cnt) { 100 } else { [math]::Round(($buIssuedCnt / $cnt) * 100, 2) } } else { 0 }
        $buDataAnomaly = $buIssuedCnt - $cnt
        $buComplete = ($cnt -gt 0 -and $buIssuedCnt -eq $cnt)
        $buNameHtml = if ($buComplete) { "$buName <span class='done-badge'>&#10003; 100%</span>" } else { $buName }
        $buSubColor = if ($buComplete) { "#0969da" } else { "#1a7f37" }
        $buPendingCnt = if ($BuPendingCount.ContainsKey($buName)) { $BuPendingCount[$buName] } else { 0 }
        $buMathGap = $cnt - $buIssuedCnt
        if ($buMathGap -gt $buPendingCnt) { $buPendingCnt = $buMathGap }
        $buPendingBreakdown = if ($BuPendingDocType.ContainsKey($buName)) { Build-DocTypeSummary -Counts $BuPendingDocType[$buName] } else { "" }
        $buRowsHtml += @"
        <div class="bu-row$(if ($buComplete) {' bu-row-complete'})">
          <div class="bu-name">$buNameHtml</div>
          <div class="bu-bar-track">
            <div class="bu-bar-fill" style="width:$pct%;">
              <div class="bu-bar-issued" style="width:$buIssuedPctOfBar%;$(if ($buComplete) {'background:linear-gradient(90deg,#0969da,#54aeff);'})" title="Da ra hoa don: $buIssuedCnt/$cnt"></div>
            </div>
          </div>
          <div class="bu-count">$cnt <span class="bu-pct">($pct%)</span></div>
        </div>
        <div class="bu-row bu-subrow">
          <div class="bu-name"></div>
          <div style="flex:1;font-size:14px;color:$buSubColor;font-weight:800;">&#9632; Da ra hoa don: <span style="font-size:16px;">$buIssuedCnt / $cnt</span> <span style="font-size:15px;">($buIssuedPctOfBar%)</span>$(if ($buPendingCnt -gt 0) {" &middot; <span style='color:#cf222e;'>Con lai: $buPendingCnt</span>"})$(if ($buDataAnomaly -gt 0) {"<br><span style='font-size:10.5px;color:#bc4c00;font-weight:400;'>&#9888;&#65039; Du lieu bat thuong: co $buDataAnomaly billing duoc tinh 'da ra' nhung khong nam trong tong cua BU nay</span>"})</div>
          <div style="width:90px;"></div>
        </div>
"@
        if ($buPendingBreakdown) {
            $buRowsHtml += @"
        <div class="bu-row bu-subrow">
          <div class="bu-name"></div>
          <div style="flex:1;font-size:10.5px;color:#8c959f;">Con lai theo loai: $buPendingBreakdown</div>
          <div style="width:90px;"></div>
        </div>
"@
        }
        $buExcludedSummary = if ($BuExcludedDocType.ContainsKey($buName)) { Build-ExcludedSummaryWithPending -TotalCounts $BuExcludedDocType[$buName] -PendingCounts $BuExcludedPending[$buName] } else { "" }
        if ($buExcludedSummary) {
            $buExclCount = if ($BuExcludedFiles.ContainsKey($buName)) { $BuExcludedFiles[$buName].Count } else { 0 }
            $buRowsHtml += @"
        <div class="bu-row bu-subrow">
          <div class="bu-name"></div>
          <div style="flex:1;font-size:10.5px;color:#bc4c00;">Khong tinh (PXK/Return/D.Chinh/Cancel/Z3F2): $buExcludedSummary</div>
          <div style="width:90px;"></div>
        </div>
"@
            if ($buExclCount -gt 0) {
                $buExclPendingBnSet = if ($BuExcludedPendingBnSet.ContainsKey($buName)) { $BuExcludedPendingBnSet[$buName] } else { $null }
                $buExclListHtml = Build-PendingFileListHtml -Files $BuExcludedFiles[$buName].Values -PendingBnSet $buExclPendingBnSet
                $buExclPendingCnt = if ($BuExcludedPendingFiles.ContainsKey($buName)) { $BuExcludedPendingFiles[$buName].Count } else { 0 }
                $buExclSumLabel = if ($buExclPendingCnt -gt 0) { "Xem danh sach $buExclCount billing khong tinh (co $buExclPendingCnt CHUA ra hoa don)" } else { "Xem danh sach $buExclCount billing khong tinh" }
                $buRowsHtml += @"
        <div class="bu-row bu-subrow">
          <div class="bu-name"></div>
          <div style="flex:1;">
            <details>
              <summary style="cursor:pointer;font-size:11px;color:#bc4c00;font-weight:600;">$buExclSumLabel</summary>
              <div style="max-height:260px;overflow-y:auto;margin-top:6px;font-size:11px;font-family:monospace;">$buExclListHtml</div>
            </details>
          </div>
          <div style="width:90px;"></div>
        </div>
"@
            }
        }
        $buKnownPendingFiles = if ($BuPendingFiles.ContainsKey($buName)) { $BuPendingFiles[$buName] } else { @() }
        $buUntrackable = if ($BuUntrackableFiles.ContainsKey($buName)) { $BuUntrackableFiles[$buName] } else { @() }
        $buGapUnexplained = $buPendingCnt - $buKnownPendingFiles.Count
        if ($buPendingCnt -gt 0) {
            $buPendingListHtml = Build-PendingFileListHtml -Files $buKnownPendingFiles
            if ($buGapUnexplained -gt 0 -and $buUntrackable.Count -gt 0) {
                foreach ($uf in ($buUntrackable | Select-Object -First $buGapUnexplained)) {
                    $buPendingListHtml += "<div style='padding:3px 0;border-bottom:1px solid #f1f3f5;background:#fff8f6;'><span style='color:#8c959f;font-weight:600;'>$($uf.d)</span> &middot; $($uf.n) <span style='background:#ffebe9;color:#cf222e;font-size:9.5px;font-weight:700;padding:1px 6px;border-radius:8px;margin-left:4px;'>KHONG XAC DINH SO BILLING</span></div>`n"
                }
            }
            $buRowsHtml += @"
        <div class="bu-row bu-subrow">
          <div class="bu-name"></div>
          <div style="flex:1;">
            <details>
              <summary style="cursor:pointer;font-size:11px;color:#cf222e;font-weight:600;">Xem danh sach $buPendingCnt billing con lai</summary>
              <div style="max-height:220px;overflow-y:auto;margin-top:6px;font-size:11px;font-family:monospace;">$buPendingListHtml</div>
            </details>
          </div>
          <div style="width:90px;"></div>
        </div>
"@
        }
        if ($buName -eq "Khac" -and $UnmappedBuCodes.Count -gt 0) {
            $unmappedList = ($UnmappedBuCodes.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "$($_.Key): $($_.Value)" }) -join ", "
            $khacExampleNames = ($AllSapFiles | Where-Object { $_.b -eq "Khac" } | Select-Object -First 8 -ExpandProperty n) -join "<br>"
            $buRowsHtml += @"
        <div class="bu-row bu-subrow">
          <div class="bu-name"></div>
          <div style="flex:1;font-size:11.5px;color:#57606a;">Chi tiet ma chua gan BU: $unmappedList</div>
          <div style="width:90px;"></div>
        </div>
        <div class="bu-row bu-subrow">
          <div class="bu-name"></div>
          <div style="flex:1;font-size:11px;color:#8c959f;font-family:monospace;line-height:1.6;">Vi du ten file: $khacExampleNames</div>
          <div style="width:90px;"></div>
        </div>
"@
        }
    }
}
if ([string]::IsNullOrWhiteSpace($buRowsHtml)) {
    $buRowsHtml = "<div style='color:#8c959f;font-size:13px;'>Chua co du lieu BU.</div>"
}

# Xay bang thong ke theo loai chung tu
$docTypeRowsHtml = ""
$docTypeTotal = ($DocTypeCounts.Values | Measure-Object -Sum).Sum
$docTypeOrder = @("Billing Ban", "Billing Dieu Chinh", "Billing Return", "Billing Smallunit", "PXK", "Z3F2 (Khong xuat HD)", "Cancel Bill", "Khac")

# Bang tra ma loi VNPT (tham khao)
$VnptErrCodes = [ordered]@{
    "ERR:30" = "true"
    "ERR:51" = "Chung thu so bi thu hoi."
    "ERR:52" = "San luong hoa don con lai (SoLuongConLai) khong du de tiep tuc thuc hien phat hanh hoa don"
    "ERR:54" = "Loi, khong phat hanh duoc voi ngay hoa don lon hon ngay hien tai"
    "ERR:58" = "Loi, khong phat hanh duoc voi ngay dich vu nam khac dai so nam (serialYear)"
    "ERR:29" = "Chung thu qua han."
    "ERR:2901" = "Khong tim thay keystore comId"
    "ERR:2902" = "certificate null"
    "ERR:1501" = "Ngay lap hoa don khong duoc lon hon ngay hien tai"
    "ERR:1502" = "TCHDon = 5 bat buoc nhap SBKe va NBKe"
    "ERR:1503" = "Ty gia bat buoc phai la VND (CurrencyUnit)"
    "ERR:1504" = "MDVQHNSach khong dung quy dinh"
    "ERR:1505" = "MDVQHNSach co du lieu, bat buoc nhap Ten va DChi"
    "ERR:1506" = "MST co du lieu, bat buoc nhap Ten va DChi"
    "ERR:1507" = "Email (EmailDeliver) khong dung dinh dang"
    "ERR:1508" = "HDCTTChinh = 1 chi cho phep cung 1 loai TSuat = -5 hoac CTTC"
    "ERR:1509" = "vatrate == -99 vatrate khong dung dinh dang."
    "ERR:1510" = "LHHDTrung khong duoc de trong"
    "ERR:1511" = "SKhung khong duoc de trong khi LHHDTrung = 1"
    "ERR:1512" = "SMay khong duoc de trong khi LHHDTrung = 1"
    "ERR:1513" = "BKSPTVChuyen khong duoc de trong khi LHHDTrung = 2"
    "ERR:1514" = "DCNGHang khong duoc de trong khi LHHDTrung = 3"
    "ERR:1515" = "TNGHang khong duoc de trong khi LHHDTrung = 3"
    "ERR:1516" = "MSTNGHang khong duoc de trong khi LHHDTrung = 3"
    "ERR:1517" = "MDDNGHang khong duoc de trong khi LHHDTrung = 3"
}
$vnptErrRowsHtml = ""
foreach ($k in $VnptErrCodes.Keys) {
    $vnptErrRowsHtml += "<tr><td style='font-family:monospace;font-weight:600;color:#cf222e;white-space:nowrap;'>$k</td><td>$($VnptErrCodes[$k])</td></tr>`n"
}

# Bang ma loi theo tung loai thong bao (Replace / Create / Adjust / Cancel)
$VnptErrByAction = [ordered]@{
    "Replace" = [ordered]@{
        "ERR:1" = "Tai khoan dang nhap sai hoac khong co quyen"
        "ERR:2" = "Khong ton tai hoa don can thay the"
        "ERR:3" = "Du lieu xml dau vao khong dung quy dinh"
        "ERR:5" = "Co loi trong qua trinh tao moi hoa don thay the"
        "ERR:6" = "Dai hoa don cu da het"
        "ERR:7" = "User name khong phu hop, khong tim thay company tuong ung cho user."
        "ERR:8" = "Hoa don da duoc thay the roi. Khong the thay the nua."
        "ERR:9" = "Trang thai hoa don khong duoc thay the"
    }
    "Create" = [ordered]@{
        "ERR:1" = "Tai khoan dang nhap sai hoac khong co quyen them khach hang"
        "ERR:3" = "Du lieu xml dau vao khong dung quy dinh (1 hoa don loi trong chuoi XML se lam ca lo khong duoc phat hanh)"
        "ERR:7" = "Thong tin ve Username/pass khong hop le"
        "ERR:20" = "Pattern va Serial khong phu hop, hoac khong ton tai hoa don da dang ki co su dung Pattern/Serial truyen vao"
        "ERR:5" = "Khong phat hanh duoc hoa don (loi khong xac dinh, kiem tra exception tra ve - DB roll back)"
        "ERR:10" = "Lo co so hoa don vuot qua so luong cho phep"
        "ERR:6" = "Dai hoa don khong du so hoa don cho lo phat hanh"
        "ERR:13" = "Loi trung fkey (1 hoac nhieu hoa don trong lo co Fkey trung voi hoa don da phat hanh)"
        "ERR:21" = "Loi trung so hoa don"
        "ERR:29" = "Loi chung thu het han"
        "ERR:30" = "Danh sach hoa don ton tai ngay hoa don nho hon ngay hoa don da phat hanh"
    }
    "Adjust" = [ordered]@{
        "ERR:1" = "Tai khoan dang nhap sai hoac khong co quyen"
        "ERR:2" = "Hoa don can dieu chinh khong ton tai"
        "ERR:3" = "Du lieu xml dau vao khong dung quy dinh"
        "ERR:5" = "Co loi trong qua trinh tao moi hoa don dieu chinh"
        "ERR:6" = "Dai hoa don cu da het"
        "ERR:7" = "User name khong phu hop, khong tim thay company tuong ung cho user."
        "ERR:8" = "Hoa don can dieu chinh da bi thay the. Khong the dieu chinh duoc nua."
        "ERR:9" = "Trang thai hoa don khong duoc dieu chinh"
        "ERR:15" = "Loi khi thuc hien Deserialize chuoi hoa don dau vao"
        "ERR:19" = "Pattern truyen vao khong giong voi hoa don can dieu chinh"
        "ERR:20" = "Dai hoa don het, User/Account khong co quyen voi Serial/Pattern va serial khong phu hop"
    }
    "Cancel" = [ordered]@{
        "ERR:1" = "Tai khoan dang nhap sai hoac khong co quyen"
        "ERR:2" = "Khong tim thay hoa don"
        "ERR:6" = "Loi khong xac dinh"
        "ERR:7" = "Khong tim thay thong tin cong ty tuong ung, hoac loi xac dinh"
        "ERR:8" = "Hoa don da bi dieu chinh / huy / hoa don moi tao khong the huy duoc"
        "ERR:9" = "Hoa don da thanh toan, khong cho phep huy"
        "ERR:20" = "Dai hoa don het, User/Account khong co quyen voi Serial/Pattern va serial khong phu hop"
    }
}
$vnptErrByActionHtml = ""
foreach ($action in $VnptErrByAction.Keys) {
    $rows = ""
    foreach ($k in $VnptErrByAction[$action].Keys) {
        $rows += "<tr><td style='font-family:monospace;font-weight:600;color:#cf222e;white-space:nowrap;'>$k</td><td>$($VnptErrByAction[$action][$k])</td></tr>`n"
    }
    $vnptErrByActionHtml += @"
    <details style="margin-bottom:10px;">
      <summary style="cursor:pointer;font-weight:700;font-size:13.5px;color:#24292f;">Neu cau thong bao: $action....Err ($($VnptErrByAction[$action].Count) ma)</summary>
      <table style="width:100%;border-collapse:collapse;margin-top:8px;">
        <thead><tr><th style="text-align:left;padding:6px 8px;font-size:12px;color:#57606a;border-bottom:1px solid #eaeef2;">Ket qua</th><th style="text-align:left;padding:6px 8px;font-size:12px;color:#57606a;border-bottom:1px solid #eaeef2;">Mo ta</th></tr></thead>
        <tbody style="font-size:13px;">$rows</tbody>
      </table>
    </details>
"@
}

foreach ($dt in $docTypeOrder) {
    if ($DocTypeCounts.ContainsKey($dt)) {
        $cnt = $DocTypeCounts[$dt]
        $pct = if ($docTypeTotal -gt 0) { [math]::Round(($cnt / $docTypeTotal) * 100, 1) } else { 0 }
        $docTypeRowsHtml += @"
        <div class="bu-row">
          <div class="bu-name" style="width:140px;">$dt</div>
          <div class="bu-bar-track"><div class="bu-bar-fill" style="width:$pct%;background:linear-gradient(90deg,#bf3989,#ff9bce);"></div></div>
          <div class="bu-count">$cnt <span class="bu-pct">($pct%)</span></div>
        </div>
"@
        if ($dt -eq "Khac" -and $UnmappedDocCodes.Count -gt 0) {
            $unmappedDocList = ($UnmappedDocCodes.GetEnumerator() | Sort-Object Value -Descending | ForEach-Object { "$($_.Key): $($_.Value)" }) -join ", "
            $docTypeRowsHtml += @"
        <div class="bu-row bu-subrow">
          <div class="bu-name"></div>
          <div style="flex:1;font-size:11.5px;color:#57606a;">Chi tiet ma chua gan loai: $unmappedDocList</div>
          <div style="width:90px;"></div>
        </div>
"@
        }
    }
}
if ([string]::IsNullOrWhiteSpace($docTypeRowsHtml)) {
    $docTypeRowsHtml = "<div style='color:#8c959f;font-size:13px;'>Chua co du lieu.</div>"
}

# Xay bang thong ke theo server (co lop phu the hien so da ra hoa don)
$serverRowsHtml = ""
$serverTotal = ($ServerCounts.Values | Measure-Object -Sum).Sum
foreach ($server in $Config.sap_to_pp.sftp.servers) {
    $srvName = $server.name
    $cnt = if ($ServerCounts.ContainsKey($srvName)) { $ServerCounts[$srvName] } else { 0 }
    $issuedCnt = if ($ServerIssuedCounts.ContainsKey($srvName)) { $ServerIssuedCounts[$srvName] } else { 0 }
    $pct = if ($serverTotal -gt 0) { [math]::Round(($cnt / $serverTotal) * 100, 1) } else { 0 }
    $issuedPctOfBar = if ($cnt -gt 0) { if ($issuedCnt -ge $cnt) { 100 } else { [math]::Round(($issuedCnt / $cnt) * 100, 2) } } else { 0 }
    $srvDataAnomaly = $issuedCnt - $cnt
    $srvComplete = ($cnt -gt 0 -and $issuedCnt -eq $cnt)
    $srvNameHtml = if ($srvComplete) { "$srvName <span class='done-badge'>&#10003; 100%</span>" } else { $srvName }
    $srvPendingBreakdown = if ($ServerPendingDocType.ContainsKey($srvName)) { Build-DocTypeSummary -Counts $ServerPendingDocType[$srvName] } else { "" }
    $srvPendingCnt = if ($ServerPendingCount.ContainsKey($srvName)) { $ServerPendingCount[$srvName] } else { 0 }
    # Lop bao ve: neu so hoc (Tong - Da ra) khac voi danh sach chi tiet, van hien so chenh lech
    # thay vi de trong im lang (giup phat hien loi du lieu neu con sot truong hop tuong tu).
    $srvMathGap = $cnt - $issuedCnt
    if ($srvMathGap -gt $srvPendingCnt) { $srvPendingCnt = $srvMathGap }
    $serverRowsHtml += @"
    <div class="bu-row$(if ($srvComplete) {' bu-row-complete'})">
      <div class="bu-name" style="width:150px;">$srvNameHtml</div>
      <div class="bu-bar-track">
        <div class="bu-bar-fill" style="width:$pct%;background:linear-gradient(90deg,#8250df,#c297ff);">
          <div class="bu-bar-issued" style="width:$issuedPctOfBar%;$(if ($srvComplete) {'background:linear-gradient(90deg,#0969da,#54aeff);'})" title="Da ra hoa don: $issuedCnt/$cnt"></div>
        </div>
      </div>
      <div class="bu-count">$cnt <span class="bu-pct">($pct%)</span></div>
    </div>
    <div class="bu-row bu-subrow">
      <div class="bu-name" style="width:150px;"></div>
      <div style="flex:1;font-size:14px;color:$(if ($srvComplete) {'#0969da'} else {'#1a7f37'});font-weight:800;">&#9632; Da ra hoa don: <span style="font-size:16px;">$issuedCnt / $cnt</span> <span style="font-size:15px;">($issuedPctOfBar%)</span>$(if ($srvPendingCnt -gt 0) {" &middot; <span style='color:#cf222e;'>Con lai: $srvPendingCnt</span>"})$(if ($srvDataAnomaly -gt 0) {"<br><span style='font-size:10.5px;color:#bc4c00;font-weight:400;'>&#9888;&#65039; Du lieu bat thuong: co $srvDataAnomaly billing duoc tinh 'da ra' nhung khong nam trong tong cua server nay</span>"})</div>
      <div style="width:90px;"></div>
    </div>
"@
    if ($srvPendingBreakdown) {
        $serverRowsHtml += @"
    <div class="bu-row bu-subrow">
      <div class="bu-name" style="width:150px;"></div>
      <div style="flex:1;font-size:10.5px;color:#8c959f;">Con lai theo loai: $srvPendingBreakdown</div>
      <div style="width:90px;"></div>
    </div>
"@
    }
    $srvExcludedSummary = if ($ServerExcludedDocType.ContainsKey($srvName)) { Build-ExcludedSummaryWithPending -TotalCounts $ServerExcludedDocType[$srvName] -PendingCounts $ServerExcludedPending[$srvName] } else { "" }
    if ($srvExcludedSummary) {
        $srvExclCount = if ($ServerExcludedFiles.ContainsKey($srvName)) { $ServerExcludedFiles[$srvName].Count } else { 0 }
        $serverRowsHtml += @"
    <div class="bu-row bu-subrow">
      <div class="bu-name" style="width:150px;"></div>
      <div style="flex:1;font-size:10.5px;color:#bc4c00;">Khong tinh (PXK/Return/D.Chinh/Cancel/Z3F2): $srvExcludedSummary</div>
      <div style="width:90px;"></div>
    </div>
"@
        if ($srvExclCount -gt 0) {
            $srvExclPendingBnSet = if ($ServerExcludedPendingBnSet.ContainsKey($srvName)) { $ServerExcludedPendingBnSet[$srvName] } else { $null }
            $srvExclListHtml = Build-PendingFileListHtml -Files $ServerExcludedFiles[$srvName].Values -PendingBnSet $srvExclPendingBnSet
            $srvExclPendingCnt = if ($ServerExcludedPendingFiles.ContainsKey($srvName)) { $ServerExcludedPendingFiles[$srvName].Count } else { 0 }
            $srvExclSumLabel = if ($srvExclPendingCnt -gt 0) { "Xem danh sach $srvExclCount billing khong tinh (co $srvExclPendingCnt CHUA ra hoa don)" } else { "Xem danh sach $srvExclCount billing khong tinh" }
            $serverRowsHtml += @"
    <div class="bu-row bu-subrow">
      <div class="bu-name" style="width:150px;"></div>
      <div style="flex:1;">
        <details>
          <summary style="cursor:pointer;font-size:11px;color:#bc4c00;font-weight:600;">$srvExclSumLabel</summary>
          <div style="max-height:260px;overflow-y:auto;margin-top:6px;font-size:11px;font-family:monospace;">$srvExclListHtml</div>
        </details>
      </div>
      <div style="width:90px;"></div>
    </div>
"@
        }
    }
    $srvKnownPendingFiles = if ($ServerPendingFiles.ContainsKey($srvName)) { $ServerPendingFiles[$srvName] } else { @() }
    $srvUntrackable = if ($ServerUntrackableFiles.ContainsKey($srvName)) { $ServerUntrackableFiles[$srvName] } else { @() }
    $srvGapUnexplained = $srvPendingCnt - $srvKnownPendingFiles.Count
    if ($srvPendingCnt -gt 0) {
        $srvPendingListHtml = Build-PendingFileListHtml -Files $srvKnownPendingFiles
        if ($srvGapUnexplained -gt 0 -and $srvUntrackable.Count -gt 0) {
            foreach ($uf in ($srvUntrackable | Select-Object -First $srvGapUnexplained)) {
                $srvPendingListHtml += "<div style='padding:3px 0;border-bottom:1px solid #f1f3f5;background:#fff8f6;'><span style='color:#8c959f;font-weight:600;'>$($uf.d)</span> &middot; $($uf.n) <span style='background:#ffebe9;color:#cf222e;font-size:9.5px;font-weight:700;padding:1px 6px;border-radius:8px;margin-left:4px;'>KHONG XAC DINH SO BILLING</span></div>`n"
            }
        }
        $serverRowsHtml += @"
    <div class="bu-row bu-subrow">
      <div class="bu-name" style="width:150px;"></div>
      <div style="flex:1;">
        <details>
          <summary style="cursor:pointer;font-size:11px;color:#cf222e;font-weight:600;">Xem danh sach $srvPendingCnt billing con lai</summary>
          <div style="max-height:220px;overflow-y:auto;margin-top:6px;font-size:11px;font-family:monospace;">$srvPendingListHtml</div>
        </details>
      </div>
      <div style="width:90px;"></div>
    </div>
"@
    }
}
$workflowSvg = @"
<svg width="100%" viewBox="0 0 850 380" style="max-width:850px;">
<defs>
<marker id="arrow" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M2 1L8 5L2 9" fill="none" stroke="#8c959f" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"/></marker>
<filter id="wfGlow" x="-50%" y="-50%" width="200%" height="200%">
  <feGaussianBlur stdDeviation="4" result="blur"/>
  <feMerge><feMergeNode in="blur"/><feMergeNode in="SourceGraphic"/></feMerge>
</filter>
</defs>

<rect x="60" y="70" width="160" height="70" rx="10" fill="#f6f8fa" stroke="#8c959f" stroke-width="1.2"/>
<text x="140" y="96" text-anchor="middle" dominant-baseline="central" font-size="17" font-weight="700" fill="#24292f">SAP</text>
<text x="140" y="118" text-anchor="middle" dominant-baseline="central" font-size="13" fill="#57606a">Xuat file ZVRD</text>

<rect x="250" y="70" width="160" height="70" rx="10" fill="$($cSapPP.fill)" stroke="$($cSapPP.stroke)" stroke-width="2" filter="$(if ($cSapPP.stroke -eq '#1a7f37') {'url(#wfGlow)'})"/>
<text x="330" y="96" text-anchor="middle" dominant-baseline="central" font-size="17" font-weight="700" fill="$($cSapPP.text)">PP</text>
<text x="330" y="118" text-anchor="middle" dominant-baseline="central" font-size="13" fill="$($cSapPP.text)">SAP -&gt; PP (SFTP)</text>

<rect x="440" y="70" width="160" height="70" rx="10" fill="$($cVnpt.fill)" stroke="$($cVnpt.stroke)" stroke-width="2" filter="$(if ($cVnpt.stroke -eq '#1a7f37') {'url(#wfGlow)'})"/>
<text x="520" y="96" text-anchor="middle" dominant-baseline="central" font-size="17" font-weight="700" fill="$($cVnpt.text)">VNPT</text>
<text x="520" y="118" text-anchor="middle" dominant-baseline="central" font-size="13" fill="$($cVnpt.text)">Ky &amp; phat hanh</text>

<rect x="630" y="70" width="160" height="70" rx="10" fill="#f6f8fa" stroke="#8c959f" stroke-width="1.2"/>
<text x="710" y="96" text-anchor="middle" dominant-baseline="central" font-size="17" font-weight="700" fill="#24292f">Tax GOV</text>
<text x="710" y="118" text-anchor="middle" dominant-baseline="central" font-size="13" fill="#57606a">Nhan hoa don</text>

<rect x="250" y="164" width="160" height="38" rx="8" fill="$($cStaging.fill)" stroke="$($cStaging.stroke)" stroke-width="2"/>
<text x="330" y="183" text-anchor="middle" dominant-baseline="central" font-size="14" font-weight="700" fill="$($cStaging.text)">PP Processing</text>

<rect x="440" y="256" width="160" height="70" rx="10" fill="#f6f8fa" stroke="#8c959f" stroke-width="1.2"/>
<text x="520" y="282" text-anchor="middle" dominant-baseline="central" font-size="17" font-weight="700" fill="#24292f">Khach hang</text>
<text x="520" y="304" text-anchor="middle" dominant-baseline="central" font-size="13" fill="#57606a">Nhan hoa don</text>

<line id="wfL1" x1="220" y1="105" x2="248" y2="105" stroke="#8c959f" stroke-width="2" marker-end="url(#arrow)"/>
<line id="wfL2" x1="410" y1="105" x2="438" y2="105" stroke="#8c959f" stroke-width="2" marker-end="url(#arrow)"/>
<line id="wfL3" x1="600" y1="105" x2="628" y2="105" stroke="#8c959f" stroke-width="2" marker-end="url(#arrow)"/>
<line id="wfL4" x1="520" y1="140" x2="520" y2="254" stroke="#8c959f" stroke-width="2" marker-end="url(#arrow)"/>

<path d="M460 140 L460 220 L140 220 L140 142" fill="none" stroke="$($cCallback.stroke)" stroke-width="2" stroke-dasharray="5 4" marker-end="url(#arrow)"/>
<text x="475" y="224" text-anchor="start" font-size="14" fill="$($cCallback.text)" font-weight="600">Callback: eSeri/eForm/eRunning</text>

<circle r="4" fill="#54aeff"><animateMotion dur="2.2s" repeatCount="indefinite" path="M220,105 L248,105"/></circle>
<circle r="4" fill="#8250df"><animateMotion dur="2.2s" begin="0.5s" repeatCount="indefinite" path="M410,105 L438,105"/></circle>
<circle r="4" fill="#bf3989"><animateMotion dur="2.2s" begin="1s" repeatCount="indefinite" path="M600,105 L628,105"/></circle>
<circle r="4" fill="#1a7f37"><animateMotion dur="2.8s" repeatCount="indefinite" path="M520,140 L520,254"/></circle>
<circle r="4" fill="$($cCallback.stroke)"><animateMotion dur="3.5s" repeatCount="indefinite" path="M460,140 L460,220 L140,220 L140,142"/></circle>

<g>
  <animateMotion dur="4.5s" repeatCount="indefinite" path="M226,105 L784,105"/>
  <rect x="-26" y="-11" width="52" height="22" rx="11" fill="#0969da" stroke="white" stroke-width="1.5"/>
  <text x="0" y="4" text-anchor="middle" font-size="11" font-weight="700" fill="white">Billing</text>
</g>
<g>
  <animateMotion dur="4.5s" begin="2.25s" repeatCount="indefinite" path="M226,105 L784,105"/>
  <rect x="-20" y="-11" width="40" height="22" rx="11" fill="#8250df" stroke="white" stroke-width="1.5"/>
  <text x="0" y="4" text-anchor="middle" font-size="11" font-weight="700" fill="white">PXK</text>
</g>
</svg>
"@

$sapFilesJson = ($AllSapFiles | ConvertTo-Json -Compress -Depth 3)
if (-not $sapFilesJson) { $sapFilesJson = "[]" }
# Dung -InputObject @(...) de luon ra mang JSON, ke ca khi chi co dung 1 file
$stagingFilesJson = (ConvertTo-Json -InputObject @($AllStagingFiles) -Compress -Depth 3)
if (-not $stagingFilesJson) { $stagingFilesJson = "[]" }
$scannedFoldersJson = (ConvertTo-Json -InputObject @($ScannedFolders) -Compress -Depth 3)
# Nguong "xu ly cham" cho bang Billing xu ly hon N phut (lay theo max_stuck_minutes cua staging, mac dinh 30)
$SlowThresholdMinutes = if ($MaxStuckMinutes) { $MaxStuckMinutes } else { 30 }
$excludedTypesJson = (ConvertTo-Json -InputObject @($ExcludedFromBuServerCount) -Compress)
$runTimeIso = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
if (-not $scannedFoldersJson) { $scannedFoldersJson = "[]" }
$bakupFilesJson = ($AllBakupFiles | ConvertTo-Json -Compress -Depth 3)
if (-not $bakupFilesJson) { $bakupFilesJson = "[]" }

$html = @"
<!DOCTYPE html>
<html lang="vi">
<head>
<meta charset="UTF-8">
<title>Billing Pipeline Monitor</title>
<meta http-equiv="refresh" content="300">
<style>
  * { box-sizing: border-box; }
  @keyframes fadeInUp { from { opacity: 0; transform: translateY(14px); } to { opacity: 1; transform: translateY(0); } }
  @keyframes pulseDot { 0% { box-shadow: 0 0 0 0 rgba(63,209,116,0.6); } 70% { box-shadow: 0 0 0 7px rgba(63,209,116,0); } 100% { box-shadow: 0 0 0 0 rgba(63,209,116,0); } }
  @keyframes growBar { from { width: 0 !important; } }
  body {
    font-family: 'Segoe UI', -apple-system, Arial, sans-serif;
    background: linear-gradient(180deg, #f0f4f9 0%, #f6f8fa 260px);
    margin: 0; padding: 0 0 48px 0; color: #1f2328;
  }
  .topbar {
    background: linear-gradient(120deg, #0b3d91 0%, #0969da 100%),
      radial-gradient(circle at 90% 10%, rgba(255,255,255,0.10) 0%, transparent 50%),
      radial-gradient(circle at 15% 90%, rgba(255,255,255,0.08) 0%, transparent 40%);
    color: white; padding: 30px 40px; margin-bottom: -60px;
    background-size: 200% 200%, auto, auto; animation: gradientShift 8s ease infinite;
    position: relative; overflow: hidden;
  }
  .topbar::after {
    content: ""; position: absolute; right: -40px; top: -40px; width: 180px; height: 180px;
    border-radius: 50%; background: rgba(255,255,255,0.06); pointer-events: none;
  }
  .topbar h1 { display: flex; align-items: center; gap: 10px; font-size: 22px; margin: 0 0 6px 0; font-weight: 700; letter-spacing: 0.2px; width: 100%; }
  .topbar h1 button:hover { background: rgba(255,255,255,0.3) !important; }
  .topbar h1 button:active #refreshIcon { animation: spin360 0.5s linear; }
  @keyframes spin360 { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }
  .topbar h1 .logo-dot { width: 10px; height: 10px; border-radius: 3px; background: #54aeff; display: inline-block; transform: rotate(45deg); }
  @keyframes gradientShift { 0% { background-position: 0% 50%; } 50% { background-position: 100% 50%; } 100% { background-position: 0% 50%; } }
  .live-dot { display: inline-block; width: 8px; height: 8px; border-radius: 50%; background: #3fd174; margin-right: 6px; animation: pulseDot 2s infinite; vertical-align: middle; }
  .topbar .subtitle { color: #cfe0fb; font-size: 13px; margin: 2px 0; }
  .topbar .subtitle b { color: white; }
  .container { max-width: 1040px; margin: 0 auto; padding: 0 24px; }
  .summary { display: flex; gap: 16px; margin-bottom: 20px; }
  .card {
    flex: 1; background: white; border-radius: 12px; padding: 18px 16px;
    text-align: center; box-shadow: 0 2px 10px rgba(15,45,90,0.08);
    border: 1px solid #eef1f5; animation: fadeInUp 0.5s ease backwards;
    transition: transform 0.18s ease, box-shadow 0.18s ease;
  }
  .card:hover { transform: translateY(-3px); box-shadow: 0 8px 20px rgba(15,45,90,0.14); }
  .hero-summary { gap: 14px; }
  .svc-tile-row { display: flex; gap: 14px; margin-bottom: 20px; flex-wrap: wrap; }
  .svc-tile {
    flex: 0 1 160px; aspect-ratio: 1.2 / 1; border-radius: 14px;
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    color: white; box-shadow: 0 4px 14px rgba(0,0,0,0.12); transition: transform 0.18s ease;
    position: relative; overflow: hidden;
  }
  .svc-tile:hover { transform: translateY(-3px) scale(1.02); }
  .svc-tile-ok { background: linear-gradient(155deg, #1a7f37 0%, #2da44e 60%, #3fd174 100%); }
  .svc-tile-error { background: linear-gradient(155deg, #82231f 0%, #cf222e 60%, #ff6b6b 100%); animation: svcPulse 1.6s ease-in-out infinite; }
  .svc-tile-unknown { background: linear-gradient(155deg, #57606a 0%, #8c959f 100%); }
  @keyframes svcPulse { 0%,100% { box-shadow: 0 4px 14px rgba(207,34,46,0.25); } 50% { box-shadow: 0 4px 26px rgba(207,34,46,0.55); } }
  .svc-tile-icon { font-size: 30px; margin-bottom: 4px; }
  .svc-tile-name { font-size: 13px; font-weight: 700; opacity: 0.95; }
  .svc-tile-status { font-size: 19px; font-weight: 900; letter-spacing: 0.4px; margin-top: 2px; }
  .svc-tile-sub { font-size: 11px; opacity: 0.85; margin-top: 4px; }
  .hero-card {
    padding: 24px 16px 20px; position: relative; overflow: hidden;
    background: linear-gradient(160deg, #ffffff 0%, #fbfcfe 100%);
    border: 1.5px solid #e3e9f0;
  }
  .hero-card::before {
    content: ""; position: absolute; left: 0; top: 0; width: 100%; height: 4px;
    background: linear-gradient(90deg, #0969da, #54aeff, #8250df, #bf3989, #cf222e);
    background-size: 300% 100%; animation: gradientShift 6s ease infinite;
  }
  .hero-card .hero-icon { font-size: 22px; margin-bottom: 6px; filter: drop-shadow(0 2px 3px rgba(0,0,0,0.12)); }
  .hero-card .num { font-size: 38px; font-weight: 900; letter-spacing: -0.5px; }
  .hero-card .label { font-size: 13px; font-weight: 600; margin-top: 8px; }
  .hero-card .hero-sub { font-size: 10.5px; color: #8c959f; margin-top: 6px; line-height: 1.5; }
  .hero-card:hover { transform: translateY(-5px) scale(1.02); box-shadow: 0 12px 28px rgba(15,45,90,0.18); }
  .card .num { font-size: 30px; font-weight: 800; line-height: 1.1; }
  .card .label { color: #57606a; font-size: 12.5px; margin-top: 6px; font-weight: 500; }
  .panel {
    background: white; border-radius: 12px; padding: 20px;
    margin-bottom: 20px; box-shadow: 0 2px 10px rgba(15,45,90,0.06);
    border: 1px solid #eef1f5; animation: fadeInUp 0.5s ease backwards;
    transition: box-shadow 0.18s ease;
  }
  .panel:hover { box-shadow: 0 6px 18px rgba(15,45,90,0.10); }
  .panel-title { font-size: 15px; font-weight: 700; margin: 0 0 14px 0; color: #24292f; }
  .bu-row { display: flex; align-items: center; gap: 12px; margin-bottom: 10px; }
  .bu-row:last-child { margin-bottom: 0; }
  .bu-name { width: 90px; font-weight: 700; font-size: 13px; color: #24292f; flex-shrink: 0; }
  .bu-bar-track { flex: 1; background: #eef1f5; border-radius: 999px; height: 15px; overflow: hidden; box-shadow: inset 0 1px 3px rgba(0,0,0,0.08); }
  .bu-bar-fill { background: linear-gradient(90deg,#0969da,#54aeff); height: 100%; border-radius: 999px; position: relative; animation: growBar 1s ease-out; box-shadow: inset 0 6px 6px -3px rgba(255,255,255,0.45), inset 0 -5px 6px -3px rgba(0,0,0,0.18); }
  .bu-bar-issued { position: absolute; left: 0; top: 0; height: 100%; background: repeating-linear-gradient(45deg,#1a7f37,#1a7f37 4px,#2da44e 4px,#2da44e 8px); border-radius: 999px; opacity: 0.9; box-shadow: inset 0 6px 6px -3px rgba(255,255,255,0.35), inset 0 -5px 6px -3px rgba(0,0,0,0.18); }
  .bu-subrow { margin-top: -4px; margin-bottom: 10px !important; }
  .done-badge { display: inline-block; background: linear-gradient(90deg,#0969da,#54aeff); color: white; font-size: 10.5px; font-weight: 800; padding: 1px 8px; border-radius: 10px; margin-left: 4px; animation: pulseDot 2.2s infinite; }
  .bu-row-complete .bu-name { color: #0969da; }
  .bu-count { width: 90px; text-align: right; font-size: 13px; font-weight: 600; flex-shrink: 0; }
  .bu-pct { color: #8c959f; font-weight: 400; }
  table { width: 100%; border-collapse: collapse; }
  th { text-align: left; background: #f6f8fa; padding: 10px 12px; font-size: 12.5px; color: #57606a; border-bottom: 1px solid #eaeef2; }
  td { padding: 10px 12px; border-bottom: 1px solid #f1f3f5; font-size: 13.5px; vertical-align: top; }
  tr:last-child td { border-bottom: none; }
  tr:hover td { background: #fafbfc; }
  .badge { padding: 2px 10px; border-radius: 12px; font-weight: 600; font-size: 12px; }
  .search-hint {
    background: #eef6ff; border: 1px solid #cfe0fb; border-radius: 10px;
    padding: 10px 16px; font-size: 13px; color: #0b3d91; margin-bottom: 20px;
  }
  .search-hint b { color: #0969da; }
  .live-search {
    background: white; border-radius: 12px; padding: 20px; margin-bottom: 20px;
    box-shadow: 0 2px 10px rgba(15,45,90,0.08); border: 1px solid #eef1f5;
  }
  .live-search input {
    width: 100%; padding: 12px 16px; border: 1.5px solid #d0d7de; border-radius: 8px;
    font-size: 15px; box-sizing: border-box; outline: none; transition: border-color 0.15s;
  }
  .live-search input:focus { border-color: #0969da; }
  .live-search-result { margin-top: 14px; }
  .ls-row { display: flex; align-items: center; gap: 10px; padding: 8px 0; border-bottom: 1px solid #f1f3f5; font-size: 13.5px; }
  .ls-row:last-child { border-bottom: none; }
  .ls-badge { padding: 2px 10px; border-radius: 12px; font-weight: 700; font-size: 11.5px; flex-shrink: 0; }
  .ls-found { background: #dafbe1; color: #1a7f37; }
  .ls-missing { background: #ffebe9; color: #cf222e; }
  .ls-name { font-family: 'Consolas', monospace; font-size: 12.5px; color: #57606a; word-break: break-all; }
  .ls-empty { color: #8c959f; font-size: 13px; padding: 6px 0; }
  .ls-stage-grid { display: flex; gap: 10px; margin-bottom: 12px; flex-wrap: wrap; }
  .ls-stage { flex: 1; min-width: 140px; border-radius: 8px; padding: 10px 14px; }
  .ls-stage-ok { background: #dafbe1; }
  .ls-stage-missing { background: #ffebe9; }
  .ls-stage-wait { background: #fff8c5; }
  .slow-tabs { display: flex; gap: 8px; flex-wrap: wrap; margin-bottom: 12px; }
  .slow-tab { background: #f6f8fa; border: 1px solid #d0d7de; color: #57606a; padding: 6px 14px; border-radius: 999px; font-size: 13px; font-weight: 600; cursor: pointer; }
  .slow-tab-active { background: #0969da; border-color: #0969da; color: white; }
  .slow-dl { margin-left: auto; background: white; border: 1px solid #d0d7de; color: #0969da; padding: 6px 14px; border-radius: 8px; font-size: 13px; font-weight: 600; cursor: pointer; }
  .slow-wrap { max-height: 420px; overflow: auto; border: 1px solid #eaeef2; border-radius: 8px; }
  .slow-wrap table td, .slow-wrap table th { font-size: 12.5px; padding: 7px 10px; }
  .slow-wrap th { position: sticky; top: 0; }
  .slow-bn { font-family: 'Consolas', monospace; color: #0969da; font-weight: 700; cursor: pointer; text-decoration: underline dotted; }
  .slow-dur { font-weight: 800; white-space: nowrap; }
  .slow-stats { display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 14px; }
  .slow-stat { flex: 1; min-width: 180px; border-radius: 10px; padding: 12px 16px; background: #f6f8fa; }
  .slow-stat .v { font-size: 26px; font-weight: 800; }
  .slow-stat .l { font-size: 12px; color: #57606a; font-weight: 600; }
  .dr-table th, .dr-table td { font-size: 13px; padding: 8px 10px; }
  .dr-table tfoot td { font-weight: 700; background: #f6f8fa; border-top: 2px solid #d0d7de; }
  .dr-row-today td { background: #eef6ff; }
  .dr-today { display: inline-block; background: #0969da; color: white; font-size: 10.5px; font-weight: 700; padding: 1px 8px; border-radius: 10px; margin-left: 4px; }
  .dr-bar { background: #eaeef2; border-radius: 999px; height: 10px; overflow: hidden; }
  .dr-bar-fill { background: linear-gradient(90deg,#0969da,#54aeff); height: 100%; border-radius: 999px; }
  .slow-note { font-size: 11px; color: #8c959f; margin-top: 8px; }
  .ls-status { border-radius: 8px; padding: 12px 16px; margin-bottom: 12px; font-size: 14px; line-height: 1.5; }
  .ls-status-ok { background: #dafbe1; color: #116329; border: 1px solid #4ac26b; }
  .ls-status-wait { background: #fff8c5; color: #7d5700; border: 1px solid #eac54f; }
  .ls-status-bad { background: #ffebe9; color: #82231f; border: 1px solid #ff8182; }
  .ls-scanned { font-size: 12.5px; color: #57606a; margin-bottom: 12px; }
  .ls-scanned summary { cursor: pointer; font-weight: 600; color: #0969da; }
  .ls-scanned div { margin-top: 4px; }
  .ls-stage-wait .ls-stage-val { color: #9a6700; font-weight: 700; font-size: 13.5px; }
  .ls-location { background: #fff8c5; border: 1px solid #eac54f; color: #7d5700; border-radius: 8px; padding: 12px 16px; margin-bottom: 12px; font-size: 13px; }
  .ls-location-stuck { background: #ffebe9; border-color: #ff8182; color: #82231f; }
  .ls-path { font-family: 'Consolas', monospace; font-size: 11.5px; color: #57606a; word-break: break-all; margin-top: 2px; }
  .ls-stuck { display: inline-block; background: #cf222e; color: white; font-size: 10.5px; font-weight: 700; padding: 1px 8px; border-radius: 10px; margin-left: 6px; }
  .ls-stage-label { font-size: 11.5px; font-weight: 700; color: #57606a; }
  .ls-stage-ok .ls-stage-val { color: #1a7f37; font-weight: 700; font-size: 13.5px; }
  .ls-stage-missing .ls-stage-val { color: #cf222e; font-weight: 700; font-size: 13.5px; }
  .ls-elapsed { background: #eef6ff; border: 1px solid #cfe0fb; border-radius: 8px; padding: 12px 16px; margin-bottom: 12px; }
  .ls-hint { background: #fff8c5; color: #7d5700; border-radius: 8px; padding: 10px 14px; font-size: 13px; margin-bottom: 12px; }
  .ls-badge.ls-doctype { background: #ddf4ff; color: #0969da; }
  .ls-filelist { max-height: 220px; overflow-y: auto; }
  .ls-trigger { display: inline-block; background: #fff1e5; color: #bc4c00; font-size: 10.5px; font-weight: 700; padding: 1px 8px; border-radius: 10px; margin-left: 6px; }
  .ls-diff-type { display: inline-block; background: #ddf4ff; color: #0969da; font-size: 10.5px; font-weight: 700; padding: 1px 8px; border-radius: 10px; margin-left: 6px; }
  .card:nth-child(1) { animation-delay: 0.02s; } .card:nth-child(2) { animation-delay: 0.08s; } .card:nth-child(3) { animation-delay: 0.14s; }
  .panel:nth-of-type(1) { animation-delay: 0.05s; } .panel:nth-of-type(2) { animation-delay: 0.12s; } .panel:nth-of-type(3) { animation-delay: 0.19s; } .panel:nth-of-type(4) { animation-delay: 0.26s; }
</style>
</head>
<body>
<div class="topbar">
  <h1><span class="logo-dot"></span>Billing Pipeline Monitor
    <button onclick="location.reload(true)" style="margin-left:auto;background:rgba(255,255,255,0.18);border:1px solid rgba(255,255,255,0.4);color:white;padding:6px 14px;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer;display:flex;align-items:center;gap:6px;">
      <span id="refreshIcon" style="display:inline-block;">&#128260;</span> Tai lai trang
    </button>
  </h1>
  <div class="subtitle">SAP &rarr; PP &rarr; VNPT &rarr; Thue</div>
  <div class="subtitle"><span class="live-dot"></span>Lan chay gan nhat: <b>$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</b> &middot; tu refresh moi 5 phut</div>
  $(if ($CheckDate -ne (Get-Date -Format "yyyyMMdd")) { "<div class=`"subtitle`" style=`"background:rgba(255,255,255,0.15);display:inline-block;padding:4px 12px;border-radius:8px;margin-top:6px;`">&#128197; Dang xem du lieu NGAY QUA KHU: <b>$CheckDate</b></div>" })
  <div class="subtitle" style="font-size:11px;opacity:0.85;">Luu y: nut nay chi tai lai trang (xem ban da luu gan nhat) -- muon co du lieu moi that su, can chay lai run.bat (hoac dat Task Scheduler tu dong)</div>
</div>
<div class="container" style="padding-top:76px;">
  <div class="live-search">
    <div class="panel-title" style="margin-bottom:10px;">&#128269; Tim billing nhanh (trong du lieu lan chay nay)</div>
    <input type="text" id="billingSearchBox" placeholder="Nhap mot phan so hoa don / billing, vi du: 2483512784" autocomplete="off" onkeyup="doLiveSearch()">
    <div class="live-search-result" id="billingSearchResult"></div>
  </div>
  <div class="search-hint">Muon tim chi tiet hon (ca file .txt goc, ngay khac)? Bam dup file <b>run_search.bat</b> trong cung thu muc.</div>

  <div class="summary hero-summary">
    <div class="card hero-card"><div class="hero-icon">&#128230;</div><div class="num countup" data-target="$SapBillingCount" style="color:#0969da">0</div><div class="label">Billing hom nay (tu SAP)</div><div class="hero-sub">$SapDocTypeSummary</div></div>
    <div class="card hero-card" style="background:$IssuedBg;position:relative;">$ringSvg<div class="hero-icon">&#9989;</div><div class="num $(if ($null -ne $IssuedCount) {'countup'}) " data-target="$(if ($null -ne $IssuedCount) {$IssuedCount} else {0})" style="color:$IssuedColor">$(if ($null -ne $IssuedCount) {'0'} else {$IssuedDisplay})</div><div class="label">Da ra hoa don $IssuedNote</div></div>
    <div class="card hero-card"><div class="hero-icon">&#128176;</div><div class="num countup" data-target="$BillingBanCount" style="color:#bf3989">0</div><div class="label">Billing Ban (ZF2)</div></div>
    <div class="card hero-card"><div class="hero-icon">&#9203;</div><div class="num" style="color:#8250df">$AvgProcessingDisplay</div><div class="label">Thoi gian xu ly TB $AvgProcessingNote</div><div class="hero-sub">Tong cong don: $TotalProcessingDisplay<br>Chua doi chieu duoc gio: $UnmatchedIssuedCount billing</div>
      <details style="margin-top:8px;text-align:left;">
        <summary style="cursor:pointer;font-size:10px;color:#8250df;font-weight:600;text-align:center;">&#8505;&#65039; Cach tinh</summary>
        <div style="font-size:10px;color:#57606a;margin-top:6px;line-height:1.5;">
          Voi moi billing: <b>Gio Bakup</b> (luc ra hoa don) tru <b>Gio SAP</b> (luc xuat file), roi lay trung binh cong tat ca.<br>
          &middot; Gio SAP: trich tu ten file SAP.<br>
          &middot; Gio Bakup: trich tu ten file xac nhan (server 06P duoc tru 1 tieng do dong ho lech).<br>
          &middot; Chi tinh khi chenh lech tu 0 den duoi 24 gio.
        </div>
      </details>
    </div>
    <div class="card hero-card"><div class="hero-icon">&#128337;</div><div class="num countup" data-target="$PendingTotal" style="color:#cf222e">0</div><div class="label">Con lai chua ra hoa don</div><div class="hero-sub">$PendingDocTypeSummary</div></div>
  </div>

  <div class="panel" id="slowPanel">
    <div class="panel-title">&#9201;&#65039; Billing xu ly hon $SlowThresholdMinutes phut</div>
    <div class="slow-stats" id="slowStats"></div>
    <div class="slow-tabs">
      <button id="slowTabPending" class="slow-tab slow-tab-active" onclick="showSlowTab('pending')">Chua xong, da cho &gt; $SlowThresholdMinutes phut (<span id="slowCntPending">0</span>)</button>
      <button id="slowTabDone" class="slow-tab" onclick="showSlowTab('done')">Da ra hoa don nhung mat &gt; $SlowThresholdMinutes phut (<span id="slowCntDone">0</span>)</button>
      <button class="slow-dl" onclick="downloadSlowList()">&#128190; Tai danh sach (.csv)</button>
    </div>
    <div id="slowTableWrap" class="slow-wrap"></div>
    <div class="slow-note">Tinh tu gio xuat file SAP den luc ra hoa don (hoac den luc chay tool: <b>$(Get-Date -Format 'HH:mm:ss')</b> neu chua xong). Khong gom PXK, Return, Dieu Chinh, Cancel, Z3F2 (khong co xac nhan ra hoa don). Bam vao so billing de xem chi tiet.</div>
  </div>

  <div class="panel">
    <div class="panel-title">&#128197; Tong billing $DailyReportDays ngay gan nhat (trong PRODATA)</div>
    $DailyReportHtml
    <div class="slow-note">Dem so billing KHONG TRUNG (1 billing trigger nhieu lan chi tinh 1) trong: $DailyReportPaths. Cot server dem rieng tung server; "--" = khong truy cap duoc hoac khong co thu muc ngay do. Loai khac = PXK, Return, Dieu chinh, Cancel, Z3F2...</div>
  </div>

  <div class="panel">
    <div class="panel-title">&#128421;&#65039; Billing hom nay theo Server</div>
    <div class="svc-tile-row" style="margin-bottom:16px;">
      $serviceTilesHtml
    </div>
    <div style="font-size:11px;color:#8c959f;margin:-8px 0 12px 0;">Chi tinh: Billing Ban, Billing Smallunit (khong gom PXK, Return, Dieu Chinh, Cancel, Z3F2 -- khong xac nhan duoc trang thai ra hoa don)</div>
    $serverRowsHtml
  </div>

  <div class="panel">
    <details>
      <summary class="panel-title" style="cursor:pointer;display:inline-block;">&#128200; Lich su service stopped ($ServiceStopIncidentCount lan) -- Click de xem</summary>
      <div style="max-height:400px;overflow-y:auto;margin-top:12px;">
        $ServiceHistoryRowsHtml
      </div>
    </details>
  </div>

  <div class="panel">
    <div class="panel-title">&#128202; Billing hom nay theo BU</div>
    <div style="font-size:11px;color:#8c959f;margin:-8px 0 12px 0;">Chi tinh: Billing Ban, Billing Smallunit (khong gom PXK, Return, Dieu Chinh, Cancel, Z3F2 -- khong xac nhan duoc trang thai ra hoa don)</div>
    $buRowsHtml
  </div>

  <div class="panel">
    <div class="panel-title">&#128196; Billing hom nay theo loai chung tu</div>
    $docTypeRowsHtml
  </div>

  <div class="panel">
    <div class="panel-title">&#9203; Billing chua ra hoa don ($PendingTotal)</div>
    <div class="ls-filelist" style="max-height:500px;">
      $pendingRowsHtml
    </div>
  </div>

  <div class="panel">
    <details>
      <summary class="panel-title" style="cursor:pointer;display:inline-block;">&#8987; Da ra hoa don nhung chua tinh duoc thoi gian xu ly ($UnmatchedIssuedCount, bam de mo)</summary>
      <div class="ls-filelist" style="max-height:400px;margin-top:12px;font-size:12px;">
        ###UNMATCHED_TIME_LIST_PLACEHOLDER###
      </div>
    </details>
  </div>

  <div class="panel" style="border:1.5px solid #ffd8b0;">
    <div class="panel-title" style="color:#bc4c00;">&#9888;&#65039; Billing Ban (ZF2) tu output ZVRD chua ra hoa don ($PendingZF2Total)</div>
    <div class="ls-filelist" style="max-height:400px;">
      $pendingZF2RowsHtml
    </div>
  </div>

  <div class="panel">
    <div class="panel-title">&#128260; So do quy trinh</div>
    $workflowSvg
  </div>

  <div class="panel">
    <div class="panel-title">&#127760; Tra cuu hoa don tren VNPT</div>
    <div style="display:flex;gap:10px;flex-wrap:wrap;">
      <a href="https://vnbizbox.vnpt-invoice.com.vn" target="_blank" style="display:inline-block;background:#0969da;color:white;text-decoration:none;padding:8px 16px;border-radius:6px;font-size:14px;">DKSH</a>
      <a href="https://0316942021-tt78.vnpt-invoice.com.vn" target="_blank" style="display:inline-block;background:#0969da;color:white;text-decoration:none;padding:8px 16px;border-radius:6px;font-size:14px;">Meta HCM</a>
      <a href="https://metahealthcaredanang-tt78.vnpt-invoice.com.vn" target="_blank" style="display:inline-block;background:#0969da;color:white;text-decoration:none;padding:8px 16px;border-radius:6px;font-size:14px;">Meta Da Nang</a>
    </div>
  </div>

  <div class="panel">
    <details>
      <summary class="panel-title" style="cursor:pointer;display:inline-block;">&#128220; Bang tra ma loi VNPT (bam de mo)</summary>
      <div style="max-height:320px;overflow-y:auto;margin-top:12px;">
        <table style="width:100%;border-collapse:collapse;">
          <thead><tr><th style="text-align:left;padding:6px 8px;font-size:12px;color:#57606a;border-bottom:1px solid #eaeef2;">Ma loi</th><th style="text-align:left;padding:6px 8px;font-size:12px;color:#57606a;border-bottom:1px solid #eaeef2;">Y nghia</th></tr></thead>
          <tbody style="font-size:13px;">
            $vnptErrRowsHtml
          </tbody>
        </table>
      </div>
    </details>
  </div>

  <div class="panel">
    <div class="panel-title" style="margin-bottom:12px;">&#128203; Ma loi theo tung nghiep vu (Replace / Create / Adjust / Cancel)</div>
    $vnptErrByActionHtml
  </div>

  <div class="summary">
    <div class="card"><div class="num countup" data-target="$okCount" style="color:#1a7f37">0</div><div class="label">OK</div></div>
    <div class="card"><div class="num countup" data-target="$warnCount" style="color:#9a6700">0</div><div class="label">Canh bao</div></div>
    <div class="card"><div class="num countup" data-target="$errCount" style="color:#cf222e">0</div><div class="label">Loi</div></div>
  </div>

  <div class="panel" style="padding:0;overflow:hidden;">
    <table>
      <thead><tr><th>Buoc</th><th>Kiem tra</th><th>Trang thai</th><th>Chi tiet</th><th>Thoi diem</th></tr></thead>
      <tbody>
        $rowsHtml
      </tbody>
    </table>
  </div>
</div>
<script>
var sapFiles = $sapFilesJson;

function animateCountUp() {
  document.querySelectorAll('.countup').forEach(function(el) {
    var target = parseInt(el.getAttribute('data-target'), 10) || 0;
    var start = 0;
    var duration = 700;
    var startTime = null;
    function step(ts) {
      if (!startTime) startTime = ts;
      var progress = Math.min((ts - startTime) / duration, 1);
      var eased = 1 - Math.pow(1 - progress, 3);
      el.textContent = Math.floor(eased * target).toLocaleString('vi-VN');
      if (progress < 1) { requestAnimationFrame(step); } else { el.textContent = target.toLocaleString('vi-VN'); }
    }
    requestAnimationFrame(step);
  });
}
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', animateCountUp);
} else {
  animateCountUp();
}

var stagingFiles = $stagingFilesJson;
var scannedFolders = $scannedFoldersJson;
var bakupFiles = $bakupFilesJson;

function fmtDuration(ms) {
  var totalSec = Math.round(ms / 1000);
  var m = Math.floor(totalSec / 60);
  var s = totalSec % 60;
  return m + ' phut ' + s + ' giay';
}
function parseLocal(s) {
  // Doc chuoi "yyyy-MM-ddTHH:mm:ss" theo dung tung con so, KHONG de trinh duyet tu suy doan mui gio
  var p = s.split(/[-T:]/);
  return new Date(parseInt(p[0],10), parseInt(p[1],10)-1, parseInt(p[2],10), parseInt(p[3],10), parseInt(p[4],10), parseInt(p[5],10));
}
function fmtTime(d) {
  return d.toLocaleString('vi-VN');
}

function stagingLocation(f) {
  // Mo ta vi tri file trong staging: loai (INV / DO-PXK) + hang doi (thu muc con)
  return '<b>' + (f.k || 'Staging') + '</b>' + (f.q ? ' &rarr; ' + (f.e ? 'thu muc con' : 'hang doi') + ' <b>' + f.q + '</b>' : '');
}

function doLiveSearch() {
  var term = document.getElementById('billingSearchBox').value.trim();
  var resultDiv = document.getElementById('billingSearchResult');
  if (term.length < 3) {
    resultDiv.innerHTML = '<div class="ls-empty">Go it nhat 3 ky tu de tim...</div>';
    return;
  }
  var sapMatches = sapFiles.filter(function(f) { return f.n.indexOf(term) !== -1; });
  var stagingMatches = stagingFiles.filter(function(f) { return f.n.indexOf(term) !== -1 || (f.c && f.c.indexOf(term) !== -1); });
  var bakupMatches = bakupFiles.filter(function(f) { return f.n.indexOf(term) !== -1; });

  if (sapMatches.length === 0 && stagingMatches.length === 0 && bakupMatches.length === 0) {
    resultDiv.innerHTML = '<div class="ls-empty">Khong tim thay trong du lieu lan chay nay. Co the billing thuoc ngay khac, hoac thu file run_search.bat de tim sau hon (ca noi dung file).</div>';
    return;
  }

  var html = '<div class="ls-stage-grid">';
  html += '<div class="ls-stage ' + (sapMatches.length ? 'ls-stage-ok' : 'ls-stage-missing') + '"><div class="ls-stage-label">SAP</div><div class="ls-stage-val">' + (sapMatches.length ? 'TIM THAY (' + sapMatches.length + ')' : 'KHONG THAY') + '</div></div>';
  var ppDone = !stagingMatches.length && bakupMatches.length;
  var ppErr = stagingMatches.some(function(f){ return f.e; });
  var ppStuck = ppErr || stagingMatches.some(function(f){ return f.x; });
  var ppLost = !stagingMatches.length && !bakupMatches.length && sapMatches.length;
  var ppClass = stagingMatches.length ? (ppStuck ? 'ls-stage-missing' : 'ls-stage-wait') : (ppDone ? 'ls-stage-ok' : (ppLost ? 'ls-stage-wait' : 'ls-stage-missing'));
  var ppVal = stagingMatches.length ? ((ppErr ? 'BI LOI' : (ppStuck ? 'BI KET' : 'DANG CHO XU LY')) + ' (' + stagingMatches.length + ')') : (ppDone ? 'DA XU LY XONG' : (ppLost ? 'KHONG CON O STAGING' : 'KHONG THAY'));
  html += '<div class="ls-stage ' + ppClass + '"><div class="ls-stage-label">PP Processing</div><div class="ls-stage-val">' + ppVal + '</div></div>';
  html += '<div class="ls-stage ' + (bakupMatches.length ? 'ls-stage-ok' : 'ls-stage-missing') + '"><div class="ls-stage-label">Da ra hoa don</div><div class="ls-stage-val">' + (bakupMatches.length ? 'TIM THAY (' + bakupMatches.length + ')' : 'KHONG THAY') + '</div></div>';
  html += '</div>';

  // Dong ket luan: billing dang o dau / co dang xu ly khong
  var st;
  if (bakupMatches.length) {
    st = { c: 'ls-status-ok', t: '&#9989; DA XU LY XONG -- da co xac nhan ra hoa don.' };
  } else if (ppErr) {
    st = { c: 'ls-status-bad', t: '&#10060; BI LOI -- file dang nam o thu muc loi <b>' + stagingMatches.filter(function(f){return f.e;}).map(stagingLocation).join(', ') + '</b>, can xu ly lai.' };
  } else if (ppStuck) {
    st = { c: 'ls-status-bad', t: '&#9888;&#65039; BI KET -- file nam qua lau o <b>' + stagingMatches.filter(function(f){return f.x;}).map(stagingLocation).join(', ') + '</b>, chua duoc xu ly.' };
  } else if (stagingMatches.length) {
    st = { c: 'ls-status-wait', t: '&#9203; DANG XU LY -- file dang nam o <b>' + stagingMatches.map(stagingLocation).join(', ') + '</b>, cho PP lay xu ly.' };
  } else if (sapMatches.length) {
    var lastPp = sapMatches.filter(function(f){ return f.tpp; }).map(function(f){ return f.tpp; }).sort().pop();
    st = { c: 'ls-status-wait', t: '&#10067; KHONG XAC DINH -- PP da nhan file tu SAP' + (lastPp ? ' luc <b>' + fmtTime(parseLocal(lastPp)) + '</b>' : '') + ' nhung file KHONG con nam trong cac thu muc dang quet va CHUA co xac nhan ra hoa don. '
         + 'Kha nang: dang gui sang VNPT / cho VNPT tra ve, hoac nam o thu muc chua khai bao (VD: staging cua server khac, thu muc Error). Them thu muc do vao <b>search_folders</b> trong config.json de tool tim duoc.' };
  } else {
    st = { c: 'ls-status-wait', t: 'Chi thay o buoc sau, khong thay file goc tu SAP trong ngay nay.' };
  }
  html += '<div class="ls-status ' + st.c + '">' + st.t + '</div>';

  if (!stagingMatches.length && !bakupMatches.length) {
    html += '<details class="ls-scanned"><summary>Cac thu muc da quet (' + scannedFolders.length + ')</summary>';
    scannedFolders.forEach(function(sf) {
      html += '<div><b>' + sf.k + '</b> (' + (sf.ok ? sf.c + ' file' : '<span style="color:#cf222e">khong truy cap duoc</span>') + ') <span class="ls-path" style="display:inline">' + sf.p + '</span></div>';
    });
    html += '</details>';
  }

  if (stagingMatches.length) {
    html += '<div class="ls-location' + (ppStuck ? ' ls-location-stuck' : '') + '">';
    html += '<div style="font-weight:700;margin-bottom:6px;">&#128205; Billing dang nam o:</div>';
    sortByTime(stagingMatches, 't').forEach(function(f) {
      html += '<div style="margin-top:4px;">' + stagingLocation(f) + ' &middot; da cho <b>' + (f.a != null ? f.a + ' phut' : '?') + '</b>'
            + (f.e ? ' <span class="ls-stuck">THU MUC LOI</span>' : (f.x ? ' <span class="ls-stuck">BI KET</span>' : ''))
            + (f.n.indexOf(term) === -1 ? ' <span class="ls-diff-type">tim thay trong noi dung file</span>' : '')
            + '<div class="ls-path">' + (f.p || f.n) + '</div></div>';
    });
    html += '</div>';
  }

  // Tinh thoi gian xu ly qua 3 moc: SAP -> PP xu ly -> Ra hoa don
  // Neu co nhieu lan trigger (nhieu SAP / nhieu Bakup), ghep theo thu tu thoi gian tung cap.
  var sapWithTime = sortByTime(sapMatches.filter(function(f){ return f.t; }), 't');
  var ppWithTime = sapMatches.filter(function(f){ return f.tpp; });
  var bakupSorted = sortByTime(bakupMatches, 't');

  if (sapWithTime.length && bakupSorted.length) {
    var pairCount = Math.max(sapWithTime.length, bakupSorted.length);
    html += '<div class="ls-elapsed">';
    if (pairCount > 1) {
      html += '<div style="font-weight:700;color:#116329;margin-bottom:6px;">Co ' + pairCount + ' lan trigger:</div>';
    }
    var totalDiffSum = 0, totalDiffCount = 0;
    for (var i = 0; i < pairCount; i++) {
      var sapF = sapWithTime[i] || sapWithTime[sapWithTime.length - 1];
      var bakF = bakupSorted[i] || bakupSorted[bakupSorted.length - 1];
      var sTime = parseLocal(sapF.t);
      var bTime = parseLocal(bakF.t);
      var d = bTime.getTime() - sTime.getTime();
      var label = pairCount > 1 ? ('Trigger ' + (i + 1) + ': ') : 'Xuat tu SAP: ';
      html += '<div style="font-size:12px;color:#116329;' + (i > 0 ? 'margin-top:8px;' : '') + '">' + label + '<b>' + fmtTime(sTime) + '</b> &nbsp;&rarr;&nbsp; Ra hoa don: <b>' + fmtTime(bTime) + '</b></div>';
      if (d >= 0) {
        html += '<div style="font-size:' + (pairCount > 1 ? '16px' : '20px') + ';font-weight:800;color:#116329;margin-top:2px;">Mat: ' + fmtDuration(d) + '</div>';
        totalDiffSum += d; totalDiffCount++;
      } else {
        html += '<div style="font-size:13px;color:#9a6700;margin-top:2px;">(Ra hoa don truoc SAP -- co the khac cap trigger)</div>';
      }
    }
    if (pairCount > 1 && totalDiffCount > 0) {
      html += '<div style="font-size:12px;color:#57606a;margin-top:8px;border-top:1px solid #cfe0fb;padding-top:6px;">Trung binh: ' + fmtDuration(totalDiffSum / totalDiffCount) + '</div>';
    }
    html += '</div>';
  } else if (sapMatches.length && !bakupMatches.length) {
    if (sapWithTime.length) {
      var onlySapTime = parseLocal(sapWithTime[sapWithTime.length - 1].t);
      var minutesAgo = Math.round((Date.now() - onlySapTime.getTime()) / 60000);
      html += '<div class="ls-elapsed">';
      html += '<div style="font-size:12px;color:#116329;">Xuat tu SAP luc: <b>' + fmtTime(onlySapTime) + '</b></div>';
      html += '<div style="font-size:14px;font-weight:700;color:#9a6700;margin-top:4px;">Da cho: khoang ' + (minutesAgo >= 0 ? minutesAgo + ' phut' : 'khong xac dinh') + '</div>';
      html += '</div>';
    }
  }

  html += '<div class="ls-filelist">';
  function sortByTime(list, timeField) {
    return list.slice().sort(function(a, b) {
      var ta = a[timeField] ? parseLocal(a[timeField]).getTime() : 0;
      var tb = b[timeField] ? parseLocal(b[timeField]).getTime() : 0;
      return ta - tb;
    });
  }
  function triggerLabel(list, idx, item) {
    if (list.length <= 1) return '';
    // Chi goi la "Trigger" khi cung 1 loai chung tu (doc type) -- neu khac loai (VD: 1 cai BVK/Smallunit,
    // 1 cai thuong) thi khong phai thu lai, ma la 2 luong khac nhau -- khong gan nhan Trigger.
    if (item && item.d) {
      var sameTypeList = list.filter(function(x){ return x.d === item.d; });
      if (sameTypeList.length <= 1) return ' <span class="ls-diff-type">Rieng (' + item.d + ')</span>';
      var sameTypeIdx = sameTypeList.indexOf(item);
      return ' <span class="ls-trigger">Trigger lan ' + (sameTypeIdx + 1) + '</span>';
    }
    return ' <span class="ls-trigger">Trigger lan ' + (idx + 1) + '</span>';
  }

  sortByTime(sapMatches, 't').slice(0, 5).forEach(function(f, idx) {
    var timeStr = f.t ? fmtTime(parseLocal(f.t)) : '';
    html += '<div class="ls-row"><span class="ls-badge ls-found">SAP (' + f.s + (f.b ? (' - ' + f.b) : '') + ')</span>' + (f.d ? '<span class="ls-badge ls-doctype">' + f.d + '</span>' : '') + '<span class="ls-name">' + f.n + (timeStr ? ' <b>[' + timeStr + ']</b>' : '') + '</span>' + triggerLabel(sapMatches, idx, f) + '</div>';
  });
  sortByTime(stagingMatches, 't').slice(0, 5).forEach(function(f, idx) {
    var timeStr = f.t ? fmtTime(parseLocal(f.t)) : '';
    html += '<div class="ls-row"><span class="ls-badge ' + (f.x ? 'ls-missing' : 'ls-found') + '">PP Processing' + (f.k ? (' - ' + f.k) : '') + (f.q ? ((f.e ? ' - thu muc con ' : ' - hang doi ') + f.q) : '') + '</span><span class="ls-name">' + f.n + (timeStr ? ' <b>[' + timeStr + ']</b>' : '') + '</span>' + triggerLabel(stagingMatches, idx) + '</div>';
  });
  sortByTime(bakupMatches, 't').slice(0, 5).forEach(function(f, idx) {
    var timeStr = f.t ? fmtTime(parseLocal(f.t)) : '';
    html += '<div class="ls-row"><span class="ls-badge ls-found">Da ra hoa don (' + f.s + ')</span><span class="ls-name">' + f.n + (timeStr ? ' <b>[' + timeStr + ']</b>' : '') + '</span>' + triggerLabel(bakupMatches, idx) + '</div>';
  });
  html += '</div>';

  html += '<button onclick="downloadSearchResult(\'' + term.replace(/'/g, "\\'") + '\')" style="margin-top:12px;background:#0969da;color:white;border:none;padding:8px 16px;border-radius:6px;font-size:13px;font-weight:600;cursor:pointer;">&#128190; Tai ket qua tim kiem (.txt)</button>';

  resultDiv.innerHTML = html;
  window._lastSearch = { term: term, sap: sapMatches, staging: stagingMatches, bakup: bakupMatches };
}

function downloadSearchResult(term) {
  var d = window._lastSearch;
  if (!d) return;
  var lines = [];
  lines.push('KET QUA TIM KIEM BILLING: ' + term);
  lines.push('Xuat luc: ' + new Date().toLocaleString('vi-VN'));
  lines.push('=======================================================');
  lines.push('');
  lines.push('--- SAP (' + d.sap.length + ' file) ---');
  d.sap.forEach(function(f) {
    lines.push((f.t ? '[' + fmtTime(parseLocal(f.t)) + '] ' : '') + f.s + (f.b ? ' - ' + f.b : '') + (f.d ? ' - ' + f.d : '') + ' - ' + f.n);
  });
  lines.push('');
  lines.push('--- PP Processing (' + d.staging.length + ' file) ---');
  d.staging.forEach(function(f) {
    lines.push((f.t ? '[' + fmtTime(parseLocal(f.t)) + '] ' : '') + stagingLocation(f).replace(/<[^>]+>/g, '') + (f.a != null ? ' - da cho ' + f.a + ' phut' : '') + (f.e ? ' - THU MUC LOI' : (f.x ? ' - BI KET' : '')) + ' - ' + (f.p || f.n));
  });
  lines.push('');
  lines.push('--- Da ra hoa don / Bakup (' + d.bakup.length + ' file) ---');
  d.bakup.forEach(function(f) {
    lines.push((f.t ? '[' + fmtTime(parseLocal(f.t)) + '] ' : '') + f.s + ' - ' + f.n);
  });
  var blob = new Blob([lines.join('\n')], { type: 'text/plain;charset=utf-8' });
  var url = URL.createObjectURL(blob);
  var a = document.createElement('a');
  a.href = url;
  a.download = 'ket_qua_tim_kiem_' + term.replace(/[^a-zA-Z0-9]/g, '') + '.txt';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

// ============================================================
// Bang: Billing xu ly hon N phut
// ============================================================
var slowThreshold = $SlowThresholdMinutes;
var slowRunTime = parseLocal('$runTimeIso');
var slowExcluded = $excludedTypesJson;
var slowData = { pending: [], done: [] };
var slowTab = 'pending';

function buildSlowList() {
  var bnRe = /(\d{10})/;
  // Gio ra hoa don theo so billing
  var bakByBn = {};
  bakupFiles.forEach(function(f) {
    var m = bnRe.exec(f.n); if (!m || !f.t) return;
    (bakByBn[m[1]] = bakByBn[m[1]] || []).push(parseLocal(f.t));
  });
  // Vi tri hien tai theo so billing (staging / thu muc loi / PRODATA...)
  var locByBn = {};
  stagingFiles.forEach(function(f) {
    var list = (f.c ? f.c.split(',') : []);
    var m = bnRe.exec(f.n); if (m) list.push(m[1]);
    list.forEach(function(bn) { if (bn && !(bn in locByBn)) locByBn[bn] = f; });
  });
  // Gom file SAP theo so billing, lay gio SAP som nhat
  var sapByBn = {};
  sapFiles.forEach(function(f) {
    if (!f.bn || !f.t) return;
    if (slowExcluded.indexOf(f.d) !== -1) return;
    var t = parseLocal(f.t);
    var cur = sapByBn[f.bn];
    if (!cur || t < cur.t) sapByBn[f.bn] = { bn: f.bn, t: t, d: f.d, s: f.s, b: f.b };
  });
  slowData = { pending: [], done: [], issuedTotal: 0 };
  Object.keys(sapByBn).forEach(function(bn) {
    var r = sapByBn[bn];
    var baks = (bakByBn[bn] || []).filter(function(b) { return b >= r.t; }).sort(function(a, b) { return a - b; });
    if (baks.length) {
      slowData.issuedTotal++;
      var mins = Math.round((baks[0] - r.t) / 60000);
      if (mins > slowThreshold) slowData.done.push({ bn: bn, d: r.d, s: r.s, b: r.b, t: r.t, end: baks[0], mins: mins });
    } else if ((bakByBn[bn] || []).length === 0) {
      var age = Math.round((slowRunTime - r.t) / 60000);
      if (age > slowThreshold) slowData.pending.push({ bn: bn, d: r.d, s: r.s, b: r.b, t: r.t, mins: age, loc: locByBn[bn] || null });
    }
  });
  slowData.pending.sort(function(a, b) { return b.mins - a.mins; });
  slowData.done.sort(function(a, b) { return b.mins - a.mins; });
  document.getElementById('slowCntPending').textContent = slowData.pending.length;
  document.getElementById('slowCntDone').textContent = slowData.done.length;
  var tot = slowData.issuedTotal, nd = slowData.done.length;
  var pct = tot ? Math.round(nd * 1000 / tot) / 10 : 0;
  document.getElementById('slowStats').innerHTML =
      '<div class="slow-stat" style="background:#fff1e5"><div class="v" style="color:#bc4c00">' + nd + ' <span style="font-size:15px;color:#57606a">/ ' + tot + '</span></div>'
    + '<div class="l">Hoa don DA RA nhung mat hon ' + slowThreshold + ' phut (' + pct + '%)</div></div>'
    + '<div class="slow-stat" style="background:#dafbe1"><div class="v" style="color:#1a7f37">' + (tot - nd) + '</div><div class="l">Hoa don da ra trong vong ' + slowThreshold + ' phut</div></div>'
    + '<div class="slow-stat" style="background:#ffebe9"><div class="v" style="color:#cf222e">' + slowData.pending.length + '</div><div class="l">Chua ra hoa don, da cho hon ' + slowThreshold + ' phut</div></div>';
}

function fmtMins(m) {
  if (m < 60) return m + ' phut';
  return Math.floor(m / 60) + ' gio ' + (m % 60) + ' phut';
}
function slowLocText(r) {
  if (!r.loc) return '<span style="color:#9a6700">Khong thay o thu muc nao dang quet (co the dang gui VNPT / cho xac nhan)</span>';
  var f = r.loc;
  var col = f.e ? '#cf222e' : (f.x ? '#cf222e' : '#9a6700');
  return '<span style="color:' + col + '">' + stagingLocation(f) + (f.e ? ' <span class="ls-stuck">THU MUC LOI</span>' : (f.x ? ' <span class="ls-stuck">BI KET</span>' : '')) + '</span>';
}

function showSlowTab(tab) {
  slowTab = tab;
  document.getElementById('slowTabPending').className = 'slow-tab' + (tab === 'pending' ? ' slow-tab-active' : '');
  document.getElementById('slowTabDone').className = 'slow-tab' + (tab === 'done' ? ' slow-tab-active' : '');
  var rows = slowData[tab];
  var wrap = document.getElementById('slowTableWrap');
  if (!rows.length) {
    wrap.innerHTML = '<div class="ls-empty" style="padding:14px;">' + (tab === 'pending' ? '&#9989; Khong co billing nao cho qua ' + slowThreshold + ' phut.' : 'Khong co billing nao mat hon ' + slowThreshold + ' phut de ra hoa don.') + '</div>';
    return;
  }
  var h = '<table><thead><tr><th>#</th><th>So billing</th><th>Loai</th><th>Server - BU</th><th>Xuat SAP</th><th>'
        + (tab === 'pending' ? 'Dang nam o' : 'Ra hoa don') + '</th><th>' + (tab === 'pending' ? 'Da cho' : 'Mat') + '</th></tr></thead><tbody>';
  rows.forEach(function(r, i) {
    var durCol = r.mins > slowThreshold * 4 ? '#cf222e' : '#9a6700';
    h += '<tr><td>' + (i + 1) + '</td>'
       + '<td><span class="slow-bn" onclick="jumpToBilling(\'' + r.bn + '\')">' + r.bn + '</span></td>'
       + '<td>' + (r.d || '') + '</td>'
       + '<td>' + (r.s || '') + (r.b ? ' - ' + r.b : '') + '</td>'
       + '<td>' + fmtTime(r.t) + '</td>'
       + '<td>' + (tab === 'pending' ? slowLocText(r) : fmtTime(r.end)) + '</td>'
       + '<td class="slow-dur" style="color:' + durCol + '">' + fmtMins(r.mins) + '</td></tr>';
  });
  h += '</tbody></table>';
  wrap.innerHTML = h;
}

function jumpToBilling(bn) {
  var box = document.getElementById('billingSearchBox');
  box.value = bn;
  doLiveSearch();
  box.scrollIntoView({ behavior: 'smooth', block: 'center' });
}

function downloadSlowList() {
  var lines = ['Trang thai,So billing,Loai,Server,BU,Xuat SAP,Ra hoa don / Dang nam o,Phut'];
  function q(v) { return '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"'; }
  slowData.pending.forEach(function(r) {
    var loc = r.loc ? ((r.loc.k || '') + (r.loc.q ? ' / ' + r.loc.q : '') + (r.loc.e ? ' (THU MUC LOI)' : '') + ' - ' + (r.loc.p || r.loc.n)) : 'Khong thay o thu muc dang quet';
    lines.push([q('Chua xong'), q(r.bn), q(r.d), q(r.s), q(r.b), q(fmtTime(r.t)), q(loc), r.mins].join(','));
  });
  slowData.done.forEach(function(r) {
    lines.push([q('Da xong - cham'), q(r.bn), q(r.d), q(r.s), q(r.b), q(fmtTime(r.t)), q(fmtTime(r.end)), r.mins].join(','));
  });
  var blob = new Blob(['﻿' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' });
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'billing_xu_ly_hon_' + slowThreshold + '_phut.csv';
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
}

try { buildSlowList(); showSlowTab('pending'); } catch (e) {
  document.getElementById('slowTableWrap').innerHTML = '<div class="ls-empty" style="padding:14px;">Loi khi tinh danh sach: ' + e.message + '</div>';
}
</script>
<div style="text-align:center;color:#8c959f;font-size:12px;padding:20px 0 8px 0;">
  Billing Pipeline Monitor &middot; Tac gia: <b>thanhnv</b>
</div>
</body>
</html>
"@

# Thay placeholder bang danh sach thuc te (dung .Replace() thay vi interpolation
# de tranh cac loi la ve here-string voi noi dung dai/phuc tap)
$html = $html.Replace("###UNMATCHED_TIME_LIST_PLACEHOLDER###", $unmatchedTimeRowsHtml)

$dashOutputPath = $Config.dashboard.output_path
if ($CheckDate -ne (Get-Date -Format "yyyyMMdd")) {
    # Ngay qua khu -- ghi ra file rieng, khong de len dashboard hom nay
    $ext = [System.IO.Path]::GetExtension($dashOutputPath)
    $base = $dashOutputPath.Substring(0, $dashOutputPath.Length - $ext.Length)
    $dashOutputPath = "${base}_${CheckDate}${ext}"
}
if ($dashOutputPath -match '^[A-Za-z]:\\' -or $dashOutputPath -match '^\\\\') {
    # Duong dan tuyet doi (o dia C:\... hoac \\server\share\...) -- dung thang, khong ghep voi ScriptDir
    $dashboardPath = $dashOutputPath
    $dashOutputDir = Split-Path $dashboardPath -Parent
    if ($dashOutputDir -and -not (Test-Path $dashOutputDir -ErrorAction SilentlyContinue)) {
        try { New-Item -ItemType Directory -Path $dashOutputDir -Force -ErrorAction Stop | Out-Null } catch {}
    }
} else {
    $dashboardPath = Join-Path $ScriptDir $dashOutputPath
}
$html | Out-File -FilePath $dashboardPath -Encoding UTF8
# Ghi lai duong dan thuc te vua xuat ra, de run.bat biet chinh xac file nao can mo
# (tranh truong hop output_path da doi sang noi khac nhung run.bat van mo file cu).
# Dung ASCII (khong BOM) de batch (set /p) doc dung, khong bi lech ky tu dau do BOM cua UTF8
[System.IO.File]::WriteAllText((Join-Path $ScriptDir "dashboard_path.txt"), $dashboardPath, [System.Text.Encoding]::ASCII)

Write-Host ""
Write-Host "Tong ket: $okCount OK, $warnCount canh bao, $errCount loi"
Write-Host "Dashboard da cap nhat: $dashboardPath"
