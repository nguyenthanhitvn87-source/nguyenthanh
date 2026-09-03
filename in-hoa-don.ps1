<#
    IN HÓA ĐƠN HÀNG LOẠT — Nguyễn Thanh
    ------------------------------------------------------------------
      • Quét một thư mục (kể cả thư mục con), tách số hóa đơn trong tên file.
      • Sắp xếp theo số hóa đơn, tên file hoặc ngày sửa; hoặc nạp file
        "thu-tu-in.txt" do công cụ sắp xếp tạo ra để in đúng thứ tự trong Excel.
      • Tick chọn từng hóa đơn, tìm nhanh, hoặc chọn theo khoảng số.
      • In lần lượt ra máy in đã chọn, đúng số bản, đúng thứ tự; có nút Dừng
        và chế độ in thử để chạy nháp mà không tốn giấy.

    Cách chạy:
        Nhấp đúp vào "in-hoa-don.bat"
        hoặc: powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\in-hoa-don.ps1

    Sắp xếp hóa đơn vào thư mục theo danh sách Excel là công cụ riêng:
    "sap-xep-hoa-don.ps1".

    Yêu cầu: Windows 7 trở lên, Windows PowerShell 5.1 (có sẵn trong Windows).
             Máy phải có ứng dụng mở được loại file đó (PDF cần Adobe Reader
             hoặc phần mềm đọc PDF hỗ trợ lệnh in của Windows).
#>

#Requires -Version 3.0

# ----------------------------------------------------------------------------
#  BẮT LỖI: có sự cố thì ghi ra "loi-chay.txt" cạnh script và hiện lên màn hình,
#  thay vì để cửa sổ tắt mất không kịp đọc.
# ----------------------------------------------------------------------------
trap {
    $err = $_

    $apartment = 'không rõ'
    try { $apartment = [System.Threading.Thread]::CurrentThread.GetApartmentState() } catch { }
    $where = ''
    try { $where = $err.InvocationInfo.PositionMessage } catch { }
    $kind = ''
    try { $kind = $err.Exception.GetType().FullName } catch { }

    $info = @(
        'LỖI KHI CHẠY CÔNG CỤ',
        ('Lúc: ' + (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')),
        ('PowerShell: ' + $PSVersionTable.PSVersion),
        ('Chế độ luồng: ' + $apartment),
        '',
        ('Nội dung lỗi: ' + $err.Exception.Message),
        '',
        ('Chỗ lỗi: ' + $where),
        '',
        ('Loại lỗi: ' + $kind)
    ) -join [Environment]::NewLine

    try {
        $root = (Get-Location).Path
        if ($PSScriptRoot) { $root = $PSScriptRoot }
        [System.IO.File]::WriteAllText((Join-Path $root 'loi-chay.txt'), $info, (New-Object System.Text.UTF8Encoding($true)))
    } catch { }

    try {
        Write-Host ''
        Write-Host $info -ForegroundColor Red
        Write-Host ''
        Write-Host 'Nội dung trên đã được ghi vào file loi-chay.txt cạnh script — gửi file đó là biết lỗi gì.' -ForegroundColor Yellow
    } catch { }

    try { [void][System.Windows.Forms.MessageBox]::Show($info, 'Lỗi khi chạy công cụ', 'OK', 'Error') } catch { }
    try { [void](Read-Host 'Nhấn Enter để đóng') } catch { }
    exit 1
}

# ----------------------------------------------------------------------------
#  NHẬT KÝ KHỞI ĐỘNG: ghi từng bước ra "khoi-dong.txt" cạnh script, để khi cửa sổ
#  tắt mất vẫn biết công cụ chạy tới đâu thì hỏng.
# ----------------------------------------------------------------------------
$script:TraceFile = $null
try {
    $traceRoot = (Get-Location).Path
    if ($PSScriptRoot) { $traceRoot = $PSScriptRoot }
    $script:TraceFile = Join-Path $traceRoot 'khoi-dong.txt'
    [System.IO.File]::WriteAllText($script:TraceFile, '', (New-Object System.Text.UTF8Encoding($true)))
} catch { $script:TraceFile = $null }

function Write-Trace {
    param([string]$Step)
    if (-not $script:TraceFile) { return }
    try {
        $line = '{0}  {1}{2}' -f (Get-Date -Format 'HH:mm:ss'), $Step, [Environment]::NewLine
        [System.IO.File]::AppendAllText($script:TraceFile, $line, (New-Object System.Text.UTF8Encoding($true)))
    } catch { }
}

Write-Trace ('Bắt đầu — PowerShell ' + $PSVersionTable.PSVersion)


Write-Trace 'Nạp thư viện giao diện Windows'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()
Write-Trace 'Đã nạp xong thư viện giao diện'

# ============================================================================
#  BIẾN DÙNG CHUNG
# ============================================================================
$script:AllFiles     = @()
$script:CheckedPaths = New-Object 'System.Collections.Generic.HashSet[string]'
$script:Populating   = $false
$script:Cancel       = $false
$script:Busy         = $false

$script:FileGroups = [ordered]@{
    'PDF'                = @('.pdf')
    'Word (.doc/.docx)'  = @('.doc', '.docx', '.rtf')
    'Excel (.xls/.xlsx)' = @('.xls', '.xlsx', '.xlsm', '.csv')
    'Ảnh (JPG/PNG/TIF)'  = @('.jpg', '.jpeg', '.png', '.tif', '.tiff', '.bmp')
    'Tất cả các file'    = @()
}

# ============================================================================
#  HÀM PHỤ TRỢ
# ============================================================================
function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '[{0}] {1}  {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level.PadRight(4), $Message
    $txtLog.AppendText($line + [Environment]::NewLine)
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Start-UiSleep {
    <# Nghỉ nhưng vẫn để giao diện phản hồi (còn bấm được nút Dừng). #>
    param([int]$Milliseconds)
    $end = (Get-Date).AddMilliseconds($Milliseconds)
    while ((Get-Date) -lt $end) {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 120
        if ($script:Cancel) { break }
    }
}

# ============================================================================
#  IN
# ============================================================================
function Get-PrinterNames {
    try   { return @(Get-CimInstance -ClassName Win32_Printer -ErrorAction Stop | Select-Object -ExpandProperty Name) }
    catch {
        try   { return @(Get-WmiObject -Class Win32_Printer -ErrorAction Stop | Select-Object -ExpandProperty Name) }
        catch { return @() }
    }
}

function Get-DefaultPrinterName {
    try {
        $p = Get-CimInstance -ClassName Win32_Printer -Filter 'Default = TRUE' -ErrorAction Stop | Select-Object -First 1
        if ($p) { return $p.Name }
    } catch { }
    return $null
}

function Set-DefaultPrinterName {
    param([string]$Name)
    try {
        (New-Object -ComObject WScript.Network).SetDefaultPrinter($Name)
        return $true
    } catch { return $false }
}

function Wait-PrintQueue {
    <# Chờ hàng đợi máy in vơi bớt để các hóa đơn ra đúng thứ tự. #>
    param([string]$PrinterName, [int]$TimeoutSeconds = 60)
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        if ($script:Cancel) { return }
        try {
            $jobs = @(Get-CimInstance -ClassName Win32_PrintJob -ErrorAction Stop |
                      Where-Object { $_.Name -like ('{0},*' -f $PrinterName) }).Count
        } catch { return }
        if ($jobs -le 0) { return }
        Start-UiSleep -Milliseconds 700
    }
}

function Invoke-PrintOneFile {
    <# Gửi một file ra máy in: ưu tiên lệnh PrintTo (chỉ định đúng máy in),
       nếu ứng dụng không hỗ trợ thì dùng lệnh Print với máy in mặc định. #>
    param([string]$Path, [string]$PrinterName, [int]$WaitSeconds = 6, [bool]$CloseApp = $true)

    $proc = $null
    $verb = 'PrintTo'
    try {
        $proc = Start-Process -FilePath $Path -Verb PrintTo -ArgumentList ('"{0}"' -f $PrinterName) `
                              -PassThru -WindowStyle Hidden -ErrorAction Stop
    } catch {
        $verb = 'Print'
        $proc = Start-Process -FilePath $Path -Verb Print -PassThru -WindowStyle Hidden -ErrorAction Stop
    }

    Start-UiSleep -Milliseconds ($WaitSeconds * 1000)

    if ($CloseApp -and $proc -and -not $proc.HasExited) {
        try {
            [void]$proc.CloseMainWindow()
            Start-UiSleep -Milliseconds 1200
            if (-not $proc.HasExited) { $proc.Kill() }
        } catch { }
    }
    return $verb
}

Write-Trace 'Bắt đầu dựng cửa sổ'

# ============================================================================
#  GIAO DIỆN
# ============================================================================
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'In hóa đơn hàng loạt — Nguyễn Thanh'
$form.Size          = New-Object System.Drawing.Size(1024, 800)
$form.MinimumSize   = New-Object System.Drawing.Size(1024, 740)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

$grpFolder          = New-Object System.Windows.Forms.GroupBox
$grpFolder.Text     = '1. Thư mục chứa hóa đơn cần in'
$grpFolder.Location = New-Object System.Drawing.Point(8, 6)
$grpFolder.Size     = New-Object System.Drawing.Size(984, 120)
$form.Controls.Add($grpFolder)

$lblFolder          = New-Object System.Windows.Forms.Label
$lblFolder.Text     = 'Thư mục'
$lblFolder.Location = New-Object System.Drawing.Point(12, 28)
$lblFolder.Size     = New-Object System.Drawing.Size(70, 20)
$grpFolder.Controls.Add($lblFolder)

$txtFolder          = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(86, 25)
$txtFolder.Size     = New-Object System.Drawing.Size(660, 24)
$grpFolder.Controls.Add($txtFolder)

$btnBrowse          = New-Object System.Windows.Forms.Button
$btnBrowse.Text     = 'Chọn...'
$btnBrowse.Location = New-Object System.Drawing.Point(752, 23)
$btnBrowse.Size     = New-Object System.Drawing.Size(110, 28)
$grpFolder.Controls.Add($btnBrowse)

$btnScan          = New-Object System.Windows.Forms.Button
$btnScan.Text     = 'Quét folder'
$btnScan.Location = New-Object System.Drawing.Point(868, 23)
$btnScan.Size     = New-Object System.Drawing.Size(104, 28)
$grpFolder.Controls.Add($btnScan)

$lblType          = New-Object System.Windows.Forms.Label
$lblType.Text     = 'Loại file'
$lblType.Location = New-Object System.Drawing.Point(12, 62)
$lblType.Size     = New-Object System.Drawing.Size(60, 20)
$grpFolder.Controls.Add($lblType)

$cboType               = New-Object System.Windows.Forms.ComboBox
$cboType.Location      = New-Object System.Drawing.Point(76, 59)
$cboType.Size          = New-Object System.Drawing.Size(160, 24)
$cboType.DropDownStyle = 'DropDownList'
[void]$cboType.Items.AddRange(@($script:FileGroups.Keys))
$cboType.SelectedIndex = 0
$grpFolder.Controls.Add($cboType)

$chkSub          = New-Object System.Windows.Forms.CheckBox
$chkSub.Text     = 'Gồm thư mục con'
$chkSub.Location = New-Object System.Drawing.Point(248, 61)
$chkSub.Size     = New-Object System.Drawing.Size(150, 22)
$chkSub.Checked  = $true
$grpFolder.Controls.Add($chkSub)

$lblSort          = New-Object System.Windows.Forms.Label
$lblSort.Text     = 'Sắp xếp'
$lblSort.Location = New-Object System.Drawing.Point(412, 62)
$lblSort.Size     = New-Object System.Drawing.Size(55, 20)
$grpFolder.Controls.Add($lblSort)

$cboSort               = New-Object System.Windows.Forms.ComboBox
$cboSort.Location      = New-Object System.Drawing.Point(470, 59)
$cboSort.Size          = New-Object System.Drawing.Size(276, 24)
$cboSort.DropDownStyle = 'DropDownList'
[void]$cboSort.Items.AddRange(@(
    'Số hóa đơn tăng dần (1, 2, 3...)',
    'Số hóa đơn giảm dần',
    'Tên file A → Z (đúng thứ tự đã đánh số)',
    'Tên file Z → A',
    'Ngày sửa: cũ trước',
    'Ngày sửa: mới trước',
    'Theo danh sách thứ tự in đã nạp'
))
$cboSort.SelectedIndex = 0
$grpFolder.Controls.Add($cboSort)

$lblRegex          = New-Object System.Windows.Forms.Label
$lblRegex.Text     = 'Mẫu số hóa đơn (regex)'
$lblRegex.Location = New-Object System.Drawing.Point(12, 94)
$lblRegex.Size     = New-Object System.Drawing.Size(148, 20)
$grpFolder.Controls.Add($lblRegex)

$txtRegex          = New-Object System.Windows.Forms.TextBox
$txtRegex.Location = New-Object System.Drawing.Point(164, 91)
$txtRegex.Size     = New-Object System.Drawing.Size(140, 24)
$grpFolder.Controls.Add($txtRegex)

$lblRegexHint           = New-Object System.Windows.Forms.Label
$lblRegexHint.Text      = 'Để trống = lấy dãy số cuối trong tên file (001_K25TAA_618585.pdf → 618585). Ví dụ khác: K\d{2}[A-Z]+.?(\d+)'
$lblRegexHint.Location  = New-Object System.Drawing.Point(314, 94)
$lblRegexHint.Size      = New-Object System.Drawing.Size(420, 20)
$lblRegexHint.ForeColor = [System.Drawing.Color]::DimGray
$grpFolder.Controls.Add($lblRegexHint)

$btnLoadOrder          = New-Object System.Windows.Forms.Button
$btnLoadOrder.Text     = 'Nạp danh sách thứ tự in...'
$btnLoadOrder.Location = New-Object System.Drawing.Point(742, 89)
$btnLoadOrder.Size     = New-Object System.Drawing.Size(230, 28)
$grpFolder.Controls.Add($btnLoadOrder)

$grpPick          = New-Object System.Windows.Forms.GroupBox
$grpPick.Text     = '2. Chọn hóa đơn cần in'
$grpPick.Location = New-Object System.Drawing.Point(8, 132)
$grpPick.Size     = New-Object System.Drawing.Size(984, 330)
$form.Controls.Add($grpPick)

$lblSearch          = New-Object System.Windows.Forms.Label
$lblSearch.Text     = 'Tìm nhanh'
$lblSearch.Location = New-Object System.Drawing.Point(12, 26)
$lblSearch.Size     = New-Object System.Drawing.Size(70, 20)
$grpPick.Controls.Add($lblSearch)

$txtSearch          = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(85, 23)
$txtSearch.Size     = New-Object System.Drawing.Size(190, 24)
$grpPick.Controls.Add($txtSearch)

$lblFrom          = New-Object System.Windows.Forms.Label
$lblFrom.Text     = 'Từ số'
$lblFrom.Location = New-Object System.Drawing.Point(292, 26)
$lblFrom.Size     = New-Object System.Drawing.Size(45, 20)
$grpPick.Controls.Add($lblFrom)

$txtFrom          = New-Object System.Windows.Forms.TextBox
$txtFrom.Location = New-Object System.Drawing.Point(338, 23)
$txtFrom.Size     = New-Object System.Drawing.Size(75, 24)
$grpPick.Controls.Add($txtFrom)

$lblTo          = New-Object System.Windows.Forms.Label
$lblTo.Text     = 'đến số'
$lblTo.Location = New-Object System.Drawing.Point(420, 26)
$lblTo.Size     = New-Object System.Drawing.Size(48, 20)
$grpPick.Controls.Add($lblTo)

$txtTo          = New-Object System.Windows.Forms.TextBox
$txtTo.Location = New-Object System.Drawing.Point(470, 23)
$txtTo.Size     = New-Object System.Drawing.Size(75, 24)
$grpPick.Controls.Add($txtTo)

$btnRange          = New-Object System.Windows.Forms.Button
$btnRange.Text     = 'Chọn theo khoảng'
$btnRange.Location = New-Object System.Drawing.Point(555, 21)
$btnRange.Size     = New-Object System.Drawing.Size(140, 28)
$grpPick.Controls.Add($btnRange)

$btnAll          = New-Object System.Windows.Forms.Button
$btnAll.Text     = 'Chọn hết'
$btnAll.Location = New-Object System.Drawing.Point(701, 21)
$btnAll.Size     = New-Object System.Drawing.Size(80, 28)
$grpPick.Controls.Add($btnAll)

$btnNone          = New-Object System.Windows.Forms.Button
$btnNone.Text     = 'Bỏ chọn'
$btnNone.Location = New-Object System.Drawing.Point(787, 21)
$btnNone.Size     = New-Object System.Drawing.Size(80, 28)
$grpPick.Controls.Add($btnNone)

$btnInvert          = New-Object System.Windows.Forms.Button
$btnInvert.Text     = 'Đảo chọn'
$btnInvert.Location = New-Object System.Drawing.Point(873, 21)
$btnInvert.Size     = New-Object System.Drawing.Size(99, 28)
$grpPick.Controls.Add($btnInvert)

$lvFiles               = New-Object System.Windows.Forms.ListView
$lvFiles.Location      = New-Object System.Drawing.Point(12, 56)
$lvFiles.Size          = New-Object System.Drawing.Size(960, 230)
$lvFiles.View          = 'Details'
$lvFiles.CheckBoxes    = $true
$lvFiles.FullRowSelect = $true
$lvFiles.GridLines     = $true
$lvFiles.HideSelection = $false
[void]$lvFiles.Columns.Add('STT', 45)
[void]$lvFiles.Columns.Add('Số HĐ', 90)
[void]$lvFiles.Columns.Add('Tên file', 430)
[void]$lvFiles.Columns.Add('Ngày sửa', 135)
[void]$lvFiles.Columns.Add('Dung lượng', 90)
[void]$lvFiles.Columns.Add('Thư mục', 145)
$grpPick.Controls.Add($lvFiles)

$lblCount          = New-Object System.Windows.Forms.Label
$lblCount.Text     = 'Chưa quét thư mục nào.'
$lblCount.Location = New-Object System.Drawing.Point(12, 296)
$lblCount.Size     = New-Object System.Drawing.Size(700, 20)
$grpPick.Controls.Add($lblCount)

$btnExport          = New-Object System.Windows.Forms.Button
$btnExport.Text     = 'Xuất danh sách CSV'
$btnExport.Location = New-Object System.Drawing.Point(822, 292)
$btnExport.Size     = New-Object System.Drawing.Size(150, 28)
$grpPick.Controls.Add($btnExport)

$grpPrinter          = New-Object System.Windows.Forms.GroupBox
$grpPrinter.Text     = '3. Máy in'
$grpPrinter.Location = New-Object System.Drawing.Point(8, 468)
$grpPrinter.Size     = New-Object System.Drawing.Size(984, 110)
$form.Controls.Add($grpPrinter)

$lblPrinter          = New-Object System.Windows.Forms.Label
$lblPrinter.Text     = 'Máy in'
$lblPrinter.Location = New-Object System.Drawing.Point(12, 28)
$lblPrinter.Size     = New-Object System.Drawing.Size(60, 20)
$grpPrinter.Controls.Add($lblPrinter)

$cboPrinter               = New-Object System.Windows.Forms.ComboBox
$cboPrinter.Location      = New-Object System.Drawing.Point(76, 25)
$cboPrinter.Size          = New-Object System.Drawing.Size(390, 24)
$cboPrinter.DropDownStyle = 'DropDownList'
$grpPrinter.Controls.Add($cboPrinter)

$btnRefreshPrinters          = New-Object System.Windows.Forms.Button
$btnRefreshPrinters.Text     = 'Làm mới'
$btnRefreshPrinters.Location = New-Object System.Drawing.Point(474, 23)
$btnRefreshPrinters.Size     = New-Object System.Drawing.Size(90, 28)
$grpPrinter.Controls.Add($btnRefreshPrinters)

$lblCopies          = New-Object System.Windows.Forms.Label
$lblCopies.Text     = 'Số bản'
$lblCopies.Location = New-Object System.Drawing.Point(578, 28)
$lblCopies.Size     = New-Object System.Drawing.Size(50, 20)
$grpPrinter.Controls.Add($lblCopies)

$numCopies          = New-Object System.Windows.Forms.NumericUpDown
$numCopies.Location = New-Object System.Drawing.Point(630, 25)
$numCopies.Size     = New-Object System.Drawing.Size(55, 24)
$numCopies.Minimum  = 1
$numCopies.Maximum  = 50
$numCopies.Value    = 1
$grpPrinter.Controls.Add($numCopies)

$lblDelay          = New-Object System.Windows.Forms.Label
$lblDelay.Text     = 'Chờ mỗi file (giây)'
$lblDelay.Location = New-Object System.Drawing.Point(700, 28)
$lblDelay.Size     = New-Object System.Drawing.Size(120, 20)
$grpPrinter.Controls.Add($lblDelay)

$numDelay          = New-Object System.Windows.Forms.NumericUpDown
$numDelay.Location = New-Object System.Drawing.Point(826, 25)
$numDelay.Size     = New-Object System.Drawing.Size(55, 24)
$numDelay.Minimum  = 1
$numDelay.Maximum  = 120
$numDelay.Value    = 6
$grpPrinter.Controls.Add($numDelay)

$chkClose          = New-Object System.Windows.Forms.CheckBox
$chkClose.Text     = 'Đóng ứng dụng sau khi in xong từng file'
$chkClose.Location = New-Object System.Drawing.Point(14, 64)
$chkClose.Size     = New-Object System.Drawing.Size(280, 22)
$chkClose.Checked  = $true
$grpPrinter.Controls.Add($chkClose)

$chkQueue          = New-Object System.Windows.Forms.CheckBox
$chkQueue.Text     = 'Chờ hàng đợi máy in trống rồi mới in file tiếp'
$chkQueue.Location = New-Object System.Drawing.Point(300, 64)
$chkQueue.Size     = New-Object System.Drawing.Size(300, 22)
$chkQueue.Checked  = $true
$grpPrinter.Controls.Add($chkQueue)

$chkDryRun          = New-Object System.Windows.Forms.CheckBox
$chkDryRun.Text     = 'In thử (KHÔNG gửi máy in)'
$chkDryRun.Location = New-Object System.Drawing.Point(606, 64)
$chkDryRun.Size     = New-Object System.Drawing.Size(180, 22)
$chkDryRun.ForeColor = [System.Drawing.Color]::Firebrick
$chkDryRun.Font      = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grpPrinter.Controls.Add($chkDryRun)

$btnPrint          = New-Object System.Windows.Forms.Button
$btnPrint.Text     = 'IN'
$btnPrint.Location = New-Object System.Drawing.Point(790, 60)
$btnPrint.Size     = New-Object System.Drawing.Size(80, 30)
$btnPrint.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$grpPrinter.Controls.Add($btnPrint)

$btnStop          = New-Object System.Windows.Forms.Button
$btnStop.Text     = 'Dừng'
$btnStop.Location = New-Object System.Drawing.Point(876, 60)
$btnStop.Size     = New-Object System.Drawing.Size(96, 30)
$btnStop.Enabled  = $false
$grpPrinter.Controls.Add($btnStop)

$bar          = New-Object System.Windows.Forms.ProgressBar
$bar.Location = New-Object System.Drawing.Point(8, 590)
$bar.Size     = New-Object System.Drawing.Size(996, 18)
$form.Controls.Add($bar)

$txtLog            = New-Object System.Windows.Forms.TextBox
$txtLog.Location   = New-Object System.Drawing.Point(8, 614)
$txtLog.Size       = New-Object System.Drawing.Size(996, 132)
$txtLog.Multiline  = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly   = $true
$txtLog.BackColor  = [System.Drawing.Color]::White
$txtLog.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

# ============================================================================
#  LOGIC
# ============================================================================
function Update-CountLabel {
    $lblCount.Text = 'Hiển thị {0} file — đang chọn {1} hóa đơn để in.' -f $lvFiles.Items.Count, $lvFiles.CheckedItems.Count
}

function Get-InvoiceNumberFromName {
    param([string]$BaseName, [string]$Pattern)
    $text = $null
    if ($Pattern) {
        try {
            $m = [regex]::Match($BaseName, $Pattern)
            if ($m.Success) {
                if ($m.Groups.Count -gt 1 -and $m.Groups[1].Success) { $text = $m.Groups[1].Value } else { $text = $m.Value }
            }
        } catch { $text = $null }
    }
    if (-not $text) {
        $all = [regex]::Matches($BaseName, '\d+')
        if ($all.Count -gt 0) { $text = $all[$all.Count - 1].Value }
    }
    if (-not $text) { return $null }

    $digits = ($text -replace '\D', '')
    if (-not $digits) { return $null }
    $value = 0.0
    if ([double]::TryParse($digits, [ref]$value)) { return [pscustomobject]@{ Text = $text; Value = $value } }
    return $null
}

function Get-DisplayList {
    $keyword = $txtSearch.Text.Trim()
    $items = $script:AllFiles
    if ($keyword) { $items = @($items | Where-Object { $_.Name -like ('*{0}*' -f $keyword) }) }

    $big = [double]::MaxValue
    switch ($cboSort.SelectedIndex) {
        0 { $items = @($items | Sort-Object @{ Expression = { if ($null -eq $_.Number) { $big } else { $_.Number } } }, Name) }
        1 { $items = @($items | Sort-Object @{ Expression = { if ($null -eq $_.Number) { -$big } else { $_.Number } }; Descending = $true }, Name) }
        2 { $items = @($items | Sort-Object Name, FullName) }
        3 { $items = @($items | Sort-Object Name, FullName -Descending) }
        4 { $items = @($items | Sort-Object Modified) }
        5 { $items = @($items | Sort-Object Modified -Descending) }
        6 { $items = @($items | Sort-Object Order) }
    }
    return $items
}

function Update-FileList {
    $script:Populating = $true
    $lvFiles.BeginUpdate()
    $lvFiles.Items.Clear()

    $i = 0
    foreach ($f in (Get-DisplayList)) {
        $i++
        $row = New-Object System.Windows.Forms.ListViewItem([string]$i)
        [void]$row.SubItems.Add($(if ($f.NumberText) { $f.NumberText } else { '(không có)' }))
        [void]$row.SubItems.Add($f.Name)
        [void]$row.SubItems.Add($f.Modified.ToString('dd/MM/yyyy HH:mm'))
        [void]$row.SubItems.Add('{0:N0} KB' -f [math]::Ceiling($f.Size / 1KB))
        [void]$row.SubItems.Add($f.FolderName)
        $row.Tag = $f
        if ($null -eq $f.Number) { $row.ForeColor = [System.Drawing.Color]::Firebrick }
        $row.Checked = $script:CheckedPaths.Contains($f.FullName)
        [void]$lvFiles.Items.Add($row)
    }

    $lvFiles.EndUpdate()
    $script:Populating = $false
    Update-CountLabel
}

function Invoke-Scan {
    $folder = $txtFolder.Text.Trim()
    if (-not $folder -or -not (Test-Path -LiteralPath $folder -PathType Container)) {
        [void][System.Windows.Forms.MessageBox]::Show('Chưa chọn thư mục hợp lệ.', 'Thiếu thư mục',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $exts    = @($script:FileGroups[$cboType.SelectedItem.ToString()])
    $pattern = $txtRegex.Text.Trim()

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $raw = @(Get-ChildItem -LiteralPath $folder -File -Recurse:$chkSub.Checked -ErrorAction SilentlyContinue)
        if ($exts.Count -gt 0) { $raw = @($raw | Where-Object { $exts -contains $_.Extension.ToLower() }) }

        $order = 0
        $script:AllFiles = @(
            foreach ($f in $raw) {
                $num = Get-InvoiceNumberFromName -BaseName $f.BaseName -Pattern $pattern
                $order++
                [pscustomobject]@{
                    Order      = $order
                    Name       = $f.Name
                    FullName   = $f.FullName
                    FolderName = $f.Directory.Name
                    Modified   = $f.LastWriteTime
                    Size       = $f.Length
                    Number     = $(if ($num) { $num.Value } else { $null })
                    NumberText = $(if ($num) { $num.Text } else { $null })
                }
            }
        )

        $script:CheckedPaths.Clear()
        foreach ($f in $script:AllFiles) { [void]$script:CheckedPaths.Add($f.FullName) }
        Update-FileList

        Write-Log ('Quét xong: {0} file trong "{1}".' -f $script:AllFiles.Count, $folder)
        $noNumber = @($script:AllFiles | Where-Object { $null -eq $_.Number }).Count
        if ($noNumber -gt 0) { Write-Log ('{0} file không tách được số hóa đơn (dòng màu đỏ).' -f $noNumber) 'WARN' }
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Set-AllChecked {
    param([string]$Mode)
    $script:Populating = $true
    $lvFiles.BeginUpdate()
    foreach ($row in $lvFiles.Items) {
        switch ($Mode) {
            'all'    { $row.Checked = $true }
            'none'   { $row.Checked = $false }
            'invert' { $row.Checked = -not $row.Checked }
        }
        if ($row.Checked) { [void]$script:CheckedPaths.Add($row.Tag.FullName) }
        else              { [void]$script:CheckedPaths.Remove($row.Tag.FullName) }
    }
    $lvFiles.EndUpdate()
    $script:Populating = $false
    Update-CountLabel
}

function Invoke-PrintJobs {
    <# Lõi in dùng chung: nhận danh sách file ĐÃ ĐÚNG THỨ TỰ và in lần lượt. #>
    param([array]$Jobs)

    if ($Jobs.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('Không có hóa đơn nào để in.', 'Chưa có gì để in',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    if (-not $cboPrinter.SelectedItem) {
        [void][System.Windows.Forms.MessageBox]::Show('Chưa chọn máy in.', 'Chưa chọn máy in',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $printer = $cboPrinter.SelectedItem.ToString()
    $copies  = [int]$numCopies.Value
    $delay   = [int]$numDelay.Value
    $dry     = $chkDryRun.Checked

    $msg = "In {0} hóa đơn × {1} bản ra máy in:`n{2}`n`nTiếp tục?" -f $Jobs.Count, $copies, $printer
    if ($dry) { $msg = "CHẾ ĐỘ IN THỬ — chỉ ghi nhật ký, KHÔNG gửi gì ra máy in.`nMuốn in thật thì bấm Không, bỏ tick ""In thử"" rồi bấm IN lại.`n`n" + $msg }
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'Xác nhận in',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $script:Cancel = $false
    $script:Busy   = $true
    $btnPrint.Enabled = $false; $btnStop.Enabled = $true; $btnScan.Enabled = $false
    $bar.Minimum = 0; $bar.Maximum = $Jobs.Count; $bar.Value = 0

    $prevDefault = $null
    $ok = 0; $fail = 0
    try {
        if (-not $dry) {
            $prevDefault = Get-DefaultPrinterName
            if ($prevDefault -ne $printer) {
                if (Set-DefaultPrinterName $printer) {
                    Write-Log ('Tạm đặt máy in mặc định: "{0}" (sẽ trả lại "{1}" khi xong).' -f $printer, $prevDefault)
                } else {
                    Write-Log 'Không đổi được máy in mặc định — vẫn thử in bằng lệnh PrintTo.' 'WARN'
                    $prevDefault = $null
                }
            } else { $prevDefault = $null }
        }

        Write-Log ('=== Bắt đầu in {0} hóa đơn, {1} bản mỗi hóa đơn ===' -f $Jobs.Count, $copies)
        if (-not $dry -and ($printer -match 'PDF' -or $printer -match 'XPS' -or $printer -match 'OneNote' -or $printer -match 'Fax')) {
            Write-Log ('"{0}" là máy in ảo — Windows sẽ hỏi chỗ lưu cho từng file chứ không ra giấy. Chọn máy in thật nếu muốn in giấy.' -f $printer) 'WARN'
        }
        $index = 0
        foreach ($job in $Jobs) {
            if ($script:Cancel) { Write-Log 'Đã dừng theo yêu cầu.' 'STOP'; break }
            $index++
            $label = '{0}/{1}  {2}' -f $index, $Jobs.Count, $job.Label

            if (-not (Test-Path -LiteralPath $job.Path)) {
                Write-Log ('KHÔNG THẤY FILE — {0}' -f $label) 'ERR'
                $fail++
            } elseif ($dry) {
                Write-Log ('[IN THỬ] {0}' -f $label)
                $ok++
                Start-UiSleep -Milliseconds 200
            } else {
                try {
                    for ($c = 1; $c -le $copies; $c++) {
                        if ($script:Cancel) { break }
                        $verb = Invoke-PrintOneFile -Path $job.Path -PrinterName $printer -WaitSeconds $delay -CloseApp $chkClose.Checked
                        Write-Log ('Đã gửi ({0}) bản {1}/{2} — {3}' -f $verb, $c, $copies, $label)
                        if ($chkQueue.Checked) { Wait-PrintQueue -PrinterName $printer -TimeoutSeconds ($delay * 10) }
                    }
                    $ok++
                } catch {
                    Write-Log ('LỖI khi in {0} — {1}' -f $label, $_.Exception.Message) 'ERR'
                    $fail++
                }
            }

            $bar.Value = [math]::Min($index, $bar.Maximum)
            [System.Windows.Forms.Application]::DoEvents()
        }
        Write-Log ('=== Xong: {0} hóa đơn đã gửi, {1} lỗi ===' -f $ok, $fail)
        if ($dry) {
            Write-Log 'VỪA CHẠY IN THỬ — chưa có tờ nào được in. Bỏ tick "In thử (KHÔNG gửi máy in)" rồi bấm IN để in thật.' 'WARN'
        }
    } finally {
        if ($prevDefault) {
            if (Set-DefaultPrinterName $prevDefault) { Write-Log ('Đã trả máy in mặc định về "{0}".' -f $prevDefault) }
        }
        $script:Busy = $false
        $btnPrint.Enabled = $true; $btnStop.Enabled = $false; $btnScan.Enabled = $true
    }
}

function Import-PrintOrder {
    <# Nạp file "thu-tu-in.txt" do công cụ sắp xếp tạo ra: mỗi dòng một đường dẫn,
       đúng thứ tự trong Excel. #>
    param([string]$Path)

    $pattern = $txtRegex.Text.Trim()
    $lines = @(Get-Content -LiteralPath $Path -Encoding UTF8 |
               ForEach-Object { $_.Trim() } |
               Where-Object { $_ -and -not $_.StartsWith('#') })

    $files   = @()
    $missing = 0
    $order   = 0
    foreach ($line in $lines) {
        if (-not (Test-Path -LiteralPath $line -PathType Leaf)) { $missing++; continue }
        $f = Get-Item -LiteralPath $line
        $num = Get-InvoiceNumberFromName -BaseName $f.BaseName -Pattern $pattern
        $order++
        $files += [pscustomobject]@{
            Order      = $order
            Name       = $f.Name
            FullName   = $f.FullName
            FolderName = $f.Directory.Name
            Modified   = $f.LastWriteTime
            Size       = $f.Length
            Number     = $(if ($num) { $num.Value } else { $null })
            NumberText = $(if ($num) { $num.Text } else { $null })
        }
    }

    if ($files.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('Không có file nào trong danh sách còn tồn tại.', 'Danh sách rỗng',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $script:AllFiles = $files
    $script:CheckedPaths.Clear()
    foreach ($f in $files) { [void]$script:CheckedPaths.Add($f.FullName) }
    $txtSearch.Text = ''
    $cboSort.SelectedIndex = 6          # giữ đúng thứ tự trong danh sách
    Update-FileList

    Write-Log ('Đã nạp danh sách thứ tự in: {0} hóa đơn từ "{1}".' -f $files.Count, $Path)
    if ($missing -gt 0) { Write-Log ('{0} dòng trong danh sách không còn file — đã bỏ qua.' -f $missing) 'WARN' }
}

# ============================================================================
#  SỰ KIỆN
# ============================================================================
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Chọn thư mục chứa file hóa đơn cần in'
    if ($txtFolder.Text -and (Test-Path -LiteralPath $txtFolder.Text)) { $dlg.SelectedPath = $txtFolder.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFolder.Text = $dlg.SelectedPath
        Invoke-Scan
    }
})

$btnScan.Add_Click({ Invoke-Scan })

$btnLoadOrder.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter   = 'Danh sách thứ tự in (*.txt)|*.txt|Tất cả các file (*.*)|*.*'
    $dlg.Title    = 'Chọn file thu-tu-in.txt do công cụ sắp xếp tạo ra'
    $dlg.FileName = 'thu-tu-in.txt'
    if ($txtFolder.Text -and (Test-Path -LiteralPath $txtFolder.Text)) { $dlg.InitialDirectory = $txtFolder.Text }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Import-PrintOrder -Path $dlg.FileName }
})

$cboSort.Add_SelectedIndexChanged({ if ($script:AllFiles.Count -gt 0) { Update-FileList } })
$txtSearch.Add_TextChanged({ if ($script:AllFiles.Count -gt 0) { Update-FileList } })

$lvFiles.Add_ItemChecked({
    param($sender, $e)
    if ($script:Populating) { return }
    if ($e.Item.Checked) { [void]$script:CheckedPaths.Add($e.Item.Tag.FullName) }
    else                 { [void]$script:CheckedPaths.Remove($e.Item.Tag.FullName) }
    Update-CountLabel
})

$btnAll.Add_Click({ Set-AllChecked 'all' })
$btnNone.Add_Click({ Set-AllChecked 'none' })
$btnInvert.Add_Click({ Set-AllChecked 'invert' })

$btnRange.Add_Click({
    $from = 0.0; $to = 0.0
    $fromOk = [double]::TryParse(($txtFrom.Text -replace '\D', ''), [ref]$from)
    $toOk   = [double]::TryParse(($txtTo.Text   -replace '\D', ''), [ref]$to)
    if (-not $fromOk -and -not $toOk) {
        [void][System.Windows.Forms.MessageBox]::Show('Nhập "Từ số" và/hoặc "đến số" bằng chữ số.', 'Khoảng số hóa đơn',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    if (-not $fromOk) { $from = [double]::MinValue }
    if (-not $toOk)   { $to   = [double]::MaxValue }

    $script:Populating = $true
    $lvFiles.BeginUpdate()
    $hit = 0
    foreach ($row in $lvFiles.Items) {
        $n = $row.Tag.Number
        $inRange = ($null -ne $n) -and ($n -ge $from) -and ($n -le $to)
        $row.Checked = $inRange
        if ($inRange) { [void]$script:CheckedPaths.Add($row.Tag.FullName); $hit++ }
        else          { [void]$script:CheckedPaths.Remove($row.Tag.FullName) }
    }
    $lvFiles.EndUpdate()
    $script:Populating = $false
    Update-CountLabel
    Write-Log ('Chọn theo khoảng số hóa đơn: {0} hóa đơn được chọn.' -f $hit)
})

$btnRefreshPrinters.Add_Click({
    $current = $cboPrinter.SelectedItem
    $cboPrinter.Items.Clear()
    $names = Get-PrinterNames
    if ($names.Count -eq 0) { Write-Log 'Không tìm thấy máy in nào trên máy tính này.' 'WARN'; return }
    foreach ($n in $names) { [void]$cboPrinter.Items.Add($n) }
    $default = Get-DefaultPrinterName
    if ($current -and $cboPrinter.Items.Contains($current))     { $cboPrinter.SelectedItem = $current }
    elseif ($default -and $cboPrinter.Items.Contains($default)) { $cboPrinter.SelectedItem = $default }
    else { $cboPrinter.SelectedIndex = 0 }
    Write-Log ('Tìm thấy {0} máy in. Đang chọn: {1}' -f $names.Count, $cboPrinter.SelectedItem)
})

$btnExport.Add_Click({
    if ($lvFiles.Items.Count -eq 0) { return }
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'File CSV (*.csv)|*.csv'
    $dlg.FileName = 'danh-sach-hoa-don.csv'
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $rows = foreach ($row in $lvFiles.Items) {
        [pscustomobject]@{
            STT      = $row.Text
            SoHoaDon = $row.SubItems[1].Text
            DaChon   = $(if ($row.Checked) { 'x' } else { '' })
            TenFile  = $row.Tag.Name
            NgaySua  = $row.Tag.Modified.ToString('dd/MM/yyyy HH:mm')
            DuongDan = $row.Tag.FullName
        }
    }
    $rows | Export-Csv -LiteralPath $dlg.FileName -NoTypeInformation -Encoding UTF8
    Write-Log ('Đã xuất danh sách ra "{0}".' -f $dlg.FileName)
})

$btnPrint.Add_Click({
    $jobs = @()
    foreach ($row in $lvFiles.CheckedItems) {
        $jobs += @{
            Path  = $row.Tag.FullName
            Label = 'số HĐ {0}  {1}' -f $(if ($row.Tag.NumberText) { $row.Tag.NumberText } else { '-' }), $row.Tag.Name
        }
    }
    Invoke-PrintJobs -Jobs $jobs
})

$btnStop.Add_Click({ $script:Cancel = $true; Write-Log 'Đang dừng sau khi in xong file hiện tại...' 'STOP' })

$form.Add_FormClosing({
    param($sender, $e)
    if ($script:Busy) {
        $r = [System.Windows.Forms.MessageBox]::Show('Đang in. Thoát luôn?', 'Đang in',
            [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { $e.Cancel = $true; return }
        $script:Cancel = $true
    }
})

# ============================================================================
#  KHỞI ĐỘNG
# ============================================================================
$form.Add_Shown({
    $form.Activate()
    $start = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $txtFolder.Text = $start
    Write-Log 'Sẵn sàng. Chọn thư mục chứa hóa đơn rồi bấm "Quét folder".'
    Write-Log 'Đã sắp xếp bằng công cụ kia thì bấm "Nạp danh sách thứ tự in..." và chọn file thu-tu-in.txt để in đúng thứ tự trong Excel.'
    $btnRefreshPrinters.PerformClick()
})

Write-Trace 'Dựng xong, mở cửa sổ'
[void]$form.ShowDialog()
Write-Trace 'Người dùng đóng cửa sổ — kết thúc bình thường'
$form.Dispose()
