<#
    SẮP XẾP & IN HÓA ĐƠN — Nguyễn Thanh
    ------------------------------------------------------------------
    Công cụ PowerShell có giao diện, gồm 2 tab:

      Tab 1 — Sắp xếp theo danh sách Excel
        • Đọc file Excel danh sách hóa đơn (mỗi sheet là một nhóm).
        • Trong mỗi sheet tự dò các bảng "Ký hiệu / Số hóa đơn / Ngày hóa đơn",
          tên mẫu nằm phía trên bảng (VD: TELMA 80 H PLUS (TABLET B/100)).
        • Dò tìm file hóa đơn trong các thư mục nguồn (kể cả thư mục con,
          hóa đơn lộn xộn nhiều năm 2023–2025) theo KÝ HIỆU + SỐ HÓA ĐƠN,
          lọc thêm theo NĂM nếu cần.
        • Chép (hoặc di chuyển) sang thư mục đích:
              <Thư mục đích>\<Tên sheet>\<Tên mẫu>\001_K25TAA_618585.pdf
        • Báo cáo rõ hóa đơn nào thiếu file, hóa đơn nào khớp nhiều file.

      Tab 2 — In hóa đơn
        • Quét một thư mục, sắp xếp theo số hóa đơn / tên file / ngày.
        • Tick chọn từng hóa đơn hoặc chọn theo khoảng số.
        • In lần lượt ra máy in đã chọn, đúng số bản, đúng thứ tự.

    Cách chạy:
        Nhấp đúp vào "sap-xep-in-hoa-don.bat"
        hoặc: powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\sap-xep-in-hoa-don.ps1

    Yêu cầu: Windows 7 trở lên, Windows PowerShell 5.1 (có sẵn trong Windows).
             File Excel nên là .xlsx / .xlsm (đọc trực tiếp, không cần cài Excel).
             File .xls cũ cần có Excel trên máy.
             Muốn in PDF thì máy phải có Adobe Reader (hoặc phần mềm đọc PDF
             hỗ trợ lệnh in của Windows).
#>

#Requires -Version 3.0

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()

# ============================================================================
#  BIẾN DÙNG CHUNG
# ============================================================================
$script:Sheets       = @()          # các sheet đọc được từ Excel
$script:Invoices     = @()          # danh sách hóa đơn đã dò (sau khi đối chiếu)
$script:AllFiles     = @()          # file quét được ở tab In
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
#  HÀM PHỤ TRỢ CHUNG
# ============================================================================

function Write-Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = '[{0}] {1}  {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level.PadRight(4), $Message
    $txtLog.AppendText($line + [Environment]::NewLine)
    $txtLog.SelectionStart = $txtLog.TextLength
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Remove-Diacritics {
    param([string]$Text)
    if (-not $Text) { return '' }
    $norm = $Text.Normalize([System.Text.NormalizationForm]::FormD)
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $norm.ToCharArray()) {
        if ([Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString() -replace 'đ', 'd' -replace 'Đ', 'D'
}

function Get-HeaderKey {
    <# Chuẩn hóa tiêu đề cột để so sánh: bỏ dấu, bỏ ký tự lạ, viết thường. #>
    param([string]$Text)
    $t = Remove-Diacritics $Text
    $t = $t -replace '[^A-Za-z0-9]', ' '
    $t = ($t -replace '\s+', ' ').Trim().ToLower()
    return $t
}

function Get-SafeName {
    <# Đổi tên sheet / tên mẫu thành tên thư mục hợp lệ trên Windows. #>
    param([string]$Text, [int]$MaxLength = 80)
    $t = $Text.Trim()
    foreach ($c in [System.IO.Path]::GetInvalidFileNameChars()) { $t = $t.Replace($c, '-') }
    $t = ($t -replace '\s+', ' ').Trim(' ', '.', '-')
    if (-not $t) { $t = 'Khong-ten' }
    if ($t.Length -gt $MaxLength) { $t = $t.Substring(0, $MaxLength).Trim() }
    return $t
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

function ConvertTo-InvoiceDate {
    <# Ô "Ngày hóa đơn" có thể là số sê-ri của Excel hoặc chuỗi dd/MM/yyyy. #>
    param([string]$Value)
    if (-not $Value) { return $null }
    $v = $Value.Trim()

    $serial = 0.0
    if ([double]::TryParse($v, [ref]$serial) -and $serial -ge 20000 -and $serial -le 80000) {
        try { return [datetime]::FromOADate($serial) } catch { }
    }
    $formats = @('dd/MM/yyyy', 'd/M/yyyy', 'dd-MM-yyyy', 'yyyy-MM-dd', 'dd/MM/yy', 'MM/dd/yyyy')
    $dt = [datetime]::MinValue
    foreach ($f in $formats) {
        if ([datetime]::TryParseExact($v, $f, [Globalization.CultureInfo]::InvariantCulture,
                                      [Globalization.DateTimeStyles]::None, [ref]$dt)) { return $dt }
    }
    if ([datetime]::TryParse($v, [ref]$dt)) { return $dt }
    return $null
}

function Get-YearFromSymbol {
    <# Ký hiệu hóa đơn kiểu K25TAA: hai chữ số sau chữ cái đầu là năm (25 -> 2025). #>
    param([string]$Symbol)
    if (-not $Symbol) { return $null }
    $m = [regex]::Match($Symbol.Trim(), '^[A-Za-z]?(\d{2})')
    if ($m.Success) { return 2000 + [int]$m.Groups[1].Value }
    return $null
}

# ============================================================================
#  ĐỌC FILE EXCEL
# ============================================================================

function Convert-RefToColumn {
    <# "H4" -> 8 #>
    param([string]$CellRef)
    $letters = ([regex]::Match($CellRef, '^[A-Z]+')).Value
    $col = 0
    foreach ($ch in $letters.ToCharArray()) { $col = $col * 26 + ([int][char]$ch - 64) }
    return $col
}

function Read-ZipXml {
    param($Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { return $null }
    $stream = $entry.Open()
    $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
    try   { return [xml]$reader.ReadToEnd() }
    finally { $reader.Dispose(); $stream.Dispose() }
}

function Read-XlsxWorkbook {
    <# Đọc .xlsx/.xlsm bằng cách mở gói ZIP — không cần cài Excel.
       Trả về mảng sheet, mỗi sheet có bảng ô: khóa "hàng_cột" -> nội dung. #>
    param([string]$Path)

    $zip = [System.IO.Compression.ZipFile]::OpenRead($Path)
    try {
        # --- chuỗi dùng chung ---
        $shared = New-Object System.Collections.Generic.List[string]
        $ssXml = Read-ZipXml $zip 'xl/sharedStrings.xml'
        if ($ssXml) {
            foreach ($si in $ssXml.DocumentElement.ChildNodes) {
                $parts = @($si.SelectNodes(".//*[local-name()='t']") | ForEach-Object { $_.InnerText })
                $shared.Add(($parts -join ''))
            }
        }

        # --- danh sách sheet + đường dẫn XML tương ứng ---
        $relXml = Read-ZipXml $zip 'xl/_rels/workbook.xml.rels'
        $relMap = @{}
        if ($relXml) {
            foreach ($r in $relXml.DocumentElement.ChildNodes) { $relMap[$r.Id] = $r.Target }
        }

        $wbXml = Read-ZipXml $zip 'xl/workbook.xml'
        if (-not $wbXml) { throw 'Không đọc được cấu trúc file Excel (thiếu xl/workbook.xml).' }

        $result = @()
        $nsRel = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        foreach ($s in $wbXml.DocumentElement.SelectNodes("*[local-name()='sheets']/*[local-name()='sheet']")) {
            $name   = $s.GetAttribute('name')
            $rid    = $s.GetAttribute('id', $nsRel)
            $target = $relMap[$rid]
            if (-not $target) { continue }
            $target = $target -replace '^/', ''
            if ($target -notlike 'xl/*') { $target = 'xl/' + $target }

            $shXml = Read-ZipXml $zip $target
            if (-not $shXml) { continue }

            $cells  = @{}
            $maxRow = 0
            $maxCol = 0
            foreach ($row in $shXml.DocumentElement.SelectNodes("*[local-name()='sheetData']/*[local-name()='row']")) {
                $rowIndex = [int]$row.GetAttribute('r')
                foreach ($c in $row.ChildNodes) {
                    $ref = $c.GetAttribute('r')
                    if (-not $ref) { continue }
                    $colIndex = Convert-RefToColumn $ref
                    $type = $c.GetAttribute('t')

                    $text = $null
                    if ($type -eq 's') {
                        $v = $c.SelectSingleNode("*[local-name()='v']")
                        if ($v) {
                            $idx = [int]$v.InnerText
                            if ($idx -ge 0 -and $idx -lt $shared.Count) { $text = $shared[$idx] }
                        }
                    } elseif ($type -eq 'inlineStr') {
                        $node = $c.SelectSingleNode(".//*[local-name()='t']")
                        if ($node) { $text = $node.InnerText }
                    } else {
                        $v = $c.SelectSingleNode("*[local-name()='v']")
                        if ($v) { $text = $v.InnerText }
                    }

                    if ($null -ne $text -and $text.Trim() -ne '') {
                        $cells['{0}_{1}' -f $rowIndex, $colIndex] = $text.Trim()
                        if ($rowIndex -gt $maxRow) { $maxRow = $rowIndex }
                        if ($colIndex -gt $maxCol) { $maxCol = $colIndex }
                    }
                }
            }

            $result += [pscustomobject]@{ Name = $name; Cells = $cells; MaxRow = $maxRow; MaxCol = $maxCol }
        }
        return $result
    } finally {
        $zip.Dispose()
    }
}

function Read-WorkbookWithExcel {
    <# Dự phòng cho file .xls cũ: dùng Excel trên máy để đọc. #>
    param([string]$Path)

    $excel = $null
    $book  = $null
    try {
        $excel = New-Object -ComObject Excel.Application
    } catch {
        throw 'File .xls cũ cần có Microsoft Excel trên máy. Hãy mở file và lưu lại thành .xlsx.'
    }

    try {
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $book = $excel.Workbooks.Open($Path, 0, $true)

        $result = @()
        foreach ($ws in $book.Worksheets) {
            $used = $ws.UsedRange
            $rows = $used.Rows.Count
            $cols = $used.Columns.Count
            $r0   = $used.Row
            $c0   = $used.Column

            $cells = @{}
            $data = $used.Value2
            if ($rows -eq 1 -and $cols -eq 1) {
                if ($null -ne $data -and "$data".Trim() -ne '') { $cells['{0}_{1}' -f $r0, $c0] = "$data".Trim() }
            } elseif ($null -ne $data) {
                for ($i = 1; $i -le $rows; $i++) {
                    for ($j = 1; $j -le $cols; $j++) {
                        $v = $data.GetValue($i, $j)
                        if ($null -ne $v -and "$v".Trim() -ne '') {
                            $cells['{0}_{1}' -f ($r0 + $i - 1), ($c0 + $j - 1)] = "$v".Trim()
                        }
                    }
                }
            }
            $result += [pscustomobject]@{
                Name = $ws.Name; Cells = $cells; MaxRow = ($r0 + $rows - 1); MaxCol = ($c0 + $cols - 1)
            }
        }
        return $result
    } finally {
        if ($book)  { $book.Close($false)  | Out-Null }
        if ($excel) { $excel.Quit()        | Out-Null }
        if ($book)  { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
        if ($excel) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
    }
}

function Read-Workbook {
    param([string]$Path)
    $ext = [System.IO.Path]::GetExtension($Path).ToLower()
    if ($ext -in @('.xlsx', '.xlsm')) {
        try { return Read-XlsxWorkbook -Path $Path }
        catch { Write-Log ('Đọc trực tiếp không được ({0}), thử mở bằng Excel...' -f $_.Exception.Message) 'WARN' }
    }
    return Read-WorkbookWithExcel -Path $Path
}

# ============================================================================
#  DÒ CÁC BẢNG HÓA ĐƠN TRONG SHEET
# ============================================================================

$script:HeaderSymbol = @('ky hieu', 'ki hieu', 'ky hieu hoa don', 'serial', 'mau so ky hieu')
$script:HeaderNumber = @('so hoa don', 'so hd', 'so hoa don gtgt', 'invoice no', 'so')
$script:HeaderDate   = @('ngay hoa don', 'ngay hd', 'ngay', 'ngay lap hoa don')

function Get-Cell {
    param($Sheet, [int]$Row, [int]$Col)
    $v = $Sheet.Cells['{0}_{1}' -f $Row, $Col]
    if ($null -eq $v) { return '' }
    return [string]$v
}

function Find-InvoiceBlocks {
    <# Tìm mọi bảng hóa đơn trong một sheet (nhiều bảng nằm cạnh nhau cũng được).
       Mỗi bảng trả về: tên sản phẩm/mẫu + danh sách hóa đơn theo đúng thứ tự trong Excel. #>
    param($Sheet)

    # chỉ duyệt các ô có nội dung để nhanh với sheet lớn
    $headers = @()
    foreach ($key in $Sheet.Cells.Keys) {
        if ($script:HeaderNumber -contains (Get-HeaderKey $Sheet.Cells[$key])) {
            $parts = $key.Split('_')
            $headers += [pscustomobject]@{ Row = [int]$parts[0]; Col = [int]$parts[1] }
        }
    }
    $headers = @($headers | Sort-Object Row, Col)

    $blocks = @()
    foreach ($h in $headers) {
        $r = $h.Row
        $c = $h.Col

        # cột "Ký hiệu" nằm bên trái tiêu đề "Số hóa đơn"
        $symbolCol = 0
        for ($k = 1; $k -le 8; $k++) {
            if ($c - $k -lt 1) { break }
            if ($script:HeaderSymbol -contains (Get-HeaderKey (Get-Cell $Sheet $r ($c - $k)))) { $symbolCol = $c - $k; break }
        }
        if ($symbolCol -eq 0) { continue }

        # cột "Ngày hóa đơn" nằm bên phải
        $dateCol = 0
        for ($k = 1; $k -le 5; $k++) {
            if ($script:HeaderDate -contains (Get-HeaderKey (Get-Cell $Sheet $r ($c + $k)))) { $dateCol = $c + $k; break }
        }

        # tên sản phẩm/mẫu: ô có chữ gần nhất phía trên bảng
        $skipWords = $script:HeaderSymbol + $script:HeaderNumber + $script:HeaderDate + @('stt')
        $title = ''
        for ($up = 1; $up -le 4 -and -not $title; $up++) {
            if ($r - $up -lt 1) { break }
            for ($cc = [math]::Max(1, $symbolCol - 2); $cc -le $c + 3; $cc++) {
                $t = Get-Cell $Sheet ($r - $up) $cc
                if ($t -and (Get-HeaderKey $t) -notin $skipWords) { $title = $t; break }
            }
        }
        if (-not $title) { $title = 'Không rõ sản phẩm' }

        # đọc dữ liệu xuống dưới, dừng khi gặp dòng trống
        $rows = @()
        $rr = $r + 1
        while ($rr -le $Sheet.MaxRow) {
            $symbol = Get-Cell $Sheet $rr $symbolCol
            $number = Get-Cell $Sheet $rr $c
            if (-not $symbol -and -not $number) { break }
            if ($number) {
                $dateText = if ($dateCol -gt 0) { Get-Cell $Sheet $rr $dateCol } else { '' }
                $date     = ConvertTo-InvoiceDate $dateText
                $year     = if ($date) { $date.Year } else { Get-YearFromSymbol $symbol }
                $rows += [pscustomobject]@{
                    Symbol    = $symbol.Trim()
                    Number    = ($number -replace '[^\d]', '')
                    RawNumber = $number.Trim()
                    Date      = $date
                    DateText  = if ($date) { $date.ToString('dd/MM/yyyy') } else { $dateText }
                    Year      = $year
                }
            }
            $rr++
        }

        if ($rows.Count -gt 0) {
            $blocks += [pscustomobject]@{ Title = $title.Trim(); Rows = $rows }
        }
    }
    return $blocks
}

# ============================================================================
#  DÒ TÌM FILE HÓA ĐƠN TRONG THƯ MỤC NGUỒN
# ============================================================================

function Build-FileIndex {
    <# Quét các thư mục nguồn, chuẩn hóa tên file để dò tìm nhanh. #>
    param([string[]]$Folders, [bool]$Recurse, [string[]]$Extensions)

    $index = @()
    foreach ($folder in $Folders) {
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
            Write-Log ('Bỏ qua thư mục không tồn tại: {0}' -f $folder) 'WARN'
            continue
        }
        $files = @(Get-ChildItem -LiteralPath $folder -File -Recurse:$Recurse -ErrorAction SilentlyContinue)
        if ($Extensions -and $Extensions.Count -gt 0) {
            $files = @($files | Where-Object { $Extensions -contains $_.Extension.ToLower() })
        }
        foreach ($f in $files) {
            $norm = (Remove-Diacritics $f.BaseName).ToUpper() -replace '[^A-Z0-9]', ''
            $digits = @([regex]::Matches($f.BaseName, '\d+') | ForEach-Object { $_.Value.TrimStart('0') })
            $index += [pscustomobject]@{
                File   = $f
                Norm   = $norm
                Digits = $digits
                Year   = $f.LastWriteTime.Year
            }
        }
        Write-Log ('Đã quét {0} file trong "{1}".' -f $files.Count, $folder)
    }
    return $index
}

function Find-InvoiceFile {
    <# Tìm file cho một hóa đơn: ưu tiên khớp cả KÝ HIỆU lẫn SỐ HÓA ĐƠN. #>
    param($Index, [string]$Symbol, [string]$Number, $Year)

    $symNorm = (Remove-Diacritics $Symbol).ToUpper() -replace '[^A-Z0-9]', ''
    $numKey  = $Number.TrimStart('0')
    if (-not $numKey) { $numKey = '0' }

    $strong = @($Index | Where-Object { $_.Digits -contains $numKey -and $symNorm -and $_.Norm.Contains($symNorm) })
    if ($strong.Count -eq 1) { return @{ Files = $strong; Status = 'Khớp ký hiệu + số' } }
    if ($strong.Count -gt 1) { return @{ Files = $strong; Status = 'Khớp nhiều file' } }

    $weak = @($Index | Where-Object { $_.Digits -contains $numKey })
    if ($Year -and $weak.Count -gt 1) {
        $byYear = @($weak | Where-Object { $_.Norm -match [string]$Year -or $_.File.FullName -match [string]$Year -or $_.Year -eq $Year })
        if ($byYear.Count -gt 0) { $weak = $byYear }
    }
    if ($weak.Count -eq 1) { return @{ Files = $weak; Status = 'Khớp số (thiếu ký hiệu)' } }
    if ($weak.Count -gt 1) { return @{ Files = $weak; Status = 'Khớp nhiều file' } }

    return @{ Files = @(); Status = 'Không tìm thấy' }
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

# ============================================================================
#  GIAO DIỆN
# ============================================================================
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Sắp xếp & in hóa đơn theo danh sách Excel — Nguyễn Thanh'
$form.Size          = New-Object System.Drawing.Size(1044, 880)
$form.MinimumSize   = New-Object System.Drawing.Size(1044, 780)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(10, 8)
$tabs.Size     = New-Object System.Drawing.Size(1010, 660)
$form.Controls.Add($tabs)

$tabSort           = New-Object System.Windows.Forms.TabPage
$tabSort.Text      = '  1. Sắp xếp hóa đơn theo Excel  '
$tabSort.BackColor = [System.Drawing.SystemColors]::Control
$tabs.TabPages.Add($tabSort)

$tabPrint           = New-Object System.Windows.Forms.TabPage
$tabPrint.Text      = '  2. In hóa đơn  '
$tabPrint.BackColor = [System.Drawing.SystemColors]::Control
$tabs.TabPages.Add($tabPrint)

$bar          = New-Object System.Windows.Forms.ProgressBar
$bar.Location = New-Object System.Drawing.Point(10, 676)
$bar.Size     = New-Object System.Drawing.Size(1010, 18)
$form.Controls.Add($bar)

$txtLog            = New-Object System.Windows.Forms.TextBox
$txtLog.Location   = New-Object System.Drawing.Point(10, 700)
$txtLog.Size       = New-Object System.Drawing.Size(1010, 130)
$txtLog.Multiline  = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly   = $true
$txtLog.BackColor  = [System.Drawing.Color]::White
$txtLog.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

# ---------------------------------------------------------------------------
#  TAB 1 — SẮP XẾP
# ---------------------------------------------------------------------------
$grpExcel          = New-Object System.Windows.Forms.GroupBox
$grpExcel.Text     = '1. File Excel danh sách hóa đơn — mỗi sheet sẽ thành một thư mục'
$grpExcel.Location = New-Object System.Drawing.Point(8, 6)
$grpExcel.Size     = New-Object System.Drawing.Size(984, 150)
$tabSort.Controls.Add($grpExcel)

$lblExcel          = New-Object System.Windows.Forms.Label
$lblExcel.Text     = 'File Excel'
$lblExcel.Location = New-Object System.Drawing.Point(12, 28)
$lblExcel.Size     = New-Object System.Drawing.Size(70, 20)
$grpExcel.Controls.Add($lblExcel)

$txtExcel          = New-Object System.Windows.Forms.TextBox
$txtExcel.Location = New-Object System.Drawing.Point(86, 25)
$txtExcel.Size     = New-Object System.Drawing.Size(660, 24)
$grpExcel.Controls.Add($txtExcel)

$btnBrowseExcel          = New-Object System.Windows.Forms.Button
$btnBrowseExcel.Text     = 'Chọn file...'
$btnBrowseExcel.Location = New-Object System.Drawing.Point(752, 23)
$btnBrowseExcel.Size     = New-Object System.Drawing.Size(110, 28)
$grpExcel.Controls.Add($btnBrowseExcel)

$btnLoadExcel          = New-Object System.Windows.Forms.Button
$btnLoadExcel.Text     = 'Đọc danh sách'
$btnLoadExcel.Location = New-Object System.Drawing.Point(868, 23)
$btnLoadExcel.Size     = New-Object System.Drawing.Size(104, 28)
$grpExcel.Controls.Add($btnLoadExcel)

$lblSheets          = New-Object System.Windows.Forms.Label
$lblSheets.Text     = 'Sheet cần xử lý (tick chọn):'
$lblSheets.Location = New-Object System.Drawing.Point(12, 60)
$lblSheets.Size     = New-Object System.Drawing.Size(250, 20)
$grpExcel.Controls.Add($lblSheets)

$clbSheets              = New-Object System.Windows.Forms.CheckedListBox
$clbSheets.Location     = New-Object System.Drawing.Point(12, 82)
$clbSheets.Size         = New-Object System.Drawing.Size(960, 58)
$clbSheets.CheckOnClick = $true
$clbSheets.MultiColumn  = $true
$clbSheets.ColumnWidth  = 310
$grpExcel.Controls.Add($clbSheets)

$grpSource          = New-Object System.Windows.Forms.GroupBox
$grpSource.Text     = '2. Thư mục nguồn chứa file hóa đơn (lộn xộn, nhiều năm — có thể thêm nhiều thư mục)'
$grpSource.Location = New-Object System.Drawing.Point(8, 162)
$grpSource.Size     = New-Object System.Drawing.Size(984, 150)
$tabSort.Controls.Add($grpSource)

$lstSources          = New-Object System.Windows.Forms.ListBox
$lstSources.Location = New-Object System.Drawing.Point(12, 26)
$lstSources.Size     = New-Object System.Drawing.Size(730, 80)
$grpSource.Controls.Add($lstSources)

$btnAddSrc          = New-Object System.Windows.Forms.Button
$btnAddSrc.Text     = 'Thêm thư mục nguồn...'
$btnAddSrc.Location = New-Object System.Drawing.Point(752, 26)
$btnAddSrc.Size     = New-Object System.Drawing.Size(220, 30)
$grpSource.Controls.Add($btnAddSrc)

$btnRemoveSrc          = New-Object System.Windows.Forms.Button
$btnRemoveSrc.Text     = 'Bỏ thư mục đang chọn'
$btnRemoveSrc.Location = New-Object System.Drawing.Point(752, 62)
$btnRemoveSrc.Size     = New-Object System.Drawing.Size(220, 30)
$grpSource.Controls.Add($btnRemoveSrc)

$chkSubSrc          = New-Object System.Windows.Forms.CheckBox
$chkSubSrc.Text     = 'Gồm thư mục con'
$chkSubSrc.Location = New-Object System.Drawing.Point(12, 114)
$chkSubSrc.Size     = New-Object System.Drawing.Size(150, 22)
$chkSubSrc.Checked  = $true
$grpSource.Controls.Add($chkSubSrc)

$lblSrcType          = New-Object System.Windows.Forms.Label
$lblSrcType.Text     = 'Loại file'
$lblSrcType.Location = New-Object System.Drawing.Point(172, 116)
$lblSrcType.Size     = New-Object System.Drawing.Size(55, 20)
$grpSource.Controls.Add($lblSrcType)

$cboSrcType               = New-Object System.Windows.Forms.ComboBox
$cboSrcType.Location      = New-Object System.Drawing.Point(230, 113)
$cboSrcType.Size          = New-Object System.Drawing.Size(160, 24)
$cboSrcType.DropDownStyle = 'DropDownList'
[void]$cboSrcType.Items.AddRange(@($script:FileGroups.Keys))
$cboSrcType.SelectedIndex = 0
$grpSource.Controls.Add($cboSrcType)

$lblYears          = New-Object System.Windows.Forms.Label
$lblYears.Text     = 'Chỉ lấy năm'
$lblYears.Location = New-Object System.Drawing.Point(404, 116)
$lblYears.Size     = New-Object System.Drawing.Size(75, 20)
$grpSource.Controls.Add($lblYears)

$txtYears          = New-Object System.Windows.Forms.TextBox
$txtYears.Location = New-Object System.Drawing.Point(482, 113)
$txtYears.Size     = New-Object System.Drawing.Size(140, 24)
$grpSource.Controls.Add($txtYears)

$lblYearsHint           = New-Object System.Windows.Forms.Label
$lblYearsHint.Text      = 'VD: 2023,2024,2025 — để trống là lấy tất cả các năm trong danh sách'
$lblYearsHint.Location  = New-Object System.Drawing.Point(632, 116)
$lblYearsHint.Size      = New-Object System.Drawing.Size(340, 20)
$lblYearsHint.ForeColor = [System.Drawing.Color]::DimGray
$grpSource.Controls.Add($lblYearsHint)

$grpDest          = New-Object System.Windows.Forms.GroupBox
$grpDest.Text     = '3. Thư mục đích — sẽ tạo: <thư mục đích>\<tên sheet>\<tên sản phẩm>\file'
$grpDest.Location = New-Object System.Drawing.Point(8, 318)
$grpDest.Size     = New-Object System.Drawing.Size(984, 90)
$tabSort.Controls.Add($grpDest)

$lblDest          = New-Object System.Windows.Forms.Label
$lblDest.Text     = 'Thư mục đích'
$lblDest.Location = New-Object System.Drawing.Point(12, 28)
$lblDest.Size     = New-Object System.Drawing.Size(90, 20)
$grpDest.Controls.Add($lblDest)

$txtDest          = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object System.Drawing.Point(106, 25)
$txtDest.Size     = New-Object System.Drawing.Size(640, 24)
$grpDest.Controls.Add($txtDest)

$btnBrowseDest          = New-Object System.Windows.Forms.Button
$btnBrowseDest.Text     = 'Chọn thư mục đích...'
$btnBrowseDest.Location = New-Object System.Drawing.Point(752, 23)
$btnBrowseDest.Size     = New-Object System.Drawing.Size(220, 28)
$grpDest.Controls.Add($btnBrowseDest)

$rdoCopy          = New-Object System.Windows.Forms.RadioButton
$rdoCopy.Text     = 'Chép file'
$rdoCopy.Location = New-Object System.Drawing.Point(12, 58)
$rdoCopy.Size     = New-Object System.Drawing.Size(90, 22)
$rdoCopy.Checked  = $true
$grpDest.Controls.Add($rdoCopy)

$rdoMove          = New-Object System.Windows.Forms.RadioButton
$rdoMove.Text     = 'Di chuyển file'
$rdoMove.Location = New-Object System.Drawing.Point(106, 58)
$rdoMove.Size     = New-Object System.Drawing.Size(110, 22)
$grpDest.Controls.Add($rdoMove)

$chkPerProduct          = New-Object System.Windows.Forms.CheckBox
$chkPerProduct.Text     = 'Tạo thư mục con theo từng sản phẩm/mẫu'
$chkPerProduct.Location = New-Object System.Drawing.Point(230, 58)
$chkPerProduct.Size     = New-Object System.Drawing.Size(280, 22)
$chkPerProduct.Checked  = $true
$grpDest.Controls.Add($chkPerProduct)

$chkPrefix          = New-Object System.Windows.Forms.CheckBox
$chkPrefix.Text     = 'Đánh số thứ tự vào đầu tên file (001_, 002_...)'
$chkPrefix.Location = New-Object System.Drawing.Point(520, 58)
$chkPrefix.Size     = New-Object System.Drawing.Size(310, 22)
$chkPrefix.Checked  = $true
$grpDest.Controls.Add($chkPrefix)

$btnMatch          = New-Object System.Windows.Forms.Button
$btnMatch.Text     = 'Đối chiếu danh sách'
$btnMatch.Location = New-Object System.Drawing.Point(8, 414)
$btnMatch.Size     = New-Object System.Drawing.Size(190, 32)
$btnMatch.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$tabSort.Controls.Add($btnMatch)

$btnOrganize          = New-Object System.Windows.Forms.Button
$btnOrganize.Text     = 'Tạo folder && chép file'
$btnOrganize.Location = New-Object System.Drawing.Point(206, 414)
$btnOrganize.Size     = New-Object System.Drawing.Size(220, 32)
$btnOrganize.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnOrganize.Enabled  = $false
$tabSort.Controls.Add($btnOrganize)

$btnPrintList          = New-Object System.Windows.Forms.Button
$btnPrintList.Text     = 'In ngay theo thứ tự Excel'
$btnPrintList.Location = New-Object System.Drawing.Point(434, 414)
$btnPrintList.Size     = New-Object System.Drawing.Size(230, 32)
$btnPrintList.Enabled  = $false
$tabSort.Controls.Add($btnPrintList)

$btnReport          = New-Object System.Windows.Forms.Button
$btnReport.Text     = 'Xuất báo cáo CSV'
$btnReport.Location = New-Object System.Drawing.Point(672, 414)
$btnReport.Size     = New-Object System.Drawing.Size(160, 32)
$btnReport.Enabled  = $false
$tabSort.Controls.Add($btnReport)

$btnGotoPrint          = New-Object System.Windows.Forms.Button
$btnGotoPrint.Text     = 'Sang tab In →'
$btnGotoPrint.Location = New-Object System.Drawing.Point(840, 414)
$btnGotoPrint.Size     = New-Object System.Drawing.Size(152, 32)
$tabSort.Controls.Add($btnGotoPrint)

$lblMatchInfo          = New-Object System.Windows.Forms.Label
$lblMatchInfo.Text     = 'Chưa đối chiếu.'
$lblMatchInfo.Location = New-Object System.Drawing.Point(8, 606)
$lblMatchInfo.Size     = New-Object System.Drawing.Size(984, 20)
$tabSort.Controls.Add($lblMatchInfo)

$lvMatch               = New-Object System.Windows.Forms.ListView
$lvMatch.Location      = New-Object System.Drawing.Point(8, 452)
$lvMatch.Size          = New-Object System.Drawing.Size(984, 150)
$lvMatch.View          = 'Details'
$lvMatch.FullRowSelect = $true
$lvMatch.GridLines     = $true
[void]$lvMatch.Columns.Add('Sheet (thư mục)', 150)
[void]$lvMatch.Columns.Add('Sản phẩm', 200)
[void]$lvMatch.Columns.Add('Ký hiệu', 80)
[void]$lvMatch.Columns.Add('Số hóa đơn', 90)
[void]$lvMatch.Columns.Add('Ngày', 85)
[void]$lvMatch.Columns.Add('Trạng thái', 150)
[void]$lvMatch.Columns.Add('File tìm được', 200)
$tabSort.Controls.Add($lvMatch)

# ---------------------------------------------------------------------------
#  TAB 2 — IN
# ---------------------------------------------------------------------------
$grpFolder          = New-Object System.Windows.Forms.GroupBox
$grpFolder.Text     = '1. Thư mục chứa hóa đơn cần in'
$grpFolder.Location = New-Object System.Drawing.Point(8, 6)
$grpFolder.Size     = New-Object System.Drawing.Size(984, 120)
$tabPrint.Controls.Add($grpFolder)

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
    'Ngày sửa: mới trước'
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
$lblRegexHint.Size      = New-Object System.Drawing.Size(660, 20)
$lblRegexHint.ForeColor = [System.Drawing.Color]::DimGray
$grpFolder.Controls.Add($lblRegexHint)

$grpPick          = New-Object System.Windows.Forms.GroupBox
$grpPick.Text     = '2. Chọn hóa đơn cần in'
$grpPick.Location = New-Object System.Drawing.Point(8, 132)
$grpPick.Size     = New-Object System.Drawing.Size(984, 330)
$tabPrint.Controls.Add($grpPick)

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
$tabPrint.Controls.Add($grpPrinter)

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
$chkDryRun.Text     = 'In thử (không gửi máy in)'
$chkDryRun.Location = New-Object System.Drawing.Point(606, 64)
$chkDryRun.Size     = New-Object System.Drawing.Size(180, 22)
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

# ============================================================================
#  LOGIC TAB 1 — SẮP XẾP THEO EXCEL
# ============================================================================

function Get-YearFilter {
    $text = $txtYears.Text.Trim()
    if (-not $text) { return @() }
    return @([regex]::Matches($text, '\d{4}') | ForEach-Object { [int]$_.Value })
}

function Invoke-LoadExcel {
    $path = $txtExcel.Text.Trim()
    if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) {
        [void][System.Windows.Forms.MessageBox]::Show('Chưa chọn file Excel hợp lệ.', 'Thiếu file Excel',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $path = (Resolve-Path -LiteralPath $path).Path
        Write-Log ('Đang đọc file Excel: {0}' -f $path)
        $script:Sheets = @(Read-Workbook -Path $path)
        $clbSheets.Items.Clear()

        foreach ($sheet in $script:Sheets) {
            $blocks = Find-InvoiceBlocks $sheet
            $count  = ($blocks | ForEach-Object { $_.Rows.Count } | Measure-Object -Sum).Sum
            if (-not $count) { $count = 0 }
            $label = '{0}  ({1} hóa đơn / {2} sản phẩm)' -f $sheet.Name, $count, $blocks.Count
            $index = $clbSheets.Items.Add($label)
            if ($count -gt 0) { $clbSheets.SetItemChecked($index, $true) }
            Write-Log ('  • Sheet "{0}": {1} sản phẩm, {2} hóa đơn.' -f $sheet.Name, $blocks.Count, $count)
        }

        if ($clbSheets.Items.Count -eq 0) { Write-Log 'File Excel không có sheet nào đọc được.' 'WARN' }
        else { Write-Log 'Đọc xong. Hãy tick chọn các sheet cần làm rồi bấm "Đối chiếu danh sách".' }
    } catch {
        Write-Log ('Lỗi đọc Excel: {0}' -f $_.Exception.Message) 'ERR'
        [void][System.Windows.Forms.MessageBox]::Show($_.Exception.Message, 'Lỗi đọc Excel',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    } finally {
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Get-CheckedSheets {
    $names = @()
    foreach ($i in $clbSheets.CheckedIndices) {
        $label = [string]$clbSheets.Items[$i]
        $names += ($label -replace '\s+\(\d+ hóa đơn.*$', '')
    }
    return $names
}

function Update-MatchList {
    $lvMatch.BeginUpdate()
    $lvMatch.Items.Clear()
    foreach ($inv in $script:Invoices) {
        $row = New-Object System.Windows.Forms.ListViewItem($inv.Sheet)
        [void]$row.SubItems.Add($inv.Product)
        [void]$row.SubItems.Add($inv.Symbol)
        [void]$row.SubItems.Add($inv.Number)
        [void]$row.SubItems.Add($inv.DateText)
        [void]$row.SubItems.Add($inv.Status)
        [void]$row.SubItems.Add($inv.SourceName)
        switch -Wildcard ($inv.Status) {
            'Không tìm thấy*'  { $row.ForeColor = [System.Drawing.Color]::Firebrick }
            'Khớp nhiều file*' { $row.ForeColor = [System.Drawing.Color]::DarkOrange }
            'Khớp số*'         { $row.ForeColor = [System.Drawing.Color]::DarkGoldenrod }
            'Đã chép*'         { $row.ForeColor = [System.Drawing.Color]::ForestGreen }
            'Đã chuyển*'       { $row.ForeColor = [System.Drawing.Color]::ForestGreen }
        }
        $row.Tag = $inv
        [void]$lvMatch.Items.Add($row)
    }
    $lvMatch.EndUpdate()
}

function Invoke-Match {
    if ($script:Sheets.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('Chưa đọc file Excel.', 'Thiếu danh sách',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $sheetNames = Get-CheckedSheets
    if ($sheetNames.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('Chưa tick chọn sheet nào.', 'Thiếu sheet',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $sources = @($lstSources.Items | ForEach-Object { [string]$_ })
    if ($sources.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('Chưa thêm thư mục nguồn chứa file hóa đơn.', 'Thiếu thư mục nguồn',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $btnMatch.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        $exts  = @($script:FileGroups[$cboSrcType.SelectedItem.ToString()])
        $years = Get-YearFilter
        if ($years.Count -gt 0) { Write-Log ('Chỉ lấy hóa đơn năm: {0}' -f ($years -join ', ')) }

        $index = Build-FileIndex -Folders $sources -Recurse $chkSubSrc.Checked -Extensions $exts
        Write-Log ('Tổng cộng {0} file trong kho nguồn.' -f $index.Count)

        # gom trước để biết tổng số dòng (chạy thanh tiến trình)
        $plan = @()
        foreach ($name in $sheetNames) {
            $sheet = $script:Sheets | Where-Object { $_.Name -eq $name } | Select-Object -First 1
            if (-not $sheet) { continue }
            $plan += [pscustomobject]@{ Sheet = $name; Blocks = (Find-InvoiceBlocks $sheet) }
        }
        $total = ($plan | ForEach-Object { $_.Blocks | ForEach-Object { $_.Rows.Count } } | Measure-Object -Sum).Sum
        if (-not $total) { $total = 0 }

        $bar.Minimum = 0
        $bar.Maximum = [math]::Max(1, $total)
        $bar.Value   = 0

        $script:Invoices = @()
        $done = 0
        foreach ($p in $plan) {
            $seq = 0     # số thứ tự chạy liên tục trong một sheet, đúng thứ tự trong Excel
            foreach ($block in $p.Blocks) {
                foreach ($r in $block.Rows) {
                    $done++
                    if ($years.Count -gt 0 -and $r.Year -and ($years -notcontains $r.Year)) { continue }
                    $seq++
                    $found = Find-InvoiceFile -Index $index -Symbol $r.Symbol -Number $r.Number -Year $r.Year
                    $file  = if ($found.Files.Count -gt 0) { $found.Files[0].File } else { $null }
                    $status = $found.Status
                    if ($found.Files.Count -gt 1) {
                        $status = '{0} ({1} file)' -f $found.Status, $found.Files.Count
                    }
                    $script:Invoices += [pscustomobject]@{
                        Seq        = $seq
                        Sheet      = $p.Sheet
                        Product    = $block.Title
                        Symbol     = $r.Symbol
                        Number     = $r.Number
                        DateText   = $r.DateText
                        Year       = $r.Year
                        Status     = $status
                        MatchCount = $found.Files.Count
                        SourcePath = if ($file) { $file.FullName } else { '' }
                        SourceName = if ($file) { $file.Name } else { '' }
                        DestPath   = ''
                    }
                    if (($done % 25) -eq 0) { $bar.Value = [math]::Min($done, $bar.Maximum); [System.Windows.Forms.Application]::DoEvents() }
                }
            }
        }
        $bar.Value = $bar.Maximum

        Update-MatchList

        $ok      = @($script:Invoices | Where-Object { $_.SourcePath }).Count
        $missing = @($script:Invoices | Where-Object { -not $_.SourcePath }).Count
        $many    = @($script:Invoices | Where-Object { $_.MatchCount -gt 1 }).Count
        $lblMatchInfo.Text = 'Đối chiếu: {0} hóa đơn — tìm thấy {1}, thiếu {2}, khớp nhiều file {3}.' -f $script:Invoices.Count, $ok, $missing, $many
        Write-Log $lblMatchInfo.Text
        if ($missing -gt 0) { Write-Log ('{0} hóa đơn chưa tìm thấy file (dòng màu đỏ) — kiểm tra lại thư mục nguồn hoặc cách đặt tên file.' -f $missing) 'WARN' }

        $btnOrganize.Enabled  = ($ok -gt 0)
        $btnPrintList.Enabled = ($ok -gt 0)
        $btnReport.Enabled    = ($script:Invoices.Count -gt 0)
    } catch {
        Write-Log ('Lỗi khi đối chiếu: {0}' -f $_.Exception.Message) 'ERR'
    } finally {
        $btnMatch.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Invoke-Organize {
    $dest = $txtDest.Text.Trim()
    if (-not $dest) {
        [void][System.Windows.Forms.MessageBox]::Show('Chưa chọn thư mục đích.', 'Thiếu thư mục đích',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $withFile = @($script:Invoices | Where-Object { $_.SourcePath })
    if ($withFile.Count -eq 0) { return }

    $move = $rdoMove.Checked
    $verb = if ($move) { 'DI CHUYỂN' } else { 'CHÉP' }
    $answer = [System.Windows.Forms.MessageBox]::Show(
        ("{0} {1} file hóa đơn sang:`n{2}`n`nTiếp tục?" -f $verb, $withFile.Count, $dest), 'Xác nhận',
        [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Question)
    if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $btnOrganize.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try {
        if (-not (Test-Path -LiteralPath $dest)) {
            [void](New-Item -ItemType Directory -Path $dest -Force)
            Write-Log ('Đã tạo thư mục đích "{0}".' -f $dest)
        }

        $bar.Minimum = 0
        $bar.Maximum = $withFile.Count
        $bar.Value   = 0

        # bề rộng số thứ tự: đủ chữ số cho sheet nhiều hóa đơn nhất (001..999, 0001..9999)
        $widths = @{}
        foreach ($g in ($script:Invoices | Group-Object Sheet)) {
            $max = ($g.Group | Measure-Object Seq -Maximum).Maximum
            $widths[$g.Name] = [math]::Max(3, ([string]$max).Length)
        }

        $done = 0; $ok = 0; $skip = 0; $fail = 0
        foreach ($inv in $withFile) {
            $done++
            try {
                $dir = Join-Path $dest (Get-SafeName $inv.Sheet)
                if ($chkPerProduct.Checked) { $dir = Join-Path $dir (Get-SafeName $inv.Product) }
                if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }

                $ext  = [System.IO.Path]::GetExtension($inv.SourcePath)
                $w = $widths[$inv.Sheet]
                if (-not $w) { $w = 3 }
                $name = if ($chkPrefix.Checked) { '{0}_{1}_{2}' -f ([string]$inv.Seq).PadLeft($w, '0'), $inv.Symbol, $inv.Number }
                        else { '{0}_{1}' -f $inv.Symbol, $inv.Number }
                $target = Join-Path $dir ((Get-SafeName $name) + $ext)

                if (Test-Path -LiteralPath $target) {
                    $src = Get-Item -LiteralPath $inv.SourcePath
                    $dst = Get-Item -LiteralPath $target
                    if ($src.Length -eq $dst.Length) {
                        $inv.DestPath = $target
                        $inv.Status   = 'Đã có sẵn'
                        $skip++
                        continue
                    }
                    for ($k = 2; $k -le 20; $k++) {
                        $try = Join-Path $dir ((Get-SafeName $name) + ('_{0}' -f $k) + $ext)
                        if (-not (Test-Path -LiteralPath $try)) { $target = $try; break }
                    }
                }

                if ($move) { Move-Item -LiteralPath $inv.SourcePath -Destination $target -Force }
                else       { Copy-Item -LiteralPath $inv.SourcePath -Destination $target -Force }

                $inv.DestPath = $target
                $inv.Status   = if ($move) { 'Đã chuyển' } else { 'Đã chép' }
                $ok++
            } catch {
                $inv.Status = 'Lỗi: ' + $_.Exception.Message
                $fail++
                Write-Log ('Lỗi với hóa đơn {0} {1}: {2}' -f $inv.Symbol, $inv.Number, $_.Exception.Message) 'ERR'
            }
            $bar.Value = [math]::Min($done, $bar.Maximum)
            if (($done % 20) -eq 0) { [System.Windows.Forms.Application]::DoEvents() }
        }

        Update-MatchList
        Write-Log ('Sắp xếp xong: {0} file, bỏ qua {1} file đã có, {2} lỗi.' -f $ok, $skip, $fail)

        $reportPath = Join-Path $dest 'bao-cao-doi-chieu.csv'
        Export-MatchReport -Path $reportPath
        Write-Log ('Báo cáo đối chiếu: {0}' -f $reportPath)

        # chuẩn bị sẵn tab In: đúng thư mục, đúng thứ tự tên file (001_, 002_...)
        $txtFolder.Text = $dest
        $chkSub.Checked = $true
        $cboSort.SelectedIndex = 2
        Write-Log 'Tab "2. In hóa đơn" đã trỏ sẵn vào thư mục đích, sắp theo đúng thứ tự trong Excel.'
    } finally {
        $btnOrganize.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
}

function Export-MatchReport {
    param([string]$Path)
    $rows = foreach ($inv in $script:Invoices) {
        [pscustomobject]@{
            STT        = $inv.Seq
            Sheet      = $inv.Sheet
            SanPham    = $inv.Product
            KyHieu     = $inv.Symbol
            SoHoaDon   = $inv.Number
            NgayHoaDon = $inv.DateText
            Nam        = $inv.Year
            TrangThai  = $inv.Status
            FileNguon  = $inv.SourcePath
            FileDich   = $inv.DestPath
        }
    }
    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

# ============================================================================
#  LOGIC TAB 2 — IN
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

        $script:AllFiles = @(
            foreach ($f in $raw) {
                $num = Get-InvoiceNumberFromName -BaseName $f.BaseName -Pattern $pattern
                [pscustomobject]@{
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
        [void][System.Windows.Forms.MessageBox]::Show('Chưa chọn máy in (tab "2. In hóa đơn").', 'Chưa chọn máy in',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $printer = $cboPrinter.SelectedItem.ToString()
    $copies  = [int]$numCopies.Value
    $delay   = [int]$numDelay.Value
    $dry     = $chkDryRun.Checked

    $msg = "In {0} hóa đơn × {1} bản ra máy in:`n{2}`n`nTiếp tục?" -f $Jobs.Count, $copies, $printer
    if ($dry) { $msg = "CHẾ ĐỘ IN THỬ — không gửi gì ra máy in.`n`n" + $msg }
    if ([System.Windows.Forms.MessageBox]::Show($msg, 'Xác nhận in',
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question) -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    $script:Cancel = $false
    $script:Busy   = $true
    $btnPrint.Enabled = $false; $btnPrintList.Enabled = $false; $btnStop.Enabled = $true; $btnScan.Enabled = $false
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
    } finally {
        if ($prevDefault) {
            if (Set-DefaultPrinterName $prevDefault) { Write-Log ('Đã trả máy in mặc định về "{0}".' -f $prevDefault) }
        }
        $script:Busy = $false
        $btnPrint.Enabled = $true; $btnStop.Enabled = $false; $btnScan.Enabled = $true
        $btnPrintList.Enabled = (@($script:Invoices | Where-Object { $_.SourcePath }).Count -gt 0)
    }
}

# ============================================================================
#  SỰ KIỆN
# ============================================================================

# --- Tab 1 -------------------------------------------------------------------
$btnBrowseExcel.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'File Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|Tất cả các file (*.*)|*.*'
    $dlg.Title  = 'Chọn file Excel danh sách hóa đơn'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtExcel.Text = $dlg.FileName
        Invoke-LoadExcel
    }
})

$btnLoadExcel.Add_Click({ Invoke-LoadExcel })

$btnAddSrc.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Chọn thư mục chứa file hóa đơn'
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($lstSources.Items -notcontains $dlg.SelectedPath) {
            [void]$lstSources.Items.Add($dlg.SelectedPath)
            Write-Log ('Thêm thư mục nguồn: {0}' -f $dlg.SelectedPath)
        }
    }
})

$btnRemoveSrc.Add_Click({
    if ($lstSources.SelectedIndex -ge 0) { $lstSources.Items.RemoveAt($lstSources.SelectedIndex) }
})

$btnBrowseDest.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Chọn thư mục đích để sắp hóa đơn vào'
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtDest.Text = $dlg.SelectedPath }
})

$btnMatch.Add_Click({ Invoke-Match })
$btnOrganize.Add_Click({ Invoke-Organize })

$btnPrintList.Add_Click({
    # In đúng thứ tự trong file Excel: dùng file đã sắp nếu có, chưa sắp thì in file gốc.
    $jobs = @()
    foreach ($inv in $script:Invoices) {
        $path = if ($inv.DestPath -and (Test-Path -LiteralPath $inv.DestPath)) { $inv.DestPath } else { $inv.SourcePath }
        if (-not $path) { continue }
        $jobs += @{
            Path  = $path
            Label = '{0} | {1} | {2} {3}' -f $inv.Sheet, $inv.Product, $inv.Symbol, $inv.Number
        }
    }
    $missing = @($script:Invoices | Where-Object { -not $_.SourcePath }).Count
    if ($missing -gt 0) {
        Write-Log ('Bỏ qua {0} hóa đơn chưa tìm thấy file.' -f $missing) 'WARN'
    }
    Invoke-PrintJobs -Jobs $jobs
})

$btnReport.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'File CSV (*.csv)|*.csv'
    $dlg.FileName = 'bao-cao-doi-chieu.csv'
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    Export-MatchReport -Path $dlg.FileName
    Write-Log ('Đã xuất báo cáo ra "{0}".' -f $dlg.FileName)
})

$btnGotoPrint.Add_Click({ $tabs.SelectedTab = $tabPrint })

$lvMatch.Add_DoubleClick({
    $sel = $lvMatch.SelectedItems
    if ($sel.Count -eq 0) { return }
    $inv = $sel[0].Tag
    $path = if ($inv.DestPath) { $inv.DestPath } else { $inv.SourcePath }
    if ($path -and (Test-Path -LiteralPath $path)) { Start-Process -FilePath $path }
})

# --- Tab 2 -------------------------------------------------------------------
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
    $txtDest.Text   = Join-Path $start 'Hoa-don-da-sap-xep'
    Write-Log 'Sẵn sàng.'
    Write-Log 'Tab 1: chọn file Excel → tick sheet → thêm thư mục nguồn → "Đối chiếu danh sách" → "Tạo folder && chép file".'
    Write-Log 'Tab 2: in theo thứ tự đã sắp, hoặc bấm "In ngay theo thứ tự Excel" ở tab 1.'
    $btnRefreshPrinters.PerformClick()
})

[void]$form.ShowDialog()
$form.Dispose()
