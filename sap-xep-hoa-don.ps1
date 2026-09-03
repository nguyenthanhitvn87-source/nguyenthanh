<#
    SẮP XẾP HÓA ĐƠN THEO DANH SÁCH EXCEL — Nguyễn Thanh
    ------------------------------------------------------------------
    Gom các file hóa đơn nằm rải rác nhiều năm vào đúng thư mục theo danh sách Excel:

      • Đọc file Excel danh sách hóa đơn — mỗi sheet là một thư mục.
      • Trong mỗi sheet tự dò các bảng "Ký hiệu / Số hóa đơn / Ngày hóa đơn",
        tên sản phẩm lấy từ dòng tiêu đề phía trên bảng; mỗi sản phẩm là một
        thư mục con, trong đó có thể gồm nhiều ký hiệu hóa đơn.
      • Có cột "tên file" (mẫu kiểu *K25TAA*618585) thì tìm file theo mẫu đó trước.
      • Dò tìm trong các thư mục nguồn (kể cả thư mục con), lọc theo năm nếu cần.
      • Chép hoặc di chuyển sang: <thư mục đích>\<tên sheet>\<tên sản phẩm>\file
        Thư mục đích đã có sẵn thư mục tên ngắn hơn thì tự ghép, sửa tay được.
      • Ghi báo cáo đối chiếu (bao-cao-doi-chieu.csv) và danh sách thứ tự in
        (thu-tu-in.txt) để công cụ in đọc lại đúng thứ tự trong Excel.

    Cách chạy:
        Nhấp đúp vào "sap-xep-hoa-don.bat"
        hoặc: powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\sap-xep-hoa-don.ps1

    In hóa đơn là công cụ riêng: "in-hoa-don.ps1" (bấm nút "Mở công cụ in").

    Yêu cầu: Windows 7 trở lên, Windows PowerShell 5.1 (có sẵn trong Windows).
             File .xlsx/.xlsm đọc thẳng, không cần cài Excel; file .xls cũ cần có Excel.
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
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()
Write-Trace 'Đã nạp xong thư viện giao diện'

# ============================================================================
#  BIẾN DÙNG CHUNG
# ============================================================================
$script:Sheets           = @()    # các sheet đọc được từ Excel
$script:Invoices         = @()    # danh sách hóa đơn đã dò (sau khi đối chiếu)
$script:SheetFolderMap   = @{}    # tên sheet -> tên thư mục trong thư mục đích
$script:ProductFolderMap = @{}    # "tên sheet|tên sản phẩm" -> tên thư mục con
$script:Cancel           = $false
$script:BlocksBySheet    = @{}    # tên sheet -> các bảng hóa đơn đã dò (khỏi dò lại)
$script:FileIndex        = $null  # kết quả quét kho nguồn, dùng lại cho các lần đối chiếu sau
$script:FileIndexKey     = ''     # dấu hiệu của lần quét (thư mục + loại file + thư mục con)

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

function Get-MatchKey {
    <# Chuẩn hóa tên để so khớp: bỏ dấu, bỏ khoảng trắng và ký tự lạ, viết thường. #>
    param([string]$Text)
    $t = Remove-Diacritics $Text
    return ($t -replace '[^A-Za-z0-9]', '').ToLower()
}

function Get-NameSimilarity {
    <# Độ giống nhau giữa hai tên (hệ số Dice trên các cặp ký tự liền nhau), 0..1. #>
    param([string]$A, [string]$B)
    if (-not $A -or -not $B) { return 0.0 }
    if ($A -eq $B) { return 1.0 }
    if ($A.Length -lt 2 -or $B.Length -lt 2) { return 0.0 }

    $pairsA = @()
    for ($i = 0; $i -lt $A.Length - 1; $i++) { $pairsA += $A.Substring($i, 2) }
    $pairsB = New-Object 'System.Collections.Generic.List[string]'
    for ($i = 0; $i -lt $B.Length - 1; $i++) { $pairsB.Add($B.Substring($i, 2)) }

    $countB = $pairsB.Count
    $hit = 0
    foreach ($pair in $pairsA) {
        $idx = $pairsB.IndexOf($pair)
        if ($idx -ge 0) { $hit++; $pairsB.RemoveAt($idx) }
    }
    return (2.0 * $hit) / ($pairsA.Count + $countB)
}

function Get-FolderMatchScore {
    <# Chấm điểm mức hợp giữa một tên trong Excel và một thư mục có sẵn.
       Thư mục hay đặt tên ngắn: "TELMA 80 H PLUS (TABLET B/100)" -> "TELMA",
       nên tên thư mục là phần đầu hoặc nằm trong tên Excel được chấm điểm cao. #>
    param([string]$NameKey, [string]$FolderKey)

    if (-not $NameKey -or -not $FolderKey) { return 0.0 }
    if ($NameKey -eq $FolderKey) { return 1.0 }

    $ratio = [double]$FolderKey.Length / [double]$NameKey.Length
    if ($ratio -gt 1) { $ratio = 1.0 / $ratio }

    if ($NameKey.StartsWith($FolderKey))  { return 0.90 + 0.09 * $ratio }
    if ($FolderKey.StartsWith($NameKey))  { return 0.85 + 0.09 * $ratio }
    if ($NameKey.Contains($FolderKey))    { return 0.80 + 0.09 * $ratio }
    if ($FolderKey.Contains($NameKey))    { return 0.75 + 0.09 * $ratio }

    return (Get-NameSimilarity -A $NameKey -B $FolderKey)
}

function Resolve-FolderMap {
    <# Ghép mỗi tên (sheet hoặc sản phẩm) với một thư mục có sẵn, mỗi thư mục chỉ
       nhận một tên: chấm điểm mọi cặp rồi lấy cặp điểm cao trước. Tên nào không
       hợp thư mục nào thì lấy chính nó làm thư mục mới. #>
    param([string[]]$Names, [string[]]$Existing)

    $map = @{}
    $names    = @($Names    | Where-Object { $_ })
    $existing = @($Existing | Where-Object { $_ })
    if ($names.Count -eq 0) { return $map }

    $pairs = @()
    foreach ($name in $names) {
        $nameKey = Get-MatchKey $name
        foreach ($folder in $existing) {
            $pairs += [pscustomobject]@{
                Name   = $name
                Folder = $folder
                Score  = (Get-FolderMatchScore -NameKey $nameKey -FolderKey (Get-MatchKey $folder))
            }
        }
    }
    $pairs = @($pairs | Sort-Object Score -Descending)

    $takenName   = @{}
    $takenFolder = @{}
    $assign = {
        param([double]$MinScore)
        foreach ($pair in $pairs) {
            if ($pair.Score -lt $MinScore) { break }
            if ($takenName.ContainsKey($pair.Name) -or $takenFolder.ContainsKey($pair.Folder)) { continue }
            $map[$pair.Name]           = $pair.Folder
            $takenName[$pair.Name]     = $true
            $takenFolder[$pair.Folder] = $true
        }
    }

    & $assign 0.6      # vòng chắc chắn: trùng tên, hoặc tên thư mục là phần đầu của tên Excel
    & $assign 0.4      # vòng nới tay cho các tên viết tắt khác kiểu

    # còn đúng một tên và một thư mục chưa ghép thì ghép nốt (vẫn sửa lại được ở hộp thoại)
    $leftName   = @($names    | Where-Object { -not $takenName.ContainsKey($_) })
    $leftFolder = @($existing | Where-Object { -not $takenFolder.ContainsKey($_) })
    if ($leftName.Count -eq 1 -and $leftFolder.Count -eq 1) {
        $score = Get-FolderMatchScore -NameKey (Get-MatchKey $leftName[0]) -FolderKey (Get-MatchKey $leftFolder[0])
        if ($score -ge 0.25) {
            $map[$leftName[0]]        = $leftFolder[0]
            $takenName[$leftName[0]]  = $true
        }
    }

    foreach ($name in $names) {
        if (-not $map.ContainsKey($name)) { $map[$name] = Get-SafeName $name }
    }
    return $map
}

function Get-SubFolderNames {
    param([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path -PathType Container)) { return ,@() }
    return ,@(Get-ChildItem -LiteralPath $Path -Directory -ErrorAction SilentlyContinue |
              Select-Object -ExpandProperty Name)
}

function Resolve-SheetFolders {
    <# Ghép mỗi sheet với thư mục có sẵn trong thư mục đích. #>
    param([string[]]$SheetNames, [string]$Dest)
    return Resolve-FolderMap -Names $SheetNames -Existing (Get-SubFolderNames $Dest)
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
        $shared = New-Object 'System.Collections.Generic.List[string]'
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
$script:HeaderSymbol = @('ky hieu', 'ki hieu', 'serial')
$script:HeaderNumber = @('so hoa don', 'so hd', 'invoice no')
$script:HeaderDate   = @('ngay hoa don', 'ngay hd', 'ngay lap hoa don', 'ngay')
$script:HeaderFile   = @('ten file', 'ten tap tin', 'file')

function Test-HeaderMatch {
    <# Tiêu đề cột trong thực tế hay có đuôi lạ ("Số hóa đơn2", "Ký hiệu HĐ") nên so
       khớp theo phần đầu; các từ ngắn (dưới 5 ký tự) thì phải trùng khít. #>
    param([string]$Text, [string[]]$Keys)

    $key = Get-HeaderKey $Text
    if (-not $key) { return $false }
    foreach ($k in $Keys) {
        if ($key -eq $k) { return $true }
        if ($k.Length -ge 5 -and $key.StartsWith($k)) { return $true }
    }
    return $false
}

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
        if (Test-HeaderMatch $Sheet.Cells[$key] $script:HeaderNumber) {
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
            if (Test-HeaderMatch (Get-Cell $Sheet $r ($c - $k)) $script:HeaderSymbol) { $symbolCol = $c - $k; break }
        }
        if ($symbolCol -eq 0) { continue }

        # cột "Ngày hóa đơn" nằm bên phải
        $dateCol = 0
        for ($k = 1; $k -le 5; $k++) {
            if (Test-HeaderMatch (Get-Cell $Sheet $r ($c + $k)) $script:HeaderDate) { $dateCol = $c + $k; break }
        }

        # cột "tên file" (mẫu tìm file, VD *K25TAA*618585) cũng nằm bên phải
        $fileCol = 0
        for ($k = 1; $k -le 6; $k++) {
            if (Test-HeaderMatch (Get-Cell $Sheet $r ($c + $k)) $script:HeaderFile) { $fileCol = $c + $k; break }
        }

        # cột mẫu tên file có thể không có tiêu đề — nhận ra bằng dấu * trong dữ liệu
        if ($fileCol -eq 0) {
            $lastRow = [math]::Min($r + 6, $Sheet.MaxRow)
            for ($k = 1; $k -le 6 -and $fileCol -eq 0; $k++) {
                $cc = $c + $k
                if ($cc -eq $dateCol) { continue }
                for ($rr2 = $r + 1; $rr2 -le $lastRow; $rr2++) {
                    $v = Get-Cell $Sheet $rr2 $cc
                    if ($v -and $v.Contains('*')) { $fileCol = $cc; break }
                }
            }
        }

        # tên sản phẩm/mẫu: ô có chữ gần nhất phía trên bảng
        $skipWords = $script:HeaderSymbol + $script:HeaderNumber + $script:HeaderDate + $script:HeaderFile + @('stt')
        $title = ''
        for ($up = 1; $up -le 4 -and -not $title; $up++) {
            if ($r - $up -lt 1) { break }
            for ($cc = [math]::Max(1, $symbolCol - 2); $cc -le $c + 3; $cc++) {
                $t = Get-Cell $Sheet ($r - $up) $cc
                if ($t -and -not (Test-HeaderMatch $t $skipWords)) { $title = $t; break }
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
                    Pattern   = if ($fileCol -gt 0) { (Get-Cell $Sheet $rr $fileCol).Trim() } else { '' }
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
function Test-CloudOnlyFile {
    <# File OneDrive/SharePoint mới chỉ có trên mây, chưa tải về máy.
       Offline = 0x1000, RecallOnOpen = 0x40000, RecallOnDataAccess = 0x400000 #>
    param($File)
    $attr = [int]$File.Attributes
    return (($attr -band 0x1000) -ne 0) -or (($attr -band 0x40000) -ne 0) -or (($attr -band 0x400000) -ne 0)
}

function Test-PathInside {
    <# $Child có nằm bên trong $Parent không (để cảnh báo thư mục đích nằm trong nguồn). #>
    param([string]$Child, [string]$Parent)
    try {
        $c = [System.IO.Path]::GetFullPath($Child).TrimEnd('\') + '\'
        $p = [System.IO.Path]::GetFullPath($Parent).TrimEnd('\') + '\'
        return $c.StartsWith($p, [StringComparison]::OrdinalIgnoreCase)
    } catch { return $false }
}

function Build-FileIndex {
    <# Quét các thư mục nguồn một lần, chuẩn hóa tên file và lập chỉ mục theo số
       trong tên để dò từng hóa đơn cho nhanh (không phải duyệt lại cả kho). #>
    param([string[]]$Folders, [bool]$Recurse, [string[]]$Extensions)

    $files     = New-Object 'System.Collections.Generic.List[object]'
    $byNumber  = @{}
    $cloudOnly = 0
    $started   = Get-Date

    foreach ($folder in $Folders) {
        if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
            Write-Log ('Bỏ qua thư mục không tồn tại: {0}' -f $folder) 'WARN'
            continue
        }

        $t0 = Get-Date
        $raw = New-Object 'System.Collections.Generic.List[object]'
        $readErrors = @()
        if ($Extensions -and $Extensions.Count -gt 0) {
            # lọc ngay khi liệt kê — nhanh hơn nhiều so với lấy hết rồi lọc sau
            foreach ($ext in $Extensions) {
                foreach ($f in @(Get-ChildItem -LiteralPath $folder -File -Recurse:$Recurse -Filter ('*' + $ext) -ErrorAction SilentlyContinue -ErrorVariable +readErrors)) {
                    if ($Extensions -contains $f.Extension.ToLower()) { $raw.Add($f) }
                }
                [System.Windows.Forms.Application]::DoEvents()
            }
        } else {
            foreach ($f in @(Get-ChildItem -LiteralPath $folder -File -Recurse:$Recurse -ErrorAction SilentlyContinue -ErrorVariable +readErrors)) {
                $raw.Add($f)
            }
        }

        if ($readErrors.Count -gt 0) {
            Write-Log ('{0} thư mục con không đọc được (thiếu quyền hoặc đường dẫn quá dài). Ví dụ: {1}' -f $readErrors.Count, $readErrors[0].Exception.Message) 'WARN'
        }
        if ($raw.Count -eq 0) {
            $what = if ($Extensions -and $Extensions.Count -gt 0) { ($Extensions -join ', ') } else { 'bất kỳ loại nào' }
            Write-Log ('Không thấy file {0} nào trong "{1}"{2}.' -f $what, $folder, $(if ($Recurse) { ' (đã tính cả thư mục con)' } else { ' — chưa bật "Gồm thư mục con"' })) 'WARN'
        }

        foreach ($f in $raw) {
            if (Test-CloudOnlyFile $f) { $cloudOnly++ }
            $entry = [pscustomobject]@{
                File   = $f
                Norm   = ((Remove-Diacritics $f.BaseName).ToUpper() -replace '[^A-Z0-9]', '')
                Digits = @([regex]::Matches($f.BaseName, '\d+') | ForEach-Object { $_.Value.TrimStart('0') })
                Year   = $f.LastWriteTime.Year
            }
            $files.Add($entry)
            foreach ($d in $entry.Digits) {
                $key = if ($d) { $d } else { '0' }
                if (-not $byNumber.ContainsKey($key)) { $byNumber[$key] = New-Object 'System.Collections.Generic.List[object]' }
                $byNumber[$key].Add($entry)
            }
        }

        Write-Log ('Đã quét {0} file trong "{1}" ({2:N1} giây).' -f $raw.Count, $folder, ((Get-Date) - $t0).TotalSeconds)
        [System.Windows.Forms.Application]::DoEvents()
    }

    if ($cloudOnly -gt 0) {
        Write-Log ('{0} file đang ở dạng đám mây (OneDrive chưa tải về máy). Windows sẽ tự tải khi chép/di chuyển — bước sắp xếp sẽ chậm hơn và cần có mạng.' -f $cloudOnly) 'WARN'
    }
    Write-Log ('Tổng cộng {0} file, lập chỉ mục xong sau {1:N1} giây.' -f $files.Count, ((Get-Date) - $started).TotalSeconds)
    if ($files.Count -eq 0) {
        Write-Log 'KHO NGUỒN KHÔNG CÓ FILE NÀO. Kiểm tra: đúng thư mục chưa, đã bật "Gồm thư mục con" chưa, và "Loại file" có đúng với file hóa đơn của bạn không.' 'WARN'
    } else {
        $samples = @($files | Select-Object -First 3 | ForEach-Object { $_.File.Name })
        Write-Log ('Ví dụ tên file trong kho: {0}' -f ($samples -join '  |  '))
    }

    return [pscustomobject]@{ Files = $files.ToArray(); ByNumber = $byNumber }
}

function Select-BestMatches {
    <# Nhiều file cùng khớp thì ưu tiên file nằm trong thư mục/tên có đúng năm hóa đơn
       (kho chia theo "HCM - 2023", "HCM - 2024", "HCM - 2025", "Ha Noi"...),
       sau đó ưu tiên đường dẫn ngắn hơn (bản gốc thay vì bản sao chép lồng nhau).
       Viết bằng vòng lặp thường để chạy chắc trên mọi bản PowerShell. #>
    param($Files, $Year)

    $list = @()
    foreach ($f in $Files) { if ($f) { $list += $f } }
    if ($list.Count -le 1) { return $list }

    if ($Year) {
        $yearText = [string]$Year
        $inYear = @()
        foreach ($f in $list) {
            if ($f.File.FullName.Contains($yearText)) { $inYear += $f }
        }
        if ($inYear.Count -gt 0) { $list = $inYear }
    }
    if ($list.Count -le 1) { return $list }

    # sắp theo độ dài đường dẫn, bằng vòng lặp chèn — không dùng Sort-Object
    $sorted = @()
    foreach ($f in $list) {
        $placed = $false
        $out = @()
        foreach ($x in $sorted) {
            if (-not $placed -and $f.File.FullName.Length -lt $x.File.FullName.Length) {
                $out += $f
                $placed = $true
            }
            $out += $x
        }
        if (-not $placed) { $out += $f }
        $sorted = $out
    }
    return $sorted
}

function Find-InvoiceFile {
    <# Tìm file cho một hóa đơn. Chỉ xét những file có chứa đúng số hóa đơn trong tên
       (lấy từ chỉ mục), ưu tiên cột "tên file" trong Excel rồi mới tới ký hiệu + số. #>
    param($Index, [string]$Symbol, [string]$Number, $Year, [string]$Pattern)

    $numKey = $Number.TrimStart('0')
    if (-not $numKey) { $numKey = '0' }

    $candidates = @()
    if ($Index.ByNumber.ContainsKey($numKey)) { $candidates = $Index.ByNumber[$numKey].ToArray() }

    if ($Pattern) {
        $pat = $Pattern.Trim()
        if (-not $pat.StartsWith('*')) { $pat = '*' + $pat }
        if (-not $pat.EndsWith('*'))   { $pat = $pat + '*' }

        $byPattern = @()
        foreach ($f in $candidates) { if ($f.File.Name -like $pat) { $byPattern += $f } }
        if ($byPattern.Count -eq 0 -and $candidates.Count -eq 0) {
            # mẫu không kèm số hóa đơn thì mới phải duyệt cả kho
            foreach ($f in $Index.Files) { if ($f.File.Name -like $pat) { $byPattern += $f } }
        }
        if ($byPattern.Count -gt 0) {
            $best = @(Select-BestMatches -Files $byPattern -Year $Year)
            if ($best.Count -eq 0) { $best = $byPattern }      # không bao giờ trả về rỗng khi đã có file khớp
            $status = 'Khớp nhiều file'
            if ($best.Count -eq 1) { $status = 'Khớp theo cột tên file' }
            return @{ Files = $best; Status = $status }
        }
    }

    if ($candidates.Count -eq 0) { return @{ Files = @(); Status = 'Không tìm thấy' } }

    $symNorm = (Remove-Diacritics $Symbol).ToUpper() -replace '[^A-Z0-9]', ''
    $strong = @()
    if ($symNorm) {
        foreach ($f in $candidates) { if ($f.Norm.Contains($symNorm)) { $strong += $f } }
    }
    if ($strong.Count -gt 0) {
        $best = @(Select-BestMatches -Files $strong -Year $Year)
        if ($best.Count -eq 0) { $best = $strong }
        $status = 'Khớp nhiều file'
        if ($best.Count -eq 1) { $status = 'Khớp ký hiệu + số' }
        return @{ Files = $best; Status = $status }
    }

    $best = @(Select-BestMatches -Files $candidates -Year $Year)
    if ($best.Count -eq 0) { $best = $candidates }
    if ($best.Count -gt 1 -and $Year) {
        $byTime = @()
        foreach ($f in $best) { if ($f.Year -eq $Year) { $byTime += $f } }
        if ($byTime.Count -gt 0) { $best = $byTime }
    }
    $status = 'Khớp nhiều file'
    if ($best.Count -eq 1) { $status = 'Khớp số (thiếu ký hiệu)' }
    return @{ Files = $best; Status = $status }
}

# ============================================================================
#  HỘP THOẠI GÁN THƯ MỤC (SHEET VÀ SẢN PHẨM)
# ============================================================================
function Show-FolderMapDialog {
    <# Chọn mỗi sheet vào thư mục nào, và mỗi sản phẩm vào thư mục con nào —
       dùng khi thư mục đích đã có sẵn thư mục đặt tên ngắn hơn tên trong Excel
       (TELMA 80 H PLUS (TABLET B/100) -> TELMA). #>
    param([string[]]$SheetNames, [hashtable]$ProductsBySheet, [string]$Dest, $SheetMap, $ProductMap)

    $dlg                 = New-Object System.Windows.Forms.Form
    $dlg.Text            = 'Thư mục đích cho từng sheet và sản phẩm'
    $dlg.Size            = New-Object System.Drawing.Size(820, 560)
    $dlg.StartPosition   = 'CenterParent'
    $dlg.Font            = New-Object System.Drawing.Font('Segoe UI', 9)
    $dlg.FormBorderStyle = 'FixedDialog'
    $dlg.MaximizeBox     = $false
    $dlg.MinimizeBox     = $false

    $lblTop          = New-Object System.Windows.Forms.Label
    $lblTop.Text     = if ($Dest) { 'Thư mục đích: ' + $Dest } else { 'Chưa chọn thư mục đích — các thư mục dưới đây sẽ được tạo mới.' }
    $lblTop.Location = New-Object System.Drawing.Point(12, 12)
    $lblTop.Size     = New-Object System.Drawing.Size(780, 20)
    $dlg.Controls.Add($lblTop)

    $lv               = New-Object System.Windows.Forms.ListView
    $lv.Location      = New-Object System.Drawing.Point(12, 38)
    $lv.Size          = New-Object System.Drawing.Size(780, 320)
    $lv.View          = 'Details'
    $lv.FullRowSelect = $true
    $lv.GridLines     = $true
    $lv.HideSelection = $false
    [void]$lv.Columns.Add('Sheet / sản phẩm trong Excel', 360)
    [void]$lv.Columns.Add('Thư mục đích', 300)
    [void]$lv.Columns.Add('Trạng thái', 100)
    $dlg.Controls.Add($lv)

    $boldFont = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)

    foreach ($sheet in $SheetNames) {
        $folder = $SheetMap[$sheet]
        if (-not $folder) { $folder = Get-SafeName $sheet }
        $row = New-Object System.Windows.Forms.ListViewItem($sheet)
        [void]$row.SubItems.Add($folder)
        [void]$row.SubItems.Add('')
        $row.Font = $boldFont
        $row.Tag  = [pscustomobject]@{ Kind = 'Sheet'; Sheet = $sheet; Product = '' }
        [void]$lv.Items.Add($row)

        foreach ($product in @($ProductsBySheet[$sheet])) {
            if (-not $product) { continue }
            $key = '{0}|{1}' -f $sheet, $product
            $pf  = $ProductMap[$key]
            if (-not $pf) { $pf = Get-SafeName $product }
            $prow = New-Object System.Windows.Forms.ListViewItem('        ' + $product)
            [void]$prow.SubItems.Add($pf)
            [void]$prow.SubItems.Add('')
            $prow.Tag = [pscustomobject]@{ Kind = 'Product'; Sheet = $sheet; Product = $product }
            [void]$lv.Items.Add($prow)
        }
    }

    # thư mục cha của một dòng: sheet thì nằm trong thư mục đích, sản phẩm thì nằm trong thư mục của sheet
    $getParent = {
        param($Row)
        if ($Row.Tag.Kind -eq 'Sheet') { return $Dest }
        foreach ($item in $lv.Items) {
            if ($item.Tag.Kind -eq 'Sheet' -and $item.Tag.Sheet -eq $Row.Tag.Sheet) {
                if (-not $Dest) { return '' }
                return (Join-Path $Dest $item.SubItems[1].Text)
            }
        }
        return ''
    }.GetNewClosure()

    $updateStatus = {
        param($Row)
        $parent = & $getParent $Row
        $has = $false
        if ($parent) { $has = ((Get-SubFolderNames $parent) -contains $Row.SubItems[1].Text) }
        $Row.SubItems[2].Text = if ($has) { 'Đã có sẵn' } else { 'Sẽ tạo mới' }
        $Row.ForeColor = if ($has) { [System.Drawing.SystemColors]::WindowText } else { [System.Drawing.Color]::DarkOrange }
    }.GetNewClosure()

    foreach ($row in $lv.Items) { & $updateStatus $row }

    $lblPick          = New-Object System.Windows.Forms.Label
    $lblPick.Text     = 'Thư mục có sẵn'
    $lblPick.Location = New-Object System.Drawing.Point(12, 374)
    $lblPick.Size     = New-Object System.Drawing.Size(100, 20)
    $dlg.Controls.Add($lblPick)

    $cbo               = New-Object System.Windows.Forms.ComboBox
    $cbo.Location      = New-Object System.Drawing.Point(114, 371)
    $cbo.Size          = New-Object System.Drawing.Size(400, 24)
    $cbo.DropDownStyle = 'DropDownList'
    $dlg.Controls.Add($cbo)

    $btnAssign          = New-Object System.Windows.Forms.Button
    $btnAssign.Text     = 'Gán cho dòng đang chọn'
    $btnAssign.Location = New-Object System.Drawing.Point(522, 369)
    $btnAssign.Size     = New-Object System.Drawing.Size(190, 28)
    $dlg.Controls.Add($btnAssign)

    $btnNew          = New-Object System.Windows.Forms.Button
    $btnNew.Text     = 'Tạo thư mục mới theo tên trong Excel'
    $btnNew.Location = New-Object System.Drawing.Point(12, 406)
    $btnNew.Size     = New-Object System.Drawing.Size(270, 28)
    $dlg.Controls.Add($btnNew)

    $lblHint           = New-Object System.Windows.Forms.Label
    $lblHint.Text      = 'Công cụ đã tự đoán theo tên gần giống. Chọn một dòng, chọn thư mục ở ô bên trên rồi bấm "Gán" nếu cần sửa. Dòng in đậm là sheet, dòng thụt vào là sản phẩm.'
    $lblHint.Location  = New-Object System.Drawing.Point(12, 440)
    $lblHint.Size      = New-Object System.Drawing.Size(780, 36)
    $lblHint.ForeColor = [System.Drawing.Color]::DimGray
    $dlg.Controls.Add($lblHint)

    $btnOk              = New-Object System.Windows.Forms.Button
    $btnOk.Text         = 'Xong'
    $btnOk.Location     = New-Object System.Drawing.Point(596, 480)
    $btnOk.Size         = New-Object System.Drawing.Size(96, 30)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $dlg.Controls.Add($btnOk)
    $dlg.AcceptButton   = $btnOk

    $btnCancel              = New-Object System.Windows.Forms.Button
    $btnCancel.Text         = 'Hủy'
    $btnCancel.Location     = New-Object System.Drawing.Point(696, 480)
    $btnCancel.Size         = New-Object System.Drawing.Size(96, 30)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $dlg.Controls.Add($btnCancel)
    $dlg.CancelButton       = $btnCancel

    $setFolder = {
        param([string]$Folder)
        if ($lv.SelectedItems.Count -eq 0 -or -not $Folder) { return }
        $row = $lv.SelectedItems[0]
        $row.SubItems[1].Text = $Folder
        & $updateStatus $row
        if ($row.Tag.Kind -eq 'Sheet') {
            foreach ($item in $lv.Items) {
                if ($item.Tag.Kind -eq 'Product' -and $item.Tag.Sheet -eq $row.Tag.Sheet) { & $updateStatus $item }
            }
        }
    }.GetNewClosure()

    $lv.Add_SelectedIndexChanged({
        $cbo.Items.Clear()
        if ($lv.SelectedItems.Count -eq 0) { return }
        foreach ($name in (Get-SubFolderNames (& $getParent $lv.SelectedItems[0]))) { [void]$cbo.Items.Add($name) }
        if ($cbo.Items.Count -gt 0) { $cbo.SelectedIndex = 0 }
    }.GetNewClosure())

    $btnAssign.Add_Click({ if ($cbo.SelectedItem) { & $setFolder ([string]$cbo.SelectedItem) } }.GetNewClosure())
    $lv.Add_DoubleClick({  if ($cbo.SelectedItem) { & $setFolder ([string]$cbo.SelectedItem) } }.GetNewClosure())

    $btnNew.Add_Click({
        if ($lv.SelectedItems.Count -eq 0) { return }
        $row = $lv.SelectedItems[0]
        $name = if ($row.Tag.Kind -eq 'Sheet') { $row.Tag.Sheet } else { $row.Tag.Product }
        & $setFolder (Get-SafeName $name)
    }.GetNewClosure())

    if ($lv.Items.Count -gt 0) { $lv.Items[0].Selected = $true }

    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { $dlg.Dispose(); return $null }

    $sheets   = @{}
    $products = @{}
    foreach ($row in $lv.Items) {
        if ($row.Tag.Kind -eq 'Sheet') { $sheets[$row.Tag.Sheet] = $row.SubItems[1].Text }
        else { $products['{0}|{1}' -f $row.Tag.Sheet, $row.Tag.Product] = $row.SubItems[1].Text }
    }
    $dlg.Dispose()
    return @{ Sheets = $sheets; Products = $products }
}

Write-Trace 'Bắt đầu dựng cửa sổ'

# ============================================================================
#  GIAO DIỆN
# ============================================================================
$form               = New-Object System.Windows.Forms.Form
$form.Text          = 'Sắp xếp hóa đơn theo danh sách Excel — Nguyễn Thanh'
$form.Size          = New-Object System.Drawing.Size(1024, 850)
$form.MinimumSize   = New-Object System.Drawing.Size(1024, 780)
$form.StartPosition = 'CenterScreen'
$form.Font          = New-Object System.Drawing.Font('Segoe UI', 9)

$grpExcel          = New-Object System.Windows.Forms.GroupBox
$grpExcel.Text     = '1. File Excel danh sách hóa đơn — mỗi sheet sẽ thành một thư mục'
$grpExcel.Location = New-Object System.Drawing.Point(8, 6)
$grpExcel.Size     = New-Object System.Drawing.Size(984, 150)
$form.Controls.Add($grpExcel)

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
$form.Controls.Add($grpSource)

$lstSources          = New-Object System.Windows.Forms.ListBox
$lstSources.Location = New-Object System.Drawing.Point(12, 26)
$lstSources.Size     = New-Object System.Drawing.Size(730, 52)
$grpSource.Controls.Add($lstSources)

$lblSourcePath          = New-Object System.Windows.Forms.Label
$lblSourcePath.Text     = 'Gõ / dán đường dẫn'
$lblSourcePath.Location = New-Object System.Drawing.Point(12, 86)
$lblSourcePath.Size     = New-Object System.Drawing.Size(120, 20)
$grpSource.Controls.Add($lblSourcePath)

$txtSourcePath          = New-Object System.Windows.Forms.TextBox
$txtSourcePath.Location = New-Object System.Drawing.Point(136, 83)
$txtSourcePath.Size     = New-Object System.Drawing.Size(606, 24)
$grpSource.Controls.Add($txtSourcePath)

$btnAddSrc          = New-Object System.Windows.Forms.Button
$btnAddSrc.Text     = 'Chọn thư mục nguồn...'
$btnAddSrc.Location = New-Object System.Drawing.Point(752, 24)
$btnAddSrc.Size     = New-Object System.Drawing.Size(220, 26)
$grpSource.Controls.Add($btnAddSrc)

$btnRemoveSrc          = New-Object System.Windows.Forms.Button
$btnRemoveSrc.Text     = 'Bỏ thư mục đang chọn'
$btnRemoveSrc.Location = New-Object System.Drawing.Point(752, 54)
$btnRemoveSrc.Size     = New-Object System.Drawing.Size(220, 26)
$grpSource.Controls.Add($btnRemoveSrc)

$btnAddPath          = New-Object System.Windows.Forms.Button
$btnAddPath.Text     = 'Thêm đường dẫn đã gõ'
$btnAddPath.Location = New-Object System.Drawing.Point(752, 83)
$btnAddPath.Size     = New-Object System.Drawing.Size(220, 26)
$grpSource.Controls.Add($btnAddPath)

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
$lblYearsHint.Text      = 'Để trống là lấy tất cả các năm. Trỏ thẳng vào thư mục cần sắp cho nhanh.'
$lblYearsHint.Location  = New-Object System.Drawing.Point(632, 116)
$lblYearsHint.Size      = New-Object System.Drawing.Size(340, 20)
$lblYearsHint.ForeColor = [System.Drawing.Color]::DimGray
$grpSource.Controls.Add($lblYearsHint)

$grpDest          = New-Object System.Windows.Forms.GroupBox
$grpDest.Text     = '3. Thư mục đích — sẽ tạo: <thư mục đích>\<tên sheet>\<tên sản phẩm>\file'
$grpDest.Location = New-Object System.Drawing.Point(8, 318)
$grpDest.Size     = New-Object System.Drawing.Size(984, 90)
$form.Controls.Add($grpDest)

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
$chkPrefix.Text     = 'Thêm số thứ tự vào đầu tên file, giữ nguyên tên gốc'
$chkPrefix.Location = New-Object System.Drawing.Point(520, 58)
$chkPrefix.Size     = New-Object System.Drawing.Size(300, 22)
$chkPrefix.Checked  = $true
$grpDest.Controls.Add($chkPrefix)

$btnMapFolders          = New-Object System.Windows.Forms.Button
$btnMapFolders.Text     = 'Thư mục cho từng sheet...'
$btnMapFolders.Location = New-Object System.Drawing.Point(826, 55)
$btnMapFolders.Size     = New-Object System.Drawing.Size(146, 28)
$grpDest.Controls.Add($btnMapFolders)

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
$form.Controls.Add($lvMatch)


$btnMatch          = New-Object System.Windows.Forms.Button
$btnMatch.Text     = 'Đối chiếu danh sách'
$btnMatch.Location = New-Object System.Drawing.Point(8, 414)
$btnMatch.Size     = New-Object System.Drawing.Size(190, 32)
$btnMatch.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$form.Controls.Add($btnMatch)

$btnOrganize          = New-Object System.Windows.Forms.Button
$btnOrganize.Text     = 'Tạo folder && chép file'
$btnOrganize.Location = New-Object System.Drawing.Point(206, 414)
$btnOrganize.Size     = New-Object System.Drawing.Size(220, 32)
$btnOrganize.Font     = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnOrganize.Enabled  = $false
$form.Controls.Add($btnOrganize)

$btnReport          = New-Object System.Windows.Forms.Button
$btnReport.Text     = 'Xuất báo cáo CSV'
$btnReport.Location = New-Object System.Drawing.Point(434, 414)
$btnReport.Size     = New-Object System.Drawing.Size(170, 32)
$btnReport.Enabled  = $false
$form.Controls.Add($btnReport)

$btnOpenPrint          = New-Object System.Windows.Forms.Button
$btnOpenPrint.Text     = 'Mở công cụ in →'
$btnOpenPrint.Location = New-Object System.Drawing.Point(612, 414)
$btnOpenPrint.Size     = New-Object System.Drawing.Size(180, 32)
$form.Controls.Add($btnOpenPrint)

$btnScanSource          = New-Object System.Windows.Forms.Button
$btnScanSource.Text     = 'Quét lại kho nguồn'
$btnScanSource.Location = New-Object System.Drawing.Point(800, 414)
$btnScanSource.Size     = New-Object System.Drawing.Size(196, 32)
$form.Controls.Add($btnScanSource)

$lblMatchInfo          = New-Object System.Windows.Forms.Label
$lblMatchInfo.Text     = 'Chưa đối chiếu.'
$lblMatchInfo.Location = New-Object System.Drawing.Point(8, 606)
$lblMatchInfo.Size     = New-Object System.Drawing.Size(984, 20)
$form.Controls.Add($lblMatchInfo)

$bar          = New-Object System.Windows.Forms.ProgressBar
$bar.Location = New-Object System.Drawing.Point(8, 632)
$bar.Size     = New-Object System.Drawing.Size(996, 18)
$form.Controls.Add($bar)

$txtLog            = New-Object System.Windows.Forms.TextBox
$txtLog.Location   = New-Object System.Drawing.Point(8, 656)
$txtLog.Size       = New-Object System.Drawing.Size(996, 140)
$txtLog.Multiline  = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly   = $true
$txtLog.BackColor  = [System.Drawing.Color]::White
$txtLog.Font       = New-Object System.Drawing.Font('Consolas', 9)
$form.Controls.Add($txtLog)

# ============================================================================
#  LOGIC
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
        $script:BlocksBySheet = @{}
        $clbSheets.Items.Clear()

        foreach ($sheet in $script:Sheets) {
            $blocks = Find-InvoiceBlocks $sheet
            $script:BlocksBySheet[$sheet.Name] = $blocks
            $count  = ($blocks | ForEach-Object { $_.Rows.Count } | Measure-Object -Sum).Sum
            if (-not $count) { $count = 0 }
            $label = '{0}  ({1} hóa đơn / {2} sản phẩm)' -f $sheet.Name, $count, $blocks.Count
            $index = $clbSheets.Items.Add($label)
            if ($count -gt 0) { $clbSheets.SetItemChecked($index, $true) }

            $withPattern = 0
            foreach ($b in $blocks) { $withPattern += @($b.Rows | Where-Object { $_.Pattern }).Count }
            $note = ''
            if ($count -gt 0 -and $withPattern -eq 0) { $note = ' — không có cột mẫu tên file, sẽ dò theo ký hiệu + số' }
            elseif ($withPattern -lt $count)          { $note = ' — {0} dòng có mẫu tên file' -f $withPattern }
            Write-Log ('  • Sheet "{0}": {1} sản phẩm, {2} hóa đơn{3}.' -f $sheet.Name, $blocks.Count, $count, $note)
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

function Get-SheetBlocks {
    <# Các bảng hóa đơn của một sheet — dò một lần lúc đọc Excel rồi dùng lại. #>
    param([string]$Name)
    if ($script:BlocksBySheet.ContainsKey($Name)) { return $script:BlocksBySheet[$Name] }
    $sheet = $script:Sheets | Where-Object { $_.Name -eq $Name } | Select-Object -First 1
    if (-not $sheet) { return @() }
    $script:BlocksBySheet[$Name] = Find-InvoiceBlocks $sheet
    return $script:BlocksBySheet[$Name]
}

function Get-FileIndex {
    <# Quét kho nguồn. Cùng thư mục + cùng loại file thì dùng lại kết quả lần trước,
       khỏi phải quét lại (kho OneDrive nhiều file quét rất lâu). #>
    param([switch]$Force)

    $sources = @($lstSources.Items | ForEach-Object { [string]$_ })
    $exts    = @($script:FileGroups[$cboSrcType.SelectedItem.ToString()])
    $key     = '{0}|{1}|{2}' -f ($sources -join ';'), ($exts -join ','), $chkSubSrc.Checked

    if (-not $Force -and $script:FileIndex -and $script:FileIndexKey -eq $key) {
        Write-Log ('Dùng lại kết quả quét trước: {0} file (bấm "Quét lại kho nguồn" nếu vừa thêm file mới).' -f $script:FileIndex.Files.Count)
        return $script:FileIndex
    }

    $script:FileIndex    = Build-FileIndex -Folders $sources -Recurse $chkSubSrc.Checked -Extensions $exts
    $script:FileIndexKey = $key
    return $script:FileIndex
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
        $years = Get-YearFilter
        if ($years.Count -gt 0) { Write-Log ('Chỉ lấy hóa đơn năm: {0}' -f ($years -join ', ')) }

        $index = Get-FileIndex

        # gom trước để biết tổng số dòng (chạy thanh tiến trình)
        $plan = @()
        foreach ($name in $sheetNames) {
            $sheet = $script:Sheets | Where-Object { $_.Name -eq $name } | Select-Object -First 1
            if (-not $sheet) { continue }
            $plan += [pscustomobject]@{ Sheet = $name; Blocks = (Get-SheetBlocks $name) }
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
                    $found = Find-InvoiceFile -Index $index -Symbol $r.Symbol -Number $r.Number -Year $r.Year -Pattern $r.Pattern
                    $hits  = @($found.Files)
                    $file  = $null
                    if ($hits.Count -gt 0) { $file = $hits[0].File }
                    $status = $found.Status
                    if ($hits.Count -gt 1)  { $status = '{0} ({1} file)' -f $found.Status, $hits.Count }
                    if ($hits.Count -eq 0)  { $status = 'Không tìm thấy' }

                    if ($done -le 3) {
                        Write-Log ('   [dò thử] {0} {1} — mẫu "{2}" → {3} file khớp{4}' -f $r.Symbol, $r.Number, $r.Pattern, $hits.Count,
                                   $(if ($file) { ': ' + $file.Name } else { '' }))
                    }
                    $script:Invoices += [pscustomobject]@{
                        Seq        = $seq
                        Sheet      = $p.Sheet
                        Product    = $block.Title
                        Symbol     = $r.Symbol
                        Number     = $r.Number
                        DateText   = $r.DateText
                        Year       = $r.Year
                        Pattern    = $r.Pattern
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
        if ($missing -gt 0) {
            Write-Log ('{0} hóa đơn chưa tìm thấy file (dòng màu đỏ). Vài ví dụ và lý do:' -f $missing) 'WARN'
            foreach ($inv in (@($script:Invoices | Where-Object { -not $_.SourcePath }) | Select-Object -First 5)) {
                $numKey = $inv.Number.TrimStart('0')
                if (-not $numKey) { $numKey = '0' }
                $symNorm = (Remove-Diacritics $inv.Symbol).ToUpper() -replace '[^A-Z0-9]', ''

                if ($index.ByNumber.ContainsKey($numKey)) {
                    $names = @($index.ByNumber[$numKey].ToArray() | Select-Object -First 2 | ForEach-Object { $_.File.Name })
                    Write-Log ('   {0} {1}: trong kho CÓ file chứa số này ({2}) nhưng không khớp ký hiệu/mẫu "{3}".' -f $inv.Symbol, $inv.Number, ($names -join ' | '), $inv.Pattern) 'WARN'
                } else {
                    $sameSymbol = @($index.Files | Where-Object { $symNorm -and $_.Norm.Contains($symNorm) } | Select-Object -First 2 | ForEach-Object { $_.File.Name })
                    if ($sameSymbol.Count -gt 0) {
                        Write-Log ('   {0} {1}: kho có file cùng ký hiệu ({2}) nhưng KHÔNG có file nào chứa số {1}.' -f $inv.Symbol, $inv.Number, ($sameSymbol -join ' | ')) 'WARN'
                    } else {
                        Write-Log ('   {0} {1}: kho KHÔNG có file nào chứa ký hiệu này lẫn số này.' -f $inv.Symbol, $inv.Number) 'WARN'
                    }
                }
            }
            Write-Log '   → So tên file thật với mẫu ở trên để biết cần sửa mẫu hay tìm sai thư mục.' 'WARN'
        }

        $btnOrganize.Enabled = ($ok -gt 0)
        $btnReport.Enabled   = ($script:Invoices.Count -gt 0)
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

    $sources = @($lstSources.Items | ForEach-Object { [string]$_ })
    foreach ($src in $sources) {
        if ((Test-PathInside -Child $dest -Parent $src) -and $chkSubSrc.Checked) {
            $warn = [System.Windows.Forms.MessageBox]::Show(
                ("Thư mục đích nằm bên trong thư mục nguồn:`n{0}`n`nLần đối chiếu sau sẽ quét trúng cả file vừa sắp xếp. Nên chọn thư mục đích nằm ngoài kho hóa đơn.`n`nVẫn tiếp tục?" -f $src),
                'Thư mục đích nằm trong thư mục nguồn',
                [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
            if ($warn -ne [System.Windows.Forms.DialogResult]::Yes) { return }
            break
        }
    }

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

        # sheet nào vào thư mục nào (tự đoán cho sheet chưa được gán tay)
        $sheetNames = @($withFile | ForEach-Object { $_.Sheet } | Sort-Object -Unique)
        $auto = Resolve-SheetFolders -SheetNames $sheetNames -Dest $dest
        foreach ($name in $sheetNames) {
            if (-not $script:SheetFolderMap.ContainsKey($name)) { $script:SheetFolderMap[$name] = $auto[$name] }
            Write-Log ('Sheet "{0}" → thư mục "{1}"' -f $name, $script:SheetFolderMap[$name])

            if ($chkPerProduct.Checked) {
                $sheetDir = Join-Path $dest $script:SheetFolderMap[$name]
                $products = @($withFile | Where-Object { $_.Sheet -eq $name } |
                              ForEach-Object { $_.Product } | Select-Object -Unique)
                $autoProduct = Resolve-FolderMap -Names $products -Existing (Get-SubFolderNames $sheetDir)
                foreach ($product in $products) {
                    $key = '{0}|{1}' -f $name, $product
                    if (-not $script:ProductFolderMap.ContainsKey($key)) { $script:ProductFolderMap[$key] = $autoProduct[$product] }
                    Write-Log ('      {0} → {1}' -f $product, $script:ProductFolderMap[$key])
                }
            }
        }

        # bề rộng số thứ tự: đủ chữ số cho sheet nhiều hóa đơn nhất (001..999, 0001..9999)
        $widths = @{}
        foreach ($g in ($script:Invoices | Group-Object Sheet)) {
            $max = ($g.Group | Measure-Object Seq -Maximum).Maximum
            $widths[$g.Name] = [math]::Max(3, ([string]$max).Length)
        }

        # một file nguồn có thể phục vụ nhiều dòng (cùng số hóa đơn ở hai sản phẩm):
        # chép cho từng dòng, chỉ bỏ bản gốc sau khi dòng cuối cùng dùng xong
        $useCount = @{}
        foreach ($inv in $withFile) {
            if ($useCount.ContainsKey($inv.SourcePath)) { $useCount[$inv.SourcePath]++ }
            else { $useCount[$inv.SourcePath] = 1 }
        }
        $shared = @($useCount.Keys | Where-Object { $useCount[$_] -gt 1 })
        if ($shared.Count -gt 0) {
            Write-Log ('{0} file được dùng cho nhiều dòng (cùng hóa đơn ở nhiều sản phẩm) — sẽ chép cho từng thư mục.' -f $shared.Count)
        }
        $usedSoFar = @{}

        $done = 0; $ok = 0; $skip = 0; $fail = 0
        foreach ($inv in $withFile) {
            $done++
            try {
                $sheetFolder = $script:SheetFolderMap[$inv.Sheet]
                if (-not $sheetFolder) { $sheetFolder = Get-SafeName $inv.Sheet }
                $dir = Join-Path $dest $sheetFolder
                if ($chkPerProduct.Checked) {
                    $productFolder = $script:ProductFolderMap['{0}|{1}' -f $inv.Sheet, $inv.Product]
                    if (-not $productFolder) { $productFolder = Get-SafeName $inv.Product }
                    $dir = Join-Path $dir $productFolder
                }
                if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }

                # giữ nguyên tên file gốc, chỉ thêm số thứ tự vào đầu nếu có tick
                $ext  = [System.IO.Path]::GetExtension($inv.SourcePath)
                $base = [System.IO.Path]::GetFileNameWithoutExtension($inv.SourcePath)
                $w = $widths[$inv.Sheet]
                if (-not $w) { $w = 3 }
                $name = $base
                if ($chkPrefix.Checked) { $name = '{0}_{1}' -f ([string]$inv.Seq).PadLeft($w, '0'), $base }
                $target = Join-Path $dir ((Get-SafeName $name) + $ext)

                if (Test-Path -LiteralPath $target) {
                    $src = Get-Item -LiteralPath $inv.SourcePath
                    $dst = Get-Item -LiteralPath $target
                    if ($src.Length -eq $dst.Length) {
                        $inv.DestPath = $target
                        $inv.Status   = 'Đã có sẵn'
                        if ($usedSoFar.ContainsKey($inv.SourcePath)) { $usedSoFar[$inv.SourcePath]++ }
                        else { $usedSoFar[$inv.SourcePath] = 1 }
                        if ($move -and $usedSoFar[$inv.SourcePath] -ge $useCount[$inv.SourcePath]) {
                            Remove-Item -LiteralPath $inv.SourcePath -Force
                        }
                        $skip++
                        continue
                    }
                    for ($k = 2; $k -le 20; $k++) {
                        $try = Join-Path $dir ((Get-SafeName $name) + ('_{0}' -f $k) + $ext)
                        if (-not (Test-Path -LiteralPath $try)) { $target = $try; break }
                    }
                }

                # luôn chép trước, so lại dung lượng; chế độ di chuyển thì bỏ bản gốc
                # sau khi dòng cuối cùng dùng file đó đã chép xong — an toàn với OneDrive
                $srcLen = (Get-Item -LiteralPath $inv.SourcePath).Length
                Copy-Item -LiteralPath $inv.SourcePath -Destination $target -Force
                if ((Get-Item -LiteralPath $target).Length -ne $srcLen) {
                    throw 'Chép chưa đủ dung lượng nên giữ nguyên file gốc.'
                }

                if ($usedSoFar.ContainsKey($inv.SourcePath)) { $usedSoFar[$inv.SourcePath]++ }
                else { $usedSoFar[$inv.SourcePath] = 1 }

                if ($move -and $usedSoFar[$inv.SourcePath] -ge $useCount[$inv.SourcePath]) {
                    Remove-Item -LiteralPath $inv.SourcePath -Force
                }

                $inv.DestPath = $target
                $inv.Status = 'Đã chép'
                if ($useCount[$inv.SourcePath] -gt 1) { $inv.Status = 'Đã chép (dùng chung file)' }
                elseif ($move)                        { $inv.Status = 'Đã chuyển' }
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

        $orderPath = Join-Path $dest 'thu-tu-in.txt'
        Export-PrintOrder -Path $orderPath
        Write-Log ('Danh sách thứ tự in: {0}' -f $orderPath)
        Write-Log 'Mở "in-hoa-don.ps1", bấm "Nạp danh sách thứ tự in..." rồi chọn file này để in đúng thứ tự trong Excel.'
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
            MauTenFile = $inv.Pattern
            TrangThai  = $inv.Status
            FileNguon  = $inv.SourcePath
            FileDich   = $inv.DestPath
        }
    }
    $rows | Export-Csv -LiteralPath $Path -NoTypeInformation -Encoding UTF8
}

function Export-PrintOrder {
    <# Ghi danh sách đường dẫn file theo đúng thứ tự trong Excel để công cụ in đọc lại. #>
    param([string]$Path)

    $lines = New-Object 'System.Collections.Generic.List[string]'
    $lines.Add('# Danh sách thứ tự in — tạo bởi sap-xep-hoa-don.ps1 lúc ' + (Get-Date -Format 'dd/MM/yyyy HH:mm'))
    foreach ($inv in $script:Invoices) {
        if ($inv.DestPath -and (Test-Path -LiteralPath $inv.DestPath)) { $lines.Add($inv.DestPath) }
    }
    [System.IO.File]::WriteAllLines($Path, $lines, (New-Object System.Text.UTF8Encoding($true)))
}

# ============================================================================
#  SỰ KIỆN
# ============================================================================
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

function Add-SourceFolder {
    <# Thêm một thư mục nguồn vào danh sách (nhận cả đường dẫn gõ tay hoặc dán vào). #>
    param([string]$Path)

    $folder = $Path.Trim().Trim('"')
    if (-not $folder) { return }
    if (-not (Test-Path -LiteralPath $folder -PathType Container)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            ("Không thấy thư mục:`n{0}" -f $folder), 'Đường dẫn không đúng',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    $folder = (Resolve-Path -LiteralPath $folder).Path
    if ($lstSources.Items -contains $folder) {
        Write-Log ('Thư mục này đã có trong danh sách: {0}' -f $folder) 'WARN'
        return
    }
    [void]$lstSources.Items.Add($folder)
    $script:FileIndexKey = ''      # danh sách nguồn đổi thì phải quét lại
    Write-Log ('Thêm thư mục nguồn: {0}' -f $folder)
}

$btnAddSrc.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Chọn thư mục chứa file hóa đơn'
    if ($txtSourcePath.Text -and (Test-Path -LiteralPath $txtSourcePath.Text.Trim('"'))) {
        $dlg.SelectedPath = $txtSourcePath.Text.Trim('"')
    }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { Add-SourceFolder $dlg.SelectedPath }
})

$btnAddPath.Add_Click({
    Add-SourceFolder $txtSourcePath.Text
    $txtSourcePath.Clear()
})

$txtSourcePath.Add_KeyDown({
    param($sender, $e)
    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::Enter) {
        $e.SuppressKeyPress = $true
        Add-SourceFolder $txtSourcePath.Text
        $txtSourcePath.Clear()
    }
})

$btnScanSource.Add_Click({
    if ($lstSources.Items.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('Chưa có thư mục nguồn nào.', 'Thiếu thư mục nguồn',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $btnScanSource.Enabled = $false
    $form.Cursor = [System.Windows.Forms.Cursors]::WaitCursor
    try { [void](Get-FileIndex -Force) }
    finally {
        $btnScanSource.Enabled = $true
        $form.Cursor = [System.Windows.Forms.Cursors]::Default
    }
})

$btnRemoveSrc.Add_Click({
    if ($lstSources.SelectedIndex -ge 0) {
        $lstSources.Items.RemoveAt($lstSources.SelectedIndex)
        $script:FileIndexKey = ''
    }
})

$chkSubSrc.Add_CheckedChanged({ $script:FileIndexKey = '' })
$cboSrcType.Add_SelectedIndexChanged({ $script:FileIndexKey = '' })

$btnBrowseDest.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Chọn thư mục đích để sắp hóa đơn vào'
    $dlg.ShowNewFolderButton = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $txtDest.Text = $dlg.SelectedPath }
})

$txtDest.Add_TextChanged({ $script:SheetFolderMap = @{}; $script:ProductFolderMap = @{} })

$btnMapFolders.Add_Click({
    $names = Get-CheckedSheets
    if ($names.Count -eq 0) {
        [void][System.Windows.Forms.MessageBox]::Show('Chưa tick chọn sheet nào ở mục 1.', 'Chưa chọn sheet',
            [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        return
    }
    $dest = $txtDest.Text.Trim()

    # danh sách sản phẩm của từng sheet, lấy thẳng từ Excel
    $productsBySheet = @{}
    foreach ($name in $names) {
        $productsBySheet[$name] = @((Get-SheetBlocks $name) | ForEach-Object { $_.Title } | Select-Object -Unique)
    }

    $sheetMap = Resolve-SheetFolders -SheetNames $names -Dest $dest
    foreach ($name in $names) {
        if ($script:SheetFolderMap.ContainsKey($name)) { $sheetMap[$name] = $script:SheetFolderMap[$name] }
    }

    $productMap = @{}
    foreach ($name in $names) {
        $sheetDir = if ($dest) { Join-Path $dest $sheetMap[$name] } else { '' }
        $auto = Resolve-FolderMap -Names $productsBySheet[$name] -Existing (Get-SubFolderNames $sheetDir)
        foreach ($product in $productsBySheet[$name]) {
            $key = '{0}|{1}' -f $name, $product
            $productMap[$key] = if ($script:ProductFolderMap.ContainsKey($key)) { $script:ProductFolderMap[$key] } else { $auto[$product] }
        }
    }

    $result = Show-FolderMapDialog -SheetNames $names -ProductsBySheet $productsBySheet -Dest $dest `
                                   -SheetMap $sheetMap -ProductMap $productMap
    if ($result) {
        $script:SheetFolderMap   = $result.Sheets
        $script:ProductFolderMap = $result.Products
        foreach ($name in $names) { Write-Log ('Sheet "{0}" → thư mục "{1}"' -f $name, $result.Sheets[$name]) }
        Write-Log ('Đã gán thư mục cho {0} sản phẩm.' -f $result.Products.Count)
    }
})

$btnMatch.Add_Click({ Invoke-Match })
$btnOrganize.Add_Click({ Invoke-Organize })

$btnReport.Add_Click({
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter   = 'File CSV (*.csv)|*.csv'
    $dlg.FileName = 'bao-cao-doi-chieu.csv'
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    Export-MatchReport -Path $dlg.FileName
    Write-Log ('Đã xuất báo cáo ra "{0}".' -f $dlg.FileName)
})

$btnOpenPrint.Add_Click({
    $root = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $tool = Join-Path $root 'in-hoa-don.ps1'
    if (-not (Test-Path -LiteralPath $tool)) {
        [void][System.Windows.Forms.MessageBox]::Show(
            ("Không thấy file in-hoa-don.ps1 trong thư mục:`n{0}`n`nHãy để hai công cụ trong cùng một thư mục." -f $root),
            'Chưa có công cụ in', [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    Start-Process -FilePath 'powershell.exe' `
                  -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-STA', '-File', ('"{0}"' -f $tool))
    Write-Log 'Đã mở công cụ in.'
})

$lvMatch.Add_DoubleClick({
    $sel = $lvMatch.SelectedItems
    if ($sel.Count -eq 0) { return }
    $inv = $sel[0].Tag
    $path = if ($inv.DestPath) { $inv.DestPath } else { $inv.SourcePath }
    if ($path -and (Test-Path -LiteralPath $path)) { Start-Process -FilePath $path }
})

# ============================================================================
#  KHỞI ĐỘNG
# ============================================================================
$form.Add_Shown({
    $form.Activate()
    $start = if ($PSScriptRoot) { $PSScriptRoot } else { (Get-Location).Path }
    $txtDest.Text = Join-Path $start 'Hoa-don-da-sap-xep'
    Write-Log 'Sẵn sàng.'
    Write-Log 'Chọn file Excel → tick sheet → thêm thư mục nguồn → chọn thư mục đích → "Đối chiếu danh sách" → "Tạo folder && chép file".'
})

Write-Trace 'Dựng xong, mở cửa sổ'
[void]$form.ShowDialog()
Write-Trace 'Người dùng đóng cửa sổ — kết thúc bình thường'
$form.Dispose()
