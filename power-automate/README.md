# Chạy Power Automate Desktop tự động mỗi ngày

Bộ script này tạo một Windows Scheduled Task gọi flow Power Automate Desktop (PAD)
vào đúng giờ mỗi ngày. Không cần license Premium.

## File trong thư mục

| File | Việc |
|---|---|
| `tao-job-pad-hangngay.ps1` | Chạy **một lần** để tạo Scheduled Task |
| `chay-flow-pad.ps1` | Task gọi file này mỗi ngày — nó tự dò đường dẫn PAD rồi kích hoạt flow |

## Cách dùng

1. Copy cả **hai file** vào cùng một thư mục trên máy Windows, ví dụ `C:\PAD\`.
2. Mở `tao-job-pad-hangngay.ps1`, sửa phần `CAU HINH`:
   ```powershell
   $GioChay  = "15:55"       # giờ chạy, định dạng HH:mm 24h
   $UriFlow  = "ms-powerautomate:/console/flow/run?environmentid=...&workflowid=...&source=Other"
   $QuyenCao = $false        # chỉ bật $true nếu flow phải điều khiển app chạy quyền admin
   ```

   **Lấy `$UriFlow` ở đâu:** trong Power Automate Desktop, chuột phải flow →
   *Create desktop shortcut*. Chuột phải shortcut vừa tạo → *Properties* → copy ô **Target**.

   URI này định danh flow bằng `workflowid` (GUID) và kèm sẵn `environmentid`, nên
   đổi tên flow về sau cũng không làm hỏng lịch chạy, và không lo gõ sai tên.
   Nếu không lấy được URI, để `$UriFlow = ""` rồi điền `$TenFlow` — script sẽ tự ghép
   URI theo tên (kém chắc chắn hơn).
3. Mở **PowerShell với quyền Administrator** (chuột phải → *Run as Administrator*), rồi:
   ```powershell
   cd C:\PAD
   powershell -ExecutionPolicy Bypass -File .\tao-job-pad-hangngay.ps1
   ```
4. Chạy thử ngay, không cần đợi tới giờ:
   ```powershell
   Start-ScheduledTask -TaskName "ChayPowerAutomateFlowHangNgay"
   ```
5. Xem log nếu có gì đó không chạy:
   ```powershell
   notepad "$env:LOCALAPPDATA\PAD-Scheduler\chay-flow.log"
   ```

Gỡ task ra:
```powershell
Unregister-ScheduledTask -TaskName "ChayPowerAutomateFlowHangNgay" -Confirm:$false
```

## Ba lỗi đã sửa so với bản script đầu

### 1. `-flow "ten flow"` không phải là tham số hợp lệ

`PAD.Console.Host.exe` không có switch `-flow`. Nó nhận URI:

```
ms-powerautomate:/console/flow/run?workflowName=<ten-flow-da-url-encode>
```

Tên `run job` có dấu cách nên phải encode thành `run%20job`, nếu không PAD chỉ
nhận được `run`. Launcher xử lý việc này bằng `[uri]::EscapeDataString()` — nhưng
tốt nhất là dùng thẳng `$UriFlow` lấy từ PAD (dạng `workflowid`), khỏi cần ghép tay.

Muốn truyền tham số vào flow:
```powershell
.\chay-flow-pad.ps1 -UriFlow "ms-powerautomate:/console/flow/run?..." `
                    -InputArguments '{"NgayChay":"2026-08-11"}'
```

### 2. `LogonType S4U` làm flow không chạy được

`S4U` chính là tuỳ chọn *"Run whether user is logged on or not"* trong Task Scheduler.
Nó chạy task ở một session **không có desktop tương tác**, nên mọi bước UI automation
của PAD (click chuột, gõ phím, mở ứng dụng, đọc nội dung cửa sổ) sẽ treo hoặc lỗi.

Bản attended (miễn phí) bắt buộc dùng `Interactive`. Script đã đổi sang:

```powershell
New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
                           -LogonType Interactive -RunLevel Highest
```

**Hệ quả cần biết:** tới giờ chạy, máy phải đang bật và user phải đang đăng nhập.
Nếu **màn hình bị khoá**, các bước UI automation vẫn sẽ lỗi — Windows không cho
process thường thao tác lên màn hình khoá. Cách xử lý:

- Để máy không tự khoá: Settings → Accounts → Sign-in options → *Require sign-in*: **Never**;
  và Screen saver → bỏ tick *On resume, display logon screen*.
- Hoặc bật auto-login (`netplwiz`) để máy tự đăng nhập lại sau khi khởi động.
- Hoặc nâng lên **Unattended RPA** (bản trả phí) — chỉ bản này mới chạy được khi khoá máy.

Nếu flow của bạn **không** đụng giao diện (chỉ gọi API, đọc/ghi file, chạy SQL,
xử lý Excel qua COM) thì màn hình khoá không ảnh hưởng.

### 3. Đường dẫn cứng tới `WindowsApps` sẽ chết khi PAD cập nhật

```
C:\Program Files\WindowsApps\Microsoft.PowerAutomateDesktop_11.2607.187.0_x64__8wekyb3d8bbwe\
```

`11.2607.187.0` là số version. PAD bản Store tự cập nhật, version đổi → thư mục đổi
→ task im lặng không chạy nữa và không báo gì cả.

`chay-flow-pad.ps1` xử lý theo thứ tự:

1. **Kích hoạt qua protocol handler** — `Start-Process "ms-powerautomate:/console/flow/run?..."`.
   Đây là cách chính. Với bản MSIX, Windows sẽ kích hoạt đúng gói ứng dụng; chạy thẳng
   file `.exe` nằm trong `WindowsApps` hay bị chặn bởi ACL của thư mục đó.
2. Nếu protocol lỗi, dò lại đường dẫn `.exe` **mỗi lần chạy**:
   `Get-AppxPackage Microsoft.PowerAutomateDesktop*` → `C:\Program Files (x86)\Power Automate Desktop\`
   → quét `WindowsApps` lấy version mới nhất.

Thêm nữa: `Test-Path` vào `WindowsApps` có thể trả `$false` dù file có thật, vì thư mục
đó bị hạn chế ACL — nên script đầu có khả năng báo "KHONG TIM THAY" sai ngay từ đầu.

## Các chỉnh sửa nhỏ khác

- Thêm `#Requires -RunAsAdministrator` — báo lỗi rõ ràng thay vì fail giữa chừng.
- Đổi mặc định `RunLevel` từ `Highest` sang `Limited`. Ứng dụng MSIX (bản Store) thường
  không khởi chạy được khi bị elevate; chạy PAD như lúc bạn tự mở bằng tay là hợp lý nhất.
  Đặt `$QuyenCao = $true` nếu flow cần điều khiển app chạy quyền admin.
- Đổi tên biến `$Args` → `$ThamSo`. `$args` là **biến tự động** của PowerShell, gán đè lên
  nó là nguồn lỗi khó lần.
- `MultipleInstances IgnoreNew` — flow hôm nay chạy chưa xong thì lần sau bỏ qua,
  không chồng hai instance lên nhau.
- `ExecutionTimeLimit 2 giờ` — mặc định của Windows là 3 ngày, flow treo sẽ treo mãi.
- `RestartCount 2` — lỗi thì thử lại 2 lần, cách nhau 5 phút.
- Giờ được parse bằng `InvariantCulture` — không phụ thuộc định dạng ngày giờ của Windows.
- Ghi log ra `%LOCALAPPDATA%\PAD-Scheduler\chay-flow.log` để biết task có chạy không.

## Kiểm tra khi flow không chạy

```powershell
# Task có tồn tại và đang bật không, lần chạy gần nhất kết quả ra sao
Get-ScheduledTaskInfo -TaskName "ChayPowerAutomateFlowHangNgay"
```

`LastTaskResult`:

| Mã | Nghĩa |
|---|---|
| `0` | Task chạy OK (launcher đã gọi PAD — chưa chắc flow bên trong PAD thành công) |
| `1` | Launcher lỗi — mở file log ra xem |
| `267011` | Task chưa từng chạy lần nào |
| `0x41303` | Chưa tới giờ / chưa chạy lần nào |

Nếu log ghi "Da gui lenh chay flow ... thanh cong" mà flow vẫn không làm gì, thì vấn đề
nằm trong PAD chứ không phải ở task. Kiểm tra:

- Đã **đăng nhập tài khoản Microsoft trong PAD** chưa — chưa đăng nhập thì không mở được flow.
- Tên flow gõ có **đúng tuyệt đối** không (kể cả hoa/thường và dấu cách).
- Mở PAD → tab **Run history** (Lịch sử chạy) xem flow lỗi ở bước nào.
