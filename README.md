# 🐍 Rắn Săn Mồi

**Tác giả: Nguyễn Thanh**

Một game web đơn giản viết bằng HTML + CSS + JavaScript thuần, gói gọn trong **một file duy nhất** (`index.html`). Không cần cài đặt, không cần build, không phụ thuộc thư viện nào.

## Cách chơi

Mở file `index.html` bằng trình duyệt bất kỳ (nhấp đúp vào file là được).

Hoặc chạy qua máy chủ cục bộ:

```bash
npx http-server .
# rồi mở http://localhost:8080
```

## Điều khiển

| Thao tác | Phím / cử chỉ |
| --- | --- |
| Di chuyển | `↑` `↓` `←` `→` hoặc `W` `A` `S` `D` |
| Bắt đầu / Tạm dừng | `Space`, nút **Tạm dừng**, hoặc chạm vào bàn chơi |
| Chơi lại | `R` hoặc nút **Chơi lại** |
| Trên điện thoại | Vuốt trên bàn chơi, hoặc bấm các nút mũi tên |

## Luật chơi

- Ăn mồi màu hồng để dài ra và ghi điểm.
- Mỗi mồi được `10 × cấp độ` điểm.
- Cứ ăn 5 mồi thì lên một cấp, rắn bò nhanh hơn một chút (nhanh nhất là 70ms/bước).
- Thua khi đâm vào tường hoặc cắn phải thân mình.
- Kỷ lục được lưu trong `localStorage` của trình duyệt.

## Cấu trúc

```
index.html               # toàn bộ game: giao diện, style và logic
lich-bieu.html           # lịch biểu chăm Bé Na hàng ngày (xem bên dưới)
huong-dan-dung-chung.html # hướng dẫn từng bước để cả nhà dùng chung
kiem-tra.html            # trang tự chẩn đoán khi nối không được
dong-bo-google-sheet.gs  # mã Apps Script để cả nhà dùng chung một lịch
README.md
```

Các hằng số dễ chỉnh nằm ở đầu phần `<script>`:

```js
const CELLS = 20;          // kích thước lưới
const BASE_SPEED = 160;    // ms mỗi bước ở cấp 1
const MIN_SPEED = 70;      // tốc độ nhanh nhất
const SPEED_STEP = 9;      // mỗi cấp nhanh thêm bao nhiêu ms
const FOOD_PER_LEVEL = 5;  // số mồi cần ăn để lên cấp
```

---

# 📅 Chăm Bé Na — lịch biểu hàng ngày

Mở `lich-bieu.html` bằng trình duyệt. Cũng là **một file HTML duy nhất**, không cần build.

## Tính năng

- Bảng theo tuần: **Ngày · Sáng · Trưa · Chiều · Tối · Notes · Special Note**.
- Bấm thẳng vào ô để sửa, nội dung **tự lưu** ngay (không cần nút Lưu).
- **Gửi lịch bằng đường link**: bấm 📤 Gửi lịch để gói cả lịch vào một đường link rồi
  gửi qua Zalo hay tin nhắn. Người nhận bấm vào là thấy, không cài gì, không đăng nhập.
  Không cần máy chủ hay dịch vụ nào — dữ liệu nằm sau dấu `#` nên không đi lên mạng.
- **Dùng chung cho cả nhà** qua một Google Sheet: nhiều người cùng sửa, ai cũng nhìn thấy.
  Xem mục hướng dẫn bên dưới.
- Chuyển **Tuần trước / Tuần sau**, nút **Hôm nay** quay về tuần hiện tại; dòng của ngày hôm nay được tô sáng.
- **Người làm cho từng buổi**: mỗi ô Sáng/Trưa/Chiều/Tối gán được **nhiều người**,
  bấm `×` trên tên để bỏ. Cả ô được nhuộm theo màu người làm cho dễ nhìn; hai người
  thì ô pha màu của cả hai và vệt bên trái chia đôi.
  Mỗi người một màu riêng, lấy theo thứ tự trong danh sách nên các màu cách xa nhau.
  Danh sách người sửa ở khung **Người làm**; ai đã bị xoá khỏi danh sách nhưng còn được
  gán ở đâu đó thì vẫn giữ nguyên, không mất dữ liệu.
- **Thời gian biểu trong ngày của Bé Na** hiển thị dạng timeline, sửa được ở ô văn bản bên cạnh
  (mỗi dòng một mốc, dạng `giờ: việc cần làm | người làm`, nhiều người thì ngăn bằng dấu phẩy:
  `7:30: ăn sáng | Mẹ, Bà`; chấp nhận cả `7h30 - 8h00:`).
  Mốc đang tới giờ được tô sáng, các mốc đã qua thì mờ đi.
- **Nhắc giờ**: bật nút 🔔 để trình duyệt báo mỗi khi tới một mốc trong thời gian biểu,
  kèm tên người làm. Chỉ nhắc trong vòng 5 phút kể từ mốc nên mở trang muộn không bị dội
  thông báo cũ, và mỗi mốc chỉ nhắc một lần trong ngày.
- **Xuất / Nhập JSON** để sao lưu hoặc chuyển sang máy khác, và **In / PDF**.
- Tự đổi màu theo giao diện sáng/tối của hệ thống, dùng được trên điện thoại.

## Phím tắt

| Thao tác | Phím |
| --- | --- |
| Xuống dòng trong ô | `Enter` |
| Nhảy sang ô kế tiếp | `Ctrl` + `Enter` (macOS: `⌘` + `Enter`) |

## Gửi lịch bằng đường link

Cách nhanh nhất, không phải cài gì. Bấm **📤 Gửi lịch** → **Sao chép link** → dán vào
Zalo gửi đi. Cả lịch được nén rồi nhét vào phần sau dấu `#` của địa chỉ; phần đó không
bao giờ được gửi lên máy chủ nào, nên cách này chạy mà không cần dịch vụ lưu trữ.
Một tuần đầy đủ ra link chừng 600 ký tự.

Người nhận mở link sẽ thấy băng báo *đây là lịch người khác gửi*, xem thử được ngay,
và tự chọn **Lưu vào máy này** hay **Bỏ, giữ lịch cũ của tôi** — lịch sẵn có của họ
không bị ghi đè khi chưa đồng ý. Trong lúc xem thử, mọi đường ghi xuống máy đều bị
chặn, kể cả lần lưu tự động khi rời trang.

Mỗi lần ghi đè lớn (nhận lịch từ link, nhập JSON, nối vào sheet đã có dữ liệu) đều cất
bản cũ lại. Link **Khôi phục lịch trước đó** ở cuối trang đổi qua đổi lại giữa hai bản,
nên bấm nhầm vẫn lấy lại được.

Điểm yếu: link chụp lại lịch ở thời điểm gửi, sửa xong muốn người kia thấy thì gửi
link mới. Muốn hai bên tự thấy của nhau thì dùng cách dưới đây.

## Dùng chung cho cả nhà

Mặc định mỗi máy giữ một bản riêng. Muốn nhiều người cùng sửa và ai cũng nhìn thấy
thì nối trang với một Google Sheet của bạn — làm một lần, mất chừng 5 phút.

Mở [`huong-dan-dung-chung.html`](huong-dan-dung-chung.html) để vừa xem vừa làm theo,
có cả phần xử lý mấy chỗ hay vướng. Nối không được thì mở
[`kiem-tra.html`](kiem-tra.html): trang này chạy bốn phép thử đọc và ghi theo cả hai
cách gọi, rồi chỉ ra đúng chỗ phải sửa. Tóm tắt các bước:

1. Vào [sheets.new](https://sheets.new) tạo một bảng tính mới, đặt tên tuỳ ý.
2. Trong bảng tính, chọn **Tiện ích mở rộng → Apps Script**.
3. Xoá hết nội dung có sẵn, dán toàn bộ file [`dong-bo-google-sheet.gs`](dong-bo-google-sheet.gs) vào, bấm lưu.
4. Bấm **Triển khai → Tạo bản triển khai mới**, chọn loại **Ứng dụng web**, rồi đặt:
   - *Thực thi với tư cách*: **Tôi**
   - *Người có quyền truy cập*: **Bất kỳ ai**
5. Bấm **Triển khai**, cho phép quyền khi Google hỏi, rồi **sao chép URL ứng dụng web**
   (dạng `https://script.google.com/macros/s/..../exec`).
6. Mở trang lịch biểu, bấm **🔗 Dùng chung**, dán URL đó vào, bấm **Kiểm tra và nối**.
7. Hộp thoại hiện một **mã lịch chung**. Gửi mã đó cho người khác; họ mở trang lịch,
   bấm **🔗 Dùng chung**, dán mã vào là xong — không cần biết địa chỉ đầy đủ.

Xong. Ai sửa gì thì khoảng 10 giây sau máy khác thấy, và một chấm trạng thái hiện
bên cạnh **Đã lưu** cho biết đang đồng bộ hay mất kết nối. Dữ liệu nằm thẳng trong
Google Sheet dưới dạng bảng, nên mở sheet ra xem hoặc sửa tay cũng được.

Vài điều cần biết:

- **Ai có URL đều sửa được**, không cần đăng nhập. Chỉ gửi URL cho người trong nhà.
- Trình duyệt hay chặn cuộc gọi thẳng sang Google vì khác tên miền. Trang tự nhận ra và
  chuyển sang gọi vòng qua thẻ `script`, nên vẫn chạy; cách đang dùng được nhớ lại cho
  những lần sau. Đường vòng chỉ gửi được bằng GET nên mỗi lần đẩy một mẩu nhỏ.
- Nối vào một sheet đã có dữ liệu thì bản trên sheet được giữ làm gốc, máy mới không đè
  lên. Ngày nào máy đó có mà sheet chưa có thì được thêm vào, miễn là người dùng thực sự
  đã sửa chứ không phải mấy dòng điền sẵn.
- Đang gõ dở ở ô nào thì bản của người khác không giật mất ô đó; con trỏ vẫn nằm nguyên chỗ cũ.
- Mất mạng thì vẫn gõ được, dữ liệu nằm trong máy và tự gửi lên khi có mạng lại.
- URL chỉ lưu trên máy đang dùng, không nằm trong file **Xuất JSON**.
- Bấm **🔗 Dùng chung** rồi để trống, bấm OK là ngưng dùng chung, quay về lưu riêng.

## Lưu ý

**Mỗi địa chỉ web giữ một kho riêng.** Lịch nhập ở `nguyenthanhitvn87-source.github.io`
không hiện sang bản mở bằng đường dẫn khác, bản tải về máy, hay trình duyệt khác — đó
là quy tắc của trình duyệt, không phải mất dữ liệu. Chân trang ghi rõ đang lưu cho địa
chỉ nào, và khi một địa chỉ chưa có lịch nào thì trang nhắc luôn kèm nút dán link.

Mang lịch sang nơi khác: ở nơi đang có lịch bấm **📤 Gửi lịch** để lấy link, ở nơi mới
bấm **📥 Nhận lịch** rồi dán link vào. Cách này chuyển được giữa hai địa chỉ, hai trình
duyệt hay hai máy bất kỳ.

Khi chưa nối Google Sheet, dữ liệu nằm trong trình duyệt đang dùng — xoá dữ liệu
duyệt web hoặc đổi máy sẽ mất. Nhớ bấm **Xuất JSON** khi cần sao lưu.

Nhắc giờ chỉ chạy khi trang đang mở trong trình duyệt. Trên iPhone, Safari chỉ hiện
thông báo khi trang đã được **Thêm vào màn hình chính**.
