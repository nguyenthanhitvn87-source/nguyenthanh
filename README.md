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
index.html      # toàn bộ game: giao diện, style và logic
lich-bieu.html  # lịch biểu chăm Bé Na hàng ngày (xem bên dưới)
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
- Bấm thẳng vào ô để sửa, nội dung **tự lưu** vào `localStorage` (không cần nút Lưu).
- Chuyển **Tuần trước / Tuần sau**, nút **Hôm nay** quay về tuần hiện tại; dòng của ngày hôm nay được tô sáng.
- **Người làm cho từng buổi**: mỗi ô Sáng/Trưa/Chiều/Tối có ô chọn người phụ trách.
  Mỗi người một màu riêng, lấy theo thứ tự trong danh sách nên các màu cách xa nhau.
  Danh sách người sửa ở khung **Người làm**; ai đã bị xoá khỏi danh sách nhưng còn được
  gán ở đâu đó thì vẫn giữ nguyên, không mất dữ liệu.
- **Thời gian biểu trong ngày của Bé Na** hiển thị dạng timeline, sửa được ở ô văn bản bên cạnh
  (mỗi dòng một mốc, dạng `giờ: việc cần làm | người làm`, chấp nhận cả `7h30 - 8h00:`).
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

## Lưu ý

Dữ liệu nằm trong trình duyệt đang dùng — xoá dữ liệu duyệt web hoặc đổi máy sẽ mất.
Nhớ bấm **Xuất JSON** khi cần sao lưu.

**Chưa dùng chung được giữa nhiều người.** Mỗi thiết bị giữ bản riêng, sửa ở máy này
máy kia không thấy. Muốn nhiều người cùng nhìn một lịch thì cần thêm nơi lưu chung
(máy chủ hoặc dịch vụ lưu trữ), chưa có trong bản này.

Nhắc giờ chỉ chạy khi trang đang mở trong trình duyệt. Trên iPhone, Safari chỉ hiện
thông báo khi trang đã được **Thêm vào màn hình chính**.
