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
quan-ly-tre-con.html     # ứng dụng quản lý trẻ con cho iPhone (xem bên dưới)
cong-viec.html           # theo dõi công việc của team (xem bên dưới)
lich-bieu.html           # lịch biểu chăm Bé Na hàng ngày (xem bên dưới)
huong-dan-dung-chung.html # hướng dẫn từng bước để cả nhà dùng chung
kiem-tra.html            # trang tự chẩn đoán khi nối không được
dong-bo-google-sheet.gs  # mã Apps Script để cả nhà dùng chung một lịch
nhac-qua-mail.gs         # mã Apps Script tự gửi báo cáo công việc qua mail
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
- **Thống kê tuần**: đếm mỗi người bao nhiêu buổi trong tuần đang xem, xếp từ nhiều
  xuống ít, kèm thanh so sánh và một dòng nói ai làm nhiều nhất. Một buổi hai người thì
  mỗi người tính một. Buổi có việc mà chưa gán ai được đếm riêng. Rê chuột lên một dòng
  hiện chi tiết Sáng/Trưa/Chiều/Tối.
- **Hai biểu đồ** ở cuối trang: *Ai lo buổi nào* xếp chồng theo Sáng/Trưa/Chiều/Tối của tuần
  đang xem, và *Sáu tuần gần đây* dạng cột nhóm cho thấy mỗi người lo bao nhiêu buổi qua từng
  tuần. Rê chuột lên một đoạn hay một cột hiện số cụ thể. Có chú thích màu, và tên với số
  luôn hiện thành chữ nên không phải chỉ dựa vào màu để đọc.
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

Mặc định mỗi máy giữ một bản riêng. Muốn nhiều người cùng sửa, ai cũng thấy — kể cả
những gì người trước đã nhập — thì nối tất cả các máy vào một kho chung.

### Gắn sẵn kho vào trang

Cách gọn nhất cho người dùng: điền địa chỉ kho vào hằng số `KHO_MAC_DINH` ở đầu phần
`<script>` của `lich-bieu.html`.

```js
var KHO_MAC_DINH = "https://ten-cua-ban-default-rtdb.asia-southeast1.firebasedatabase.app";
```

Điền rồi thì **ai mở trang cũng tự vào chung một lịch**, không phải dán mã, không phải
bấm gì — giống một trang nội bộ dùng chung. Máy nào tự tay bấm **Ngưng dùng chung** thì
được nhớ lại và không bị nối lại.

Để trống hằng số này thì trang chạy riêng từng máy như cũ, ai muốn dùng chung thì tự
dán mã.

### Kho dùng được

Có hai loại, chọn một:

**Firebase (khuyến nghị, không phải dán mã).** Tạo project trên
[console.firebase.google.com](https://console.firebase.google.com), tạo Realtime
Database ở chế độ test mode, sửa hai dòng `.read`/`.write` trong thẻ Rules thành `true`,
rồi chép địa chỉ kho dán vào nút **🔗 Dùng chung**. Toàn bấm nút, khoảng 4 phút. Trang
đọc ghi thẳng vào kho qua REST nên không cần máy chủ trung gian; ghi nhiều chỗ một lần
bằng `PATCH` với đường dẫn làm khoá, và `rev` lấy dấu thời gian của máy chủ.

**Google Sheet.** Chỉ nên chọn nếu bạn muốn dữ liệu nằm trong một bảng tính mở ra xem
được. Phải dán một đoạn mã vào Apps Script — làm một lần, mất chừng 5 phút.

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

---

# 🗂 Team Task Tracker — làm gì, xong chưa, tiếp theo làm gì

Mở `cong-viec.html` bằng trình duyệt. Cũng là **một file HTML duy nhất**, không cần build,
không phụ thuộc thư viện nào.

**Giao diện tiếng Anh** cho hợp môi trường công ty; hướng dẫn dưới đây vẫn tiếng Việt.
Tên ở đầu trang bấm vào sửa được ngay, đổi thành tên phòng ban của mình cũng được, và tên
đó đi theo kho chung nên cả team thấy như nhau. Ngày viết dạng `11 Aug` chứ không phải
`11/8` — kiểu số dễ bị người đọc tiếng Anh hiểu thành 8 November.

Trang trả lời đúng ba câu hỏi hay phải hỏi nhau trong team:

| Câu hỏi | Chỗ trả lời |
| --- | --- |
| Đang làm gì? | Bốn cột **To do · In progress · Blocked · Done** |
| Xong chưa? | Hàng đếm ở đầu trang, và nhãn hạn tự đổi màu khi tới hạn hay trễ |
| Tiếp theo làm gì? | Cột **Next steps** bên phải, gom mọi bước kế tiếp và xếp theo hạn |

Màn hình rộng thì bảng việc nằm bên trái, **Next steps** bám bên phải nên cuộn tới đâu vẫn
nhìn thấy. Màn hình hẹp thì xếp dọc, bảng việc lên trước.

## Tính năng

- **Bảng bốn cột**: kéo thẻ từ cột này sang cột kia, hoặc bấm `←` `→` trên thẻ
  (dùng được trên điện thoại, không cần kéo).
- **Mỗi việc có một bước tiếp theo riêng** kèm hạn riêng. Việc nào chưa ghi bước
  tiếp theo thì thẻ hiện cảnh báo và có ô đếm riêng ở đầu trang — không để việc
  nào nằm im mà không ai biết kế tiếp phải làm gì.
- **Khung Kế hoạch tiếp theo** xếp mọi bước kế tiếp theo hạn, gần nhất lên đầu.
- **Ô đếm bấm được**: Đang làm, Đang vướng, Trễ hạn, Chưa có bước tiếp, Xong trong
  7 ngày — bấm vào là lọc luôn theo mục đó.
- **Người làm**: mỗi việc gán được nhiều người, mỗi người một màu riêng. Lọc theo
  người bằng một cú bấm. Sửa danh sách ở nút **👥 Thành viên**; ai bị xoá khỏi
  danh sách nhưng còn được gán ở việc nào thì vẫn giữ nguyên, không mất dữ liệu.
- **Nhật ký cập nhật** cho từng việc: mỗi lần đổi trạng thái đều được ghi lại kèm
  giờ, và tự ghi thêm được dòng của mình.
- **Hai kiểu xem**: bảng bốn cột, hoặc bảng danh sách sửa được ngay tại ô
  (trạng thái, ưu tiên, bước tiếp theo).
- **Báo cáo**: nút **📋 Báo cáo** dựng sẵn bản tóm tắt *Đã xong / Đang làm / Đang vướng /
  Chưa làm / Kế hoạch tiếp theo / Đánh giá thành viên* theo đúng bộ lọc đang xem, chép
  một phát dán thẳng vào Zalo hay email.
- **Bốn biểu đồ** ở cuối trang — xem mục *Báo cáo bằng biểu đồ* bên dưới.
- **Chấm điểm người làm** khi việc xong — xem mục *Đánh giá thành viên* bên dưới.
- **Nhắc việc qua mail** — xem mục bên dưới.
- **Kho chung cho cả team** qua Firebase Realtime Database — xem mục dưới.
- **Gửi bằng đường link**, **Xuất / Nhập JSON**, **In / PDF** (bản in có cả biểu đồ).
- Tự đổi màu theo giao diện sáng/tối của hệ thống, dùng được trên điện thoại.

## Báo cáo bằng biểu đồ

Cuối trang có bốn khối, đều ăn theo bộ lọc đang xem — lọc riêng một người thì biểu đồ
cũng chỉ tính người đó:

- **Tiến độ chung** — một thanh xếp chồng cho thấy tỉ lệ bốn trạng thái.
- **Ai đang gánh việc gì** — mỗi người một thanh, chia theo trạng thái, xếp từ nhiều
  xuống ít. Quá tám người thì phần đuôi gộp thành *Người khác*.
- **Đánh giá thành viên** — bảng xếp hạng, xem mục dưới.
- **Xong theo tuần** — số việc xong trong sáu tuần gần đây.

Rê chuột lên một mảnh hiện số cụ thể. Tên và số luôn hiện thành chữ nên không phải chỉ
dựa vào màu mới đọc được.

Về màu: bốn trạng thái xếp theo thứ tự **Chưa làm · Vướng · Đang làm · Xong**, và thứ
tự này không phải cho đẹp. Màu đỏ của *Vướng* đứng cạnh màu lục của *Xong* thì người
mù màu đọc không ra (ΔE 5.8 — dưới ngưỡng 8), nên *Chưa làm* chen vào giữa để hai màu
đó không kề nhau. Bộ màu đã chạy qua bộ kiểm và đạt cả sáu phép ở nền sáng lẫn nền tối.
Đổi màu thì phải kiểm lại, đừng chọn bằng mắt.

## Đánh giá thành viên

Mở một việc **đã xong**, kéo xuống mục **Đánh giá**, chấm 1–5 sao kèm nhận xét ngắn.
Việc chưa xong không có mục này — chấm việc đang dở thì điểm không nói lên gì.

Điểm cộng cho tất cả những người được gán ở việc đó. Bảng xếp hạng ở cuối trang cho
biết mỗi người:

| Cột | Nghĩa |
| --- | --- |
| Việc xong | Số việc đã xong có tên người đó |
| Đúng hạn | Tỉ lệ xong trước hoặc trong ngày hạn, chỉ tính việc có đặt hạn (không đặt hạn thì hiện `—`) |
| Điểm trung bình | Trung bình số sao của những việc đã chấm |

Số sao hiện luôn trên thẻ việc và trong bảng danh sách. Bỏ chấm được bất cứ lúc nào.
Mọi thay đổi đều ghi vào nhật ký của việc và đồng bộ qua kho chung như các trường khác.

## Nhắc việc qua mail

Bấm **✉️ Nhắc mail**. Có hai đường, dùng cái nào cũng được, hoặc cả hai:

**Cách 1 — soạn thư ngay.** Không cần cài gì, không cần dựng gì. Trang mở sẵn một lá
thư trong hòm thư của máy, đã điền người nhận, tiêu đề và toàn bộ báo cáo; anh chỉ việc
bấm Gửi. Báo cáo dài quá một lá thư thì được cắt bớt kèm ghi chú.

**Cách 2 — tự gửi hằng ngày, không cần ai mở trang.** Cần kho chung và một Apps Script:

1. Vào [script.google.com](https://script.google.com), tạo project mới, xoá hết rồi dán
   toàn bộ [`nhac-qua-mail.gs`](nhac-qua-mail.gs) vào.
2. Sửa hai dòng `KHO` và `PHONG` ở đầu file cho khớp kho của team.
3. Chạy hàm `guiNgayBayGio` một lần, Google hỏi quyền thì cho phép, rồi kiểm hộp thư.
4. Chạy hàm `datLichHangNgay` một lần để nó tự gửi mỗi sáng (đổi giờ ở `GIO_GUI`).
5. Muốn bấm gửi ngay từ trang thì **Deploy → New deployment → Web app**, *Execute as: Me*,
   *Who has access: Anyone*, rồi dán URL vào ô trong hộp ✉️ Nhắc mail.

Thư gửi đi có bản HTML kèm thanh tiến độ, chia sẵn các mục *Trễ hạn · Đến hạn hôm nay ·
Đang vướng · Kế hoạch tiếp theo · Xong trong 7 ngày*. Ngày nào không có gì đáng nhắc thì
im lặng, không gửi (đổi ở `IM_KHI_KHONG_CO_GI`).

Danh sách người nhận lấy thẳng từ kho chung, nên thêm bớt người thì sửa trong trang là
xong, không phải mở lại Apps Script. Địa chỉ Apps Script chỉ lưu trên máy đang dùng,
không nằm trong link **Gửi** cũng không nằm trong file **Xuất JSON**.

## Dùng chung một cơ sở dữ liệu

Mặc định mỗi máy giữ một bản riêng trong trình duyệt. Muốn cả team nhìn chung một
bảng — ai sửa gì người khác cũng thấy — thì nối tất cả các máy vào một
**Firebase Realtime Database**. Trang đọc ghi thẳng vào kho qua REST nên không phải
dựng máy chủ nào, và mất mạng vẫn gõ được, có mạng lại thì tự gửi lên.

Các bước, chừng 4 phút, toàn bấm nút:

1. Vào [console.firebase.google.com](https://console.firebase.google.com), tạo một project.
2. Chọn **Build → Realtime Database → Create Database**, chọn vùng gần (Singapore),
   chọn **Start in test mode**.
3. Chép địa chỉ kho (dạng `https://….asia-southeast1.firebasedatabase.app`), mở trang
   `cong-viec.html`, bấm **🗄 Kho chung**, dán vào rồi bấm **Kiểm tra và nối**.
4. Hộp thoại hiện một **mã kho**. Gửi mã đó cho người trong team; họ mở trang, bấm
   **🗄 Kho chung**, dán mã vào là xong — không cần biết địa chỉ đầy đủ.
5. Test mode tự khoá sau 30 ngày. Vào thẻ **Rules**, sửa hai dòng `.read` và `.write`
   thành `true` rồi Publish là dùng được lâu dài.

### Gắn sẵn kho vào trang

Muốn ai mở trang cũng tự vào chung một kho, không phải dán mã, thì điền địa chỉ vào
hằng số `KHO_MAC_DINH` ở đầu phần `<script>` của `cong-viec.html`:

```js
var KHO_MAC_DINH = "https://ten-cua-ban-default-rtdb.asia-southeast1.firebasedatabase.app";
var PHONG_MAC_DINH = "viec-team";   // đổi tên này để một kho chứa nhiều bảng việc
```

Máy nào tự tay bấm **Ngưng dùng chung** thì được nhớ lại và không bị nối lại.

### Đồng bộ chạy thế nào

- Mỗi việc là một nút riêng trong kho, nên hai người sửa hai việc khác nhau không đè
  lên nhau. Cùng đụng một việc thì bản có mốc `updated` mới hơn thắng.
- Việc bị xoá để lại một **dấu mộ** trong kho — không có nó thì máy khác chưa kịp đồng
  bộ sẽ đẩy việc đã xoá quay về. Dấu mộ tự dọn sau 30 ngày.
- Máy khác kéo về mỗi 8 giây, và kéo ngay khi bạn quay lại tab. Chấm trạng thái cạnh
  **Đã lưu** cho biết đang gửi, đã đồng bộ hay mất kết nối.
- Đang mở hộp sửa một việc thì bản của người khác không giật mất chỗ đang gõ.
- Nối vào kho đã có dữ liệu thì hai bên được trộn vào nhau, không bên nào bị đè.
  Mấy việc mẫu điền sẵn không bao giờ được đẩy lên kho.
- Địa chỉ kho chỉ lưu trên máy đang dùng, không nằm trong file **Xuất JSON** cũng
  không nằm trong link **Gửi**.

**Ai có địa chỉ kho đều đọc và sửa được**, không cần đăng nhập — chỉ gửi cho người
trong team.

## Lưu ý

Khi chưa nối kho chung, dữ liệu nằm trong trình duyệt đang dùng, và **mỗi địa chỉ web
giữ một kho riêng** — bảng nhập ở `nguyenthanhitvn87-source.github.io` không hiện sang
bản mở bằng đường dẫn khác hay trình duyệt khác. Mang sang nơi khác thì bấm **📤 Gửi**
để lấy link rồi **📥 Nhận** ở nơi mới, hoặc nối cả hai vào cùng một kho chung.

Mỗi lần ghi đè lớn (nhận bảng từ link, nhập JSON, nối vào kho) đều cất bản cũ lại;
link **Khôi phục bản trước đó** ở chân trang đổi qua đổi lại giữa hai bản, nên bấm
nhầm vẫn lấy lại được.

---

# 👶 Quản lý bé — ứng dụng cho iPhone

Mở `quan-ly-tre-con.html`. Vẫn là **một file HTML duy nhất**, không cần cài, không cần
build, không tài khoản, không máy chủ. Cài lên Màn hình chính iPhone là chạy như một
app thật: toàn màn hình, không thanh địa chỉ, mất mạng vẫn dùng được.

## Cài lên iPhone

1. Mở trang này bằng **Safari** (Chrome trên iPhone không cài được).
2. Bấm nút **Chia sẻ** ⬆︎ ở thanh dưới.
3. Kéo xuống chọn **Thêm vào MH chính** → **Thêm**.
4. Từ đó mở bằng icon ngoài màn hình.

## Năm tab

**☀️ Hôm nay** — ba con số mở đầu: sao kiếm được hôm nay, việc đã xong trên tổng số,
số phút dùng máy còn lại. Dưới đó là nề nếp trong ngày dạng timeline (mốc đang tới giờ
được tô sáng, mốc đã qua thì mờ đi), việc nhà của đúng thứ hôm nay, mục **Sắp tới** gom
lịch khám — mũi tiêm — hạn bài tập trong 14 ngày, và một dòng nhật ký kèm tâm trạng.

**⭐️ Việc & Sao** — ví sao của bé, biểu đồ bảy ngày gần đây, danh sách việc được giao
(mỗi việc đặt riêng số sao và những thứ nào trong tuần phải làm), quầy đổi thưởng
(đủ sao thì nút **Đổi** sáng lên) và sổ sao ghi từng lần cộng trừ. Bấm ✓ vào một việc
là tự cộng sao; bỏ tick thì sao cũng được thu lại, không bị cộng dồn.

**🩺 Sức khoẻ** — chiều cao, cân nặng, BMI kèm mức chênh so với lần đo trước, biểu đồ
hai đường theo thời gian; danh sách tiêm chủng đánh dấu đã tiêm; thuốc và dị ứng;
lịch khám.

**📚 Học tập** — hạn mức giờ dùng máy mỗi ngày với thanh tiến độ đỏ lên khi vượt,
nút ＋15/＋30 phút và **đồng hồ bấm giờ** cho lúc bé bắt đầu xem; biểu đồ bảy ngày
dùng máy; bài tập kèm hạn nộp (quá hạn hiện nhãn đỏ); thời khoá biểu từng thứ.

**👶 Hồ sơ** — sửa tên, ảnh đại diện, màu, ngày sinh (tuổi tự tính ra năm và tháng);
sửa các mốc nề nếp; bật nhắc giờ; và khu dữ liệu.

## Nhiều bé

Dải chip ngay dưới tiêu đề chuyển qua lại giữa các bé, mỗi bé một màu và một bộ dữ liệu
riêng hoàn toàn. Thêm bé mới thì được hỏi có **chép nề nếp, việc nhà và phần thưởng**
của bé đang xem sang không, khỏi phải nhập lại từ đầu.

## Dữ liệu để ở đâu

Tất cả nằm trong `localStorage` của máy đang dùng — không gửi lên mạng, không có tài
khoản. Ba cách mang đi:

- **📤 Gửi qua link** — gói cả hồ sơ vào phần sau dấu `#` của địa chỉ rồi gửi qua Zalo.
  Phần sau dấu `#` không bao giờ được gửi lên máy chủ nào. Người nhận được hỏi có lưu
  vào máy không, dữ liệu cũ của họ chỉ bị thay khi họ đồng ý và vẫn được cất lại.
- **⬇︎ Xuất file** JSON để sao lưu.
- **⬆︎ Nhập file** ở máy mới.

Xoá Safari khỏi máy hay xoá dữ liệu web thì mất, nên thỉnh thoảng nên bấm **Xuất file**.

## Nhắc giờ

Bật ở tab Hồ sơ. Mỗi khi tới một mốc nề nếp, máy báo kèm tên bé và người lo việc đó.
Chỉ nhắc trong vòng 5 phút kể từ mốc nên mở app muộn không bị dội thông báo cũ, và mỗi
mốc chỉ nhắc một lần trong ngày. Trên iPhone phải mở sẵn app (đã cài lên Màn hình chính)
thì mới nhắc được.

## Lưu ý

BMI hiện ở tab Sức khoẻ chỉ để tham khảo. Ở trẻ con phải so với biểu đồ tăng trưởng
theo tuổi của bác sĩ, đừng tự kết luận từ một con số.
