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
cong-viec.html           # theo dõi công việc của team (xem bên dưới)
lich-bieu.html           # lịch biểu chăm Bé Na hàng ngày (xem bên dưới)
quan-ly-tre-em.html      # Bé Ngoan — quản lý trẻ em trên iPhone (xem bên dưới)
huong-dan-dung-chung.html # hướng dẫn từng bước để cả nhà dùng chung
kiem-tra.html            # trang tự chẩn đoán khi nối không được
dong-bo-google-sheet.gs  # mã Apps Script để cả nhà dùng chung một lịch
nhac-qua-mail.gs         # mã Apps Script tự gửi báo cáo công việc qua mail
sap-xep-hoa-don.ps1      # sắp xếp hóa đơn vào thư mục theo danh sách Excel (xem bên dưới)
sap-xep-hoa-don.bat      # bấm đúp để chạy công cụ sắp xếp
in-hoa-don.ps1           # in hóa đơn hàng loạt ra máy in (xem bên dưới)
in-hoa-don.bat           # bấm đúp để chạy công cụ in
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

# 🧸 Bé Ngoan — quản lý trẻ em trên iPhone

Mở `quan-ly-tre-em.html` bằng Safari trên iPhone. Vẫn là **một file HTML duy nhất**:
không cài đặt, không đăng nhập, không gọi ra mạng lần nào. Mọi thứ nằm trong máy.

## Đưa lên iPhone như một ứng dụng

1. Mở trang bằng **Safari** (đường dẫn kho, hoặc mở file đã tải về từ app Tệp).
2. Bấm nút **Chia sẻ** ở thanh dưới → **Thêm vào MH chính** → **Thêm**.
3. Từ đó mở bằng biểu tượng con gấu ngoài màn hình: chạy toàn màn hình, không thanh
   địa chỉ, có vùng an toàn cho tai thỏ và thanh vuốt dưới đáy — nhìn như app thật.

Không cần App Store, không cần máy Mac hay tài khoản nhà phát triển Apple. Trang cũng
chạy đúng như vậy trên Android, máy tính bảng và máy tính bàn.

## Năm khu vực

| Thẻ | Làm gì |
| --- | --- |
| **Hôm nay** | Việc trong ngày, vòng tiến độ, nhật ký, biểu đồ bảy ngày |
| **Việc** | Danh sách việc của bé: giờ, số sao, lặp lại thế nào |
| **Thưởng** | Sao đang có, kho quà đổi sao, lịch sử đã đổi |
| **Sức khoẻ** | Chiều cao cân nặng theo biểu đồ, tiêm chủng, thuốc, khám |
| **Khác** | Hồ sơ các bé, nhắc giờ, giao diện, xuất/nhập, gửi link |

## Tính năng

- **Nhiều bé**: mỗi bé một tên, một hình, một màu riêng và ngày sinh. Dải chọn bé nằm
  ngay dưới tiêu đề, đổi bé là mọi thẻ đổi theo. Có ngày sinh thì ứng dụng tự tính tuổi
  (dưới hai tuổi đếm theo tháng) và nhắc trước sinh nhật 14 ngày.
- **Việc hằng ngày**: đặt giờ, số sao, và cách lặp — *hằng ngày*, *chọn thứ trong tuần*,
  hay *một lần vào một ngày*. Chạm cả dòng là đánh dấu xong, việc quá giờ mà chưa làm
  thì giờ hiện màu đỏ. Có mười mẫu bấm một cái là thêm: đánh răng, học bài, dọn đồ chơi…
- **Sao và quà**: mỗi việc xong cộng số sao đã đặt. Thẻ Thưởng cho biết sao đã kiếm,
  đã đổi, còn lại và riêng tuần này. Quà chưa đủ sao thì hiện *còn thiếu bao nhiêu*;
  đổi nhầm thì bấm **Hoàn** để trả sao lại.
- **Vòng tiến độ và chuỗi ngày**: phần trăm việc đã xong trong ngày, cộng huy hiệu
  🔥 khi bé làm trọn vẹn nhiều ngày liền.
- **Nhật ký**: bấm một nút là ghi nhanh chuyện ăn, ngủ, thuốc, nhiệt độ, học, tâm trạng,
  ghi chú — kèm giờ, sửa lại được.
- **Sổ sức khoẻ**: mỗi lần đo chiều cao cân nặng vẽ thành hai đường trên cùng một biểu đồ
  (chiều cao trục trái, cân nặng trục phải), kèm chênh lệch so với lần đo trước và BMI.
  Mũi tiêm, đợt thuốc, lần khám ghi kèm **ngày hẹn lần sau** và đếm ngược còn bao nhiêu ngày.
- **Biểu đồ bảy ngày**: mỗi ngày một cột theo tỉ lệ việc hoàn thành — xanh lá là trọn vẹn,
  tím là quá nửa, cam là còn ít. Chạm giữ vào cột hiện số cụ thể.
- **Nhắc giờ**: bật trong thẻ Khác. Máy báo trong vòng 5 phút kể từ mốc giờ nên mở muộn
  không bị dội thông báo cũ, và mỗi việc chỉ nhắc một lần trong ngày. Trên iPhone hãy
  thêm trang vào Màn hình chính rồi bật từ đó.
- **Giao diện sáng/tối** theo máy, hoặc chọn cứng một kiểu.

## Gửi cho người nhà

Thẻ **Khác** → **Gửi cho người nhà**: cả kho dữ liệu được nén rồi nhét vào phần sau dấu
`#` của địa chỉ. Phần đó không bao giờ đi lên máy chủ nào, nên cách này chạy mà không cần
dịch vụ lưu trữ. Bấm **Sao chép** (trên iPhone hiện luôn bảng Chia sẻ của hệ thống) rồi
gửi qua Zalo hay tin nhắn.

Người nhận mở link sẽ thấy băng báo *đây là dữ liệu người khác gửi*, xem thử được ngay,
rồi tự chọn **Lưu vào máy này** hay **Bỏ, giữ dữ liệu cũ** — dữ liệu sẵn có của họ không
bị ghi đè khi chưa đồng ý. Trong lúc xem thử, mọi đường ghi xuống máy đều bị chặn.

## Sao lưu

- **Xuất tệp JSON** ra tệp `be-ngoan-<ngày>.json`, cất trong app Tệp hay iCloud Drive.
- **Nhập tệp JSON** đọc lại bản đã cất.
- Mỗi lần ghi đè lớn (nhập tệp, nhận link, xoá sạch) đều cất bản cũ lại; dòng
  **Khôi phục bản trước đó** đổi qua đổi lại giữa hai bản nên bấm nhầm vẫn lấy lại được.

## Lưu ý

Dữ liệu nằm trong `localStorage` của trình duyệt, và **mỗi địa chỉ web giữ một kho riêng** —
mở bằng đường dẫn khác hay trình duyệt khác là một kho khác. Xoá dữ liệu duyệt web của
Safari cũng mất, nên thỉnh thoảng bấm **Xuất tệp JSON** để giữ một bản.

---

# 🧾 Sắp xếp & in hóa đơn — hai công cụ PowerShell cho Windows

**Tác giả: Nguyễn Thanh**

Hai công cụ tách riêng, chạy độc lập, dùng chung một mạch việc: gom hóa đơn lộn xộn
nhiều năm vào đúng thư mục theo danh sách Excel, rồi in ra máy in **đúng thứ tự trong
file Excel**.

| File | Việc |
| --- | --- |
| `sap-xep-hoa-don.ps1` + `.bat` | Sắp hóa đơn vào thư mục theo danh sách Excel |
| `in-hoa-don.ps1` + `.bat` | In hóa đơn hàng loạt ra máy in |

Bước sắp xếp ghi ra `thu-tu-in.txt` trong thư mục đích; công cụ in nạp file đó là in
đúng thứ tự trong Excel. Cần công cụ nào thì chạy công cụ đó, không cần cái kia.

Không cần cài thêm gì: dùng Windows PowerShell 5.1 có sẵn trong Windows. File Excel
`.xlsx` / `.xlsm` được đọc thẳng (không cần mở Excel); file `.xls` cũ thì máy cần có Excel.

## Chạy thế nào

Để cả bốn file trong cùng một thư mục rồi bấm đúp `sap-xep-hoa-don.bat` hoặc
`in-hoa-don.bat`. Nếu Windows chặn script, mở PowerShell rồi gõ:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\sap-xep-hoa-don.ps1
powershell -NoProfile -ExecutionPolicy Bypass -STA -File .\in-hoa-don.ps1
```

## Công cụ 1 — sắp xếp hóa đơn theo danh sách Excel

File Excel được hiểu đúng theo cách bảng đang làm: **mỗi sheet là một thư mục**, trong
sheet có nhiều bảng, **mỗi bảng là một sản phẩm** và thành một thư mục con; một sản phẩm
có thể gồm nhiều ký hiệu hóa đơn khác nhau.

| Trong Excel | Thành cái gì |
| --- | --- |
| Tên sheet (`Mau 01 - GGM`, `Mẫu 02B - GGM`...) | Thư mục cấp 1 |
| Tiêu đề phía trên bảng (`TELMA 80 H PLUS (TABLET B/100)`) | Thư mục cấp 2 (sản phẩm) |
| Cột `Ký hiệu` + `Số hóa đơn` + `Ngày hóa đơn` | Từng file hóa đơn được tìm và chép vào |
| Cột `tên file` (nếu có), VD `*K25TAA*618585` | Mẫu dùng để tìm đúng file PDF |

Công cụ tự dò các bảng nằm cạnh nhau trong cùng một sheet (như bảng ở cột G, cột M,
cột S), tự nhận tiêu đề cột dù có dấu hay không dấu.

Các bước:

1. **Chọn file Excel** → bấm *Đọc danh sách*. Mỗi sheet hiện kèm số hóa đơn và số sản phẩm.
2. Tick những sheet cần làm.
3. **Thêm thư mục nguồn** — thư mục đang chứa hóa đơn lộn xộn. Chỉ cần thêm thư mục cha
   (ví dụ `GGM`) và bật *Gồm thư mục con* là quét hết `HCM - 2023`, `HCM - 2024`,
   `HCM - 2025`, `Ha Noi`... trong đó; hoặc thêm từng thư mục riêng cũng được.
   Muốn lọc bớt thì điền năm vào ô *Chỉ lấy năm*, ví dụ `2023,2024,2025`;
   để trống là lấy hết.
4. **Chọn thư mục đích**. Nếu thư mục này đã có sẵn các thư mục con, bấm
   *Thư mục cho từng sheet...* để xem và sửa bảng ghép tên (xem mục dưới).
5. Bấm *Đối chiếu danh sách* để xem trước từng dòng.
6. Bấm *Tạo folder & chép file*.

Kết quả:

```
HoaDon\
├── Mau01-GGM\
│   ├── TELMA\
│   │   ├── 001_K25TAA_618585.pdf
│   │   ├── 002_K25TAA_565923.pdf
│   │   └── ...
│   ├── KLENZIT\
│   │   ├── 006_K25TDA_2228.pdf
│   │   └── ...
│   ├── KLENZIT-C\
│   └── COMBIWAVE\
├── Mau02A-GGM\
├── Mau02A-GGM30day\
├── Mau02B-GGM\
└── bao-cao-doi-chieu.csv
```

Số ở đầu tên file (`001_`, `002_`...) chạy liên tục trong một sheet **theo đúng thứ tự
dòng trong Excel**, nên chỉ cần sắp theo tên file là ra đúng thứ tự in.

### Khi thư mục đích đã có sẵn folder

Thư mục có sẵn thường đặt tên ngắn hơn tên trong Excel. Công cụ tự ghép theo tên gần
giống, mỗi thư mục chỉ nhận một sheet/sản phẩm:

| Trong Excel | Thư mục có sẵn |
| --- | --- |
| Sheet `Mau 01 - GGM` | `Mau01-GGM` |
| Sheet `Mau 02A-GGM-Truoc sau 30 ngay` | `Mau02A-GGM30day` |
| `TELMA 80 H PLUS (TABLET B/100)` | `TELMA` |
| `KLENZIT MS (GEL 15G)` | `KLENZIT` |
| `KLENZIT-C (GEL 15G)` | `KLENZIT-C` |

Bấm *Thư mục cho từng sheet...* để mở bảng ghép: dòng in đậm là sheet, dòng thụt vào là
sản phẩm, cột *Trạng thái* cho biết thư mục đã có sẵn hay sẽ được tạo mới. Muốn sửa thì
chọn dòng, chọn thư mục ở ô *Thư mục có sẵn* rồi bấm *Gán cho dòng đang chọn*; hoặc bấm
*Tạo thư mục mới theo tên trong Excel*. Tên nào không hợp thư mục nào thì tạo thư mục mới
theo đúng tên trong Excel.

### Cách dò tìm file

Nếu bảng có cột **`tên file`** (mẫu kiểu `*K25TAA*618585`), công cụ dùng thẳng mẫu đó để
tìm file — dấu `*` hiểu như khi tìm kiếm trong Windows. Mẫu không ra file nào thì mới quay
sang dò theo ký hiệu và số hóa đơn, nên tên file đặt ngược kiểu `00002850_K25TDA.pdf` vẫn
tìm ra. Không có cột này thì chỉ dò theo ký hiệu và số.

| Trạng thái | Nghĩa là |
| --- | --- |
| `Khớp theo cột tên file` | Đúng mẫu ghi trong cột `tên file` — chắc chắn nhất |
| `Khớp ký hiệu + số` | Tên file có cả ký hiệu (`K25TAA`) lẫn số hóa đơn |
| `Khớp số (thiếu ký hiệu)` | Chỉ khớp số hóa đơn, nên kiểm tra lại cho chắc |
| `Khớp nhiều file` | Có nhiều file cùng khớp, công cụ lấy file đầu tiên (dòng màu cam) |
| `Không tìm thấy` | Chưa có file cho hóa đơn này (dòng màu đỏ) |

Dấu gạch ngang, khoảng trắng, số 0 ở đầu, chữ có dấu đều được bỏ qua khi so khớp, nên
`HD K25TAA 618585.pdf`, `K25TAA-618585.pdf`, `00618585_K25TAA.pdf` đều tìm ra.
Khi một hóa đơn khớp nhiều file (bản sao nằm rải ở các thư mục năm khác nhau), công cụ
ưu tiên file nằm trong thư mục có đúng năm của hóa đơn — hóa đơn năm 2023 lấy bản trong
`HCM - 2023` chứ không lấy bản sao lạc trong `HCM - 2025` — rồi mới tới đường dẫn ngắn hơn. Nhấp đúp
vào một dòng để mở file kiểm tra. Mọi thứ đều được ghi vào `bao-cao-doi-chieu.csv`
(có cả danh sách hóa đơn còn thiếu).

Mặc định là **chép** file, giữ nguyên kho gốc; muốn dọn hẳn thì chọn *Di chuyển file*.
Khi di chuyển, công cụ chép sang thư mục đích và so lại dung lượng rồi mới bỏ bản gốc,
chép hụt thì file gốc vẫn còn nguyên.

### Kho hóa đơn nằm trong OneDrive

Trỏ *thư mục nguồn* vào đúng thư mục OneDrive đã đồng bộ trên máy, ví dụ
`C:\Users\<tên máy>\OneDrive - DKSH\Thanh - DKSH\GGM\HCM - 2023`, và bật *Gồm thư mục con*.

- File có biểu tượng đám mây là **chưa tải về máy**. Công cụ đếm và báo số lượng này
  trong nhật ký; Windows sẽ tự tải khi chép hoặc di chuyển, nên bước sắp xếp cần có mạng
  và sẽ chậm hơn. Muốn chạy nhanh: chuột phải thư mục trong File Explorer →
  *Always keep on this device*, chờ tải xong rồi mới sắp xếp.
- Nên đặt **thư mục đích nằm ngoài thư mục nguồn** (ví dụ ra Desktop hoặc một thư mục
  OneDrive khác). Nếu đích nằm trong nguồn, công cụ sẽ hỏi lại trước khi chạy, vì lần
  đối chiếu sau sẽ quét trúng cả những file vừa sắp xếp.

Xong xuôi, thư mục đích có thêm `bao-cao-doi-chieu.csv` (đối chiếu từng dòng) và
`thu-tu-in.txt` (danh sách đường dẫn theo đúng thứ tự Excel). Nút *Mở công cụ in →* mở
thẳng công cụ in.

## Công cụ 2 — in hóa đơn

- Quét một thư mục (kể cả thư mục con), tách số hóa đơn trong tên file.
- Sắp xếp theo số hóa đơn, theo tên file (`Tên file A → Z` = đúng thứ tự đã đánh số),
  hoặc theo ngày sửa.
- Tick chọn từng hóa đơn, tìm nhanh theo tên, hoặc *Chọn theo khoảng* từ số ... đến số ...
- Chọn máy in, số bản, rồi bấm **IN**. Có nút *Dừng*, có ô *In thử* để chạy nháp
  trước mà không tốn giấy.

Muốn in đúng thứ tự trong Excel: bấm **Nạp danh sách thứ tự in...** rồi chọn
`thu-tu-in.txt` trong thư mục đích. Danh sách hiện đúng thứ tự Excel, tick sẵn tất cả,
dòng nào đã xóa file thì tự bỏ qua và báo trong nhật ký.

## Lưu ý

- Máy phải có ứng dụng mở được loại file đó (PDF thì cần Adobe Reader hoặc phần mềm
  đọc PDF hỗ trợ lệnh in của Windows). Công cụ dùng lệnh `PrintTo` của Windows, nếu ứng
  dụng không hỗ trợ thì tự tạm đổi máy in mặc định rồi trả lại sau khi in xong.
- *Chờ mỗi file (giây)* và *Chờ hàng đợi máy in trống* giúp các hóa đơn ra đúng thứ tự;
  máy in chậm thì tăng số giây lên.
- Nếu quét cả thư mục đích có nhiều sheet cùng lúc, số thứ tự của các sheet sẽ đan vào
  nhau. In từng thư mục sheet một, hoặc nạp `thu-tu-in.txt`.
- Ô *Mẫu số hóa đơn (regex)* để trống là lấy dãy số cuối trong tên file; muốn lấy chỗ
  khác thì điền regex có nhóm bắt, ví dụ `K\d{2}[A-Z]+.?(\d+)`.
