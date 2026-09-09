# Ảnh đặc sản

Thư mục này chứa ảnh của các món trong `dac-san.html`. Trang **không bắt buộc phải có
ảnh**: món nào chưa có tệp thì thẻ hàng vẫn đẹp — hiện ô màu theo miền kèm biểu tượng
của món. Bỏ ảnh vào dần từng tấm cũng được, không hỏng gì cả.

## Đặt tên tệp

Tên tệp lấy theo **tên món, bỏ dấu, nối bằng gạch ngang**, đuôi `.jpg`:

```
Chả mực Hạ Long          →  cha-muc-ha-long.jpg
Kẹo dừa Bến Tre          →  keo-dua-ben-tre.jpg
Bánh tráng phơi sương Trảng Bàng  →  banh-trang-phoi-suong-trang-bang.jpg
```

Danh sách đầy đủ **252 tên tệp** — xếp theo miền và theo tỉnh — nằm ở
[`danh-sach-anh.txt`](danh-sach-anh.txt). Cứ mở ra, chép đúng tên ở cột đầu là được.

## Ảnh nên như thế nào

- Khung **4:3 nằm ngang** (trang tự cắt gọn theo khung này), cỡ **800×600** là vừa đẹp.
- Nặng dưới **150 KB** mỗi tấm cho trang mở nhanh.
- Chụp món ăn thật, đủ sáng, chủ thể nằm giữa khung.

## Mấy chỗ chỉnh được

Mở `dac-san.html`, tìm khối `const ANH` ở đầu phần `<script>`:

```js
const ANH = {
  bat: true,          // để false nếu chưa có tấm ảnh nào, trang khỏi đi tìm tệp
  thuMuc: 'anh/',     // thư mục chứa ảnh
  duoi: '.jpg'        // đổi thành '.webp' hay '.png' nếu dùng định dạng khác
};
```

Muốn một món dùng **ảnh trên mạng** hoặc tên tệp khác thì ghi thẳng đường dẫn vào **ô thứ
sáu** của món đó trong `DAC_SAN` — ô này ưu tiên hơn tên tự suy ra:

```js
["Chả mực Hạ Long","🦑","Khay 500g",320000,"Mực mai giã tay…","https://.../cha-muc.jpg"],
```

## Lưu ý khi mở trang

Mở thẳng bằng cách nhấp đúp vào `dac-san.html` thì một số trình duyệt chặn đọc tệp ảnh
bên cạnh. Chạy qua máy chủ cục bộ là hiện đủ:

```bash
npx http-server .
# rồi mở http://localhost:8080/dac-san.html
```

Khi chia sẻ trang cho người khác, nhớ gửi kèm cả thư mục `anh/` này.
