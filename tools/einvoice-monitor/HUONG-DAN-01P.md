# Thêm server VNSGNEIVAP01P (sales org VN57)

`monitor.ps1` đã được sửa để:

- Quét thêm thư mục lỗi `\\vnsgneivap01p\DKSH(VN)\ERROR\INV` và `\ERROR\DO`.
- Thêm `01P` vào báo cáo tổng billing theo ngày (`\\vnsgneivap01p\PRODATA\{date}`), nếu `config.json` không tự khai báo `daily_report.folders`.
- Đọc tùy chọn mới `exclude_bu_codes` trên từng server trong `sap_to_pp.sftp.servers`. File có mã BU nằm trong danh sách đó **không được đếm** cho server ấy (bước SAP->PP và báo cáo theo ngày).

## Cần sửa trong `config.json`

1. `sap_to_pp.sftp.servers`: thêm 01P, và thêm `exclude_bu_codes` cho server .2:

```json
"servers": [
  { "name": "VNSGNEIVAP05P", "host": "..." },
  { "name": "<server .2>",   "host": "....2", "exclude_bu_codes": ["VN57"] },
  { "name": "VNSGNEIVAP01P", "host": "<IP/host cua 01P>" }
]
```

2. `pp_to_sap_unc.servers`: thêm 01P (chép mục của 05P, đổi tên host):

```json
{ "name": "VNSGNEIVAP01P",
  "bakup":   "\\\\vnsgneivap01p\\...\\Bakup",
  "arcdata": "\\\\vnsgneivap01p\\...\\Arcdata" }
```

3. `service_check.servers` (nếu có dùng): thêm

```json
{ "name": "VNSGNEIVAP01P", "folder": "\\\\vnsgneivap01p\\PRODATA" }
```

4. Nếu `config.json` có `daily_report.folders` riêng thì phải tự thêm mục `01P` vào đó.
   Mục `.2` có thể ghi `"exclude_bu_codes": ["VN57"]`. Nếu không ghi, tool lấy theo cấu hình server tương ứng.

`"name"` của 01P trong `pp_to_sap_unc.servers` phải trùng với `sap_to_pp.sftp.servers`, để tool tự chuyển sang FTP khi UNC bị từ chối quyền.
