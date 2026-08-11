/**
 * Việc của team — nhắc việc qua mail.
 *
 * Dán toàn bộ file này vào Apps Script rồi Deploy thành Web app. Nó đọc kho
 * chung (Firebase Realtime Database) mà trang `cong-viec.html` đang dùng, dựng
 * bản báo cáo rồi gửi mail cho cả team.
 *
 * Chạy được hai kiểu:
 *   - Tự động: đặt lịch chạy hằng ngày, không cần ai mở trang.
 *   - Bấm tay: trang web gọi vào đây để gửi ngay một bản.
 *
 * Các bước, làm một lần chừng 5 phút:
 *   1. Vào script.google.com, tạo project mới, xoá hết rồi dán file này vào.
 *   2. Sửa hai dòng KHO và PHONG bên dưới cho khớp kho của team.
 *   3. Chạy thử hàm `guiNgayBayGio` một lần, Google hỏi quyền thì cho phép.
 *      Kiểm tra hộp thư xem đã nhận được chưa.
 *   4. Chạy hàm `datLichHangNgay` một lần để nó tự gửi mỗi sáng.
 *   5. Muốn bấm gửi từ trang web thì Deploy → New deployment → Web app,
 *      "Execute as: Me", "Who has access: Anyone", rồi dán URL vào nút
 *      ✉️ Nhắc qua mail trong trang.
 *
 * Danh sách người nhận lấy thẳng từ kho chung (ô "Người nhận" trong trang),
 * nên sau này thêm bớt người thì sửa trong trang là xong, không phải mở lại
 * Apps Script. NGUOI_NHAN bên dưới chỉ là bản dự phòng khi kho chưa có ai.
 */

/* ================== phần cần sửa ================== */

// Địa chỉ kho Firebase, y như dán trong nút 🗄 Kho chung của trang
var KHO = 'https://ten-cua-ban-default-rtdb.asia-southeast1.firebasedatabase.app';

// Tên bảng trong kho, mặc định của trang là 'viec-team'
var PHONG = 'viec-team';

// Dự phòng khi trong kho chưa điền người nhận nào
var NGUOI_NHAN = '';

// Giờ gửi hằng ngày, 0-23 theo múi giờ của project Apps Script
var GIO_GUI = 8;

// Đặt true thì ngày nào không có việc đáng nhắc sẽ im lặng, không gửi mail
var IM_KHI_KHONG_CO_GI = true;

/* ================== chạy tay & đặt lịch ================== */

/** Gửi ngay một bản báo cáo. Chạy hàm này để thử. */
function guiNgayBayGio() {
  var kq = guiBaoCao(false);
  Logger.log(kq);
  return kq;
}

/** Hàm mà lịch hằng ngày gọi tới. */
function baoCaoHangNgay() {
  return guiBaoCao(IM_KHI_KHONG_CO_GI);
}

/** Đặt lịch tự gửi mỗi ngày. Chạy một lần là xong. */
function datLichHangNgay() {
  huyLich();
  ScriptApp.newTrigger('baoCaoHangNgay')
    .timeBased()
    .everyDays(1)
    .atHour(GIO_GUI)
    .create();
  return 'Đã đặt lịch gửi lúc ' + GIO_GUI + ' giờ mỗi ngày.';
}

/** Bỏ lịch tự gửi. */
function huyLich() {
  var ds = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ds.length; i++) {
    if (ds[i].getHandlerFunction() === 'baoCaoHangNgay') ScriptApp.deleteTrigger(ds[i]);
  }
  return 'Đã bỏ lịch tự gửi.';
}

/* ================== trang web gọi vào ================== */

/**
 * Trình duyệt hay chặn cuộc gọi thẳng sang Google vì khác tên miền, nên trang
 * gọi vòng qua thẻ script: thêm callback=tên thì trả về JavaScript thay vì JSON.
 */
function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;

  try {
    if (p.fn === 'gui') {
      out = { ok: true, ket_qua: guiBaoCao(false) };
    } else {
      // Không nêu fn thì chỉ trả lời để trang biết là nối được
      var kho = docKho();
      out = {
        ok: true,
        song: true,
        so_viec: layViec(kho).length,
        nguoi_nhan: layNguoiNhan(kho)
      };
    }
  } catch (err) {
    out = { ok: false, loi: String(err) };
  }

  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + JSON.stringify(out) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  return doGet(e);
}

/* ================== đọc kho ================== */

function docKho() {
  var url = String(KHO).replace(/\/+$/, '') + '/' + encodeURIComponent(PHONG) + '.json';
  var res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) {
    throw new Error('Kho trả về HTTP ' + res.getResponseCode() +
                    '. Kiểm tra địa chỉ KHO và phần Rules của Firebase.');
  }
  var data = JSON.parse(res.getContentText() || 'null');
  return data && typeof data === 'object' ? data : {};
}

function layNguoiNhan(kho) {
  var chu = (kho && kho.mail) ? String(kho.mail) : '';
  if (!chu.trim()) chu = NGUOI_NHAN;
  return chu.split(/[\s,;]+/)
    .map(function (s) { return s.trim(); })
    .filter(function (s) { return s.indexOf('@') > 0; });
}

/** Việc còn sống, bỏ mấy cái đã xoá (chỉ còn dấu mộ). */
function layViec(kho) {
  var ra = [];
  var tasks = kho.tasks || {};
  for (var id in tasks) {
    if (!Object.prototype.hasOwnProperty.call(tasks, id)) continue;
    var t = tasks[id];
    if (!t || typeof t !== 'object' || t.del) continue;
    ra.push({
      title: t.title || '(việc chưa đặt tên)',
      who: Array.isArray(t.who) ? t.who : [],
      status: t.status || 'todo',
      due: t.due || '',
      next: t.next || '',
      nextDue: t.nextDue || '',
      doneAt: typeof t.doneAt === 'number' ? t.doneAt : 0,
      rate: t.rate && t.rate.sao ? t.rate : null
    });
  }
  return ra;
}

/* ================== ngày tháng ================== */

function homNay() {
  var n = new Date();
  return new Date(n.getFullYear(), n.getMonth(), n.getDate());
}

function docNgay(iso) {
  var p = String(iso).split('-');
  return new Date(+p[0], +p[1] - 1, +p[2]);
}

function conMayNgay(iso) {
  return Math.round((docNgay(iso) - homNay()) / 86400000);
}

function ngayVN(iso) {
  var d = docNgay(iso);
  return d.getDate() + '/' + (d.getMonth() + 1);
}

function homNayVN() {
  var d = homNay();
  return d.getDate() + '/' + (d.getMonth() + 1) + '/' + d.getFullYear();
}

/* ================== dựng báo cáo ================== */

function dungBaoCao(viec) {
  var nhom = { todo: [], doing: [], blocked: [], done: [] };
  var tre = [], homNayHan = [], keTiep = [];

  viec.forEach(function (t) {
    if (nhom[t.status]) nhom[t.status].push(t);

    if (t.status !== 'done' && t.due) {
      var con = conMayNgay(t.due);
      if (con < 0) tre.push({ t: t, con: con });
      else if (con === 0) homNayHan.push(t);
    }
    if (t.status !== 'done' && t.next) keTiep.push(t);
  });

  tre.sort(function (a, b) { return a.con - b.con; });
  keTiep.sort(function (a, b) {
    var ha = a.nextDue || a.due || '9999-99-99';
    var hb = b.nextDue || b.due || '9999-99-99';
    return ha < hb ? -1 : (ha > hb ? 1 : 0);
  });

  // xong trong 7 ngày qua
  var moc = Date.now() - 7 * 86400000;
  var xongTuan = nhom.done.filter(function (t) { return t.doneAt && t.doneAt >= moc; });

  return {
    nhom: nhom, tre: tre, homNay: homNayHan, keTiep: keTiep, xongTuan: xongTuan,
    dangCan: tre.length + homNayHan.length + nhom.blocked.length
  };
}

function ten(t) {
  return t.who.length ? ' (' + t.who.join(', ') + ')' : '';
}

function chuThuong(bc) {
  var d = [];
  d.push('BÁO CÁO CÔNG VIỆC — ' + homNayVN());
  d.push('');

  if (bc.tre.length) {
    d.push('⚠️ TRỄ HẠN (' + bc.tre.length + ')');
    bc.tre.forEach(function (x) {
      d.push('- ' + x.t.title + ten(x.t) + ' — trễ ' + (-x.con) + ' ngày');
    });
    d.push('');
  }
  if (bc.homNay.length) {
    d.push('⏰ ĐẾN HẠN HÔM NAY (' + bc.homNay.length + ')');
    bc.homNay.forEach(function (t) { d.push('- ' + t.title + ten(t)); });
    d.push('');
  }
  if (bc.nhom.blocked.length) {
    d.push('⛔ ĐANG VƯỚNG (' + bc.nhom.blocked.length + ')');
    bc.nhom.blocked.forEach(function (t) {
      d.push('- ' + t.title + ten(t) + (t.next ? '\n  → tiếp theo: ' + t.next : ''));
    });
    d.push('');
  }
  if (bc.keTiep.length) {
    d.push('➡️ KẾ HOẠCH TIẾP THEO');
    bc.keTiep.slice(0, 10).forEach(function (t, i) {
      var khi = t.nextDue || t.due;
      d.push((i + 1) + '. ' + t.next + ten(t) + (khi ? ' — ' + ngayVN(khi) : ''));
    });
    d.push('');
  }
  if (bc.xongTuan.length) {
    d.push('✅ XONG TRONG 7 NGÀY (' + bc.xongTuan.length + ')');
    bc.xongTuan.forEach(function (t) {
      d.push('- ' + t.title + ten(t) + (t.rate ? ' — ' + t.rate.sao + '/5 sao' : ''));
    });
    d.push('');
  }

  d.push('Đang làm ' + bc.nhom.doing.length +
         ' · Chưa làm ' + bc.nhom.todo.length +
         ' · Vướng ' + bc.nhom.blocked.length +
         ' · Xong ' + bc.nhom.done.length);
  return d.join('\n');
}

function thoat(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function chuHtml(bc) {
  var h = [];
  h.push('<div style="font:15px/1.55 system-ui,-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:#1d2433">');
  h.push('<h2 style="margin:0 0 4px;font-size:18px">Báo cáo công việc</h2>');
  h.push('<div style="color:#6b7488;font-size:13px;margin-bottom:16px">' + homNayVN() + '</div>');

  // Một thanh tiến độ nho nhỏ, cùng bộ màu với trang web
  var tong = bc.nhom.todo.length + bc.nhom.blocked.length +
             bc.nhom.doing.length + bc.nhom.done.length;
  if (tong) {
    var mieng = [
      { n: bc.nhom.todo.length,    m: '#4a3aa7' },
      { n: bc.nhom.blocked.length, m: '#e34948' },
      { n: bc.nhom.doing.length,   m: '#2a78d6' },
      { n: bc.nhom.done.length,    m: '#008300' }
    ];
    h.push('<table cellpadding="0" cellspacing="0" style="width:100%;height:18px;' +
           'border-radius:9px;overflow:hidden;margin-bottom:6px"><tr>');
    mieng.forEach(function (x) {
      if (!x.n) return;
      h.push('<td style="width:' + Math.round(x.n / tong * 100) +
             '%;background:' + x.m + '">&nbsp;</td>');
    });
    h.push('</tr></table>');
    h.push('<div style="color:#6b7488;font-size:13px;margin-bottom:18px">' +
           'Chưa làm ' + bc.nhom.todo.length + ' · Vướng ' + bc.nhom.blocked.length +
           ' · Đang làm ' + bc.nhom.doing.length + ' · Xong ' + bc.nhom.done.length +
           '</div>');
  }

  function muc(nhan, mau, dong) {
    if (!dong.length) return;
    h.push('<h3 style="margin:18px 0 6px;font-size:14px;color:' + mau + '">' + nhan + '</h3>');
    h.push('<ul style="margin:0;padding-left:20px">');
    dong.forEach(function (x) { h.push('<li style="margin:3px 0">' + x + '</li>'); });
    h.push('</ul>');
  }

  muc('⚠️ Trễ hạn (' + bc.tre.length + ')', '#d63a3a', bc.tre.map(function (x) {
    return '<b>' + thoat(x.t.title) + '</b>' + thoat(ten(x.t)) +
           ' — <span style="color:#d63a3a">trễ ' + (-x.con) + ' ngày</span>';
  }));

  muc('⏰ Đến hạn hôm nay (' + bc.homNay.length + ')', '#b96908', bc.homNay.map(function (t) {
    return '<b>' + thoat(t.title) + '</b>' + thoat(ten(t));
  }));

  muc('⛔ Đang vướng (' + bc.nhom.blocked.length + ')', '#d63a3a',
    bc.nhom.blocked.map(function (t) {
      return '<b>' + thoat(t.title) + '</b>' + thoat(ten(t)) +
             (t.next ? '<br><span style="color:#6b7488">→ ' + thoat(t.next) + '</span>' : '');
    }));

  muc('➡️ Kế hoạch tiếp theo', '#2f6df6', bc.keTiep.slice(0, 10).map(function (t) {
    var khi = t.nextDue || t.due;
    return thoat(t.next) + thoat(ten(t)) +
           (khi ? ' <span style="color:#6b7488">— ' + ngayVN(khi) + '</span>' : '');
  }));

  muc('✅ Xong trong 7 ngày (' + bc.xongTuan.length + ')', '#16955f',
    bc.xongTuan.map(function (t) {
      return thoat(t.title) + thoat(ten(t)) +
             (t.rate ? ' <span style="color:#b96908">' + t.rate.sao + '/5 sao</span>' : '');
    }));

  h.push('</div>');
  return h.join('');
}

/* ================== gửi ================== */

function guiBaoCao(imKhiRong) {
  var kho = docKho();
  var nhan = layNguoiNhan(kho);
  if (!nhan.length) {
    return 'Chưa có người nhận nào. Điền địa chỉ mail ở nút ✉️ Nhắc qua mail trong trang, ' +
           'hoặc điền vào NGUOI_NHAN trong file này.';
  }

  var viec = layViec(kho);
  var bc = dungBaoCao(viec);

  if (imKhiRong && !bc.dangCan && !bc.xongTuan.length) {
    return 'Hôm nay không có gì đáng nhắc, không gửi.';
  }

  var tieuDe = 'Việc của team — ' + homNayVN();
  if (bc.tre.length) tieuDe += ' · ' + bc.tre.length + ' việc trễ hạn';
  else if (bc.homNay.length) tieuDe += ' · ' + bc.homNay.length + ' việc đến hạn hôm nay';

  MailApp.sendEmail({
    to: nhan.join(','),
    subject: tieuDe,
    body: chuThuong(bc),
    htmlBody: chuHtml(bc),
    name: 'Việc của team'
  });

  return 'Đã gửi cho ' + nhan.length + ' người: ' + nhan.join(', ');
}
