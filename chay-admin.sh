#!/usr/bin/env bash
# ============================================================
#  Chạy chương trình với quyền quản trị (Linux / macOS)
#  Cách dùng:  ./chay-admin.sh [cổng] [trang]
#  Ví dụ:      ./chay-admin.sh 8080 lich-bieu.html
# ============================================================
set -euo pipefail

PORT="${1:-8080}"
TRANG="${2:-index.html}"

# --- 1. Chưa phải quản trị thì chạy lại qua sudo ---
if [ "$(id -u)" -ne 0 ]; then
  if ! command -v sudo >/dev/null 2>&1; then
    echo "[!] Máy không có sudo, hãy chạy bằng tài khoản quản trị." >&2
    exit 1
  fi
  echo "Đang xin quyền quản trị..."
  exec sudo -E "$0" "$PORT" "$TRANG"
fi

cd "$(dirname "$0")"
echo
echo "=== Đang chạy với quyền quản trị ==="
echo "Thư mục : $PWD"
echo "Cổng    : $PORT"
echo

# --- 2. Mở cổng trên tường lửa nếu máy có ufw hoặc firewalld ---
if command -v ufw >/dev/null 2>&1 && ufw status 2>/dev/null | head -1 | grep -qi active; then
  if ufw allow "$PORT/tcp" >/dev/null 2>&1; then
    echo "[OK] Đã mở cổng $PORT trên ufw."
  else
    echo "[!] Không mở được cổng $PORT trên ufw."
  fi
elif command -v firewall-cmd >/dev/null 2>&1 && firewall-cmd --state >/dev/null 2>&1; then
  if firewall-cmd --add-port="$PORT/tcp" >/dev/null 2>&1; then
    echo "[OK] Đã mở cổng $PORT trên firewalld."
  else
    echo "[!] Không mở được cổng $PORT trên firewalld."
  fi
fi

# --- 3. Tìm một máy chủ tĩnh có sẵn trên máy ---
SERVER=()
if command -v python3 >/dev/null 2>&1; then
  SERVER=(python3 -m http.server "$PORT" --bind 0.0.0.0)
elif command -v python >/dev/null 2>&1; then
  SERVER=(python -m http.server "$PORT" --bind 0.0.0.0)
elif command -v npx >/dev/null 2>&1; then
  SERVER=(npx --yes http-server . -p "$PORT" -a 0.0.0.0)
fi

mo_trinh_duyet() {
  url="$1"
  mo=""
  if command -v xdg-open >/dev/null 2>&1; then
    mo=xdg-open
  elif command -v open >/dev/null 2>&1; then
    mo=open
  else
    return 0
  fi
  # mở bằng tài khoản người dùng chứ không phải tài khoản quản trị
  if [ -n "${SUDO_USER:-}" ] && [ "$SUDO_USER" != "root" ]; then
    sudo -u "$SUDO_USER" "$mo" "$url" >/dev/null 2>&1 || true
  else
    "$mo" "$url" >/dev/null 2>&1 || true
  fi
}

if [ ${#SERVER[@]} -eq 0 ]; then
  echo "[!] Không thấy Python lẫn Node trên máy."
  echo "    Mở thẳng file $TRANG bằng trình duyệt."
  mo_trinh_duyet "file://$PWD/$TRANG"
  exit 0
fi

# --- 4. Báo địa chỉ để máy khác trong nhà gõ vào ---
IP="$(hostname -I 2>/dev/null | awk '{print $1}')"
if [ -z "$IP" ]; then
  IP="$(ipconfig getifaddr en0 2>/dev/null || true)"
fi
echo "Mở trên máy này    : http://localhost:$PORT/$TRANG"
if [ -n "$IP" ]; then
  echo "Mở trên điện thoại : http://$IP:$PORT/$TRANG"
fi
echo
echo "Bấm Ctrl+C để dừng máy chủ."
echo

# --- 5. Đợi máy chủ lên rồi mở trình duyệt, sau đó chạy máy chủ ---
( sleep 2; mo_trinh_duyet "http://localhost:$PORT/$TRANG" ) &
exec "${SERVER[@]}"
