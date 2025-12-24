# Outlook Contacts Automation (Playwright/Python)

Script này tự động đăng nhập Outlook (web) và thực hiện thao tác xoá trong mục People/Contacts bằng Python + Playwright.

## Cài đặt

1. Tạo virtualenv nếu cần, sau đó cài phụ thuộc:

   ```bash
   python3 -m venv venv
   source venv/bin/activate 
   pip install -r requirements.txt
   python -m playwright install
   ```
2. Thiết lập biến môi trường hoặc tệp `.env`:

   ```env
   OUTLOOK_EMAIL=you@example.com
   OUTLOOK_PASSWORD=your-password
   # Tuỳ chọn: chọn browser mặc định
   PLAYWRIGHT_BROWSER=firefox
   ```

   Ghi chú: trong [src/test_outlook_contacts.py](src/test_outlook_contacts.py) hiện đang có email/password hard-code. Nếu muốn dùng `.env`, hãy uncomment 2 dòng `os.getenv(...)` và xoá hard-code.

## Chạy thử

```bash
# Chromium (mặc định)
python src/test_outlook_contacts.py

# Firefox
python src/test_outlook_contacts.py --browser firefox

# WebKit
python src/test_outlook_contacts.py --browser webkit
```

Nếu chưa cài browser tương ứng, chạy:

```bash
python -m playwright install firefox
```

## Lưu ý

- Đăng nhập Outlook có thể có CAPTCHA/2FA; khi đó script sẽ dừng cho đến khi bạn xử lý thủ công.
- Bộ chọn (selectors) có thể thay đổi tùy giao diện; nếu script không tìm thấy nút/ô nhập, mở Developer Tools và điều chỉnh bộ chọn trong mã.
- `headless=False` mặc định để dễ debug; đổi sang `True` trong hàm `run` khi cần chạy ngầm.
