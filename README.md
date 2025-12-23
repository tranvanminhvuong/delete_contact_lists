# Outlook Contacts Automation (Playwright/Python)

Script này tự động đăng nhập Outlook và xóa 1 liên hệ theo tên, chạy bằng Python + Playwright.

## Cài đặt

1. Tạo virtualenv nếu cần, sau đó cài phụ thuộc:
   ```bash
   python3 -m venv venv
   source venv/bin/activate 
   pip install -r requirements.txt
   playwright install chromium
   ```
2. Thiết lập biến môi trường hoặc tệp `.env`:
   ```env
   OUTLOOK_EMAIL=you@example.com
   OUTLOOK_PASSWORD=your-password
   ```

## Chạy thử

```bash
python src/test_outlook_contacts.py "Tên liên hệ"  # headful để theo dõi thao tác
```

## Lưu ý

- Đăng nhập Outlook có thể có CAPTCHA/2FA; khi đó script sẽ dừng cho đến khi bạn xử lý thủ công.
- Bộ chọn (selectors) có thể thay đổi tùy giao diện; nếu script không tìm thấy nút/ô nhập, mở Developer Tools và điều chỉnh bộ chọn trong mã.
- `headless=False` mặc định để dễ debug; đổi sang `True` trong hàm `run` khi cần chạy ngầm.
