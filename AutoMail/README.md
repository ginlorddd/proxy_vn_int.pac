# AutoMail

Ứng dụng Python gửi mail tự động qua Microsoft Outlook desktop, có giao diện PySide6 với trình soạn thảo rich text giống Outlook, scheduler và HTTP API để app khác (ví dụ WinForms) truyền dữ liệu gửi mail.

## Chạy

```bash
pip install -r requirements.txt
python main.py
```

Chạy nền không UI:

```bash
python main.py --no-ui
```

## API

Header bắt buộc: `X-AutoMail-Token: change-me` (đổi trong `config.json`).

- `GET /health`
- `GET /config`
- `GET /state`
- `POST /config` cập nhật cấu hình JSON.
- `POST /send` gửi ngay. Body có thể là object mail hoặc `{ "mail": {...} }`.
- `POST /scheduler/start`
- `POST /scheduler/stop`

Schedule hỗ trợ chọn kiểu `daily`, `weekly`, `interval`, giờ gửi, các ngày trong tuần (`weekdays`: T2=0 ... CN=6) và danh sách ngày cụ thể `date_schedules` để chọn ngày nào gửi nội dung/template nào.

Ví dụ WinForms/C# gửi mail:

```csharp
using var client = new HttpClient();
client.DefaultRequestHeaders.Add("X-AutoMail-Token", "change-me");
var json = "{\"mail\":{\"to\":[\"a@example.com\"],\"subject\":\"Test\",\"body\":\"<b>Hello</b>\",\"body_format\":\"html\"}}";
await client.PostAsync("http://127.0.0.1:8765/send", new StringContent(json, Encoding.UTF8, "application/json"));
```

## Lưu ý

- Giao diện rich text cần `PySide6`; gửi mail và lấy danh sách From/account từ Outlook cần Windows, Microsoft Outlook đã đăng nhập, và package `pywin32`. Nếu dùng Exchange, app sẽ thử lấy `PrimarySmtpAddress` từ AddressEntry khi `SmtpAddress` trống.
- Trên Linux/macOS vẫn có thể test đọc config/API, nhưng không gửi được Outlook COM.
