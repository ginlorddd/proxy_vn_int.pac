# AutoMail

Ứng dụng Python gửi mail tự động qua Microsoft Outlook desktop, có giao diện cấu hình bằng Tkinter, scheduler và HTTP API để app khác (ví dụ WinForms) truyền dữ liệu gửi mail.

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

Ví dụ WinForms/C# gửi mail:

```csharp
using var client = new HttpClient();
client.DefaultRequestHeaders.Add("X-AutoMail-Token", "change-me");
var json = "{\"mail\":{\"to\":[\"a@example.com\"],\"subject\":\"Test\",\"body\":\"<b>Hello</b>\",\"body_format\":\"html\"}}";
await client.PostAsync("http://127.0.0.1:8765/send", new StringContent(json, Encoding.UTF8, "application/json"));
```

## Lưu ý

- Gửi mail cần Windows, Microsoft Outlook đã đăng nhập, và package `pywin32`.
- Trên Linux/macOS vẫn có thể test đọc config/API, nhưng không gửi được Outlook COM.
