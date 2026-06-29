# Facebook Crawler — Hướng dẫn vận hành cho Agent

Dự án này: crawl bài viết từ trang web FPT University → lưu vào Excel để phân tích, theo dõi nội dung.

---

## §0. Bảng vận hành nhanh

| User muốn | Agent làm gì |
|---|---|
| "crawl tin mới" | `Chay.bat` hoặc `python scripts/fb_crawler.py` |
| "lấy toàn bộ lịch sử" | `Lay lich su.bat` hoặc `python scripts/lich_su.py` (~3-5 phút) |
| "cập nhật tóm tắt" | `python scripts/cap_nhat_tom_tat.py` |
| "đồng bộ Excel" | `python scripts/sync_xlsx.py` |
| "kiểm tra HTML" | `python scripts/check_html.py` |
| "kết quả ở đâu" | File thống nhất `Trang web FPT/fpt_news_tong_hop.xlsx` (tích lũy mọi lần chạy); snapshot ngày là `fpt_news_YYYYMMDD.xlsx` |
| "chạy test" | `pytest tests/` |
| "permissions Claude Code" | `.claude/settings.json` |

**Quy tắc cho agent:**
- Script `.bat` dùng `cd /d "%~dp0"` — khi chạy từ project root, relative path `Trang web FPT/...` vẫn đúng.
- Output Excel lưu trong `Trang web FPT/`, không xóa file cũ khi chưa được yêu cầu.
- Tất cả scripts đã chuyển vào `scripts/` — chạy bằng `python scripts/[tên].py` từ project root.

---

## Kiến trúc pipeline

```
Trang web FPT University (HTML)
  └── fb_crawler.py          → crawl bài mới → snapshot ngày + gộp vào Trang web FPT/fpt_news_tong_hop.xlsx (dedupe theo link)
  └── lich_su.py             → crawl toàn bộ lịch sử (1 lần)
  └── cap_nhat_tom_tat.py    → tóm tắt nội dung bằng AI
  └── sync_xlsx.py           → đồng bộ, deduplicate
  └── check_html.py          → kiểm tra HTML hợp lệ
```

---

## Cấu trúc thư mục

| Thư mục/File | Mục đích |
|---|---|
| `fb_crawler.py` | Crawl bài viết mới |
| `lich_su.py` | Crawl toàn bộ lịch sử |
| `cap_nhat_tom_tat.py` | Tạo tóm tắt AI |
| `sync_xlsx.py` | Đồng bộ Excel |
| `check_html.py` | Kiểm tra HTML |
| `Chay.bat` | Launcher crawl mới |
| `Lay lich su.bat` | Launcher lịch sử |
| `Trang web FPT/` | Output Excel — kết quả cuối |
| `data/` | Dữ liệu trung gian |
| `raw/` | Dữ liệu thô gốc (bất biến) |
| `scripts/` | Helper scripts tương lai (hiện rỗng) |
| `prompts/` | Prompt LLM tái sử dụng (hiện rỗng) |
| `review/` | Tài liệu review output (hiện rỗng) |
| `tests/` | Test files pytest (hiện rỗng) |
| `outputs/` | Export chính thức (hiện rỗng) |
| `wiki/` | Tài liệu nội bộ |
| `workshop/` | Thử nghiệm sandbox |
| `.claude/` | Cấu hình Claude Code: permissions |

---

## Quy tắc làm việc

1. **Không di chuyển `.py` files** — `.bat` dùng đường dẫn tương đối, sẽ hỏng nếu di chuyển.
2. **Workshop trước production** — thử ở `workshop/` trước khi sửa script chính.
3. **Data bất biến** — `raw/` chỉ đọc, output ra `Trang web FPT/`.
4. **Kế thừa, không đập lại** — mọi thay đổi phải build trên code đang chạy.
