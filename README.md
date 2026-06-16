# Facebook Crawler

Crawl bài viết từ trang web FPT University → lưu Excel → tóm tắt AI.

> Mới bắt đầu? Đọc [HUONG-DAN.md](HUONG-DAN.md). Tra lệnh? Xem [LENH.md](LENH.md).

---

## Bắt đầu trong 30 giây

```bash
pip install -r requirements.txt
# Double-click Chay.bat
# hoặc:
python fb_crawler.py
```

---

## Hai loại bước

| Loại | Ví dụ |
|---|---|
| **[CMD]** Lệnh Terminal | `python fb_crawler.py` |
| **[AGENT]** Nhờ Claude | *"Thêm nguồn crawl mới"* |

---

## Bản đồ thư mục

```
raw/                  dữ liệu gốc bất biến
data/                 dữ liệu trung gian
Trang web FPT/        kết quả Excel  ← ĐỌC Ở ĐÂY
outputs/              export chính thức
scripts/              script phụ trợ
tests/                test files
wiki/                 tài liệu nội bộ
workshop/             thử nghiệm sandbox
fb_crawler.py         crawl bài mới
lich_su.py            crawl lịch sử
Chay.bat              launcher crawl mới
Lay lich su.bat       launcher lịch sử
```
