# Lệnh Thường Dùng — Facebook Crawler

> **[CMD]** = gõ Terminal | **[AGENT]** = nhờ Claude làm

---

## 1. Crawl & thu thập

| Lệnh | Loại | Làm gì | Khi nào dùng |
|---|---|---|---|
| `Chay.bat` (double-click) | [CMD] | Crawl bài viết mới | Dùng thường xuyên |
| `python fb_crawler.py` | [CMD] | Crawl từ Terminal | Khi cần xem log |
| `Lay lich su.bat` (double-click) | [CMD] | Lấy toàn bộ lịch sử (~3-5 phút) | Chạy 1 lần duy nhất |
| `python lich_su.py` | [CMD] | Lịch sử từ Terminal | Khi cần xem log |

## 2. Xử lý & đồng bộ

| Lệnh | Loại | Làm gì | Khi nào dùng |
|---|---|---|---|
| `python cap_nhat_tom_tat.py` | [CMD] | Tạo tóm tắt AI cho bài viết | Sau crawl |
| `python sync_xlsx.py` | [CMD] | Đồng bộ, xóa trùng lặp | Khi file Excel lộn xộn |
| `python check_html.py` | [CMD] | Kiểm tra HTML hợp lệ | Khi nghi có lỗi |

## 3. Dev & test

| Lệnh | Loại | Làm gì |
|---|---|---|
| `pip install -r requirements.txt` | [CMD] | Cài thư viện |
| `pytest tests/` | [CMD] | Chạy test suite |
| *"Thêm nguồn crawl [URL]"* | [AGENT] | Agent sửa config, thêm nguồn |

---

**Output:** `Trang web FPT/*.xlsx`
