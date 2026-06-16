# Hướng Dẫn Sử Dụng — Facebook Crawler

> Tài liệu cầm tay chỉ việc — không cần biết code.

---

## 1. Dự án này làm gì?

Công cụ này **tự động crawl bài viết từ trang web FPT University**, lưu vào file Excel để phân tích nội dung, theo dõi bài đăng, tóm tắt tự động bằng AI.

---

## 2. Chuẩn bị (làm 1 lần)

```
pip install -r requirements.txt
```

---

## 3. Hai nơi thao tác

| Nơi | Là gì | Dấu hiệu |
|---|---|---|
| **Terminal** | Cửa sổ gõ lệnh | Bạn gõ `python fb_crawler.py` |
| **Khung chat Claude** | Trò chuyện với AI | Bạn nói tiếng Việt |

---

## 4. Các thao tác chính

**Crawl bài mới:** Double-click `Chay.bat`

**Lấy toàn bộ lịch sử** (chạy 1 lần duy nhất, ~3-5 phút): Double-click `Lay lich su.bat`

**Cập nhật tóm tắt AI:**
```
python cap_nhat_tom_tat.py
```

---

## 5. Kết quả ở đâu?

| Mục đích | File / Nơi |
|---|---|
| **Bài viết đã crawl** | `Trang web FPT/` — file Excel |
| **Dữ liệu trung gian** | `data/` |

---

## 6. FAQ

- **Muốn crawl từ trang khác** → Nói với AI: *"Thêm nguồn crawl [URL]"*
- **File Excel bị lỗi** → Chạy `python check_html.py` để kiểm tra
- **Dữ liệu trùng lặp** → Chạy `python sync_xlsx.py` để deduplicate
- **Không biết làm gì tiếp** → Nói với AI: *"Dự án Facebook Crawler đang ở đâu?"*
