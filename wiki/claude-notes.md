# Claude Notes — Facebook Crawler (FPT University)

## Kiến trúc script

Tên dự án hơi misleading — thực ra là **FPT University Website Crawler**, không phải Facebook.

- **`scripts/fb_crawler.py`**: cào RSS hàng ngày, output Excel → `Trang web FPT/fpt_news_YYYYMMDD.xlsx`
- **`scripts/lich_su.py`**: cào toàn bộ lịch sử (1 lần duy nhất), output → `Trang web FPT/fpt_lichsu.xlsx` và `fpt_lichsu.json`
- **`scripts/sync_xlsx.py`**: đồng bộ `fpt_lichsu.json` → `fpt_lichsu_new.xlsx` (dùng khi JSON đã update nhưng chưa xuất Excel)
- **`scripts/cap_nhat_tom_tat.py`**: cập nhật excerpt/summary qua Gemini API — đọc `fpt_lichsu.json`, ghi lại
- **`scripts/check_html.py`**: kiểm tra cấu trúc HTML của website (không ghi file)

## Thư mục data quan trọng

- **`Trang web FPT/`**: nơi lưu TẤT CẢ output (xlsx + json) — đây là convention của project, KHÔNG thay đổi
- `data/` folder trong template structure có nhưng KHÔNG dùng cho project này (dùng `Trang web FPT/` thay thế)

## Quirks quan trọng

- **Relative paths `'Trang web FPT/...'`** vẫn đúng khi chạy từ project root (`python scripts/fb_crawler.py`)
- **`os.makedirs('Trang web FPT', exist_ok=True)`** trong fb_crawler.py — tự tạo thư mục nếu chưa có
- **`cap_nhat_tom_tat.py` cần `GEMINI_API_KEY`** — các script khác không cần API key
- **`lich_su.py` chỉ chạy 1 lần** — sau đó dùng `sync_xlsx.py` để cập nhật Excel từ JSON

## Bat files

- `Chay.bat` → `python scripts\fb_crawler.py` (cào hàng ngày)
- `Lay lich su.bat` → `python scripts\lich_su.py` (cào lịch sử 1 lần)

## Lịch sử thay đổi

- 2026-06: Chuyển tất cả 5 scripts từ root → `scripts/`; bat files updated tương ứng
- Relative paths `Trang web FPT/...` KHÔNG thay đổi — vẫn đúng khi chạy từ project root
