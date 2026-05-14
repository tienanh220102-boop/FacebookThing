#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FPT University News Crawler v2
Thu thap tin tuc tu website chinh thuc Dai hoc FPT qua RSS feed.
Chay moi ngay 1 lan luc 7h sang (Gio VN), xuat bao cao Excel.
"""
import json, os, re, time
from datetime import datetime, timezone, timedelta
from email.utils import parsedate_to_datetime
import xml.etree.ElementTree as ET

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

# ── Cau hinh ──────────────────────────────────────────────────
DATA_FILE = 'data/fpt_data.json'
VN_TZ     = timezone(timedelta(hours=7))

# RSS feeds chinh thuc cua Dai hoc FPT
FEEDS = {
    'Tin tuc':    'https://daihoc.fpt.edu.vn/feed/',
    'Tuyen sinh': 'https://daihoc.fpt.edu.vn/chuyen-muc/thong-bao-tuyen-sinh/feed/',
    'Su kien':    'https://daihoc.fpt.edu.vn/event/feed/',
}

# So gio lay lui (0 = tat ca, 24 = chi lay 24h gan nhat)
HOURS_LOOKBACK = 0

# ── Lay RSS ───────────────────────────────────────────────────
def strip_html(text):
    return re.sub(r'<[^>]+>', '', text or '').strip()

def fetch_feed(name, url, since_ts=None):
    """Lay bai viet tu mot RSS feed."""
    try:
        resp = requests.get(url, timeout=15, headers={'User-Agent': 'Mozilla/5.0'})
        resp.raise_for_status()
        root = ET.fromstring(resp.content)
    except Exception as e:
        print(f'  Loi khi lay [{name}]: {e}')
        return []

    items = []
    for item in root.findall('.//item'):
        title   = strip_html(item.findtext('title') or '')
        link    = (item.findtext('link') or '').strip()
        pub_raw = (item.findtext('pubDate') or '').strip()
        desc    = strip_html(item.findtext('description') or '')[:500]

        try:
            dt_utc = parsedate_to_datetime(pub_raw)
            dt_vn  = dt_utc.astimezone(VN_TZ)
            created_display = dt_vn.strftime('%d/%m/%Y %H:%M')
            created_ts      = dt_vn.timestamp()
        except Exception:
            created_display = pub_raw
            created_ts      = 0

        if since_ts and created_ts > 0 and created_ts < since_ts:
            continue

        items.append({
            'source':     name,
            'title':      title,
            'link':       link,
            'created':    created_display,
            'created_ts': created_ts,
            'summary':    desc,
        })

    return items

# ── Xuat Excel ────────────────────────────────────────────────
HDR_FILL = PatternFill(start_color='E85B2A', end_color='E85B2A', fill_type='solid')
HDR_FONT = Font(bold=True, color='FFFFFF')
TOP_FILL = PatternFill(start_color='FFF3CD', end_color='FFF3CD', fill_type='solid')
CTR      = Alignment(horizontal='center', vertical='center', wrap_text=True)

def _header(ws, cols):
    ws.append(cols)
    for cell in ws[ws.max_row]:
        cell.font = HDR_FONT; cell.fill = HDR_FILL; cell.alignment = CTR

def _autowidth(ws, max_w=60):
    for col in ws.columns:
        w = max((len(str(c.value or '')) for c in col), default=10)
        ws.column_dimensions[col[0].column_letter].width = min(w + 3, max_w)

def export_excel(all_articles, filename):
    """Tao file Excel 3 sheet: Tong quan / Tat ca bai / 20 bai moi nhat."""
    wb = openpyxl.Workbook()

    # ── Sheet 1: Tong quan nguon ──────────────────────────────
    ws1 = wb.active
    ws1.title = 'Tong quan nguon'
    _header(ws1, ['Nguon tin', 'So bai viet', 'Bai moi nhat'])
    sources = {}
    for a in all_articles:
        sources.setdefault(a['source'], []).append(a)
    for src, arts in sources.items():
        latest = sorted(arts, key=lambda x: x['created_ts'], reverse=True)
        ws1.append([src, len(arts), latest[0]['created'] if latest else ''])
    _autowidth(ws1)

    # ── Sheet 2: Tat ca bai viet ──────────────────────────────
    ws2 = wb.create_sheet('Tat ca bai viet')
    _header(ws2, ['Nguon', 'Thoi gian', 'Tieu de', 'Tom tat', 'Link'])
    ws2.column_dimensions['C'].width = 55
    ws2.column_dimensions['D'].width = 60
    for a in sorted(all_articles, key=lambda x: x['created_ts'], reverse=True):
        ws2.append([a['source'], a['created'], a['title'], a['summary'], a['link']])
    _autowidth(ws2)

    # ── Sheet 3: 20 bai moi nhat ──────────────────────────────
    ws3 = wb.create_sheet('20 bai moi nhat')
    _header(ws3, ['STT', 'Nguon', 'Thoi gian', 'Tieu de', 'Tom tat', 'Link'])
    ws3.column_dimensions['D'].width = 55
    ws3.column_dimensions['E'].width = 60
    top20 = sorted(all_articles, key=lambda x: x['created_ts'], reverse=True)[:20]
    for i, a in enumerate(top20, 1):
        ws3.append([i, a['source'], a['created'], a['title'], a['summary'], a['link']])
        if i <= 3:
            for cell in ws3[ws3.max_row]:
                cell.fill = TOP_FILL
    _autowidth(ws3)

    wb.save(filename)
    print(f'  Xuat thanh cong: {filename}')

# ── Main ──────────────────────────────────────────────────────
def main():
    now_vn = datetime.now(VN_TZ)
    if HOURS_LOOKBACK > 0:
        since_ts = (now_vn - timedelta(hours=HOURS_LOOKBACK)).timestamp()
        lookback_text = f'{HOURS_LOOKBACK}h gan nhat'
    else:
        since_ts = None
        lookback_text = 'tat ca bai (khong gioi han thoi gian)'

    print(f'=== FPT University Crawler v2 -- {now_vn.strftime("%d/%m/%Y %H:%M")} (Gio VN) ===')
    print(f'Thu thap {lookback_text}...')

    all_articles = []

    for name, url in FEEDS.items():
        print(f'\n[{name}] Dang lay du lieu...')
        articles = fetch_feed(name, url, since_ts)
        print(f'  Lay duoc {len(articles)} bai')
        all_articles.extend(articles)
        time.sleep(0.5)

    if not all_articles:
        print('\nKhong co bai viet nao. Kiem tra ket noi mang hoac RSS feed.')
        return

    os.makedirs('data', exist_ok=True)
    date_str   = now_vn.strftime('%Y%m%d')
    excel_file = f'data/fpt_news_{date_str}.xlsx'
    print(f'\nXuat Excel...')
    export_excel(all_articles, excel_file)

    with open(DATA_FILE, 'w', encoding='utf-8') as f:
        json.dump({
            'crawled_at':   now_vn.isoformat(),
            'total':        len(all_articles),
            'articles':     all_articles,
        }, f, ensure_ascii=False, indent=2)

    print(f'\n=== Hoan thanh. {len(all_articles)} bai tu {len(FEEDS)} nguon ===')

if __name__ == '__main__':
    main()
