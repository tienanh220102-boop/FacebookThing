import requests, re, sys
sys.stdout.reconfigure(encoding='utf-8')

resp = requests.get('https://daihoc.fpt.edu.vn/chuyen-muc/tin-tuc/',
                    headers={'User-Agent': 'Mozilla/5.0'}, timeout=15)
html = resp.text

# Cat ra 1 bai viet mau de xem cau truc ben trong
start = html.find('<article class="news-list-item">')
end   = html.find('</article>', start) + len('</article>')
print(html[start:end])
