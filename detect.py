import html as htmllib
import re

from curl_cffi import requests as creq

URL = "https://www.ctflife.com.hk/tc/support/important-information/fulfillment-ratios-dividends"
WORDS = ["分紅實現率", "總價值比率", "總現金價值", "週年紅利", "周年紅利", "復歸紅利", "終期紅利",
         "終期分紅", "特別紅利", "保單年度", "不適用", "美元", "港元", "人民幣"]
LINK = r"fulfil|ratio|dividend|total-value|\.pdf|\.xlsx|\.json|/api/"


def get(url):
    r = creq.get(url, impersonate="chrome", timeout=60)
    return r.status_code, r.content.decode("utf-8", "replace")


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


st, raw = get(URL)
title = re.search(r"<title[^>]*>(.*?)</title>", raw, flags=re.S | re.I)
tables = re.findall(r"<table\b.*?</table>", raw, flags=re.S | re.I)
text = clean(raw)
print("=== 總頁 ===")
print(f"狀態 {st}｜長度 {len(raw)}｜標題 {clean(title.group(1))[:60] if title else '沒有'}")
print(f"表格 {len(tables)} 個｜表格內百分比 {sum(len(re.findall(r'[0-9]%', clean(t))) for t in tables)} 個"
      f"｜正文百分比 {len(re.findall(r'[0-9]%', text))} 個")
print("關鍵字：" + "｜".join(f"{w} {raw.count(w)}" for w in WORDS))
print(f"NEXT_DATA {'有' if '__NEXT_DATA__' in raw else '沒有'}｜NUXT {'有' if '__NUXT__' in raw else '沒有'}"
      f"｜JSON script {len(re.findall(r'<script[^>]+application/json', raw, flags=re.I))} 個"
      f"｜iframe {len(re.findall(r'<iframe', raw, flags=re.I))} 個｜select {len(re.findall(r'<select', raw, flags=re.I))} 個")

print("=== 下拉選項（首 20 個）===")
opts = [clean(o)[:30] for o in re.findall(r"<option\b[^>]*>(.*?)</option>", raw, flags=re.S | re.I)]
print(f"共 {len(opts)} 個｜" + "｜".join(opts[:20]))

print("=== iframe 及相關連結 ===")
for s in re.findall(r'<iframe[^>]+src="([^"]+)"', raw, flags=re.I)[:5]:
    print("  iframe " + s[:150])
links = list(dict.fromkeys(
    h for h in re.findall(r'(?:href|src)="([^"]+)"', raw)
    if re.search(LINK, h, flags=re.I) and not re.search(r"\.(css|js|png|jpg|svg|ico|woff2?)(\?|$)", h, flags=re.I)))
print(f"相關連結 {len(links)} 個")
for h in links[:40]:
    print("  " + h[:150])
urls = list(dict.fromkeys(re.findall(r'https?:[\\/]+[^"\'\s<>]+?(?:\.pdf|\.json|\.xlsx|/api/[^"\'\s<>]*)', raw)))
print(f"原文內的 PDF/JSON/API 網址 {len(urls)} 個")
for h in urls[:20]:
    print("  " + h[:150])

print("=== 表格樣本 ===")
pos = 0
for n, t in enumerate(tables[:4], 1):
    p = raw.find(t, pos)
    pos = p + len(t)
    rows = re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)
    print(f"-- 表 {n}｜{len(rows)} 行｜表前文字：{clean(raw[max(0, p - 1500):p])[-80:]}")
    for tr in rows[:8]:
        cells = re.findall(r"<t[hd]\b([^>]*)>(.*?)</t[hd]>", tr, flags=re.S | re.I)
        out = []
        for attr, c in cells[:16]:
            span = "".join(f"[{k[0]}{v}]" for k, v in re.findall(r'(rowspan|colspan)="?(\d+)', attr, flags=re.I) if v != "1")
            out.append((clean(c)[:16] or "∅") + span)
        print(f"   ({len(cells)}格) " + " | ".join(out))
if len(tables) > 4:
    sizes = [len(re.findall(r"<tr\b", t, flags=re.I)) for t in tables]
    print(f"全部表格行數：{sizes[:80]}")
if not tables:
    m = re.search(r"\d{2,3}(?:\.\d+)?\s*%", text)
    p = m.start() if m else text.find("分紅實現率")
    print("沒有表格，正文樣本：")
    print("  " + (text[max(0, p - 400):p + 400] if p >= 0 else text[:800]))
print("=== 完 ===")
