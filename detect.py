import html as htmllib
import re
from collections import Counter

from curl_cffi import requests as creq

BASE = "https://www.ctflife.com.hk/tc/support/important-information/"
PAGES = [("分紅實現率", BASE + "fulfillment-ratios-dividends"),
         ("總現金價值比率", BASE + "total-cash-value-ratio")]
YEAR = re.compile(r"^\s*\d{1,2}\+?\s*[（(]")
PCT = re.compile(r"^\d+(?:\.\d+)?\s*%$")
SPLIT = r"</(?:p|div|h\d|li|option|span|strong|b|td|th|section|select)>|<br\s*/?>"


def get(url):
    r = creq.get(url, impersonate="chrome", timeout=60)
    return r.status_code, r.content.decode("utf-8", "replace")


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def flat(s):
    return re.sub(r"\s+", " ", s)


def parse(t):
    rows = []
    for tr in re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I):
        cells = []
        for attr, c in re.findall(r"<t[hd]\b([^>]*)>(.*?)</t[hd]>", tr, flags=re.S | re.I):
            span = "".join(f"[{k[0].lower()}{v}]" for k, v in
                           re.findall(r'(rowspan|colspan)\s*=\s*"?(\d+)', attr, flags=re.I) if v != "1")
            cells.append((clean(c), span))
        rows.append(cells)
    return rows


def ycount(r):
    return sum(1 for c, _ in r if YEAR.match(c))


def fmt(c, n=12):
    return "｜".join(f"{k or '∅'}×{v}" for k, v in c.most_common(n)) or "沒有"


def show(t, n=10):
    for r in t[:n]:
        print(f"   ({len(r)}格) " + " | ".join((c[:14] or "∅") + s for c, s in r[:16]))
    if len(t) > n:
        print(f"   …其餘 {len(t) - n} 行略")


for label, url in PAGES:
    st, raw = get(url)
    title = re.search(r"<title[^>]*>(.*?)</title>", raw, flags=re.S | re.I)
    spans = [(m.start(), m.end()) for m in re.finditer(r"<table\b.*?</table>", raw, flags=re.S | re.I)]
    tabs = [parse(raw[a:b]) for a, b in spans]
    yidx = [i for i, t in enumerate(tabs) if any(ycount(r) >= 5 for r in t)]
    print(f"=== {label} ===")
    print(f"狀態 {st}｜長度 {len(raw)}｜標題 {clean(title.group(1))[:50] if title else '沒有'}")
    print(f"表格 {len(tabs)} 個｜年度表 {len(yidx)} 個｜內容不重複的年度表 {len({repr(tabs[i]) for i in yidx})} 個")
    for attr, body in re.findall(r"<select\b([^>]*)>(.*?)</select>", raw, flags=re.S | re.I):
        opts = re.findall(r"<option\b([^>]*)>(.*?)</option>", body, flags=re.S | re.I)
        print(f"select {flat(attr)[:60]}｜選項 {len(opts)} 個｜" +
              "｜".join(f"{clean(c)[:16]}{flat(a)[:30]}" for a, c in opts[:3]))

    nrow, ncell, ycols, lastcol, c0, c1, words = (Counter() for _ in range(7))
    pct, odd = 0, []
    for i in yidx:
        t = tabs[i]
        h = max(k for k, r in enumerate(t) if ycount(r) >= 5)
        ycols[len(t[h])] += 1
        lastcol[t[h][-1][0][:16]] += 1
        nrow[len(t) - h - 1] += 1
        for r in t[h + 1:]:
            ncell[len(r)] += 1
            full = len(r) == len(t[h]) + 2
            if (not full or any(s for _, s in r)) and len(odd) < 5:
                odd.append(f"表{i + 1} ({len(r)}格) " + " | ".join((c[:10] or "∅") + s for c, s in r[:14]))
            if full:
                c0[r[0][0][:14]] += 1
                c1[r[1][0][:14]] += 1
            for c, _ in (r[2:] if full else r):
                if PCT.match(c):
                    pct += 1
                else:
                    words[c[:18]] += 1
    print(f"-- 統計｜百分比格 {pct} 個")
    print("每表資料行數：" + fmt(nrow))
    print("每行格數：" + fmt(ncell) + "｜年份欄數：" + fmt(ycols) + "｜最後一欄：" + fmt(lastcol, 4))
    print("第 1 格（類別）：" + fmt(c0, 15))
    print("第 2 格（貨幣）：" + fmt(c1, 15))
    print("非百分比的值：" + fmt(words, 20))
    print(f"格數特別或有合併的資料行樣本 {len(odd)} 個")
    for o in odd:
        print("   " + o)

    print("-- 表格上方文字（首 8 個年度表，最後兩段）")
    for i in yidx[:8]:
        seg = raw[spans[i - 1][1] if i else 0:spans[i][0]]
        parts = [p for p in (clean(x) for x in re.split(SPLIT, seg, flags=re.I)) if p]
        print(f"   表{i + 1}：" + " ‖ ".join(p[:70] for p in parts[-2:]))
    for i in (yidx[:2] + yidx[-1:] if len(yidx) > 2 else yidx):
        seg = raw[spans[i - 1][1] if i else 0:spans[i][0]]
        print(f"-- 表{i + 1} 前的原始碼")
        print("   " + flat(seg)[-450:])

    if yidx:
        big = max(yidx, key=lambda i: len(tabs[i]))
        for i in dict.fromkeys([yidx[0], big]):
            print(f"-- 年度表樣本：表{i + 1}｜{len(tabs[i])} 行")
            show(tabs[i])
    others = [i for i in range(len(tabs)) if i not in yidx]
    print(f"-- 非年度表 {len(others)} 個")
    for i in others[:4]:
        print(f"   表{i + 1}｜{len(tabs[i])} 行")
        show(tabs[i], 3)
print("=== 完 ===")
