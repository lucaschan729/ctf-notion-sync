import datetime
import html as htmllib
import os
import re
import sys
import time
from collections import Counter

import requests
from curl_cffi import requests as creq
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

BASE = "https://www.ctflife.com.hk/tc/support/important-information/"
PAGES = [BASE + "fulfillment-ratios-dividends", BASE + "total-cash-value-ratio"]
LABELS = [str(i) for i in range(1, 11)] + ["10+"]
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ASCII_NAME = "CTF_fulfillment_auto.xlsx"
NAVY = "1F4E78"
TYPE_COLORS = {"週年紅利": "1F4E78", "終期紅利": "C55A11", "復歸紅利": "548235",
               "終期分紅": "7F6000", "特別紅利": "833C0B"}
YEAR = re.compile(r"^\s*(\d{1,2}\+?)\s*[（(]\s*([^）)]*)")
SEP = r"\s*[-–—－]\s*"
KIND_LABEL = {"tcv": "總現金價值比率", "div": "分紅實現率"}
DESC = {
    "tcv": "總現金價值比率＝實際總現金價值 ÷ 銷售時說明的總現金價值。100% 即與銷售時一致，高於 100% 即比銷售時多。定義以周大福人壽網站的說明為準。",
    "div": "分紅實現率＝實際派發的非保證紅利 ÷ 銷售時說明的金額。100% 即與預期一致，高於 100% 即比預期多。定義以周大福人壽網站的說明為準。",
}


def get(url):
    err = ""
    for _ in range(3):
        try:
            r = creq.get(url, impersonate="chrome", timeout=60)
            if r.status_code == 200:
                return r.content.decode("utf-8", "replace")
            err = f"狀態 {r.status_code}"
        except Exception as e:
            err = str(e)[:80]
        time.sleep(3)
    raise RuntimeError(err)


def clean(s):
    s = re.sub(r"<(script|style)\b.*?</\1>", " ", str(s or ""), flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", htmllib.unescape(s)).strip()


def span(attr, key):
    m = re.search(key + r"\s*=\s*[\"']?(\d+)", attr, flags=re.I)
    return max(1, int(m.group(1))) if m else 1


def grid(table):
    rows, pend = [], {}
    for tr in re.findall(r"<tr\b.*?</tr>", table, flags=re.S | re.I):
        cells = re.findall(r"<t[hd]\b([^>]*)>(.*?)</t[hd]>", tr, flags=re.S | re.I)
        row, c, k = [], 0, 0
        while k < len(cells) or c in pend:
            if c in pend:
                txt, left = pend[c]
                row.append(txt)
                if left <= 1:
                    del pend[c]
                else:
                    pend[c] = (txt, left - 1)
                c += 1
                continue
            attr, body = cells[k]
            k += 1
            txt = clean(body)
            rs, cs = span(attr, "rowspan"), span(attr, "colspan")
            for _ in range(cs):
                row.append(txt)
                if rs > 1:
                    pend[c] = (txt, rs - 1)
                c += 1
        rows.append(row)
    return rows


def parse_title(t):
    for m in re.finditer(SEP, t):
        seg = re.split(SEP, t[m.end():])[0].strip()
        if seg and "」" not in seg and re.search(r"^分紅|^萬用|保險$", seg):
            return t[:m.start()].strip(), seg
    return t.strip(), ""


def to_val(s):
    m = re.search(r"(\d+(?:\.\d+)?)\s*[%％]", s)
    return round(float(m.group(1)) / 100, 6) if m else "/"


def name_key(name):
    return re.sub(r"\s+", "", name)


def parse_page(raw, years, seen, st):
    recs, title, pos = [], "", 0
    for m in re.finditer(r"<table\b.*?</table>", raw, flags=re.S | re.I):
        seg = raw[pos:m.start()]
        pos = m.end()
        ps = [x for x in (clean(p) for p in
                          re.findall(r"<p\b[^>]*fzBold[^>]*>(.*?)</p>", seg, flags=re.S | re.I)) if x]
        fresh = bool(ps)
        if fresh:
            title = ps[-1]
        g = grid(m.group(0))
        hs = [k for k, r in enumerate(g) if sum(1 for c in r if YEAR.match(c)) >= 5]
        if not hs:
            continue
        st["tables"] += 1
        if not fresh:
            st["notitle"] += 1
        h = hs[-1]
        head = g[h]
        colmap = {}
        for i, c in enumerate(head):
            ym = YEAR.match(c)
            if not ym:
                continue
            lab = ym.group(1)
            if lab.endswith("+") or int(lab) > 10:
                lab = "10+"
            if lab in colmap.values():
                continue
            colmap[i] = lab
            years.setdefault(lab, re.sub(r"\s+", "", ym.group(2)))
        first = min(colmap)
        htext = " ".join(dict.fromkeys(c for r in g[:h] for c in r))
        rm = re.search(r"(\d{4})\s*年度", htext)
        if rm:
            st["ryear"][rm.group(1)] += 1
        name, cat = parse_title(title)
        name = name or "未識別產品"
        sig = repr((name_key(name), g[h + 1:]))
        if sig in seen:
            st["dups"] += 1
            continue
        seen.add(sig)
        for r in g[h + 1:]:
            if len(r) <= first:
                st["notes"] += 1
                continue
            r = r + [""] * (len(head) - len(r))
            if len(set(r)) == 1:
                st["notes"] += 1
                continue
            typ = r[0] if first >= 1 else ""
            if typ.startswith(("*", "＊")) or len(typ) > 10:
                st["notes"] += 1
                continue
            if "保單價值" in typ:
                st["pv"] += 1
                continue
            cur = r[first - 1] if first >= 2 else ""
            vals = {}
            for i, lab in colmap.items():
                vals[lab] = to_val(r[i])
                if vals[lab] == "/":
                    st["words"][r[i][:12] or "∅"] += 1
            if "總現金價值" in typ:
                kind = "tcv"
            elif "紅利" in typ or "分紅" in typ:
                kind = "div"
            else:
                kind = "tcv" if "總現金價值" in htext else "div"
            recs.append({"name": name, "cat": cat, "typ": typ or "其他", "cur": cur,
                         "vals": vals, "kind": kind})
    return recs


def avg(rs, lab):
    v = [r["vals"][lab] for r in rs if isinstance(r["vals"].get(lab), float)]
    return (sum(v) / len(v), len(v)) if v else (None, 0)


def suffix_avg(rs, i):
    v = [r["vals"][lab] for r in rs for lab in LABELS[i:] if isinstance(r["vals"].get(lab), float)]
    return sum(v) / len(v) if v else None


def products(recs):
    out = {}
    for x in recs:
        p = out.setdefault(name_key(x["name"]), {"name": x["name"], "cat": x["cat"], "rows": 0, "vals": []})
        p["rows"] += 1
        if not p["cat"] and x["cat"]:
            p["cat"] = x["cat"]
        p["vals"] += [v for v in x["vals"].values() if isinstance(v, float)]
    return list(out.values())


def numbers(recs):
    return [v for x in recs for v in x["vals"].values() if isinstance(v, float)]


def build_excel(sheets, years, year, path):
    wb = Workbook()
    ov = wb.active
    ov.title = "概覽"

    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    sep_border = Border(bottom=Side(style="thin", color="5B9BD5"))
    sep_fill = PatternFill("solid", fgColor="DCE6F1")
    head_fill = PatternFill("solid", fgColor=NAVY)
    green = PatternFill("solid", fgColor="E2F0D9")
    red = PatternFill("solid", fgColor="FCE4E4")
    white = PatternFill("solid", fgColor="FFFFFF")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    grey = Font(size=10, color="595959")
    today = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d")

    def pct_fmt(v):
        return "0%" if abs(v * 100 - round(v * 100)) < 1e-6 else "0.0%"

    def head_row(sh, row, heads, start=1, height=34):
        for c, h in enumerate(heads, start):
            cell = sh.cell(row=row, column=c, value=h)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = head_fill
            cell.alignment = center
            cell.border = border
        sh.row_dimensions[row].height = height

    def ratio_cell(cell, v, fmt=None):
        cell.alignment = center
        cell.border = border
        cell.fill = white
        if isinstance(v, float):
            cell.value = v
            cell.number_format = fmt or pct_fmt(v)
            if v >= 1:
                cell.fill = green
                cell.font = Font(color="375623")
            elif v < 0.8:
                cell.fill = red
                cell.font = Font(color="C00000")
            else:
                cell.font = Font(color="262626")
        else:
            cell.value = "/"
            cell.font = Font(color="BFBFBF")

    def text_row(sh, row, text, ncols, font, height=20, start=1):
        sh.merge_cells(start_row=row, start_column=start, end_row=row, end_column=ncols)
        cell = sh.cell(row=row, column=start, value=text)
        cell.font = font
        cell.alignment = left
        sh.row_dimensions[row].height = height

    def label_cell(sh, row, col, text, color="262626"):
        cell = sh.cell(row=row, column=col, value=text)
        cell.font = Font(bold=True, color=color)
        cell.alignment = center
        cell.border = border

    for sname, kind, recs in sheets:
        ws = wb.create_sheet(sname)
        is_div = kind == "div"
        label = KIND_LABEL[kind]
        meta = ["產品名稱", "產品種類", "貨幣"] + (["紅利類型"] if is_div else [])
        nm = len(meta)
        ncol = nm + len(LABELS)
        text_row(ws, 1, f"周大福{label}（報告年度 {year}）", ncol, Font(bold=True, size=14, color=NAVY), 28)
        text_row(ws, 2, DESC[kind], ncol, grey)
        text_row(ws, 3, "綠色＝100% 或以上；紅色＝低於 80%；「/」＝沒有數字（網站顯示已停售、未推出、沒有分紅、沒有保單終結等）。", ncol, grey)
        text_row(ws, 4, f"資料來源：周大福人壽網站（報告年度 {year}）｜更新日期：{today}", ncol, grey)
        head_row(ws, 5, meta + [f"保單年度{lab}\n{years.get(lab, '')}".strip() for lab in LABELS])
        r, prev = 6, None
        for rec in recs:
            key = name_key(rec["name"])
            if prev is not None and key != prev:
                for c in range(1, ncol + 1):
                    cell = ws.cell(row=r, column=c)
                    cell.fill = sep_fill
                    cell.border = sep_border
                ws.row_dimensions[r].height = 5
                r += 1
            prev = key
            row = [rec["name"], rec["cat"], rec["cur"]] + ([rec["typ"]] if is_div else [])
            for c, v in enumerate(row, 1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = border
                cell.fill = white
                cell.alignment = left if c == 1 else center
                if c == 1:
                    cell.font = Font(bold=True, color="262626")
                elif c == 2:
                    cell.font = Font(size=10, color="595959")
                elif c == 3:
                    cell.font = Font(bold=True, color="7030A0")
                else:
                    cell.font = Font(color=TYPE_COLORS.get(rec["typ"], "404040"))
            for c, lab in enumerate(LABELS, nm + 1):
                ratio_cell(ws.cell(row=r, column=c), rec["vals"].get(lab, "/"))
            n = len(rec["name"])
            ws.row_dimensions[r].height = 21 if n <= 22 else 36 if n <= 44 else 52
            r += 1
        ws.auto_filter.ref = f"A5:{get_column_letter(ncol)}{r - 1}"
        ws.freeze_panes = "B6"

        r += 1
        ws.cell(row=r, column=nm, value="平均值摘要").font = Font(bold=True, size=12, color=NAVY)
        r += 1
        head_row(ws, r, ["分紅類型" if is_div else label, "所有年份平均值"] +
                 [f"持滿{i}年平均值" for i in range(2, 10)] + ["第10年平均值", "持滿10年+平均值"], start=nm)
        if is_div:
            types = list(dict.fromkeys(x["typ"] for x in recs))
            groups = [("所有分紅", recs)] + [(g, [x for x in recs if x["typ"] == g]) for g in types]
        else:
            groups = [("所有產品", recs)]
        for g, rs in groups:
            r += 1
            label_cell(ws, r, nm, g, TYPE_COLORS.get(g, "262626"))
            for i in range(len(LABELS)):
                a = suffix_avg(rs, i)
                ratio_cell(ws.cell(row=r, column=nm + 1 + i), a if a is not None else "/", "0.00%")
            ws.row_dimensions[r].height = 21
        r += 1
        head_row(ws, r, ["當年平均值"] + [years.get(lab, lab) for lab in LABELS], start=nm, height=30)
        r += 1
        label_cell(ws, r, nm, groups[0][0])
        for i, lab in enumerate(LABELS):
            a, _ = avg(recs, lab)
            ratio_cell(ws.cell(row=r, column=nm + 1 + i), a if a is not None else "/", "0.00%")
        ws.row_dimensions[r].height = 21
        r += 1
        text_row(ws, r, "「持滿N年平均值」＝保單年度N及之後所有數字的平均；「當年平均值」＝該年度一欄所有數字的平均。平均值計入所有貨幣，為固定數字，不會隨篩選變動。",
                 ncol, Font(size=9, color="7F7F7F"), 30, start=nm)
        widths = [44, 20, 12] + ([12] if is_div else []) + [12] * len(LABELS)
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.print_title_rows = "5:5"

    # ---------- 概覽 ----------
    sec_font = Font(bold=True, size=12, color=NAVY)
    text_row(ov, 1, f"周大福分紅表現概覽（報告年度 {year}）", 6, Font(bold=True, size=14, color=NAVY), 28)
    text_row(ov, 2, f"資料來源：周大福人壽網站｜更新日期：{today}", 6, grey)
    ov.cell(row=4, column=1, value="資料概況").font = sec_font
    head_row(ov, 5, ["工作表", "產品數", "行數", "有數字的格數", "整體平均值", "低於 80% 的比例"], height=26)
    rr = 6
    for sname, kind, recs in sheets:
        nums = numbers(recs)
        n = len(nums) or 1
        row = [sname, len(products(recs)), len(recs), len(nums), sum(nums) / n, sum(1 for v in nums if v < 0.8) / n]
        for c, v in enumerate(row, 1):
            cell = ov.cell(row=rr, column=c, value=v)
            cell.border = border
            cell.alignment = center
            if c == 1:
                cell.hyperlink = f"#'{sname}'!A1"
                cell.font = Font(bold=True, color="0563C1", underline="single")
            elif c == 5:
                cell.number_format = "0.00%"
            elif c == 6:
                cell.number_format = "0.0%"
        ov.row_dimensions[rr].height = 22
        rr += 1

    rr += 1
    ov.cell(row=rr, column=1, value="按產品種類的整體平均值").font = sec_font
    rr += 1
    head_row(ov, rr, ["產品種類"] + [KIND_LABEL[k] for _, k, _ in sheets], height=26)
    prods = [products(recs) for _, _, recs in sheets]
    cats = list(dict.fromkeys(p["cat"] or "未識別種類" for ps in prods for p in ps))
    for g in cats:
        rr += 1
        cell = ov.cell(row=rr, column=1, value=g)
        cell.font = Font(bold=True, color="262626")
        cell.border = border
        cell.alignment = left
        for c, ps in enumerate(prods, 2):
            v = [x for p in ps if (p["cat"] or "未識別種類") == g for x in p["vals"]]
            ratio_cell(ov.cell(row=rr, column=c), sum(v) / len(v) if v else "/", "0.00%")
        ov.row_dimensions[rr].height = 22

    rr += 2
    ov.cell(row=rr, column=1, value="顏色說明").font = sec_font
    for v, t in [(1.0, "100% 或以上：與銷售時的說明一致或更好"),
                 (0.79, "低於 80%：明顯低於銷售時的說明"),
                 ("/", "沒有數字：網站顯示已停售、未推出、沒有分紅、沒有保單終結等")]:
        rr += 1
        ratio_cell(ov.cell(row=rr, column=1), v, "0%")
        text_row(ov, rr, t, 6, Font(size=10, color="404040"), 22, start=2)
    rr += 2
    ov.cell(row=rr, column=1, value="使用提示").font = sec_font
    for t in ["・點上方「資料概況」的工作表名稱，可直接跳到該工作表。",
              "・產品數按產品名稱計算，同名產品只算一次。",
              "・「保單年度10+」一欄即周大福網站的「11+（2014或之前）」。",
              "・網站上的「保單價值」行不屬於紅利，沒有列入。",
              "・每個工作表第 5 行有篩選箭嘴，可按產品種類、貨幣或紅利類型篩選。",
              "・產品之間有淺藍色分隔行；排序會打亂分隔行，建議只用篩選。",
              "・向右或向下捲動時，標題列和產品名稱會固定不動。",
              "・平均值摘要在每個工作表最底部。",
              "・手機預覽未必顯示全部格式，建議下載後用 Excel App 開啟。",
              "・過往表現不代表未來。"]:
        rr += 1
        text_row(ov, rr, t, 6, Font(size=10, color="404040"))
    ov.column_dimensions["A"].width = 30
    for c in "BCDEF":
        ov.column_dimensions[c].width = 18
    ov.sheet_view.showGridLines = False

    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def notion_upload(path, year):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    raw = os.environ.get("NOTION_DATABASE_ID", "").strip()
    if not token or not raw:
        print("未設定 Notion Secrets，略過上傳（只產生 Excel）")
        return True
    seg = raw.split("?")[0].rstrip("/").split("/")[-1].replace("-", "")
    dbid = seg[-32:].lower()
    if not re.fullmatch(r"[0-9a-f]{32}", dbid):
        print("Notion 上傳失敗｜NOTION_DATABASE_ID 格式不對，請貼上資料庫整頁連結")
        return False
    H = {"Authorization": f"Bearer {token}", "Notion-Version": "2025-09-03"}

    def call(method, url, what, **kw):
        r = requests.request(method, "https://api.notion.com/v1" + url, headers=H, timeout=120, **kw)
        if r.status_code >= 300:
            msg = re.sub(r"[0-9a-f]{8}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{4}-?[0-9a-f]{12}", "（已隱藏）", r.text[:300])
            raise RuntimeError(f"{what}：狀態 {r.status_code}｜{msg}")
        return r.json()

    try:
        try:
            db = call("GET", f"/databases/{dbid}", "讀取資料庫")
            ds = db["data_sources"][0]["id"]
        except RuntimeError as e:
            if "狀態 404" not in str(e):
                raise
            ds = dbid
        src = call("GET", f"/data_sources/{ds}", "讀取資料來源")
        tprop = next(k for k, v in src["properties"].items() if v.get("type") == "title")
        pages, cursor = {}, None
        while True:
            body = {"page_size": 100}
            if cursor:
                body["start_cursor"] = cursor
            res = call("POST", f"/data_sources/{ds}/query", "查詢年度頁面", json=body)
            for p in res.get("results", []):
                if p.get("archived") or p.get("in_trash"):
                    continue
                title = "".join(t.get("plain_text", "") for t in p["properties"][tprop]["title"]).strip()
                pages.setdefault(title, p["id"])
            cursor = res.get("next_cursor")
            if not res.get("has_more") or not cursor:
                break
        want = f"{year}年"
        pid = pages.get(want) or pages.get(str(year))
        if pid:
            print(f"Notion：找到「{want if want in pages else year}」頁面")
        else:
            page = call("POST", "/pages", "建立年度頁面", json={
                "parent": {"type": "data_source_id", "data_source_id": ds},
                "properties": {tprop: {"title": [{"text": {"content": want}}]}}})
            pid = page["id"]
            print(f"Notion：已建立「{want}」")
        cursor, old = None, []
        while True:
            q = "?page_size=100" + (f"&start_cursor={cursor}" if cursor else "")
            res = call("GET", f"/blocks/{pid}/children{q}", "讀取頁面內容")
            for b in res.get("results", []):
                if b.get("type") == "file":
                    nm = (b["file"].get("name") or "")
                    if "自動生成" in nm or ASCII_NAME in nm:
                        old.append(b["id"])
            cursor = res.get("next_cursor")
            if not res.get("has_more") or not cursor:
                break
        fname = os.path.basename(path)
        fu = call("POST", "/file_uploads", "建立檔案上傳", json={
            "mode": "single_part", "filename": fname, "content_type": XLSX})
        with open(path, "rb") as fh:
            call("POST", f"/file_uploads/{fu['id']}/send", "傳送檔案", files={"file": (ASCII_NAME, fh, XLSX)})
        call("PATCH", f"/blocks/{pid}/children", "放入頁面", json={"children": [{
            "object": "block", "type": "file",
            "file": {"type": "file_upload", "file_upload": {"id": fu["id"]}}}]})
        for bid in old:
            call("DELETE", f"/blocks/{bid}", "刪除舊自動檔")
        print(f"已把檔案放在頁面內容最底｜已替換舊自動檔 {len(old)} 個（手動檔不受影響）")
        return True
    except Exception as e:
        print(f"Notion 上傳失敗｜{e}")
        return False


def report(label, recs, by_type):
    nums = numbers(recs)
    print(f"-- {label}｜產品 {len(products(recs))} 個｜資料行 {len(recs)}｜數值格 {len(nums)}")
    print("   類別：" + "｜".join(f"{k}×{v}" for k, v in Counter(x["typ"] for x in recs).most_common()))
    print("   貨幣：" + "｜".join(f"{k or '∅'}×{v}" for k, v in Counter(x["cur"] for x in recs).most_common()))
    groups = [("全部", recs)]
    if by_type:
        groups += [(g, [x for x in recs if x["typ"] == g]) for g in dict.fromkeys(x["typ"] for x in recs)]
    for g, rs in groups:
        hold = [suffix_avg(rs, i) for i in range(len(LABELS))]
        print(f"   {g}（{len(rs)} 行／{len(numbers(rs))} 格）持滿平均｜" +
              "｜".join(f"{lab}:{a * 100:.2f}%" if a is not None else f"{lab}:/" for lab, a in zip(LABELS, hold)))
    parts = []
    for lab in LABELS:
        a, n = avg(recs, lab)
        parts.append(f"{lab}:{a * 100:.2f}%({n})" if a is not None else f"{lab}:/")
    print("   全部 當年平均｜" + "｜".join(parts))


def main():
    recs, years, seen = [], {}, set()
    st = {"tables": 0, "dups": 0, "notitle": 0, "notes": 0, "pv": 0,
          "words": Counter(), "ryear": Counter()}
    print("=== 讀取 ===")
    for url in PAGES:
        try:
            raw = get(url)
        except Exception as e:
            print(f"{url.rsplit('/', 1)[-1]}：讀取失敗（{e}）")
            continue
        got = parse_page(raw, years, seen, st)
        recs += got
        print(f"{url.rsplit('/', 1)[-1]}：長度 {len(raw)}｜新增資料行 {len(got)}")
        if {"tcv", "div"} <= {x["kind"] for x in recs}:
            break
    if not recs:
        print("讀不到任何資料，停止")
        sys.exit(1)
    year = st["ryear"].most_common(1)
    year = year[0][0] if year else str(datetime.date.today().year - 1)
    tcv = [x for x in recs if x["kind"] == "tcv"]
    div = [x for x in recs if x["kind"] == "div"]

    print("=== 各產品（總現金價值 行/格｜分紅 行/格）===")
    pt = {name_key(p["name"]): p for p in products(tcv)}
    pd = {name_key(p["name"]): p for p in products(div)}
    for k in dict.fromkeys(list(pt) + list(pd)):
        p = pt.get(k) or pd.get(k)
        a, b = pt.get(k), pd.get(k)
        print(f"  {p['name']}｜{p['cat'] or '未識別種類'}｜"
              f"{a['rows'] if a else 0}/{len(a['vals']) if a else 0}｜{b['rows'] if b else 0}/{len(b['vals']) if b else 0}")
    print("=== 核對 ===")
    print(f"報告年度 {year}｜年度表 {st['tables']} 個｜略過重複表格 {st['dups']} 個｜沒有標題的表 {st['notitle']} 個")
    print(f"略過備註行 {st['notes']} 行｜剔除保單價值 {st['pv']} 行")
    none = [p["name"] for p in products(recs) if not p["cat"]]
    print(f"資料行 {len(recs)}｜數值格 {len(numbers(recs))}｜未識別種類的產品 {len(none)} 個" +
          ("｜" + "、".join(none[:10]) if none else ""))
    print("年份欄：" + "｜".join(f"{lab}={years.get(lab, '?')}" for lab in LABELS))
    print("沒有數字的格：" + "｜".join(f"{k}×{v}" for k, v in st["words"].most_common(12)))
    if tcv:
        report("總現金價值比率", tcv, False)
    if div:
        report("分紅實現率", div, True)

    sheets = []
    if tcv:
        sheets.append((f"{year}總現金價值比率", "tcv", tcv))
    if div:
        sheets.append((f"{year}分紅實現率", "div", div))
    path = f"output/周大福總現金價值分紅實現率（{year}年）自動生成.xlsx"
    build_excel(sheets, years, year, path)
    print(f"已產生 {path}")
    if not notion_upload(path, year):
        sys.exit(1)


if __name__ == "__main__":
    main()
