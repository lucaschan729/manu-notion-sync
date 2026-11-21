import collections
import datetime
import html as htmllib
import os
import re
import subprocess
import sys
import time

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi", "openpyxl", "requests"])
import requests
from curl_cffi import requests as creq
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

PAGE = "https://www.manulife.com.hk/zh-hk/individual/products/understanding-your-participating-policy/fulfillment-ratio.html"
ROOT = "https://tools.manulife.com.hk/"
INTRO = ROOT + "fulfillment-ratio-2024/html/zh/intro.php"
SETS = [("總現金價值比率", ROOT + "fulfillment-tcv-2024/html/ajax/ss1.php", True, "比率類型"),
        ("分紅實現率", ROOT + "fulfillment-ratio-2024/html/ajax/ss1.php", False, "紅利類型")]
HEAD = {"Referer": PAGE, "Origin": "https://www.manulife.com.hk", "X-Requested-With": "XMLHttpRequest"}
PCT = re.compile(r"(\d+(?:\.\d+)?)\s*%")
ITEM = re.compile(r"(selectedItem\w*)\(\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*\)")
TAB = re.compile(r"""getProductDetails\w*\(\s*['"]([^'"]*)['"]\s*\)""")
NO_CUR = ["", "null", "no"]
NAVY, GREY, DARK, PURPLE = "1F3864", "8C8C8C", "262626", "7030A0"
SEP_FILL = PatternFill("solid", fgColor="EDEDED")
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ROW_HEIGHT = 34
HEAD_ROW = 3
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def decode(content):
    s = content.decode("utf-8", "replace")
    s = re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), s).replace("\\/", "/")
    return htmllib.unescape(s)


def call(url, data=None):
    last = ""
    for i in range(3):
        try:
            if data is None:
                r = creq.get(url, impersonate="chrome", headers=HEAD, timeout=60)
            else:
                r = creq.post(url, data=data, impersonate="chrome", headers=HEAD, timeout=60)
            if r.status_code == 200:
                return decode(r.content)
            last = f"狀態 {r.status_code}"
        except Exception as e:
            last = str(e)[:80]
        time.sleep(2 * (i + 1))
    raise RuntimeError(last)


def post(url, name, cur, tab):
    return call(url, {"selected1": name, "selected2": cur, "tablabel": tab, "lang": "tc"})


def clean(s):
    s = re.sub(r"<sup\b.*?</sup>", "", s, flags=re.S | re.I)
    s = re.sub(r"<[^>]+>", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def value(s):
    m = PCT.search(s)
    if m:
        return float(m.group(1)) / 100
    return "不適用" if "不適用" in s else s


def parse(body):
    out = []
    for t in re.findall(r"<table\b.*?</table>", body, flags=re.S | re.I):
        rows = []
        for tr in re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I):
            row = []
            for attrs, inner in re.findall(r"<t[hd]\b([^>]*)>(.*?)</t[hd]>", tr, flags=re.S | re.I):
                m = re.search(r"colspan\s*=\s*[\"']?(\d+)", attrs, flags=re.I)
                row += [clean(inner)] * (int(m.group(1)) if m else 1)
            rows.append(row)
        flat = [c for row in rows for c in row]
        if any("保單生效日" in c for c in flat):
            continue
        title = next((c for c in flat if "申報年度之" in c), "")
        years, leads = [], []
        for row in rows:
            idx = [i for i, c in enumerate(row) if c.startswith("保單年度")]
            if len(idx) >= 2:
                years = [row[i] for i in idx]
                leads = row[:idx[0]]
                continue
            if not years or len(row) < len(years):
                continue
            vals = row[-len(years):]
            if not any(PCT.search(v) or "不適用" in v for v in vals) or any(len(v) > 12 for v in vals):
                continue
            out.append((title, leads, years, row[:-len(years)], vals))
    return out


def fetch(label, url, prods, rows, ycols, ysubs, log):
    seen, no_cur = set(), list(NO_CUR)
    for pi, (name, curs) in enumerate(prods, 1):
        groups = [no_cur] if curs == "no" else [[c] for c in curs.split("||")]
        for group in groups:
            body, used, err = None, None, "沒有表格"
            for cur in list(group):
                try:
                    b = post(url, name, cur, "null")
                except Exception as e:
                    err = str(e)[:80]
                    continue
                if parse(b):
                    body, used = b, cur
                    break
            shown = "不分貨幣" if curs == "no" else group[0].strip()
            if body is None:
                log["empty"].append(f"{label}｜{name}｜{shown}：{err}")
                continue
            if curs == "no" and no_cur[0] != used:
                no_cur.remove(used)
                no_cur.insert(0, used)
            tabs = list(dict.fromkeys(TAB.findall(body)))
            pages = [(tabs[0] if tabs else "", body)]
            if len(tabs) > 1:
                pages = []
                for tab in tabs:
                    try:
                        pages.append((tab, post(url, name, used, tab)))
                    except Exception as e:
                        log["empty"].append(f"{label}｜{name}｜{shown}｜{tab}：{str(e)[:80]}")
                    time.sleep(0.1)
            for tab, b in pages:
                parsed = parse(b)
                if not parsed:
                    log["empty"].append(f"{label}｜{name}｜{shown}｜{tab}：沒有表格")
                for title, leads, years, extras, vals in parsed:
                    m = re.search(r"(\d{4})\s*申報年度之(.+)", title)
                    kind = m.group(2).strip() if m else title
                    if m:
                        log["years"][m.group(1)] += 1
                    kind = re.sub(r"的?分紅實現率$", "", kind) or kind
                    head = leads[0] if leads else ""
                    period = re.sub(r"\s+", " ", tab).strip()
                    if not period:
                        pm = re.search(r"[（(]於(.+?)生效的保單[）)]", head)
                        period = pm.group(1).strip() if pm else ""
                    elif head and period.replace(" ", "") not in head.replace(" ", ""):
                        msg = f"{label}｜{name}｜{period}｜網站顯示：{head[:60]}"
                        if msg not in log["mismatch"]:
                            log["mismatch"].append(msg)
                    vd = {}
                    for y, v in zip(years, vals):
                        ym = re.match(r"(保單年度[^（(]*)(.*)", y)
                        key = ym.group(1).replace(" ", "")
                        if key not in ycols:
                            ycols.append(key)
                        ysubs.setdefault(key, ym.group(2).strip())
                        vd[key] = value(v)
                    row = {"name": name.strip(), "cat": leads[1] if len(leads) > 1 else "",
                           "period": period, "cur": (leads[2] if len(leads) > 2 else used).strip() or "—",
                           "kind": kind, "note": " ".join(x for x in extras if x and x not in leads), "vals": vd}
                    sig = (row["name"], period, row["cur"], kind, row["note"], tuple(vd.items()))
                    if sig not in seen:
                        seen.add(sig)
                        rows.append(row)
            time.sleep(0.1)
        if pi % 10 == 0 or pi == len(prods):
            print(f"  {label}：已處理 {pi}/{len(prods)} 個產品｜累計 {len(rows)} 行", flush=True)


def pct_format(v):
    return "0%" if abs(v * 100 - round(v * 100)) < 1e-6 else "0.0%"


def numbers(rows):
    return [v for r in rows for v in r["vals"].values() if isinstance(v, float)]


def mean(nums):
    return sum(nums) / len(nums) if nums else None


def averages(rows, ycols):
    return [mean([r["vals"][k] for r in rows if isinstance(r["vals"].get(k), float)]) for k in ycols]


def cumulative(rows, ycols):
    return [mean([r["vals"][k] for r in rows for k in ycols[i:] if isinstance(r["vals"].get(k), float)])
            for i in range(len(ycols))]


def hold_label(i, key):
    if i == 0:
        return "所有年份平均值"
    n = key.replace("保單年度", "")
    return f"持滿 {n[:-1]} 年以上平均值" if n.endswith("+") else f"持滿 {n} 年平均值"


def head_cell(c):
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = CENTER


def write_averages(ws, start, rows, ycols, nlead):
    last = nlead + len(ycols)
    kinds = list(dict.fromkeys(x["kind"] for x in rows))
    groups = [("所有分紅" if len(kinds) > 1 else "全部", rows)]
    if len(kinds) > 1:
        groups += [(k, [x for x in rows if x["kind"] == k]) for k in kinds]
    groups = [(t, s) for t, s in groups if numbers(s)]
    r = start
    c = ws.cell(row=r, column=1, value="平均值（固定數字，不隨篩選變化；只計算有百分比的格）")
    c.font = Font(bold=True, size=11, color=NAVY)
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last)
    ws.row_dimensions[r].height = 24
    r += 1
    tables = (("持滿年期平均值（該年度及之後合併計算）", [hold_label(i, k) for i, k in enumerate(ycols)], cumulative),
              ("當年平均值（各保單年度）", ycols, averages))
    for heading, heads, fn in tables:
        head_cell(ws.cell(row=r, column=1, value=heading))
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nlead)
        for j, h in enumerate(heads):
            head_cell(ws.cell(row=r, column=nlead + 1 + j, value=h))
        ws.row_dimensions[r].height = 46
        r += 1
        for title, sub in groups:
            c = ws.cell(row=r, column=1, value=title)
            c.font = Font(bold=True, size=10, color=DARK)
            c.alignment = Alignment(horizontal="left", vertical="center")
            c.border = BOX
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=nlead)
            for j, v in enumerate(fn(sub, ycols)):
                c = ws.cell(row=r, column=nlead + 1 + j, value=v if v is not None else "/")
                c.font = Font(bold=True, size=10)
                c.alignment = CENTER
                c.border = BOX
                if v is not None:
                    c.number_format = "0.00%"
            ws.row_dimensions[r].height = 24
            r += 1
        r += 1


def write_sheet(ws, title, rows, ycols, ysubs, kind_head):
    cols = [("產品名稱", 30, "name"), ("產品種類", 14, "cat"), ("保單生效時期", 22, "period"), ("保單貨幣", 9, "cur")]
    if len(set(r["kind"] for r in rows)) > 1:
        cols.append((kind_head, 11, "kind"))
    if any(r["note"] for r in rows):
        cols.append(("備註", 16, "note"))
    heads = [c[0] for c in cols] + [(k + "\n" + ysubs.get(k, "")).strip() for k in ycols]
    widths = [c[1] for c in cols] + [13] * len(ycols)
    c = ws.cell(row=1, column=1, value=title)
    c.font = Font(bold=True, size=14, color=NAVY)
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(heads))
    ws.row_dimensions[1].height = 28
    c = ws.cell(row=2, column=1, value="100% 或以上綠底，低於 80% 紅底｜「不適用」為宏利網站原文｜固定平均值在表格底部")
    c.font = Font(size=9, color=GREY)
    c.alignment = Alignment(vertical="center")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(heads))
    ws.row_dimensions[2].height = 18
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        head_cell(ws.cell(row=HEAD_ROW, column=i, value=h))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[HEAD_ROW].height = 62
    ri, prev = HEAD_ROW, None
    for r in rows:
        if r["name"] != prev:
            if prev is not None:
                ri += 1
                ws.cell(row=ri, column=1).fill = SEP_FILL
                ws.merge_cells(start_row=ri, start_column=1, end_row=ri, end_column=len(heads))
                ws.row_dimensions[ri].height = 9
            prev = r["name"]
        ri += 1
        for ci, (_, w, key) in enumerate(cols, 1):
            c = ws.cell(row=ri, column=ci, value=r[key])
            if key == "name":
                c.font = Font(bold=True, size=10, color=DARK)
            elif key == "cur":
                c.font = Font(bold=True, size=10, color=PURPLE)
            else:
                c.font = Font(size=10)
            c.alignment = Alignment(horizontal="left" if key in ("name", "period", "note") else "center",
                                    vertical="center", wrap_text=True)
            c.border = BOX
        for j, k in enumerate(ycols):
            v = r["vals"].get(k, "")
            c = ws.cell(row=ri, column=len(cols) + 1 + j, value=v)
            c.alignment = CENTER
            c.border = BOX
            if isinstance(v, float):
                c.number_format = pct_format(v)
                if v >= 0.9995:
                    c.fill = GREEN_FILL
                    c.font = Font(size=10, color="006100")
                elif v < 0.7995:
                    c.fill = RED_FILL
                    c.font = Font(size=10, bold=True, color="9C0006")
                else:
                    c.font = Font(size=10)
            else:
                c.font = Font(size=9, color=GREY)
        ws.row_dimensions[ri].height = ROW_HEIGHT
    ws.freeze_panes = f"B{HEAD_ROW + 1}"
    ws.auto_filter.ref = f"A{HEAD_ROW}:{get_column_letter(len(heads))}{max(ri, HEAD_ROW + 1)}"
    write_averages(ws, ri + 2, rows, ycols, len(cols))
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = f"{HEAD_ROW}:{HEAD_ROW}"


def write_overview(ws, sheets, year, now):
    ncol = 6
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 38
    for i in range(2, ncol + 1):
        ws.column_dimensions[get_column_letter(i)].width = 17

    def text(r, s, bold=False, size=10, height=None, color="000000"):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        c = ws.cell(row=r, column=1, value=s)
        c.font = Font(bold=bold, size=size, color=color)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        if height:
            ws.row_dimensions[r].height = height

    def table(r, heads, lines, fmts):
        for i, h in enumerate(heads, 1):
            head_cell(ws.cell(row=r, column=i, value=h))
        ws.row_dimensions[r].height = 30
        for line in lines:
            r += 1
            for i, v in enumerate(line, 1):
                c = ws.cell(row=r, column=i, value=v if v is not None else "/")
                c.border = BOX
                c.font = Font(size=10, bold=i == 1)
                c.alignment = Alignment(horizontal="left" if i == 1 else "center", vertical="center")
                if isinstance(v, float):
                    c.number_format = fmts.get(i, "0.00%")
            ws.row_dimensions[r].height = 22
        return r + 2

    def stat(title, sub):
        nums = numbers(sub)
        low = len([v for v in nums if v < 0.7995]) / len(nums) if nums else None
        return [title, len(set(x["name"] for x in sub)), len(sub), len(nums), mean(nums), low]

    text(1, f"宏利分紅表現概覽（報告年度 {year}）", True, 14, 28, NAVY)
    text(2, f"自動生成｜擷取時間：{now}（香港時間）｜資料來源：宏利香港網站")
    text(3, PAGE, color="0563C1")
    text(5, "資料概況與整體表現", True, 11, 22, NAVY)
    lines = []
    for label, rows, _ in sheets:
        lines.append(stat(f"{year} {label}", rows))
        kinds = list(dict.fromkeys(x["kind"] for x in rows))
        if len(kinds) > 1:
            lines += [stat(f"　{k}", [x for x in rows if x["kind"] == k]) for k in kinds]
    r = table(6, ["指標項目", "產品數", "行數", "有數字格數", "整體平均值", "低於 80% 的比例"], lines, {6: "0.0%"})
    cats = list(dict.fromkeys(x["cat"] for _, rows, _ in sheets for x in rows if x["cat"]))
    if cats:
        text(r, "按產品種類的整體平均值", True, 11, 22, NAVY)
        lines = [[cat] + [mean(numbers([x for x in rows if x["cat"] == cat])) for _, rows, _ in sheets] for cat in cats]
        r = table(r + 1, ["產品種類"] + [label for label, _, _ in sheets], lines, {})
    text(r, "定義與顏色說明", True, 11, 22, NAVY)
    notes = [
        "總現金價值比率：根據所有相關保單的實際總現金價值（包括保證及非保證成份），比對銷售時建議書上所示之預計總現金價值所計算的比率。",
        "分紅實現率：所有相關保單的實際累計非保證利益，比對銷售時建議書上所示之預計累計非保證利益所計算的比率。",
        "100% 或以上（綠色）：實際與投保時的預計一致或更高。低於 80%（紅色）：實際明顯低於投保時的預計。",
        "「不適用」為宏利網站原文，表示該保單年度沒有相關數字。",
    ]
    for s in notes:
        r += 1
        text(r, s, height=32)
    r += 2
    text(r, "使用提示", True, 11, 22, NAVY)
    tips = [
        "兩張資料表第 3 行有篩選箭頭，可按產品種類、保單貨幣或紅利類型篩選；產品名稱及標題列已固定。",
        "資料表底部的「持滿 N 年平均值」包含該保單年度及之後的所有數字；「當年平均值」只計算該保單年度一欄。",
        "平均值為固定數字，已計入所有貨幣，不隨篩選變化。",
        "過往表現並不能作為分紅保險產品未來表現的指引。",
    ]
    for s in tips:
        r += 1
        text(r, s, height=22)


def notion_upload(path, year):
    token = os.environ.get("NOTION_TOKEN", "").strip()
    raw = os.environ.get("NOTION_DATABASE_ID", "").strip()
    if not token or not raw:
        print("未設定 NOTION_TOKEN 或 NOTION_DATABASE_ID，略過 Notion 上傳")
        return True
    m = re.search(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", raw)
    ids = [m.group(0)] if m else re.findall(r"[0-9a-fA-F]{32}", raw.split("?")[0])
    if not ids:
        print("Notion 上傳失敗：NOTION_DATABASE_ID 內找不到資料庫 ID")
        return False
    dbid = ids[-1]
    head = {"Authorization": f"Bearer {token}", "Notion-Version": "2025-09-03"}

    def api(step, method, path_, **kw):
        r = requests.request(method, "https://api.notion.com/v1/" + path_, headers=head, timeout=120, **kw)
        if r.status_code >= 300:
            raise RuntimeError(f"{step}：狀態 {r.status_code}｜{r.text[:200]}")
        return r.json()

    try:
        try:
            ds_id = api("讀取資料庫", "GET", f"databases/{dbid}")["data_sources"][0]["id"]
        except Exception:
            ds_id = dbid
        props = api("讀取資料來源", "GET", f"data_sources/{ds_id}")["properties"]
        title_prop = next(n for n, p in props.items() if p["type"] == "title")
        file_prop = next((n for n, p in props.items() if p["type"] == "files"), None)

        def page_title(p):
            return "".join(t.get("plain_text", "") for t in p["properties"][title_prop]["title"]).strip()

        found = api("尋找年度頁面", "POST", f"data_sources/{ds_id}/query",
                    json={"filter": {"property": title_prop, "title": {"starts_with": year}}})["results"]
        page = (next((p for p in found if page_title(p) == f"{year}年"), None)
                or next((p for p in found if page_title(p) == year), None)
                or (found[0] if found else None))
        if page:
            print(f"Notion：找到「{page_title(page)}」頁面")
        else:
            page = api("建立年度頁面", "POST", "pages",
                       json={"parent": {"type": "data_source_id", "data_source_id": ds_id},
                             "properties": {title_prop: {"title": [{"text": {"content": f"{year}年"}}]}}})
            print(f"Notion：找不到年度頁面，已建立「{year}年」")
        fname = os.path.basename(path)
        ascii_name = f"Manulife_{year}.xlsx"

        def is_auto(n):
            return n == fname or n.startswith("Manulife_")

        up = api("建立上傳", "POST", "file_uploads",
                 json={"mode": "single_part", "filename": fname, "content_type": XLSX})
        with open(path, "rb") as f:
            api("傳送檔案", "POST", f"file_uploads/{up['id']}/send", files={"file": (ascii_name, f, XLSX)})
        new = {"name": fname, "type": "file_upload", "file_upload": {"id": up["id"]}}
        if file_prop:
            keep = []
            for x in page.get("properties", {}).get(file_prop, {}).get("files", []):
                if is_auto(x.get("name", "")):
                    continue
                if x.get("type") == "file":
                    keep.append({"name": x["name"], "type": "file", "file": {"url": x["file"]["url"]}})
                elif x.get("type") == "external":
                    keep.append({"name": x["name"], "type": "external", "external": {"url": x["external"]["url"]}})
            done = api("更新檔案欄", "PATCH", f"pages/{page['id']}",
                       json={"properties": {file_prop: {"files": keep + [new]}}})
            count = len(done.get("properties", {}).get(file_prop, {}).get("files", []))
            print(f"Notion：已放入「{file_prop}」欄｜保留原有檔案 {len(keep)} 個｜現有檔案 {count} 個")
            if count < len(keep) + 1:
                print("注意：檔案數目比預期少，請到 Notion 檢查原有檔案")
        else:
            kids = api("讀取頁面內容", "GET", f"blocks/{page['id']}/children?page_size=100")["results"]
            removed = 0
            for b in kids:
                if b.get("type") == "file" and is_auto(b["file"].get("name", "")):
                    api("刪除舊檔", "DELETE", f"blocks/{b['id']}")
                    removed += 1
            api("加入檔案", "PATCH", f"blocks/{page['id']}/children",
                json={"children": [{"type": "file", "file": new}]})
            print(f"Notion：已把檔案放在頁面內容最底｜替換舊自動檔 {removed} 個｜其他內容不變")
        return True
    except Exception as e:
        print(f"Notion 上傳失敗｜{str(e)[:300]}")
        return False


def main():
    intro = call(INTRO)
    items = ITEM.findall(intro)
    log = {"empty": [], "mismatch": [], "years": collections.Counter()}
    ycols, ysubs, sheets = [], {}, []
    for label, url, is_tcv, _ in SETS:
        prods = list(dict.fromkeys((n, c) for f, a, n, c in items if a == "1" and f.endswith("TCV") == is_tcv))
        print(f"=== {label}：產品 {len(prods)} 個 ===", flush=True)
        rows = []
        fetch(label, url, prods, rows, ycols, ysubs, log)
        sheets.append((label, rows, len(prods)))
    year = log["years"].most_common(1)[0][0] if log["years"] else str(datetime.date.today().year)
    now = datetime.datetime.now(datetime.timezone(datetime.timedelta(hours=8))).strftime("%Y-%m-%d %H:%M")

    def show(vals):
        return " ".join(f"{v:.2%}" if v is not None else "/" for v in vals)

    print()
    print("=== 總結 ===")
    print(f"申報年度：{dict(log['years'])}")
    print(f"保單年度欄 {len(ycols)} 個：", "、".join(ycols))
    for label, rows, total in sheets:
        kinds = dict(collections.Counter(r["kind"] for r in rows))
        print(f"{label}：產品 {total} 個｜有資料的產品 {len(set(r['name'] for r in rows))} 個｜行數 {len(rows)}"
              f"｜保單生效時期 {len(set((r['name'], r['period']) for r in rows))} 個｜類型 {kinds}")
        print(f"    當年平均值：{show(averages(rows, ycols))}")
        print(f"    持滿平均值：{show(cumulative(rows, ycols))}")
    for title, key in (("讀不到資料", "empty"), ("分頁與網站顯示不符", "mismatch")):
        print(f"{title}：{len(log[key])} 項")
        for s in log[key][:20]:
            print("   ", s)

    if not any(rows for _, rows, _ in sheets):
        print("完全讀不到資料，沒有產生 Excel")
        sys.exit(1)
    wb = Workbook()
    write_overview(wb.active, sheets, year, now)
    wb.active.title = "概覽"
    for (label, rows, _), (_, _, _, kind_head) in zip(sheets, SETS):
        write_sheet(wb.create_sheet(f"{year} {label}"), f"宏利 {year} {label}", rows, ycols, ysubs, kind_head)
    os.makedirs("output", exist_ok=True)
    path = f"output/宏利總現金價值分紅實現率（{year}年）自動生成.xlsx"
    wb.save(path)
    print(f"已產生 {path}")
    if any(not rows for _, rows, _ in sheets):
        print("其中一類沒有資料，不上傳 Notion")
        sys.exit(1)
    if not notion_upload(path, year):
        sys.exit(1)


main()
