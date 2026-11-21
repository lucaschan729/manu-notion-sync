import collections
import datetime
import html as htmllib
import os
import re
import subprocess
import sys
import time

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi", "openpyxl"])
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


def averages(rows, ycols):
    out = []
    for k in ycols:
        nums = [r["vals"][k] for r in rows if isinstance(r["vals"].get(k), float)]
        out.append(sum(nums) / len(nums) if nums else None)
    return out


def head_cell(c):
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.fill = PatternFill("solid", fgColor=NAVY)
    c.alignment = CENTER


def write_sheet(ws, rows, ycols, ysubs, kind_head):
    cols = [("產品名稱", 30, "name"), ("產品種類", 14, "cat"), ("保單生效時期", 22, "period"), ("保單貨幣", 9, "cur")]
    if len(set(r["kind"] for r in rows)) > 1:
        cols.append((kind_head, 11, "kind"))
    if any(r["note"] for r in rows):
        cols.append(("備註", 16, "note"))
    heads = [c[0] for c in cols] + [(k + "\n" + ysubs.get(k, "")).strip() for k in ycols]
    widths = [c[1] for c in cols] + [13] * len(ycols)
    for i, (h, w) in enumerate(zip(heads, widths), 1):
        head_cell(ws.cell(row=1, column=i, value=h))
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 62
    ri, prev = 1, None
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
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{max(ri, 2)}"
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"


def write_overview(ws, sheets, ycols, year, now):
    last = 2 + len(ycols)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 9
    for i in range(len(ycols)):
        ws.column_dimensions[get_column_letter(3 + i)].width = 12

    def text(r, s, bold=False, size=10, height=None, color="000000"):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last)
        c = ws.cell(row=r, column=1, value=s)
        c.font = Font(bold=bold, size=size, color=color)
        c.alignment = Alignment(vertical="center", wrap_text=True)
        if height:
            ws.row_dimensions[r].height = height

    text(1, f"宏利 分紅實現率及總現金價值比率（{year} 申報年度）", True, 14, 28, NAVY)
    text(2, f"自動生成｜擷取時間：{now}（香港時間）｜資料來源：宏利香港網站")
    text(3, PAGE, color="0563C1")
    text(5, "各保單年度平均值（固定數字，不隨篩選變化；只計算有百分比的格）", True, 11, 22, NAVY)
    for i, h in enumerate(["項目", "行數"] + ycols, 1):
        head_cell(ws.cell(row=6, column=i, value=h))
    ws.row_dimensions[6].height = 24
    r = 7
    for label, rows, _ in sheets:
        kinds = list(dict.fromkeys(x["kind"] for x in rows))
        groups = [(f"{label}（全部）", rows)]
        if len(kinds) > 1:
            groups += [(f"{label}－{k}", [x for x in rows if x["kind"] == k]) for k in kinds]
        for title, sub in groups:
            cells = [title, len(sub)] + averages(sub, ycols)
            for i, v in enumerate(cells, 1):
                c = ws.cell(row=r, column=i, value=v)
                c.border = BOX
                c.font = Font(size=10, bold=i == 1)
                c.alignment = Alignment(horizontal="left" if i == 1 else "center", vertical="center")
                if i > 2 and v is not None:
                    c.number_format = "0.0%"
            r += 1
    notes = [
        "總現金價值比率：根據所有相關保單的實際總現金價值（包括保證及非保證成份），比對銷售時建議書上所示之預計總現金價值所計算的比率。",
        "分紅實現率：所有相關保單的實際累計非保證利益，比對銷售時建議書上所示之預計累計非保證利益所計算的比率。",
        "高於 100% 表示實際高於投保時的預計。100% 或以上以綠底顯示，低於 80% 以紅底顯示。「不適用」為宏利網站原文。",
        "過往表現並不能作為分紅保險產品未來表現的指引。",
    ]
    for i, s in enumerate(notes):
        text(r + 1 + i, s, height=32)


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

    print()
    print("=== 總結 ===")
    print(f"申報年度：{dict(log['years'])}")
    print(f"保單年度欄 {len(ycols)} 個：", "、".join(ycols))
    for label, rows, total in sheets:
        kinds = dict(collections.Counter(r["kind"] for r in rows))
        print(f"{label}：產品 {total} 個｜有資料的產品 {len(set(r['name'] for r in rows))} 個｜行數 {len(rows)}"
              f"｜保單生效時期 {len(set((r['name'], r['period']) for r in rows))} 個｜類型 {kinds}")
        for r in rows[:3]:
            vals = " ".join(f"{v:.0%}" if isinstance(v, float) else str(v) for v in r["vals"].values())
            print(f"    抽查：{r['name']}｜{r['period']}｜{r['cur']}｜{r['kind']}｜{vals}")
    for title, key in (("讀不到資料", "empty"), ("分頁與網站顯示不符", "mismatch")):
        print(f"{title}：{len(log[key])} 項")
        for s in log[key][:20]:
            print("   ", s)

    if not any(rows for _, rows, _ in sheets):
        print("完全讀不到資料，沒有產生 Excel")
        sys.exit(1)
    wb = Workbook()
    write_overview(wb.active, sheets, ycols, year, now)
    wb.active.title = "概覽"
    for (label, rows, _), (_, _, _, kind_head) in zip(sheets, SETS):
        write_sheet(wb.create_sheet(label), rows, ycols, ysubs, kind_head)
    os.makedirs("output", exist_ok=True)
    path = f"output/Manulife_{year}.xlsx"
    wb.save(path)
    print(f"已產生 {path}")
    if any(not rows for _, rows, _ in sheets):
        sys.exit(1)


main()
