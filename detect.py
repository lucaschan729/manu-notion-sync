import collections
import html as htmllib
import re
import subprocess
import sys

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

PAGE = "https://www.manulife.com.hk/zh-hk/individual/products/understanding-your-participating-policy/fulfillment-ratio.html"
ROOT = "https://tools.manulife.com.hk/"
INTRO = ROOT + "fulfillment-ratio-2024/html/zh/intro.php"
SETS = [("總現金價值比率", ROOT + "fulfillment-tcv-2024/html/", True, "lang-tcv"),
        ("分紅實現率", ROOT + "fulfillment-ratio-2024/html/", False, "lang")]
HEAD = {"Referer": PAGE, "Origin": "https://www.manulife.com.hk", "X-Requested-With": "XMLHttpRequest"}
PCT = re.compile(r"\d+(?:\.\d+)?\s*%")
ITEM = re.compile(r"(selectedItem\w*)\(\s*'([^']*)'\s*,\s*'([^']*)'\s*,\s*'([^']*)'\s*\)")
TAB = re.compile(r"""\w*Details\w*\(\s*['"]([^'"]*)['"]\s*\)""")


def decode(content):
    s = content.decode("utf-8", "replace")
    s = re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), s).replace("\\/", "/")
    return htmllib.unescape(s)


def one_line(s):
    return re.sub(r"\s+", " ", s)


def strip(s):
    return one_line(re.sub(r"<[^>]+>", " ", s)).strip()


def attr(tag, name):
    m = re.search(r"\b" + name + r"""\s*=\s*["']([^"']*)["']""", tag)
    return m.group(1) if m else ""


def get(url):
    return creq.get(url, impersonate="chrome", headers=HEAD, timeout=120)


def post(url, name, cur, tab, lang):
    data = {"selected1": name, "selected2": cur, "tablabel": tab, "lang": lang}
    r = creq.post(url, data=data, impersonate="chrome", headers=HEAD, timeout=120)
    return r.status_code, decode(r.content)


def describe(body, rows=8, raw=0):
    tables = re.findall(r"<table\b.*?</table>", body, flags=re.S | re.I)
    print(f"      長度 {len(body)} 字｜表格 {len(tables)} 個｜百分比數字 {len(PCT.findall(body))} 個")
    if raw:
        print("      原始碼：", one_line(body)[:raw])
    for ti, t in enumerate(tables[:2]):
        print(f"      表格 {ti + 1}：")
        for tr in re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)[:rows]:
            cells = re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)
            print("        ", "｜".join(strip(c) for c in cells)[:230])
    if not tables:
        print("      文字：", strip(body)[:600])
    tabs = []
    for t in TAB.findall(body):
        if t not in tabs:
            tabs.append(t)
    return tabs


r = get(INTRO)
intro = decode(r.content)
print(f"產品清單頁：狀態 {r.status_code}｜大小 {len(r.content) // 1024} KB")
inputs = re.findall(r"<input\b[^>]*>", intro, flags=re.I)
print("隱藏欄位：", "｜".join(f"{attr(t, 'id')}={attr(t, 'value')}" for t in inputs if attr(t, "id"))[:400])
items = ITEM.findall(intro)
print("選單函數：", dict(collections.Counter((f, a) for f, a, _, _ in items)))

for label, base, is_tcv, lang_id in SETS:
    print()
    print(f"=== {label} ===")
    mine = [(f, a, n, c) for f, a, n, c in items if f.endswith("TCV") == is_tcv]
    prods = []
    for f, a, n, c in mine:
        if a == "1" and (n, c) not in prods:
            prods.append((n, c))
    if not prods:
        prods = list(dict.fromkeys((n, c) for _, _, n, c in mine))
    print(f"產品 {len(prods)} 個｜貨幣組合：", dict(collections.Counter(c for _, c in prods)))
    for n, c in prods[:3]:
        print(f"    {n}｜{c}")
    try:
        js = decode(get(base + "common/cus.js").content)
        i = js.find("function selectedItem")
        print("選擇產品的程式：", one_line(js[max(i, 0):max(i, 0) + 700]))
    except Exception as e:
        print("讀取 cus.js 失敗：", str(e)[:100])
    if not prods:
        print("找不到產品清單")
        continue
    lang_tag = next((t for t in inputs if attr(t, "id") == lang_id), "")
    lang_val = attr(lang_tag, "value")
    langs = [lang_val] if lang_val else ["zh", "tc", "zh-hk", "chi", "zh_TW"]
    url = base + "ajax/ss1.php"
    name, curs = prods[0]
    first_cur = curs.split("||")[0]
    good = None
    for lang in langs:
        for cur in dict.fromkeys([first_cur, curs]):
            try:
                status, body = post(url, name, cur, "null", lang)
            except Exception as e:
                print(f"  查詢失敗：{str(e)[:100]}")
                continue
            n = len(PCT.findall(body))
            print(f"  試查 {name}｜貨幣={cur}｜lang={lang}：狀態 {status}｜長度 {len(body)}｜百分比 {n}")
            if status == 200 and (n or len(body) > 500):
                good = (lang, body)
                break
        if good:
            break
    if not good:
        print("  所有組合都沒有回傳內容")
        continue
    lang, body = good
    print(f"  成功的組合：lang={lang}")
    tabs = describe(body, rows=8, raw=1500)
    print("  分頁標籤：", tabs)
    if tabs:
        try:
            status, body2 = post(url, name, first_cur, tabs[-1], lang)
            print(f"  查詢分頁「{tabs[-1]}」：狀態 {status}")
            describe(body2, rows=6)
        except Exception as e:
            print(f"  查詢分頁失敗：{str(e)[:100]}")
    name2, curs2 = prods[len(prods) // 2]
    try:
        status, body3 = post(url, name2, curs2.split("||")[0], "null", lang)
        print(f"  另一個產品 {name2}｜{curs2.split('||')[0]}：狀態 {status}")
        tabs3 = describe(body3, rows=6)
        print("  分頁標籤：", tabs3)
    except Exception as e:
        print(f"  查詢失敗：{str(e)[:100]}")
