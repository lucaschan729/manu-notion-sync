import re
import subprocess
import sys
from urllib.parse import urljoin

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

PAGE = "https://www.manulife.com.hk/zh-hk/individual/products/understanding-your-participating-policy/fulfillment-ratio.html"
HOST = "tools.manulife.com.hk"
BASE = f"https://{HOST}/fulfillment-ratio-2024/"
JS = BASE + "intro-zh.js"
INTRO = BASE + "html/zh/intro.php"
PCT = re.compile(r"\d+(?:\.\d+)?\s*%")
REF = re.compile(r"""["'`(=]\s*([^"'`()\s<>=]+\.(?:php|js|json|html|htm|csv|xlsx|txt)(?:\?[^"'`()\s<>]*)?)""", re.I)
SKIP = re.compile(r"jquery|bootstrap|analytics|gtm|polyfill|adobe|launch", re.I)
CALL = re.compile(r"\.load\(|\.ajax\(|\$\.get|\$\.post|fetch\(|getJSON|XMLHttpRequest|\.php|\.json", re.I)
KEYS = ["富足自主", "盈傳創富", "活耀人生"]
MAX_FILES = 30


def get(url):
    r = creq.get(url, impersonate="chrome",
                 headers={"Referer": PAGE, "X-Requested-With": "XMLHttpRequest"}, timeout=120)
    if r.status_code != 200:
        r = creq.get(url, impersonate="chrome", headers={"Referer": PAGE}, timeout=120)
    return r


def decode(content):
    s = content.decode("utf-8", "replace")
    return re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), s).replace("\\/", "/")


def one_line(s):
    return re.sub(r"\s+", " ", s)


def strip(s):
    return one_line(re.sub(r"<[^>]+>", " ", s)).strip()


def attr(tag, name):
    m = re.search(name + r"""\s*=\s*["']([^"']*)["']""", tag)
    return m.group(1) if m else ""


def show(url, body):
    tables = re.findall(r"<table\b.*?</table>", body, flags=re.S | re.I)
    pcts = len(PCT.findall(body))
    hits = {k: body.count(k) for k in KEYS if k in body}
    print(f"    表格 {len(tables)} 個｜百分比數字 {pcts} 個｜產品名稱 {hits}")
    values = []
    selects = re.findall(r"<select\b([^>]*)>(.*?)</select>", body, flags=re.S | re.I)
    if selects:
        print(f"    下拉選單 {len(selects)} 個")
    for attrs, inner in selects[:6]:
        opts = re.findall(r"<option\b([^>]*)>(.*?)</option>", inner, flags=re.S | re.I)
        print(f"      選單 [{one_line(attrs)[:120]}]：選項 {len(opts)} 個")
        for oa, ot in opts[:4]:
            v = attr(oa, "value")
            values.append(v)
            print(f"        value={v[:80]}｜{strip(ot)[:40]}")
    for k in KEYS:
        i = body.find(k)
        if i >= 0:
            print(f"    「{k}」附近的原始碼：", one_line(body[max(i - 500, 0):i + 500]))
            break
    shown, last_end = 0, -1
    for m in CALL.finditer(body):
        if m.start() < last_end or shown >= 10:
            continue
        last_end = m.end() + 250
        shown += 1
        print(f"    載入位置 {shown}：", one_line(body[max(m.start() - 150, 0):m.end() + 250]))
    for t in tables:
        if PCT.search(t):
            print("    表格例子：")
            for tr in re.findall(r"<tr\b.*?</tr>", t, flags=re.S | re.I)[:6]:
                cells = re.findall(r"<t[hd]\b[^>]*>(.*?)</t[hd]>", tr, flags=re.S | re.I)
                print("      ", "｜".join(strip(c) for c in cells)[:220])
            break
    refs = []
    for ref in REF.findall(body):
        full = urljoin(url, ref)
        if HOST in full and not SKIP.search(full) and full not in refs:
            refs.append(full)
    if refs:
        print(f"    連結 {len(refs)} 個：")
        for ref in refs[:25]:
            print("      →", ref[-100:])
    return refs, values, pcts, hits


queue, seen, found, intro_values = [(JS, 0), (INTRO, 0)], {JS, INTRO}, [], []
count = 0
while queue and count < MAX_FILES:
    url, depth = queue.pop(0)
    count += 1
    print(f"[{count}] {url[-100:]}")
    try:
        r = get(url)
    except Exception as e:
        print(f"    失敗 {str(e)[:100]}")
        continue
    kind = r.headers.get("Content-Type", "")
    print(f"    狀態 {r.status_code}｜類型 {kind[:30]}｜大小 {len(r.content) // 1024} KB")
    if r.status_code != 200 or not re.search(r"html|json|javascript|text", kind):
        continue
    body = decode(r.content)
    if url == JS:
        print("    全文：", one_line(body)[:2500])
    refs, values, pcts, hits = show(url, body)
    if url == INTRO:
        intro_values = values
    if pcts >= 20:
        found.append((url, len(r.content), pcts, hits))
    if depth < 3:
        for ref in refs:
            if ref not in seen:
                seen.add(ref)
                queue.append((ref, depth + 1))

print()
print("=== 按選單的值猜測網址 ===")
tried = 0
for v in intro_values:
    if not v or not re.fullmatch(r"[\w\-./]+", v) or tried >= 3:
        continue
    tried += 1
    guesses = [urljoin(INTRO, v)] if "." in v else [urljoin(INTRO, v + ".php"), urljoin(INTRO, v + ".html")]
    for g in guesses:
        try:
            r = get(g)
            body = decode(r.content)
            n = len(PCT.findall(body))
            print(f"  {g[-90:]}：狀態 {r.status_code}｜百分比數字 {n} 個")
            if r.status_code == 200 and n >= 20:
                found.append((g, len(r.content), n, {}))
        except Exception as e:
            print(f"  {g[-90:]}：失敗 {str(e)[:80]}")
if not tried:
    print("  選單沒有可用的值")

print()
print("=== 總結 ===")
print(f"已檢查 {count} 個檔案｜尚未檢查 {len(queue)} 個")
if not found:
    print("沒有找到含大量百分比數字的檔案")
for url, size, pcts, hits in found:
    print(f"有數字的檔案：{url}")
    print(f"    大小 {size // 1024} KB｜百分比數字 {pcts} 個｜產品名稱 {hits}")
