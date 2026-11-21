import re
import subprocess
import sys
from urllib.parse import urljoin

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

PAGE = "https://www.manulife.com.hk/zh-hk/individual/products/understanding-your-participating-policy/fulfillment-ratio.html"
TOOLS = "https://tools.manulife.com.hk/"
START = TOOLS + "fulfillment-ratio-2024/intro-zh.js"
PCT = re.compile(r"\d+(?:\.\d+)?\s*%")
REF = re.compile(r"""["'`(=]\s*([^"'`()\s<>=]+\.(?:js|json|html|htm|csv|xlsx|txt)(?:\?[^"'`()\s<>]*)?)""", re.I)
SKIP = re.compile(r"jquery|bootstrap|analytics|gtm|polyfill|adobe|launch", re.I)
KEYS = ["盈傳創富", "卓越終身保", "活耀人生", "分紅實現率", "總現金價值"]
MAX_FILES = 40


def get(url):
    return creq.get(url, impersonate="chrome", headers={"Referer": PAGE}, timeout=120)


def decode(content):
    s = content.decode("utf-8", "replace")
    return re.sub(r"\\u([0-9a-fA-F]{4})", lambda m: chr(int(m.group(1), 16)), s).replace("\\/", "/")


def one_line(s):
    return re.sub(r"\s+", " ", s)


print("=== 其他年份的資料夾 ===")
for y in ("2025", "2026"):
    u = TOOLS + f"fulfillment-ratio-{y}/intro-zh.js"
    try:
        r = get(u)
        print(f"{y}：狀態 {r.status_code}｜大小 {len(r.content) // 1024} KB")
    except Exception as e:
        print(f"{y}：失敗 {str(e)[:100]}")

print()
print("=== 逐個檔案 ===")
queue, seen, found = [(START, 0)], {START}, []
count = 0
while queue and count < MAX_FILES:
    url, depth = queue.pop(0)
    count += 1
    try:
        r = get(url)
    except Exception as e:
        print(f"[{count}] {url[-90:]}：失敗 {str(e)[:100]}")
        continue
    kind = r.headers.get("Content-Type", "")
    size = len(r.content)
    print(f"[{count}] {url[-90:]}")
    print(f"    狀態 {r.status_code}｜類型 {kind[:30]}｜大小 {size // 1024} KB")
    if r.status_code != 200 or "html" not in kind and "json" not in kind and "javascript" not in kind and "text" not in kind:
        continue
    body = decode(r.content)
    pcts = len(PCT.findall(body))
    hits = {k: body.count(k) for k in KEYS if k in body}
    refs = []
    for ref in REF.findall(body):
        full = urljoin(url, ref)
        if "manulife" in full and not SKIP.search(full) and full not in refs:
            refs.append(full)
    print(f"    百分比數字 {pcts} 個｜關鍵字 {hits}｜連結 {len(refs)} 個")
    for ref in refs[:25]:
        print("      →", ref[-110:])
    if pcts >= 20 or any(k in hits for k in KEYS[:3]):
        found.append((url, size, pcts, hits))
        i = max([body.find(k) for k in KEYS[:3]] + [-1])
        if i < 0:
            m = PCT.search(body)
            i = m.start() if m else 0
        print("    內容例子：", one_line(body[max(i - 300, 0):i + 500]))
    if count == 1:
        print("    檔案開頭：", one_line(body[:500]))
    if depth < 3:
        for ref in refs:
            if ref not in seen:
                seen.add(ref)
                queue.append((ref, depth + 1))

print()
print("=== 總結 ===")
print(f"已檢查 {count} 個檔案｜尚未檢查 {len(queue)} 個")
if not found:
    print("沒有找到含產品名稱或大量百分比數字的檔案")
for url, size, pcts, hits in found:
    print(f"有數字的檔案：{url}")
    print(f"    大小 {size // 1024} KB｜百分比數字 {pcts} 個｜關鍵字 {hits}")
