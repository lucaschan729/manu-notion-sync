import re
import subprocess
import sys

import requests

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi"])
from curl_cffi import requests as creq

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
FULL = {
    "User-Agent": UA,
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "zh-HK,zh;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
}
URLS = [
    "https://www.manulife.com.hk/",
    "https://www.manulife.com.hk/zh-hk/individual/products/understanding-your-participating-policy/fulfillment-ratio.html",
    "https://www.manulife.com.hk/content/dam/insurance/hk/zh-hk/documents/products/product-services/fulfillment-ratio.pdf",
]


def plain(url):
    return requests.get(url, headers=FULL, timeout=60)


def chrome(url):
    return creq.get(url, impersonate="chrome", timeout=60)


def safari(url):
    return creq.get(url, impersonate="safari", timeout=60)


METHODS = [("A 完整瀏覽器標頭", plain), ("B 模擬 Chrome", chrome), ("C 模擬 Safari", safari)]
PCT = re.compile(r"\d+(?:\.\d+)?\s*%")
LINK = re.compile(r'["\']([^"\']+\.(?:json|pdf|xlsx|csv)[^"\']*)["\']', re.I)


def details(content):
    page = content.decode("utf-8", "replace")
    tables = re.findall(r"<table\b.*?</table>", page, flags=re.S | re.I)
    print(f"  表格 {len(tables)} 個｜百分比數字 {len(PCT.findall(page))} 個")
    print(f"  「分紅實現率」{page.count('分紅實現率')} 次｜「總現金價值」{page.count('總現金價值')} 次")
    links = sorted(set(LINK.findall(page)))
    print(f"  資料檔連結 {len(links)} 個：")
    for link in links[:40]:
        print("    ", link[:200])
    print("  內嵌頁面：", re.findall(r'<iframe\b[^>]*src="([^"]+)"', page, flags=re.I)[:10])
    for t in tables[:2]:
        print("  表格例子：", re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))[:600])
    body = re.sub(r"<(script|style)\b.*?</\1>|<[^>]+>", " ", page, flags=re.S | re.I)
    body = re.sub(r"\s+", " ", body)
    i = max(body.find("分紅實現率"), 0)
    print("  內文例子：", body[i:i + 600])


for url in URLS:
    print("=" * 30)
    print(url)
    shown = False
    for name, fn in METHODS:
        try:
            r = fn(url)
        except Exception as e:
            print(f"{name}：失敗 {str(e)[:150]}")
            continue
        kind = r.headers.get("Content-Type", "")
        content = r.content
        print(f"{name}：狀態 {r.status_code}｜類型 {kind}｜大小 {len(content) // 1024} KB")
        if r.status_code != 200 or shown:
            continue
        shown = True
        if content[:4] == b"%PDF":
            print("  這是 PDF 檔，可以下載")
        elif "html" in kind:
            details(content)
