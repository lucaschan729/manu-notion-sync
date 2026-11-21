import re
import requests

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36")
URLS = [
    "https://www.manulife.com.hk/zh-hk/individual/products/understanding-your-participating-policy/fulfillment-ratio.html",
    "https://www.manulife.com.hk/content/dam/insurance/hk/zh-hk/documents/products/product-services/fulfillment-ratio.pdf",
]

for url in URLS:
    print("=" * 30)
    print(url)
    try:
        r = requests.get(url, headers={"User-Agent": UA, "Accept-Language": "zh-HK,zh;q=0.9"}, timeout=60)
    except Exception as e:
        print("失敗：", str(e)[:200])
        continue
    kind = r.headers.get("Content-Type", "")
    print(f"狀態 {r.status_code}｜類型 {kind}｜大小 {len(r.content) // 1024} KB")
    if "html" not in kind:
        continue
    r.encoding = "utf-8"
    page = r.text
    tables = re.findall(r"<table\b.*?</table>", page, flags=re.S | re.I)
    print(f"表格 {len(tables)} 個｜百分比數字 {len(re.findall(r'\d+(?:\.\d+)?\s*%', page))} 個")
    print(f"「分紅實現率」{page.count('分紅實現率')} 次｜「總現金價值」{page.count('總現金價值')} 次")
    links = sorted(set(re.findall(r'["\']([^"\']+\.(?:json|pdf|xlsx|csv)[^"\']*)["\']', page, flags=re.I)))
    print(f"資料檔連結 {len(links)} 個：")
    for link in links[:40]:
        print("  ", link[:200])
    frames = re.findall(r'<iframe\b[^>]*src="([^"]+)"', page, flags=re.I)
    print("內嵌頁面：", frames[:10])
    for t in tables[:2]:
        plain = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", t))
        print("表格例子：", plain[:600])
    body = re.sub(r"\s+", " ", re.sub(r"<(script|style)\b.*?</\1>|<[^>]+>", " ", page, flags=re.S | re.I))
    i = body.find("分紅實現率")
    print("內文例子：", body[max(i, 0):max(i, 0) + 600])
