import html as htmllib
import io
import re
import subprocess
import sys

subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", "curl_cffi", "pdfplumber"])
import pdfplumber
from curl_cffi import requests as creq

BASE = "https://www.manulife.com.hk"
PAGE = BASE + "/zh-hk/individual/products/understanding-your-participating-policy/fulfillment-ratio.html"
PDF = BASE + "/content/dam/insurance/hk/zh-hk/documents/products/product-services/fulfillment-ratio.pdf"
PCT = re.compile(r"\d+(?:\.\d+)?\s*%")
FILES = re.compile(r"[\w/.:%~\-]+\.(?:pdf|json|xlsx|xls|csv)\b", re.I)


def get(url):
    return creq.get(url, impersonate="chrome", timeout=180)


def one_line(s):
    return re.sub(r"\s+", " ", s)


print("=== 第一部分：網頁 ===")
r = get(PAGE)
page = htmllib.unescape(r.content.decode("utf-8", "replace")).replace("\\/", "/")
print(f"狀態 {r.status_code}｜大小 {len(r.content) // 1024} KB")
files = sorted(set(FILES.findall(page)))
print(f"檔案連結 {len(files)} 個：")
for f in files[:40]:
    print("  ", f[:160])
seen = set()
for m in re.finditer("分紅實現率", page):
    snippet = one_line(page[max(m.start() - 120, 0):m.end() + 200])
    if snippet not in seen:
        seen.add(snippet)
        print(f"出現位置 {len(seen)}：", snippet)
try:
    r2 = get(PAGE.replace(".html", ".model.json"))
    body = r2.content.decode("utf-8", "replace")
    print(f"model.json：狀態 {r2.status_code}｜大小 {len(r2.content) // 1024} KB｜百分比數字 {len(PCT.findall(body))} 個")
except Exception as e:
    print("model.json 失敗：", str(e)[:150])

print()
print("=== 第二部分：PDF ===")
r = get(PDF)
data = r.content
print(f"狀態 {r.status_code}｜大小 {len(data) // 1024} KB｜是 PDF：{data[:4] == b'%PDF'}")
texts, tables_by_page = [], []
with pdfplumber.open(io.BytesIO(data)) as pdf:
    for p in pdf.pages:
        try:
            texts.append(p.extract_text() or "")
            tables_by_page.append(p.extract_tables() or [])
        except Exception as e:
            texts.append("")
            tables_by_page.append([])
            print(f"第 {len(texts)} 頁讀取失敗：{str(e)[:100]}")
n = len(texts)
pcts = [len(PCT.findall(t)) for t in texts]
print(f"頁數 {n}｜有文字的頁 {sum(1 for t in texts if t.strip())}"
      f"｜有表格的頁 {sum(1 for t in tables_by_page if t)}"
      f"｜表格總數 {sum(len(t) for t in tables_by_page)}｜百分比數字 {sum(pcts)} 個")
print("每頁（頁碼:表格數/百分比數）：",
      " ".join(f"{i + 1}:{len(tables_by_page[i])}/{pcts[i]}" for i in range(n)))
for i in range(min(2, n)):
    print(f"第 {i + 1} 頁文字：", one_line(texts[i])[:400])
target = next((i for i in range(n) if tables_by_page[i] and pcts[i] >= 5), None)
if target is None:
    print("找不到同時有表格和百分比數字的頁")
else:
    print(f"第 {target + 1} 頁（第一個有數字的表格頁）文字：", one_line(texts[target])[:600])
    for row in tables_by_page[target][0][:8]:
        print("  表格行：", "｜".join(one_line(str(c)) if c is not None else "" for c in row)[:220])
for key in ("盈傳創富", "總現金價值比率"):
    hit = next((i for i in range(n) if key in texts[i]), None)
    if hit is None:
        print(f"「{key}」：整份 PDF 找不到")
    else:
        t = one_line(texts[hit])
        j = t.find(key)
        print(f"「{key}」首次出現在第 {hit + 1} 頁：", t[max(j - 80, 0):j + 320])
