# 抓取 books.toscrape.com 的图书目录，保存书名、价格、评分、是否有货和详情链接。
# 命令：python scrape_books.py <页数> <输出.xlsx>
# 示例：python scrape_books.py 1 ./books.xlsx
import sys
import time
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

COLUMNS = ["书名", "价格", "评分", "是否有货", "详情链接"]
RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
PAGE_URL = "https://books.toscrape.com/catalogue/page-{page}.html"


def _fetch_page(page, attempts=3):
    """Return the books on one page, or [] if the page cannot be fetched."""
    url = PAGE_URL.format(page=page)
    for _ in range(attempts):
        try:
            response = requests.get(url, timeout=20)
            if response.status_code == 200:
                break
        except requests.RequestException:
            pass
        time.sleep(1)
    else:
        return []

    books = []
    soup = BeautifulSoup(response.content, "html.parser")
    for article in soup.select("article.product_pod"):
        anchor = article.select_one("h3 a")
        price = article.select_one("p.price_color").get_text(strip=True)
        classes = article.select_one("p.star-rating").get("class", [])
        stock = article.select_one("p.availability").get_text(strip=True)
        books.append({
            "书名": anchor["title"],
            "价格": float(price.replace("£", "").replace("Â", "")),
            "评分": next((RATINGS[c] for c in classes if c in RATINGS), 0),
            "是否有货": "有货" if "In stock" in stock else "无货",
            "详情链接": urljoin(url, anchor["href"]),
        })
    return books


def scrape_books(pages, output):
    """Scrape the first `pages` catalogue pages into `output`. Returns (ok, message)."""
    books, skipped = [], []
    for page in range(1, pages + 1):
        parsed = _fetch_page(page)
        if parsed:
            books.extend(parsed)
        else:
            skipped.append(str(page))
        if page != pages:
            time.sleep(1)
    if not books:
        return False, "所有页面均抓取失败"

    frame = pd.DataFrame(books, columns=COLUMNS)
    frame = frame.sort_values(["评分", "价格"], ascending=[False, True])
    try:
        frame.to_excel(output, index=False, engine="openpyxl")
    except OSError:
        return False, f"无法写入：{output}"

    text = f"成功抓取 {pages - len(skipped)} 页，保存 {len(frame)} 本书。"
    if skipped:
        text += f"跳过的页：{'、'.join(skipped)}。"
    return True, text


if __name__ == "__main__":
    if len(sys.argv) < 3 or not sys.argv[1].isdigit() or int(sys.argv[1]) < 1:
        print("用法：python scrape_books.py <页数(>=1)> <输出.xlsx>")
        sys.exit(1)
    ok, message = scrape_books(int(sys.argv[1]), sys.argv[2])
    print(message)
    sys.exit(0 if ok else 1)
