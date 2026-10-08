# 抓取 books.toscrape.com 的图书目录，保存书名、价格、评分、是否有货和详情链接。
# 命令：python scrape_books.py <页数> <输出.xlsx>
# 示例：python scrape_books.py 1 ./books.xlsx
"""Scrape books.toscrape.com catalogue pages into an Excel file."""

import errno
import os
import sys
import tempfile
import time
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup

COLUMNS = ["书名", "价格", "评分", "是否有货", "详情链接"]
_RATINGS = {"One": 1, "Two": 2, "Three": 3, "Four": 4, "Five": 5}
_PAGE_URL = "https://books.toscrape.com/catalogue/page-{page}.html"
_ATTEMPTS = 3
_TIMEOUT = 20


def scrape_books(pages, output):
    """Scrape the first `pages` catalogue pages and save them to `output`.

    Returns ``(ok, message)``. Invalid pages, failed requests, and unwritable
    destinations return ``(False, message)`` instead of raising.
    """
    if type(pages) is not int or pages < 1:
        return False, "页数必须是大于等于 1 的整数"

    path, error = _check_output(output)
    if error is not None:
        return error

    books = []
    skipped = []
    for page in range(1, pages + 1):
        parsed = _fetch_page(page)
        if parsed is None or len(parsed) == 0:
            skipped.append(page)
        else:
            books.extend(parsed)
        if page != pages:
            time.sleep(1)

    if len(skipped) == pages:
        return False, "所有页面均抓取失败"

    frame = pd.DataFrame(books, columns=COLUMNS)
    frame["价格"] = frame["价格"].astype(float)
    frame["评分"] = frame["评分"].astype(int)
    frame = frame.sort_values(
        by=["评分", "价格"],
        ascending=[False, True],
        kind="mergesort",
    ).reset_index(drop=True)

    write_error = _write_excel(frame, path, output)
    if write_error is not None:
        return write_error

    scraped = pages - len(skipped)
    return True, _success_message(scraped, len(frame), skipped)


def _success_message(scraped_pages, book_count, skipped):
    text = f"成功抓取 {scraped_pages} 页，保存 {book_count} 本书。"
    if not skipped:
        return text + "没有跳过任何页。"
    numbers = "、".join(str(number) for number in skipped)
    return text + f"跳过的页：{numbers}。"


def _check_output(output):
    cannot = f"无法写入：{output}"
    denied = f"没有权限写入：{output}"
    try:
        path = Path(output)
    except (TypeError, ValueError):
        return None, (False, cannot)
    try:
        parent = path.parent
        if path.is_dir() or not parent.exists() or not parent.is_dir():
            return None, (False, cannot)
        if not os.access(parent, os.W_OK):
            return None, (False, denied)
    except PermissionError:
        return None, (False, denied)
    except OSError:
        return None, (False, cannot)
    return path, None


def _fetch_page(page):
    url = _PAGE_URL.format(page=page)
    for attempt in range(_ATTEMPTS):
        response = None
        try:
            response = requests.get(url, timeout=_TIMEOUT)
        except requests.RequestException:
            response = None
        if response is not None and response.status_code == 200:
            return _parse_page(response.content, url)
        if attempt < _ATTEMPTS - 1:
            time.sleep(1)
    return None


def _parse_page(content, page_url):
    soup = BeautifulSoup(content, "html.parser")
    books = []
    for article in soup.select("article.product_pod"):
        book = _parse_book(article, page_url)
        if book is not None:
            books.append(book)
    return books


def _parse_book(article, page_url):
    anchor = article.select_one("h3 a")
    price_el = article.select_one("p.price_color")
    rating_el = article.select_one("p.star-rating")
    if anchor is None or price_el is None or rating_el is None:
        return None
    title = anchor.get("title")
    href = anchor.get("href")
    if not title or not href:
        return None
    rating = _rating_value(rating_el)
    if rating is None:
        return None
    try:
        price = _parse_price(price_el.get_text(strip=True))
    except ValueError:
        return None
    availability = article.select_one("p.availability")
    availability_text = availability.get_text(" ", strip=True) if availability else ""
    in_stock = "有货" if "In stock" in availability_text else "无货"
    return {
        "书名": title,
        "价格": price,
        "评分": rating,
        "是否有货": in_stock,
        "详情链接": urljoin(page_url, href),
    }


def _rating_value(element):
    for name in element.get("class") or []:
        if name in _RATINGS:
            return _RATINGS[name]
    return None


def _parse_price(text):
    cleaned = text.replace("£", "").replace("Â", "").replace(",", "").strip()
    return float(cleaned)


def _write_excel(frame, path, output):
    cannot = f"无法写入：{output}"
    denied = f"没有权限写入：{output}"
    tmp_name = None
    try:
        fd, tmp_name = tempfile.mkstemp(
            prefix=".scrape_books_",
            suffix=".xlsx",
            dir=path.parent,
        )
        os.close(fd)
        frame.to_excel(tmp_name, index=False, engine="openpyxl")
        os.replace(tmp_name, path)
        tmp_name = None
    except PermissionError:
        return False, denied
    except OSError as exc:
        if exc.errno in (errno.EACCES, errno.EPERM):
            return False, denied
        return False, cannot
    except Exception:
        return False, cannot
    finally:
        if tmp_name is not None:
            try:
                os.remove(tmp_name)
            except OSError:
                pass
    return None


def main(argv=None):
    if argv is None:
        argv = sys.argv
    if len(argv) < 3:
        print("用法：python scrape_books.py <页数> <输出.xlsx>")
        return 1
    raw_pages, output = argv[1], argv[2]
    try:
        pages = int(raw_pages)
    except (TypeError, ValueError):
        pages = raw_pages
    ok, text = scrape_books(pages, output)
    print(text)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
