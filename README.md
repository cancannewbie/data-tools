# data-tools

三个独立的命令行工具：按扩展名汇总文件夹、合并销售 Excel、抓取图书目录。

## 安装依赖

```bash
pip install -r requirements.txt
```

`file_summary.py` 只用 Python 标准库。`merge_sales.py` 和 `scrape_books.py` 需要上面列出的依赖。

## file_summary.py

统计一个文件夹里各扩展名的文件数量和大小（不进入子文件夹），并把结果写入该文件夹的 `summary.txt`。`summary.txt` 自身不计入统计。

```bash
python file_summary.py ./docs
```

## merge_sales.py

合并一个文件夹中的销售 Excel（列：日期、商品、数量、单价），写出带「明细」和「汇总」两个工作表的文件。输出文件如果就在这个文件夹里，不会被再次读入。

```bash
python merge_sales.py ./sales ./merged.xlsx
```

## scrape_books.py

抓取 books.toscrape.com 的图书目录，保存书名、价格、评分、是否有货和详情链接。某一页在多次尝试后仍失败，或者页面里没有图书时，该页记入「跳过的页」，不算作成功抓取。

```bash
python scrape_books.py 1 ./books.xlsx
```
