# 合并文件夹中的销售表，生成带明细和汇总两个工作表的 Excel。
# 命令：python merge_sales.py <文件夹> <输出.xlsx>
# 示例：python merge_sales.py ./sales ./merged.xlsx
import os
import sys

import pandas as pd

REQUIRED_COLUMNS = ["日期", "商品", "数量", "单价"]


def _read_sales_file(path, name):
    """Return (frame, None) or (None, skip reason)."""
    try:
        frame = pd.read_excel(path, engine="openpyxl")
    except Exception as exc:
        return None, f"打不开：{exc}"
    frame.columns = [str(c).strip() for c in frame.columns]
    missing = [c for c in REQUIRED_COLUMNS if c not in frame.columns]
    if missing:
        return None, "缺列：" + "、".join(missing)
    frame = frame[REQUIRED_COLUMNS].copy()
    frame["日期"] = pd.to_datetime(frame["日期"], errors="coerce", format="mixed")
    frame["数量"] = pd.to_numeric(frame["数量"], errors="coerce")
    frame["单价"] = pd.to_numeric(frame["单价"], errors="coerce")
    frame["金额"] = frame["数量"] * frame["单价"]
    frame["来源文件"] = name
    return frame, None


def merge_sales(folder, output):
    """Merge sales workbooks in one folder. Returns (ok, message)."""
    if not os.path.isdir(folder):
        return False, f"找不到文件夹：{folder}"
    output_path = os.path.abspath(output)
    names = sorted(
        n for n in os.listdir(folder)
        if n.lower().endswith(".xlsx") and not n.startswith("~$")
        and os.path.abspath(os.path.join(folder, n)) != output_path
    )

    frames, skipped = [], []
    for name in names:
        frame, reason = _read_sales_file(os.path.join(folder, name), name)
        if reason:
            skipped.append(f"- {name}：{reason}")
        else:
            frames.append(frame)
    if not frames:
        return False, "\n".join(["没有可用的 .xlsx 文件。"] + skipped)

    detail = pd.concat(frames, ignore_index=True).drop_duplicates()
    detail = detail.dropna(subset=["数量", "单价"]).reset_index(drop=True)
    summary = (
        detail.groupby("商品", as_index=False)
        .agg(总数量=("数量", "sum"), 总金额=("金额", "sum"))
        .sort_values("总金额", ascending=False)
    )

    try:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            detail.to_excel(writer, sheet_name="明细", index=False)
            summary.to_excel(writer, sheet_name="汇总", index=False)
    except OSError:
        return False, f"无法写入：{output}"

    lines = [f"已合并 {len(frames)} 个文件，写入 {len(detail)} 行明细。"]
    if skipped:
        lines += ["跳过的文件："] + skipped
    return True, "\n".join(lines)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法：python merge_sales.py <文件夹> <输出文件>")
        sys.exit(1)
    ok, message = merge_sales(sys.argv[1], sys.argv[2])
    print(message)
    sys.exit(0 if ok else 1)
