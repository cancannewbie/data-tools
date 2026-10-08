# 合并文件夹中的销售表，生成带明细和汇总两个工作表的 Excel。
# 命令：python merge_sales.py <文件夹> <输出.xlsx>
# 示例：python merge_sales.py ./sales ./merged.xlsx
import glob
import os
import sys
import tempfile

import pandas as pd

REQUIRED_COLUMNS = ["日期", "商品", "数量", "单价"]
DETAIL_COLUMNS = ["日期", "商品", "数量", "单价", "金额", "来源文件"]
SUMMARY_COLUMNS = ["商品", "总数量", "总金额"]


def _short_error(exc):
    text = " ".join(str(exc).split())
    if not text:
        text = exc.__class__.__name__
    if len(text) > 160:
        text = text[:157] + "..."
    return text


def _normalize_columns(frame):
    frame = frame.copy()
    frame.columns = [
        column.strip() if isinstance(column, str) else column
        for column in frame.columns
    ]
    return frame


def _coerce_number(series):
    """Turn blanks and non-numeric values into NaN; keep real numbers."""

    def clean(value):
        if isinstance(value, str):
            text = value.strip()
            return None if text == "" else text
        return value

    return pd.to_numeric(series.map(clean), errors="coerce")


def _path_key(path):
    """Absolute path key that matches even when the file does not exist yet."""
    return os.path.normcase(os.path.abspath(path))


def _list_workbooks(folder, output):
    """xlsx files in this folder only. Lock files (~$...) and `output` are ignored."""
    output_key = _path_key(output)
    pattern = os.path.join(glob.escape(folder), "*")
    names = []
    for path in glob.glob(pattern):
        name = os.path.basename(path)
        if name.startswith("~$"):
            continue
        if os.path.splitext(name)[1].lower() != ".xlsx":
            continue
        if not os.path.isfile(path):
            continue
        if _path_key(path) == output_key:
            continue
        names.append(name)
    names.sort()
    return names


def _read_sales_file(path, name):
    """Return (frame, None) or (None, skip reason)."""
    try:
        sheets = pd.read_excel(path, sheet_name=None, engine="openpyxl")
    except Exception as exc:
        return None, f"打不开：{_short_error(exc)}"

    if not isinstance(sheets, dict) or not sheets:
        return None, "打不开：工作簿里没有工作表"

    parts = []
    missing = []
    try:
        for frame in sheets.values():
            if not isinstance(frame, pd.DataFrame):
                continue
            frame = _normalize_columns(frame)
            if len(frame.columns) == 0 and len(frame.index) == 0:
                continue
            absent = [
                column for column in REQUIRED_COLUMNS if column not in frame.columns
            ]
            if absent:
                for column in absent:
                    if column not in missing:
                        missing.append(column)
                continue
            part = frame.loc[:, REQUIRED_COLUMNS].copy()
            part["日期"] = pd.to_datetime(part["日期"], errors="coerce", format="mixed")
            part["来源文件"] = name
            part["数量"] = _coerce_number(part["数量"])
            part["单价"] = _coerce_number(part["单价"])
            part["金额"] = part["数量"] * part["单价"]
            parts.append(part.loc[:, DETAIL_COLUMNS])
    except Exception as exc:
        return None, f"打不开：{_short_error(exc)}"

    if missing:
        return None, "缺列：" + "、".join(missing)
    if not parts:
        return None, "缺列：" + "、".join(REQUIRED_COLUMNS)
    return pd.concat(parts, ignore_index=True), None


def _skipped_lines(skipped):
    lines = ["跳过的文件："]
    for name, reason in skipped:
        lines.append(f"- {name}：{reason}")
    return lines


def _remove_quiet(path):
    try:
        if path and os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


def _write_workbook(detail, summary, output):
    """Write both sheets. On failure, leave no partial file behind."""
    if os.path.isdir(output):
        return f"无法写入：{output}"
    parent = os.path.dirname(os.path.abspath(output))
    if parent and not os.path.isdir(parent):
        return f"无法写入：{output}"

    partial = None
    try:
        # ExcelWriter accepts only an .xlsx path; replace into `output` after.
        handle = tempfile.NamedTemporaryFile(
            suffix=".xlsx", dir=parent, delete=False
        )
        partial = handle.name
        handle.close()
        with pd.ExcelWriter(partial, engine="openpyxl") as writer:
            detail.to_excel(writer, sheet_name="明细", index=False)
            summary.to_excel(writer, sheet_name="汇总", index=False)
        os.replace(partial, output)
        partial = None
    except PermissionError:
        _remove_quiet(partial)
        return f"没有权限写入：{output}"
    except OSError:
        _remove_quiet(partial)
        return f"无法写入：{output}"
    except Exception:
        _remove_quiet(partial)
        return f"无法写入：{output}"
    return None


def _merge_sales(folder, output):
    try:
        if not os.path.exists(folder):
            return False, f"找不到路径：{folder}"
        if not os.path.isdir(folder):
            return False, f"这不是文件夹：{folder}"
        os.listdir(folder)
    except PermissionError:
        return False, f"没有权限读取：{folder}"
    except OSError:
        return False, f"无法读取：{folder}"

    names = _list_workbooks(folder, output)
    if not names:
        return False, "该文件夹里没有可合并的 .xlsx 文件。"

    frames = []
    skipped = []
    for name in names:
        frame, reason = _read_sales_file(os.path.join(folder, name), name)
        if reason is not None:
            skipped.append((name, reason))
            continue
        frames.append(frame)

    if not frames:
        lines = ["没有可用的 .xlsx 文件。"]
        lines.extend(_skipped_lines(skipped))
        return False, "\n".join(lines)

    detail = pd.concat(frames, ignore_index=True)
    detail = detail.loc[:, DETAIL_COLUMNS]

    before_dup = len(detail)
    detail = detail.drop_duplicates()
    dup_removed = before_dup - len(detail)

    before_empty = len(detail)
    detail = detail.dropna(subset=["数量", "单价"])
    empty_removed = before_empty - len(detail)
    detail = detail.reset_index(drop=True)

    summary = (
        detail.groupby("商品", as_index=False, dropna=False)
        .agg(总数量=("数量", "sum"), 总金额=("金额", "sum"))
        .sort_values("总金额", ascending=False, kind="mergesort")
        .reset_index(drop=True)
    )
    summary = summary.loc[:, SUMMARY_COLUMNS]

    write_error = _write_workbook(detail, summary, output)
    if write_error is not None:
        return False, write_error

    lines = [
        f"已合并 {len(frames)} 个文件，写入 {len(detail)} 行明细。",
        f"去除完全重复行 {dup_removed} 行。",
        f"去除数量或单价为空的行 {empty_removed} 行。",
    ]
    if skipped:
        lines.extend(_skipped_lines(skipped))
    else:
        lines.append("没有跳过任何文件。")
    return True, "\n".join(lines) + "\n"


def merge_sales(folder, output):
    """Merge sales workbooks in one folder into 明细 and 汇总 sheets.

    Returns (True, report) after a successful write, or (False, message)
    when the folder is missing, not a directory, unreadable, has no usable
    workbook, or the output cannot be written. Does not raise for bad input.
    """
    try:
        return _merge_sales(folder, output)
    except Exception as exc:
        return False, f"无法完成合并：{_short_error(exc)}"


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("用法：python merge_sales.py <文件夹> <输出文件>")
        sys.exit(1)
    ok, message = merge_sales(sys.argv[1], sys.argv[2])
    print(message)
    sys.exit(0 if ok else 1)
