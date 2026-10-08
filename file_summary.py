# 按扩展名统计一个文件夹里的文件数量和大小，并写入该文件夹的 summary.txt。
# 命令：python file_summary.py <文件夹>
# 示例：python file_summary.py ./docs
import os
import sys


def _format_size(num_bytes):
    """1024-based size: integer bytes under 1024, otherwise one decimal."""
    if num_bytes < 1024:
        return f"{num_bytes} B"
    value = float(num_bytes)
    for unit in ("KB", "MB", "GB", "TB"):
        value /= 1024
        if value < 1024 or unit == "TB":
            break
    return f"{value:.1f} {unit}"


def summarize_folder(path):
    """Summarize files in one folder (no recursion) and write summary.txt.

    Returns (True, text) after a successful write, or (False, message)
    when the path is missing, not a directory, or cannot be read or written.
    """
    try:
        if not os.path.exists(path):
            return False, f"找不到路径：{path}"
        if not os.path.isdir(path):
            return False, f"这不是文件夹：{path}"
        names = os.listdir(path)
    except PermissionError:
        return False, f"没有权限读取：{path}"
    except OSError:
        return False, f"无法读取：{path}"

    groups = {}
    for name in names:
        if name == "summary.txt":
            continue
        full = os.path.join(path, name)
        try:
            if os.path.isdir(full):
                continue
            size = os.path.getsize(full)
        except (PermissionError, OSError):
            continue
        _root, ext = os.path.splitext(name)
        label = ext.lower() if ext else "（无扩展名）"
        bucket = groups.setdefault(label, {"count": 0, "size": 0})
        bucket["count"] += 1
        bucket["size"] += size

    if not groups:
        text = "该文件夹里没有文件。\n"
    else:
        rows = sorted(
            groups.items(),
            key=lambda item: (-item[1]["size"], item[0]),
        )
        lines = [
            f"{label}  {info['count']} 个, {_format_size(info['size'])}"
            for label, info in rows
        ]
        text = "\n".join(lines) + "\n"

    out_path = os.path.join(path, "summary.txt")
    try:
        with open(out_path, "w", encoding="utf-8") as handle:
            handle.write(text)
    except PermissionError:
        return False, f"没有权限写入：{out_path}"
    except OSError:
        return False, f"无法写入：{out_path}"
    return True, text


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法：python file_summary.py <文件夹>")
        sys.exit(1)
    folder = sys.argv[1]
    ok, message = summarize_folder(folder)
    print(message)
    sys.exit(0 if ok else 1)
