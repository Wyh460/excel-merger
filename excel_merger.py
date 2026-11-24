# -*- coding: utf-8 -*-
"""
Excel 批量合并工具 (excel_merger)

将指定文件夹下的所有 Excel 文件(.xlsx/.xls)合并为一个文件，
支持：自动对齐表头、去除重复行、保留来源文件名、输出 xlsx 或 csv。

用法:
    python excel_merger.py 输入文件夹 -o 输出.xlsx
    python excel_merger.py 输入文件夹 -o 输出.csv --sheet "Sheet1" --dedup --verbose

依赖: openpyxl (xlsx), xlrd (旧版 xls 可选)
"""

import argparse
import os
import sys
from collections import OrderedDict

try:
    from openpyxl import Workbook, load_workbook
except ImportError:
    sys.stderr.write("缺少依赖 openpyxl，请先执行: pip install openpyxl\n")
    sys.exit(1)


def read_excel(path, sheet_name=None):
    """读取单个 Excel 文件，返回 (表头列表, 数据行列表)。"""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    header = [str(c).strip() if c is not None else f"col{i}" for i, c in enumerate(rows[0], 1)]
    data = []
    for r in rows[1:]:
        if r is None or all(c is None or str(c).strip() == "" for c in r):
            continue  # 跳过空行
        data.append(list(r))
    wb.close()
    return header, data


def merge_files(folder, sheet_name=None, dedup=False):
    """合并文件夹下所有 Excel，返回 (最终表头, 合并行)。"""
    files = [f for f in os.listdir(folder)
             if f.lower().endswith((".xlsx", ".xls")) and not f.startswith("~$")]
    files.sort()
    if not files:
        sys.stderr.write("文件夹内没有 Excel 文件\n")
        sys.exit(1)

    all_headers = []
    all_rows = []
    for fname in files:
        path = os.path.join(folder, fname)
        header, data = read_excel(path, sheet_name)
        if not header:
            continue
        all_headers.append(header)
        for row in data:
            all_rows.append(row + [fname])  # 追加来源文件名

    # 以第一个文件的表头为基准，做表头对齐（简单策略：取并集，按首次出现顺序）
    merged_header = []
    for h in all_headers:
        for col in h:
            if col not in merged_header:
                merged_header.append(col)
    merged_header.append("来源文件")

    # 数据行对齐到最终表头（按各文件的原始表头顺序填充）
    # 这里重新按文件处理，保证列对齐正确
    final_rows = []
    seen = set()
    for fname in files:
        path = os.path.join(folder, fname)
        header, data = read_excel(path, sheet_name)
        if not header:
            continue
        col_index = {c: i for i, c in enumerate(header)}
        for row in data:
            out_row = []
            for col in merged_header[:-1]:
                out_row.append(row[col_index[col]] if col in col_index else None)
            out_row.append(fname)
            key = tuple(str(v) for v in out_row)
            if dedup and key in seen:
                continue
            seen.add(key)
            final_rows.append(out_row)

    return merged_header, final_rows


def write_output(path, header, rows):
    """写出 xlsx 或 csv。"""
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        import csv
        with open(path, "w", newline="", encoding="utf-8-sig") as f:
            w = csv.writer(f)
            w.writerow(header)
            w.writerows(rows)
    else:
        wb = Workbook()
        ws = wb.active
        ws.append(header)
        for r in rows:
            ws.append(r)
        wb.save(path)
    print(f"已输出 {len(rows)} 行到 {path}")


def main():
    p = argparse.ArgumentParser(description="Excel 批量合并工具")
    p.add_argument("folder", help="输入文件夹（含多个 Excel）")
    p.add_argument("-o", "--output", required=True, help="输出文件 (.xlsx 或 .csv)")
    p.add_argument("--sheet", default=None, help="指定工作表名（默认第一个）")
    p.add_argument("--dedup", action="store_true", help="去除完全重复的行")
    p.add_argument("--verbose", action="store_true", help="打印每个文件处理情况")
    args = p.parse_args()

    if not os.path.isdir(args.folder):
        sys.stderr.write(f"文件夹不存在: {args.folder}\n")
        sys.exit(1)

    if args.verbose:
        files = [f for f in os.listdir(args.folder) if f.lower().endswith((".xlsx", ".xls"))]
        print(f"发现 {len(files)} 个文件:", ", ".join(files))

    header, rows = merge_files(args.folder, args.sheet, args.dedup)
    write_output(args.output, header, rows)


if __name__ == "__main__":
    main()
