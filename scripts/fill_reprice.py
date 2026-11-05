#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把京喜「导出商品JX-SKU信息」写入模板底表，并按 PQ 逻辑把每个SPU(商品编码)下
京东价最低的 SKU 汇总进改嫁表。

复现的 PQ 逻辑（模板 assets/PQ模板.xlsx 里查询 “0”）:
  1. 读底表 DATA；可控采购价、京东价转数值
  2. 建议改动价格 = 可控采购价 > 1000 ? 可控采购价*1.01 + 2 : 可控采购价 + 2
  3. 可按 商品编码 分组，取组内 京东价 最小的一行
  4. 输出列: 商品编码, SKUID, 商品名称, 京东价, 商品链接, 可控采购价, 供应商店铺, 供应商公司, 建议改动价格
  结果按 可控采购价 降序（模板表 _0 的排序）。

用法:
  python fill_reprice.py <导出商品JX-SKU信息....xlsx> [-o 输出.xlsx] [--template 模板.xlsx]
默认: 模板 = assets/PQ模板.xlsx；输出 = 导出文件同目录 <原名>_改价结果.xlsx
"""

import argparse
import shutil
import sys
from pathlib import Path

import openpyxl

SHEET_RESULT = "改嫁表"   # 模板里就叫这个名字（用户原始表名，勿改）
SHEET_BASE = "底表"
RESULT_TABLE = "_0"
BASE_TABLE = "DATA"

REQUIRED_COLS = ["SKUID", "商品编码", "商品名称", "京东价", "商品链接",
                 "可控采购价", "供应商店铺", "供应商公司"]


def to_num(v):
    """把单元格值转 float；'--'/空/无法解析 -> None（对应 PQ 里转数值失败）。"""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s in ("", "--"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def suggest_price(cost):
    if cost is None:
        return None
    return round(cost * 1.01 + 2, 2) if cost > 1000 else round(cost + 2, 2)


def read_export(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        raise SystemExit(f"导出文件为空: {path}")
    headers = [str(h).strip() if h is not None else "" for h in rows[0]]
    idx = {h: i for i, h in enumerate(headers)}
    missing = [c for c in REQUIRED_COLS if c not in idx]
    if missing:
        raise SystemExit(f"导出文件缺少必需列: {missing}\n实际列: {headers}")
    data = [r for r in rows[1:] if any(v not in (None, "") for v in r)]
    return headers, idx, data


def pick_cheapest(idx, data):
    """按 商品编码 分组，每组取 京东价 最小的行；组内全部无价格则取首行。"""
    best = {}
    order = []
    for r in data:
        key = str(r[idx["商品编码"]] or "").strip()
        if not key:
            continue
        price = to_num(r[idx["京东价"]])
        if key not in best:
            best[key] = (price, r)
            order.append(key)
        else:
            cur_price, cur_row = best[key]
            if cur_price is None and price is not None:
                best[key] = (price, r)
            elif price is not None and cur_price is not None and price < cur_price:
                best[key] = (price, r)
    groups = []
    for key in order:
        price, row = best[key]
        cost = to_num(row[idx["可控采购价"]])
        groups.append({
            "商品编码": key,
            "SKUID": str(row[idx["SKUID"]] or ""),
            "商品名称": str(row[idx["商品名称"]] or ""),
            "京东价": price,
            "商品链接": str(row[idx["商品链接"]] or ""),
            "可控采购价": cost,
            "供应商店铺": str(row[idx["供应商店铺"]] or ""),
            "供应商公司": str(row[idx["供应商公司"]] or ""),
            "建议改动价格": suggest_price(cost),
        })
    # 模板表 _0 排序: 可控采购价 降序，空值排最后
    groups.sort(key=lambda g: (g["可控采购价"] is None,
                               -(g["可控采购价"] or 0)))
    return groups


def write_sheet(ws, table_name, n_cols, rows_values):
    """清空表区域并重写；同步表格 ref。"""
    if ws.max_row > 1:
        ws.delete_rows(2, ws.max_row)  # 删除第2行到末尾的所有旧数据
    for i, row in enumerate(rows_values, start=2):
        for j, v in enumerate(row, start=1):
            ws.cell(row=i, column=j, value=v)
    if table_name in ws.tables:
        last = max(1 + len(rows_values), 2)
        from openpyxl.utils import get_column_letter
        ws.tables[table_name].ref = f"A1:{get_column_letter(n_cols)}{last}"
    # 表头已有的自动筛选范围同步
    if ws.auto_filter.ref:
        from openpyxl.utils import get_column_letter
        last = max(1 + len(rows_values), 2)
        ws.auto_filter.ref = f"A1:{get_column_letter(ws.max_column)}{last}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("export", help="京喜导出商品JX-SKU信息 xlsx")
    ap.add_argument("-o", "--output", default=None)
    ap.add_argument("--template", default=str(Path(__file__).resolve().parent.parent / "assets" / "PQ模板.xlsx"))
    args = ap.parse_args()

    export = Path(args.export)
    out = Path(args.output) if args.output else export.with_name(export.stem + "_改价结果.xlsx")
    # 结果文件被 Excel 占用时换带时间戳的文件名，而不是报错退出
    if out.exists():
        try:
            with open(out, "ab"):
                pass
        except PermissionError:
            from datetime import datetime
            out = out.with_name(out.stem + "_" + datetime.now().strftime("%H%M%S") + out.suffix)
            print(f"原结果文件正被占用（可能在 Excel 里开着），改存: {out.name}")

    headers, idx, data = read_export(export)
    groups = pick_cheapest(idx, data)

    shutil.copy(args.template, out)
    wb = openpyxl.load_workbook(out)

    # 底表: 原样写入导出行（含表头行不动），模板第31列「列1」留空
    ws_base = wb[SHEET_BASE]
    write_sheet(ws_base, BASE_TABLE, 31, [tuple(r) for r in data])

    # 改嫁表: 保持模板原有 10 列表头，重写数据（旧的 L:O 草稿区一并清掉）
    ws_res = wb[SHEET_RESULT]
    result_rows = [(
        g["商品编码"], g["SKUID"], g["商品编码"], g["商品名称"], g["京东价"],
        g["商品链接"], g["可控采购价"], g["供应商店铺"], g["供应商公司"], g["建议改动价格"],
    ) for g in groups]
    write_sheet(ws_res, RESULT_TABLE, 10, result_rows)

    wb.save(out)
    wb.close()
    print(f"底表写入 {len(data)} 行 SKU；改嫁表输出 {len(groups)} 个 商品编码 的最便宜 SKU")
    print(f"输出: {out}")
    print(f"RESULT_FILE={out}")


if __name__ == "__main__":
    main()
