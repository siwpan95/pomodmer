#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cordis 表格应用 —— Python 侧 OpenPyXL 工作者（JSON-RPC over stdio）

协议纪律（见 cordis-host-python-bridge 技能）：
  - stdout 只走一行一对象的 JSON-RPC 2.0，所有日志/异常栈走 stderr；
  - 被调库若直接 print 会污染协议通道，故处理器执行期间重定向 stdout->stderr；
  - 每个方法 try/except 兜住，返回 JSON-RPC error（-32601 未知方法 / -32000 执行异常），
    保证子进程不死。

工作簿以 path 为键缓存在内存，save 时落盘；close 时保存并移出缓存。
"""

import sys
import os
import io
import json
import traceback
import contextlib

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side


def _setup_stdio_utf8():
    """三条标准流统一为 UTF-8。

    中文 Windows 控制台的 Python 默认 stdio 编码是 GBK(cp936)：本进程用
    ensure_ascii=False 写一行一 JSON-RPC（见 main），非 ASCII 字符会被编成
    GBK 字节，而 Node 侧 readline 按 UTF-8 解码，得到孤立代理项（如 \\udc8c），
    于是 openpyxl/lxml 落盘时报 `UnicodeEncodeError: surrogates not allowed`。
    显式 reconfigure 后协议两端都是 UTF-8，跨机器（含 Windows）行为一致。
    """
    for stream in (sys.stdin, sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass  # 流被重定向成不支持 reconfigure 的对象时保持原样


WORKBOOKS = {}  # path -> Workbook


def _get_wb(path):
    if path in WORKBOOKS:
        return WORKBOOKS[path]
    if os.path.exists(path):
        wb = load_workbook(path)
    else:
        wb = Workbook()
    WORKBOOKS[path] = wb
    return wb


def _sheet(wb, name):
    if name:
        if name in wb.sheetnames:
            return wb[name]
        return wb.create_sheet(title=name)
    return wb.active


def handle(method, params):
    p = params or {}
    path = p.get("path")

    if method == "ping":
        import openpyxl
        return {
            "python": sys.version.split()[0],
            "openpyxl": openpyxl.__version__,
        }

    if method == "create_workbook":
        wb = Workbook()
        if p.get("sheet"):
            wb.active.title = p["sheet"]
        WORKBOOKS[path] = wb
        return {"ok": True, "sheets": wb.sheetnames}

    if method == "add_sheet":
        wb = _get_wb(path)
        ws = wb.create_sheet(title=p.get("name", "Sheet"))
        return {"ok": True, "sheets": wb.sheetnames}

    if method == "write_rows":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        start_row = int(p.get("start_row", 1))
        start_col = int(p.get("start_col", 1))
        rows = p.get("rows", [])
        for i, row in enumerate(rows):
            for j, val in enumerate(row):
                # openpyxl 不接受 None 以外的"空"；None 即空单元格
                ws.cell(row=start_row + i, column=start_col + j, value=val)
        return {"ok": True, "rows": len(rows)}

    if method == "write_cell":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        ws.cell(row=int(p["row"]), column=int(p["col"]), value=p.get("value"))
        return {"ok": True}

    if method == "style_header":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        row = int(p.get("row", 1))
        ncols = int(p.get("ncols", 1))
        fill_hex = p.get("fill", "FF4472C4")
        font_color = p.get("font_color", "FFFFFFFF")
        fill = PatternFill(start_color=fill_hex, end_color=fill_hex, fill_type="solid")
        font = Font(bold=True, color=font_color)
        thin = Side(style="thin", color="FFB0B0B0")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row, column=c)
            cell.font = font
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        return {"ok": True}

    if method == "set_column_width":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        col = int(p["col"])
        ws.column_dimensions[get_column_letter(col)].width = float(p.get("width", 15))
        return {"ok": True}

    if method == "read_sheet":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        return {"rows": rows, "max_row": ws.max_row, "max_col": ws.max_column}

    if method == "save":
        wb = WORKBOOKS.get(path)
        if not wb:
            return {"ok": False, "error": "workbook not opened"}
        os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
        wb.save(path)
        return {"ok": True, "path": path}

    if method == "close":
        wb = WORKBOOKS.pop(path, None)
        if wb:
            os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
            wb.save(path)
        return {"ok": True}

    raise ValueError("unknown method: " + str(method))


def main():
    for line in sys.stdin:
        line = line.strip()
        if not line:
            continue
        try:
            req = json.loads(line)
        except Exception:
            continue
        msgid = req.get("id")
        method = req.get("method")
        params = req.get("params")
        try:
            with contextlib.redirect_stdout(sys.stderr):
                result = handle(method, params)
            resp = {"jsonrpc": "2.0", "id": msgid, "result": result}
        except Exception as e:
            resp = {
                "jsonrpc": "2.0",
                "id": msgid,
                "error": {"code": -32000, "message": str(e), "data": traceback.format_exc()},
            }
        sys.stdout.write(json.dumps(resp, ensure_ascii=False) + "\n")
        sys.stdout.flush()


if __name__ == "__main__":
    _setup_stdio_utf8()
    main()
