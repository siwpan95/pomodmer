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

方法分两组：
  - 表格（原）：create_workbook/add_sheet/write_rows/write_cell/style_header/
    set_column_width/read_sheet/save/close
  - 地图/渲染（新增）：set_cell_fill/merge_cells/set_cell_text/set_comment/
    set_borders/read_map/render_map_png/render_token
  - 程序化生成（P1）：gen_dungeon
单进程、单 worker：所有方法共享同一 WORKBOOKS 缓存，避免多进程对同一文件缓存冲突。
"""

import sys
import os
import io
import json
import traceback
import contextlib
import random

from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment


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


# --------------------------------------------------------------------------- #
# 颜色工具
# --------------------------------------------------------------------------- #
def _norm_argb(h):
    """归一化为 8 位 ARGB（openpyxl 接受 'FF4472C4' 或 '4472C4'）。"""
    h = (h or "").lstrip("#").upper()
    if len(h) == 6:
        return "FF" + h
    if len(h) == 8:
        return h
    return h


def _hex_to_rgb(rgb):
    """ARGB/RGB 字符串 -> (r, g, b)；无法解析时回退浅灰。"""
    if not rgb:
        return (235, 235, 235)
    h = rgb.lstrip("#")
    if len(h) == 8:
        h = h[2:]
    if len(h) != 6:
        return (235, 235, 235)
    return tuple(int(h[i : i + 2], 16) for i in (0, 2, 4))


def _cell_fill_rgb(cell):
    """读取单元格实色填充的 ARGB（无填充返回 None）。"""
    pat = cell.fill
    if pat and pat.patternType == "solid":
        rgb = pat.fgColor.rgb
        if rgb and rgb not in ("00000000", None):
            return rgb
    return None


# --------------------------------------------------------------------------- #
# 工作簿/表获取
# --------------------------------------------------------------------------- #
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


# --------------------------------------------------------------------------- #
# 方法分发
# --------------------------------------------------------------------------- #
def handle(method, params):
    p = params or {}
    path = p.get("path")

    # ----------------------------- 原表格方法 ----------------------------- #
    if method == "ping":
        import openpyxl

        return {"python": sys.version.split()[0], "openpyxl": openpyxl.__version__}

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

    # ----------------------------- 地图方法 ----------------------------- #
    if method == "set_cell_fill":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        cell = ws.cell(row=int(p["row"]), column=int(p["col"]))
        cell.fill = PatternFill(
            start_color=_norm_argb(p.get("color")),
            end_color=_norm_argb(p.get("color")),
            fill_type="solid",
        )
        return {"ok": True}

    if method == "merge_cells":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        ws.merge_cells(
            start_row=int(p["r1"]),
            start_column=int(p["c1"]),
            end_row=int(p["r2"]),
            end_column=int(p["c2"]),
        )
        return {"ok": True}

    if method == "set_cell_text":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        ws.cell(row=int(p["row"]), column=int(p["col"]), value=p.get("text"))
        return {"ok": True}

    if method == "set_comment":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        cell = ws.cell(row=int(p["row"]), column=int(p["col"]))
        cell.comment = Comment(p.get("text", ""), "DM")
        return {"ok": True}

    if method == "set_borders":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        r1, c1, r2, c2 = int(p["r1"]), int(p["c1"]), int(p["r2"]), int(p["c2"])
        color = p.get("color", "FF1F1F1F")
        style = p.get("style", "thick")
        side = Side(style=style, color=color)
        for c in range(c1, c2 + 1):
            t = ws.cell(row=r1, column=c)
            b = ws.cell(row=r2, column=c)
            t.border = Border(top=side, left=t.border.left, right=t.border.right, bottom=t.border.bottom)
            b.border = Border(bottom=side, left=b.border.left, right=b.border.right, top=b.border.top)
        for r in range(r1, r2 + 1):
            l = ws.cell(row=r, column=c1)
            rr = ws.cell(row=r, column=c2)
            l.border = Border(left=side, top=l.border.top, right=l.border.right, bottom=l.border.bottom)
            rr.border = Border(right=side, top=rr.border.top, left=rr.border.left, bottom=rr.border.bottom)
        return {"ok": True}

    if method == "read_map":
        wb = _get_wb(path)
        ws = _sheet(wb, p.get("sheet"))
        rows = []
        for r in range(1, ws.max_row + 1):
            row = []
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(row=r, column=c)
                row.append({"fill": _cell_fill_rgb(cell), "text": cell.value})
            rows.append(row)
        return {"rows": rows, "max_row": ws.max_row, "max_col": ws.max_column}

    if method == "render_map_png":
        return _render_map_png(p)

    if method == "render_token":
        return _render_token(p)

    # ----------------------------- 程序化生成（P1） ----------------------------- #
    if method == "gen_dungeon":
        return _gen_dungeon(p)

    raise ValueError("unknown method: " + str(method))


# --------------------------------------------------------------------------- #
# 渲染：网格 -> PNG（Pillow，懒导入，核心路径不依赖）
# --------------------------------------------------------------------------- #
def _render_map_png(p):
    from PIL import Image, ImageDraw

    path = p.get("path")
    wb = _get_wb(path)
    ws = _sheet(wb, p.get("sheet"))
    px = int(p.get("px", 100))
    grid = bool(p.get("grid", False))
    out = p.get("out")

    max_row = ws.max_row
    max_col = ws.max_column
    img = Image.new("RGB", (max_col * px, max_row * px), (255, 255, 255))
    draw = ImageDraw.Draw(img)
    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            rgb = _cell_fill_rgb(cell)
            color = _hex_to_rgb(rgb) if rgb else (235, 235, 235)
            x0 = (c - 1) * px
            y0 = (r - 1) * px
            draw.rectangle([x0, y0, x0 + px - 1, y0 + px - 1], fill=color)
    if grid:
        line = (150, 150, 150)
        for c in range(max_col + 1):
            x = c * px
            draw.line([(x, 0), (x, max_row * px)], fill=line, width=1)
        for r in range(max_row + 1):
            y = r * px
            draw.line([(0, y), (max_col * px, y)], fill=line, width=1)
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    img.save(out)
    return {
        "ok": True,
        "width": max_col * px,
        "height": max_row * px,
        "cols": max_col,
        "rows": max_row,
    }


def _render_token(p):
    from PIL import Image, ImageDraw, ImageFont

    text = p.get("text", "?")
    out = p.get("out")
    size = int(p.get("size", 400))
    color = _hex_to_rgb(p.get("color", "FF4F81BD"))
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse([size * 0.08, size * 0.08, size * 0.92, size * 0.92], fill=color + (255,))
    try:
        font = ImageFont.truetype("arial.ttf", size // 5)
    except Exception:
        font = ImageFont.load_default()
    draw.text((size / 2, size / 2), text, fill=(255, 255, 255, 255), font=font, anchor="mm")
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    img.save(out)
    return {"ok": True, "width": size, "height": size}


# --------------------------------------------------------------------------- #
# 程序化地牢生成（DonJon/Dungeon Flow 式；仅泛型文本，不抄官方资料）
# --------------------------------------------------------------------------- #
_ROOM_NAMES = [
    "大厅", "守卫室", "储藏间", "祭坛", "牢房", "书房", "厨房", "寝室",
    "宝库", "蛛网厅", "碑厅", "渗水窟", "王座厅", "兵械库", "静室", "裂隙厅",
]
_MONSTERS = [
    "狗头人", "哥布林", "骷髅兵", "僵尸", "史莱姆", "兽人", "食人魔",
    "暗影", "石像鬼", "巨蜘蛛", "蝠魔", "火矮人",
]
_TRAPS = ["地板夹层", "喷毒针", "落石", "魔法符文", "绊索标枪", "伪装的井"]
_TREASURE = ["金币若干", "宝石一枚", "秘银锭", "古老卷轴", "圣徽", "毒药瓶", "符文武器"]


def _cr_for_level(level):
    # 轻量 5e：CR 随队伍等级缓增（仅用于生成泛型遭遇文本，不引用官方 stat block）
    return max(0, round(level / 3) - 1)


def _gen_dungeon(p):
    path = p.get("path")
    cols = int(p.get("cols", 30))
    rows = int(p.get("rows", 20))
    n_rooms = int(p.get("rooms", 6))
    party_level = int(p.get("party_level", 3))
    sheet = p.get("sheet", "Map")

    wb = _get_wb(path)
    ws = _sheet(wb, sheet)
    # 清空既有内容
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c, value=None)
            ws.cell(row=r, column=c).fill = PatternFill(fill_type=None)

    STONE = "FF3A3A3A"
    FLOOR = "FFF2E2C0"
    CORR = "FFD9C9A0"
    DOOR = "FF8B5A2B"
    WATER = "FF4F81BD"

    # 整张填石墙
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            ws.cell(row=r, column=c).fill = PatternFill(
                start_color=STONE, end_color=STONE, fill_type="solid"
            )

    rng = random.Random(p.get("seed"))
    rooms = []
    attempts = 0
    while len(rooms) < n_rooms and attempts < n_rooms * 20:
        attempts += 1
        w = rng.randint(3, 7)
        h = rng.randint(3, 6)
        r1 = rng.randint(2, rows - h - 1)
        c1 = rng.randint(2, cols - w - 1)
        r2 = r1 + h - 1
        c2 = c1 + w - 1
        # 不与已有房间重叠（留 1 格墙）
        overlap = any(
            not (c2 + 1 < o[1] or c1 - 1 > o[3] or r2 + 1 < o[0] or r1 - 1 > o[2])
            for o in rooms
        )
        if overlap:
            continue
        rooms.append((r1, c1, r2, c2))
        # 地板
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(row=r, column=c).fill = PatternFill(
                    start_color=FLOOR, end_color=FLOOR, fill_type="solid"
                )
        # 门（随机一边中点）
        side = rng.choice(["top", "bottom", "left", "right"])
        if side == "top":
            dr, dc = r1, rng.randint(c1, c2)
        elif side == "bottom":
            dr, dc = r2, rng.randint(c1, c2)
        elif side == "left":
            dr, dc = rng.randint(r1, r2), c1
        else:
            dr, dc = rng.randint(r1, r2), c2
        ws.cell(row=dr, column=dc).fill = PatternFill(
            start_color=DOOR, end_color=DOOR, fill_type="solid"
        )

    # 走廊：按顺序把相邻房间用 L 形走廊连起来
    def carve(r0, c0, r1, c1):
        for c in range(min(c0, c1), max(c0, c1) + 1):
            ws.cell(row=r0, column=c).fill = PatternFill(
                start_color=CORR, end_color=CORR, fill_type="solid"
            )
        for r in range(min(r0, r1), max(r0, r1) + 1):
            ws.cell(row=r, column=c1).fill = PatternFill(
                start_color=CORR, end_color=CORR, fill_type="solid"
            )

    for i in range(1, len(rooms)):
        a = rooms[i - 1]
        b = rooms[i]
        carve((a[0] + a[2]) // 2, (a[1] + a[3]) // 2, (b[0] + b[2]) // 2, (b[1] + b[3]) // 2)

    # 偶尔放个水池
    if rooms:
        rr = rng.choice(rooms)
        wr = rng.randint(rr[0], rr[2])
        wc = rng.randint(rr[1], rr[3])
        ws.cell(row=wr, column=wc).fill = PatternFill(
            start_color=WATER, end_color=WATER, fill_type="solid"
        )

    # 数据表：Rooms / Encounters / Loot / NPCs
    rooms_ws = _sheet(wb, "Rooms")
    rooms_ws.append(["id", "name", "x", "y", "w", "h", "description", "encounter", "traps", "treasure"])
    enc_ws = _sheet(wb, "Encounters")
    enc_ws.append(["room_id", "level", "line"])
    loot_ws = _sheet(wb, "Loot")
    loot_ws.append(["room_id", "item"])

    cr = _cr_for_level(party_level)
    for i, (r1, c1, r2, c2) in enumerate(rooms, start=1):
        name = _ROOM_NAMES[(i - 1) % len(_ROOM_NAMES)] + " " + str(i)
        n_mon = rng.randint(1, 4)
        monsters = ", ".join(rng.sample(_MONSTERS, min(n_mon, len(_MONSTERS))))
        trap = rng.choice(_TRAPS) if rng.random() < 0.5 else "无"
        treasure = rng.choice(_TREASURE)
        desc = "一间%s，弥漫着潮湿的石腥味。" % ("昏暗的" if rng.random() < 0.5 else "宽敞的")
        rooms_ws.append([i, name, c1, r1, c2 - c1 + 1, r2 - r1 + 1, desc, monsters + " (CR %d)" % cr, trap, treasure])
        enc_ws.append([i, party_level, "遭遇：%s（约 CR %d），建议先手控制。" % (monsters, cr)])
        loot_ws.append([i, treasure])
        # 房间标签 + 批注
        ws.cell(row=r1, column=c1, value=name)
        ws.cell(row=r1, column=c1).comment = Comment(
            "遭遇：%s\n陷阱：%s\n战利品：%s" % (monsters, trap, treasure), "DM"
        )
        # 房间外墙
        _side = Side(style="thick", color="FF1F1F1F")
        for c in range(c1, c2 + 1):
            t = ws.cell(row=r1, column=c)
            b = ws.cell(row=r2, column=c)
            t.border = Border(top=_side, left=t.border.left, right=t.border.right, bottom=t.border.bottom)
            b.border = Border(bottom=_side, left=b.border.left, right=b.border.right, top=b.border.top)
        for r in range(r1, r2 + 1):
            l = ws.cell(row=r, column=c1)
            rr = ws.cell(row=r, column=c2)
            l.border = Border(left=_side, top=l.border.top, right=l.border.right, bottom=l.border.bottom)
            rr.border = Border(right=_side, top=rr.border.top, left=rr.border.left, bottom=rr.border.bottom)

    return {
        "ok": True,
        "rooms": len(rooms),
        "cols": cols,
        "rows": rows,
        "sheets": wb.sheetnames,
    }


# --------------------------------------------------------------------------- #
# 主循环
# --------------------------------------------------------------------------- #
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
