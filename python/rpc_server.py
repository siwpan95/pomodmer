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

方法分三组：
  - 表格（原）：create_workbook/add_sheet/write_rows/write_cell/style_header/
    set_column_width/read_sheet/save/close
  - 地图/渲染：set_cell_fill/merge_cells/set_cell_text/set_comment/
    set_borders/read_map/render_map_png/render_token
  - 程序化生成 + 互通（P1/P2）：gen_dungeon/export_donjon/import_donjon
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
# 颜色 / 地形工具
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
# 地形调色板（ARGB）与 DonJon 位映射 —— 互通核心
# --------------------------------------------------------------------------- #
# key -> ARGB（8 位）。STONE=墙/未开凿；其余为开凿地或特殊标记。
TERRAIN = {
    "stone":  "FF3A3A3A",  # 石墙 / 未开凿
    "floor":  "FFF2E2C0",  # 房间地板
    "corr":   "FFD9C9A0",  # 走廊
    "door":   "FF8B5A2B",  # 门
    "water":  "FF4F81BD",  # 水池
    "lava":   "FFB3471F",  # 熔岩（危害地形）
    "grass":  "FF7CA64B",  # 苔藓 / 草地面（房间变体）
    "sand":   "FFE6D98A",  # 沙地（房间变体）
    "wood":   "FF8B6B3A",  # 木桥 / 木板（房间变体）
    "rubble": "FF8A8A8A",  # 碎石地面（房间变体）
    "stair":  "FF9ACD32",  # 楼梯（上 / 下）
    "secret": "FF6A5ACD",  # 密门
    "trap":   "FFB22222",  # 陷阱标记（手动地图用，生成器一般不放）
}
ARGB_BY_TERRAIN = dict(TERRAIN)

# DonJon 真实 cell 位标志（来自 donjon.bin.sh dungeon.pl 源码，见 _export_donjon 注释）
DJ_BLOCKED  = 0x00000001
DJ_ROOM     = 0x00000002
DJ_CORR     = 0x00000004
DJ_PERIM    = 0x00000010
DJ_DOOR     = 0x00020000
DJ_SECRET   = 0x00100000
DJ_STAIR_DN = 0x00400000
DJ_STAIR_UP = 0x00800000

# 地形 -> DonJon cell 位。water/lava/grass/sand/wood/rubble 在 DonJon 无对应位，
# 归为 ROOM 开放地（颜色靠导出时的 _cta.palette 保真；其他工具按 ROOM 处理也合理）。
TERRAIN_TO_DJ = {
    "stone":  DJ_BLOCKED,
    "floor":  DJ_ROOM,
    "corr":   DJ_CORR,
    "door":   DJ_DOOR,
    "secret": DJ_DOOR | DJ_SECRET,
    "stair":  DJ_ROOM,          # 楼梯在房间内，视作开放地；方向位另记
    "water":  DJ_ROOM,
    "lava":   DJ_ROOM,
    "grass":  DJ_ROOM,
    "sand":   DJ_ROOM,
    "wood":   DJ_ROOM,
    "rubble": DJ_ROOM,
    "trap":   DJ_ROOM,
}


def _terrain_for_fill(rgb):
    """ARGB -> 地形 key；未知 / 无填充返回 'stone'。"""
    if not rgb:
        return "stone"
    up = rgb.upper()
    for k, v in ARGB_BY_TERRAIN.items():
        if v.upper() == up:
            return k
    return "stone"


def _dj_int_for_terrain(key, stair_dir=0):
    base = TERRAIN_TO_DJ.get(key, DJ_ROOM)
    if key == "stair":
        if stair_dir > 0:
            base |= DJ_STAIR_UP
        elif stair_dir < 0:
            base |= DJ_STAIR_DN
    return base


def _terrain_from_dj(cell):
    """DonJon cell int -> 地形 key（不含 _cta 时的标准回退）。"""
    cell = int(cell)
    if cell & DJ_DOOR:
        return "secret" if (cell & DJ_SECRET) else "door"
    if cell & (DJ_STAIR_UP | DJ_STAIR_DN):
        return "stair"
    if cell & DJ_CORR:
        return "corr"
    if cell & DJ_ROOM:
        return "floor"
    return "stone"


def _paint_border(ws, r1, c1, r2, c2, color="FF1F1F1F", style="thick"):
    """给矩形区域描外墙（与 set_borders 同语义，供生成 / 导入复用）。"""
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


def _clear_sheet(ws):
    """清空整张表的内容与填充（保留表本身）。"""
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(row=r, column=c)
            cell.value = None
            cell.fill = PatternFill(fill_type=None)
            cell.comment = None


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

    # ----------------------------- 程序化生成 + 互通 ----------------------------- #
    if method == "gen_dungeon":
        return _gen_dungeon(p)
    if method == "export_donjon":
        return _export_donjon(p)
    if method == "import_donjon":
        return _import_donjon(p)

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
# 主题房间模板：(名称, 修饰语, 怪物偏好, 战利品偏好, 陷阱概率)
_ROOM_TYPES = [
    ("大厅", "宽敞的", ["狗头人", "哥布林", "兽人"], ["金币若干", "秘银锭"], 0.3),
    ("守卫室", "戒备森严的", ["兽人", "哥布林", "骷髅兵"], ["兵器", "金币若干"], 0.6),
    ("储藏间", "堆满杂物的", ["鼠群", "史莱姆"], ["金币若干", "古老卷轴"], 0.3),
    ("祭坛", "弥漫香灰的", ["暗影", "石像鬼"], ["圣徽", "宝石一枚"], 0.7),
    ("牢房", "阴冷的", ["僵尸", "骷髅兵"], ["钥匙", "金币若干"], 0.5),
    ("书房", "积满灰尘的", ["暗影", "蝠魔"], ["古老卷轴", "秘银锭"], 0.4),
    ("厨房", "散发腐臭的", ["食人魔", "史莱姆"], ["食物", "金币若干"], 0.3),
    ("寝室", "凌乱的", ["哥布林", "僵尸"], ["私人物品", "金币若干"], 0.3),
    ("宝库", "布满机关的", ["火矮人", "石像鬼"], ["宝石一枚", "符文武器"], 0.8),
    ("蛛网厅", "蛛网密布的", ["巨蜘蛛", "暗影"], ["虫蛀宝箱", "毒药瓶"], 0.6),
    ("碑厅", "立满墓碑的", ["骷髅兵", "僵尸"], ["陪葬品", "古老卷轴"], 0.5),
    ("渗水窟", "渗着水的", ["史莱姆", "蝠魔"], ["湿漉漉的金币", "毒药瓶"], 0.4),
    ("王座厅", "气势恢宏的", ["兽人", "火矮人", "暗影"], ["王冠", "宝石一枚"], 0.9),
    ("兵械库", "兵器林立的", ["兽人", "哥布林"], ["符文武器", "兵器"], 0.5),
    ("静室", "异常安静的", ["暗影", "石像鬼"], ["冥想物", "圣徽"], 0.3),
    ("裂隙厅", "裂开深渊的", ["蝠魔", "暗影"], ["深渊回响", "符文武器"], 0.6),
    ("图书馆", "书架林立的", ["暗影", "石像鬼"], ["禁书", "古老卷轴"], 0.5),
    ("锻炉", "余温尚存的", ["火矮人", "兽人"], ["锻件", "秘银锭"], 0.6),
    ("圣堂", "彩窗斑驳的", ["石像鬼", "暗影"], ["圣徽", "宝石一枚"], 0.6),
    ("工坊", "工具散落的", ["哥布林", "火矮人"], ["半成品", "金币若干"], 0.4),
    ("兽栏", "腥气扑鼻的", ["兽人", "食人魔"], ["鞍具", "金币若干"], 0.4),
    ("陵墓", "棺椁排列的", ["僵尸", "骷髅兵"], ["陪葬珍宝", "圣徽"], 0.7),
]
_ROOM_NAMES_FALLBACK = ["大厅", "守卫室", "储藏间", "祭坛", "牢房", "书房", "厨房", "寝室",
                        "宝库", "蛛网厅", "碑厅", "渗水窟", "王座厅", "兵械库", "静室", "裂隙厅"]
_GENERIC_MONSTERS = [
    "狗头人", "哥布林", "骷髅兵", "僵尸", "史莱姆", "兽人", "食人魔",
    "暗影", "石像鬼", "巨蜘蛛", "蝠魔", "火矮人", "鼠群",
]
_TRAPS = ["地板夹层", "喷毒针", "落石", "魔法符文", "绊索标枪", "伪装的井", "压发弩", "滚轴巨石"]
_TREASURE = ["金币若干", "宝石一枚", "秘银锭", "古老卷轴", "圣徽", "毒药瓶", "符文武器", "王冠", "禁书", "陪葬珍宝"]
# 房间内部变异地面（纯装饰，增加视觉层次；导出时归为 ROOM 开放地，_cta 保真）
_VARIANT_TERRAINS = ["grass", "sand", "wood", "rubble"]


def _cr_for_level(level):
    # 轻量 5e：CR 随队伍等级缓增（仅用于生成泛型遭遇文本，不引用官方 stat block）
    return max(0, round(level / 3) - 1)


def _fill(ws, r, c, key):
    ws.cell(row=r, column=c).fill = PatternFill(
        start_color=_norm_argb(ARGB_BY_TERRAIN[key]),
        end_color=_norm_argb(ARGB_BY_TERRAIN[key]),
        fill_type="solid",
    )


def _gen_dungeon(p):
    path = p.get("path")
    cols = int(p.get("cols", 30))
    rows = int(p.get("rows", 20))
    n_rooms = int(p.get("rooms", 6))
    party_level = int(p.get("party_level", 3))
    sheet = p.get("sheet", "Map")

    wb = _get_wb(path)
    ws = _sheet(wb, sheet)
    _clear_sheet(ws)

    # 整张填石墙
    for r in range(1, rows + 1):
        for c in range(1, cols + 1):
            _fill(ws, r, c, "stone")

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

    # 走廊：按顺序把相邻房间用 L 形走廊连起来
    def carve(r0, c0, r1, c1):
        for c in range(min(c0, c1), max(c0, c1) + 1):
            _fill(ws, r0, c, "corr")
        for r in range(min(r0, r1), max(r0, r1) + 1):
            _fill(ws, r, c1, "corr")

    for i in range(1, len(rooms)):
        a = rooms[i - 1]
        b = rooms[i]
        carve((a[0] + a[2]) // 2, (a[1] + a[3]) // 2, (b[0] + b[2]) // 2, (b[1] + b[3]) // 2)

    # 每个房间：地板 + 门 + 主题内容 + 变异地面斑块
    cr = _cr_for_level(party_level)
    for i, (r1, c1, r2, c2) in enumerate(rooms, start=1):
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                _fill(ws, r, c, "floor")
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
        _fill(ws, dr, dc, "door")

        # 主题房间内容
        rtype = _ROOM_TYPES[(i - 1) % len(_ROOM_TYPES)]
        name = rtype[0] + " " + str(i)
        pool = list(set(rtype[2]) | set(_GENERIC_MONSTERS))
        n_mon = rng.randint(1, 4)
        monsters = rng.sample(pool, min(n_mon, len(pool)))
        trap = rng.choice(_TRAPS) if rng.random() < rtype[4] else "无"
        treasure = rng.choice(list(set(rtype[3]) | set(_TREASURE)))
        desc = "一间%s%s，弥漫着潮湿的石腥味。" % (rtype[1], rtype[0])

        # 变异地面斑块（装饰，1~2 格）
        if rng.random() < 0.5:
            vt = rng.choice(_VARIANT_TERRAINS)
            for _ in range(rng.randint(1, 2)):
                vr = rng.randint(r1, r2)
                vc = rng.randint(c1, c2)
                _fill(ws, vr, vc, vt)

        # 标签 + 批注（左上角单元格）
        ws.cell(row=r1, column=c1, value=name)
        ws.cell(row=r1, column=c1).comment = Comment(
            "遭遇：%s\n陷阱：%s\n战利品：%s" % (", ".join(monsters), trap, treasure), "DM"
        )
        _paint_border(ws, r1, c1, r2, c2)

        # 写数据表行（Rooms/Encounters/Loot 在循环结束后统一落）
        _pending_rooms.append((i, name, c1, r1, c2 - c1 + 1, r2 - r1 + 1,
                               desc, ", ".join(monsters) + " (CR %d)" % cr, trap, treasure))

    # 偶尔放个水池
    if rooms and rng.random() < 0.7:
        rr = rng.choice(rooms[:n_rooms])  # 仅真实房间（不含 _pending）
        wr = rng.randint(rr[0], rr[2])
        wc = rng.randint(rr[1], rr[3])
        _fill(ws, wr, wc, "water")

    # 楼梯：上 / 下 各一（放在不同真实房间内）
    real = rooms[:n_rooms]
    if len(real) >= 1:
        a = rng.choice(real)
        _fill(ws, rng.randint(a[0], a[2]), rng.randint(a[1], a[3]), "stair")
    if len(real) >= 2:
        b = rng.choice([x for x in real if x != a]) if len(real) > 1 else real[0]
        _fill(ws, rng.randint(b[0], b[2]), rng.randint(b[1], b[3]), "stair")

    # 密门：某房间一侧中点改密门（覆盖原门）
    if real:
        s = rng.choice(real)
        sc = rng.choice([s[1], s[3]])
        sr = rng.randint(s[0], s[2])
        _fill(ws, sr, sc, "secret")

    # 数据表：Rooms / Encounters / Loot
    rooms_ws = _sheet(wb, "Rooms")
    rooms_ws.delete_rows(1, rooms_ws.max_row)
    enc_ws = _sheet(wb, "Encounters")
    enc_ws.delete_rows(1, enc_ws.max_row)
    loot_ws = _sheet(wb, "Loot")
    loot_ws.delete_rows(1, loot_ws.max_row)
    rooms_ws.append(["id", "name", "x", "y", "w", "h", "description", "encounter", "traps", "treasure"])
    enc_ws.append(["room_id", "level", "line"])
    loot_ws.append(["room_id", "item"])
    for row in _pending_rooms:
        (i, name, x, y, w, h, desc, enc, trap, treasure) = row
        rooms_ws.append([i, name, x, y, w, h, desc, enc, trap, treasure])
        enc_ws.append([i, party_level, "遭遇：%s（约 CR %d），建议先手控制。" % (enc.split(" (CR")[0], cr)])
        loot_ws.append([i, treasure])

    # 复位模块级暂存（单 worker 复用安全：每次 gen 都是新一批）
    _pending_rooms[:] = []

    return {
        "ok": True,
        "rooms": len(real),
        "cols": cols,
        "rows": rows,
        "sheets": wb.sheetnames,
    }


# gen_dungeon 用的模块级暂存（在 _gen_dungeon 内收集，循环结束后落表）
_pending_rooms = []


# --------------------------------------------------------------------------- #
# DonJon 风格 JSON 互通（双向）
# --------------------------------------------------------------------------- #
# 字段结构对齐 donjon.bin.sh 导出的真实 dungeon JSON：
#   name / settings{name,cell_size,n_cols,n_rows,seed,level} /
#   cells[rows][cols]（整数位标志，见 DJ_* 常量）/ rooms[]{id,col,row,width,height,
#   name,contents{summary,detail{monster[],trap[],hidden_treasure[],room_features}},
#   doors{north/south/east/west[]}} / stairs[] / egress[]
# 额外扩展 _cta（cordis-table-app 私有，其他工具忽略未知键）：palette(地形->ARGB)
# 与 terrain_to_donjon，用于在我们工具间无损往返（含 DonJon 无对应位的地形）。
def _collect_rooms(wb):
    if "Rooms" not in wb.sheetnames:
        return []
    ws = wb["Rooms"]
    data = list(ws.iter_rows(values_only=True))
    if not data:
        return []
    out = []
    for r in data[1:]:
        if not r or r[0] is None:
            continue
        rid, name, x, y, w, h = r[0], r[1], r[2], r[3], r[4], r[5]
        desc = r[6] if len(r) > 6 else ""
        enc = r[7] if len(r) > 7 else ""
        trap = r[8] if len(r) > 8 else "无"
        treasure = r[9] if len(r) > 9 else ""
        monsters = [m.strip() for m in str(enc).split(",") if m.strip()][:4] if enc else []
        out.append({
            "id": int(rid),
            "col": int(x),
            "row": int(y),
            "width": int(w),
            "height": int(h),
            "name": name or "",
            "contents": {
                "summary": desc or "",
                "detail": {
                    "monster": monsters,
                    "trap": [trap] if trap and trap != "无" else [],
                    "hidden_treasure": [treasure] if treasure else [],
                    "room_features": "",
                },
            },
            "doors": {"north": [], "south": [], "east": [], "west": []},
        })
    return out


def _export_donjon(p):
    path = p.get("path")
    sheet = p.get("sheet", "Map")
    out = p.get("out")
    wb = _get_wb(path)
    ws = _sheet(wb, sheet)
    cols = ws.max_column
    rows = ws.max_row

    used = set()
    overrides = []
    cell_matrix = []
    for r in range(1, rows + 1):
        line = []
        for c in range(1, cols + 1):
            key = _terrain_for_fill(_cell_fill_rgb(ws.cell(row=r, column=c)))
            used.add(key)
            line.append(_dj_int_for_terrain(key))
            # DonJon 无对应位的地形（water/lava/grass/sand/wood/rubble/stair/trap）
            # 用 _cta.tiles 记录精确位置，导入时覆盖还原；其他工具忽略此键。
            if key not in ("stone", "floor", "corr", "door", "secret"):
                overrides.append([r - 1, c - 1, key])
        cell_matrix.append(line)

    rooms = _collect_rooms(wb)
    data = {
        "name": p.get("name", "cordis-table-app dungeon"),
        "creator": "cordis-table-app",
        "settings": {
            "name": p.get("name", "cordis-table-app dungeon"),
            "cell_size": int(p.get("px", 100)),
            "n_cols": cols,
            "n_rows": rows,
            "seed": int(p.get("seed", 0) or 0),
            "level": int(p.get("level", 0) or 0),
        },
        "cells": cell_matrix,
        "rooms": rooms,
        "stairs": [],
        "egress": [],
        "_cta": {
            "version": 1,
            "sheet": sheet,
            "palette": {k: ARGB_BY_TERRAIN[k] for k in used},
            "terrain_to_donjon": {k: TERRAIN_TO_DJ.get(k, DJ_ROOM) for k in used},
            "tiles": overrides,
        },
    }
    os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
    with open(out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return {"ok": True, "out": out, "cols": cols, "rows": rows, "rooms": len(rooms)}


def _import_donjon(p):
    path = p.get("path")
    json_path = p.get("json")
    sheet = p.get("sheet", "Map")
    with open(json_path, encoding="utf-8") as f:
        data = json.load(f)

    cells = data.get("cells") or []
    n_rows = len(cells)
    n_cols = max((len(r) for r in cells), default=0)
    cta = data.get("_cta") or {}
    palette = cta.get("palette") or {}

    wb = _get_wb(path)
    ws = _sheet(wb, sheet)
    _clear_sheet(ws)
    for r in range(n_rows):
        row = cells[r]
        for c in range(n_cols):
            cell_int = int(row[c]) if c < len(row) else 0
            key = _terrain_from_dj(cell_int)
            argb = palette.get(key) or ARGB_BY_TERRAIN.get(key, ARGB_BY_TERRAIN["stone"])
            ws.cell(row=r + 1, column=c + 1).fill = PatternFill(
                start_color=_norm_argb(argb), end_color=_norm_argb(argb), fill_type="solid"
            )

    # _cta.tiles 覆盖：还原 DonJon 无对应位的地形（water/lava/grass/sand/wood/rubble/stair/trap）
    for ov in (cta.get("tiles") or []):
        r, c, key = ov[0], ov[1], ov[2]
        argb = palette.get(key) or ARGB_BY_TERRAIN.get(key, ARGB_BY_TERRAIN["stone"])
        ws.cell(row=r + 1, column=c + 1).fill = PatternFill(
            start_color=_norm_argb(argb), end_color=_norm_argb(argb), fill_type="solid"
        )

    # 房间：恢复 Rooms/Encounters/Loot 数据表 + 标签 / 批注 / 外墙
    rooms = data.get("rooms") or []
    rooms_ws = _sheet(wb, "Rooms")
    rooms_ws.delete_rows(1, rooms_ws.max_row)
    enc_ws = _sheet(wb, "Encounters")
    enc_ws.delete_rows(1, enc_ws.max_row)
    loot_ws = _sheet(wb, "Loot")
    loot_ws.delete_rows(1, loot_ws.max_row)
    rooms_ws.append(["id", "name", "x", "y", "w", "h", "description", "encounter", "traps", "treasure"])
    enc_ws.append(["room_id", "level", "line"])
    loot_ws.append(["room_id", "item"])

    for rm in rooms:
        rid = rm.get("id")
        x = int(rm.get("col", 0))
        y = int(rm.get("row", 0))
        w = int(rm.get("width", 1))
        h = int(rm.get("height", 1))
        name = rm.get("name", "")
        contents = rm.get("contents") or {}
        summary = contents.get("summary", "")
        detail = contents.get("detail") or {}
        monsters = detail.get("monster") or []
        traps = detail.get("trap") or []
        treasures = detail.get("hidden_treasure") or []
        trap_str = traps[0] if traps else "无"
        treasure_str = treasures[0] if treasures else ""
        enc_line = "遭遇：" + ", ".join(monsters) if monsters else ""
        rooms_ws.append([rid, name, x, y, w, h, summary, ", ".join(monsters), trap_str, treasure_str])
        if enc_line:
            enc_ws.append([rid, 0, enc_line])
        if treasure_str:
            loot_ws.append([rid, treasure_str])
        if name:
            ws.cell(row=y + 1, column=x + 1, value=name)
        ws.cell(row=y + 1, column=x + 1).comment = Comment(
            "遭遇：%s\n陷阱：%s\n战利品：%s" % (", ".join(monsters), trap_str, treasure_str), "DM"
        )
        _paint_border(ws, y + 1, x + 1, y + h, x + w)

    return {"ok": True, "rows": n_rows, "cols": n_cols, "rooms": len(rooms), "sheet": sheet}


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
