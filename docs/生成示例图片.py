# -*- coding: utf-8 -*-
"""生成 README 用的示例图片（输入 / 运行 / 结果三张）。"""
import os
from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "images")
os.makedirs(OUT, exist_ok=True)

FONT = r"C:\Windows\Fonts\msyh.ttc"
FONT_B = r"C:\Windows\Fonts\msyhbd.ttc"

BG = "#F1F5F9"
CARD = "#FFFFFF"
BORDER = "#E2E8F0"
ACCENT = "#2563EB"
TEXT = "#0F172A"
MUTED = "#64748B"
STRIPE = "#F8FAFC"
DUP_BG = "#FEE2E2"
DUP_TX = "#B91C1C"


def f(size, bold=False):
    return ImageFont.truetype(FONT_B if bold else FONT, size)


def card(d, x, y, w, h, r=14):
    d.rounded_rectangle([x, y, x + w, y + h], radius=r, fill=CARD, outline=BORDER, width=2)


def table(d, x, y, cols, rows, widths, row_h=42, head_h=46, fs=18,
          highlight=None, aligns=None):
    """highlight: 需要高亮的行号集合(0起)"""
    highlight = highlight or set()
    aligns = aligns or ["left"] * len(cols)
    total = sum(widths)

    # 表头
    d.rectangle([x, y, x + total, y + head_h], fill=ACCENT)
    cx = x
    hf = f(fs, True)
    for i, c in enumerate(cols):
        tw = d.textlength(c, font=hf)
        tx = cx + (widths[i] - tw) / 2 if aligns[i] == "center" else cx + 16
        d.text((tx, y + (head_h - fs) / 2 - 2), c, font=hf, fill="#FFFFFF")
        cx += widths[i]

    # 数据行
    cy = y + head_h
    bf = f(fs)
    for ri, row in enumerate(rows):
        if ri in highlight:
            d.rectangle([x, cy, x + total, cy + row_h], fill=DUP_BG)
        elif ri % 2 == 1:
            d.rectangle([x, cy, x + total, cy + row_h], fill=STRIPE)
        d.line([x, cy + row_h, x + total, cy + row_h], fill=BORDER, width=1)
        cx = x
        for i, v in enumerate(row):
            txt = "" if v is None else str(v)
            col = DUP_TX if ri in highlight else TEXT
            if i == len(row) - 1 and ri not in highlight:
                col = MUTED
            tw = d.textlength(txt, font=bf)
            tx = cx + (widths[i] - tw) / 2 if aligns[i] == "center" else cx + 16
            d.text((tx, cy + (row_h - fs) / 2 - 2), txt, font=bf, fill=col)
            cx += widths[i]
        cy += row_h
    return cy


def title(d, x, y, text, sub=None):
    d.text((x, y), text, font=f(28, True), fill=TEXT)
    if sub:
        d.text((x, y + 42), sub, font=f(17), fill=MUTED)


# ── 图1：输入的两个文件 ─────────────────────────────
W, H = 1240, 440
img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)
title(d, 40, 34, "输入：两个表头不一致的 Excel 文件",
      "表头顺序不同、列数不同、还含重复行 —— 这是手工合并最容易出错的地方")

# 卡片1
card(d, 40, 132, 500, 268)
d.text((60, 152), "表1.xlsx", font=f(19, True), fill=TEXT)
d.text((150, 155), "3 行（含 1 行重复）", font=f(15), fill=MUTED)
table(d, 60, 192, ["姓名", "金额", "日期"],
      [["张三", 100, "2026-11-01"],
       ["李四", 200, "2026-11-02"],
       ["张三", 100, "2026-11-01"]],
      [140, 140, 180], highlight={2})
d.text((60, 348), "↑ 第 3 行与第 1 行完全相同", font=f(15), fill=DUP_TX)

# 卡片2
card(d, 580, 132, 620, 268)
d.text((600, 152), "表2.xlsx", font=f(19, True), fill=TEXT)
d.text((690, 155), "2 行（多一列「备注」，列序不同）", font=f(15), fill=MUTED)
table(d, 600, 192, ["日期", "姓名", "金额", "备注"],
      [["2026-11-03", "王五", 300, "新客户"],
       ["2026-11-04", "赵六", 400, "新客户"]],
      [160, 120, 110, 190])
d.text((600, 348), "↑ 日期在最左、多出「备注」列", font=f(15), fill=MUTED)

img.save(os.path.join(OUT, "1-输入文件.png"))

# ── 图2：命令行运行 ────────────────────────────────
W, H = 1240, 300
img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)
title(d, 40, 30, "运行：一条命令完成合并")

d.rounded_rectangle([40, 108, 1200, 268], radius=12, fill="#0F172A")
for i, c in enumerate(["#EF4444", "#F59E0B", "#22C55E"]):
    d.ellipse([62 + i * 26, 126, 74 + i * 26, 138], fill=c)

lines = [
    ("PS D:\\excel_merger> ", "#22C55E",
     "python excel_merger.py 测试数据 -o 合并结果.xlsx --dedup --verbose", "#E2E8F0"),
    ("", None, "发现 2 个文件: 表1.xlsx, 表2.xlsx", "#94A3B8"),
    ("", None, "已输出 4 行到 合并结果.xlsx", "#4ADE80"),
]
y = 158
for prompt, pc, cmd, cc in lines:
    x = 62
    if prompt:
        d.text((x, y), prompt, font=f(18), fill=pc)
        x += d.textlength(prompt, font=f(18))
    d.text((x, y), cmd, font=f(18), fill=cc)
    y += 36

img.save(os.path.join(OUT, "2-运行命令.png"))

# ── 图3：合并结果 ──────────────────────────────────
W, H = 1240, 480
img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)
title(d, 40, 34, "输出：合并后的结果（4 行，重复行已去除）",
      "表头自动对齐 · 缺失列留空 · 每行标注来源文件")

card(d, 40, 132, 1160, 268)
table(d, 60, 162, ["姓名", "金额", "日期", "备注", "来源文件"],
      [["张三", 100, "2026-11-01", None, "表1.xlsx"],
       ["李四", 200, "2026-11-02", None, "表1.xlsx"],
       ["王五", 300, "2026-11-03", "新客户", "表2.xlsx"],
       ["赵六", 400, "2026-11-04", "新客户", "表2.xlsx"]],
      [190, 170, 210, 230, 300], row_h=44, head_h=48)
d.text((60, 420), "原表1的重复行「张三 100 2026-11-01」已去重；表2的列顺序被自动对齐，表1没有的「备注」留空。",
      font=f(15), fill=MUTED)

img.save(os.path.join(OUT, "3-合并结果.png"))

print("已生成:", os.listdir(OUT))
