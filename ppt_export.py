# -*- coding: utf-8 -*-
"""把解析结果生成为 .pptx（仿保诚参考PPT版式：A4竖版、红表头、表格铺满、双行分组表头）。
支持单份与多提取方案组合：基础表(投入/正常/乐观/保证)只出一次，每个提取方案各出一页。

版式规范（依据用户提供的参考PPT）：
- 主色 C00000（红）：表头底、产品名、本币(人民币/港币)列、关键金额数字
- 正文 222222，说明 888888，缴费描述普通部分 333333，标签 0D0D0D
- 表头字号：正常/乐观表 24pt、保证vs非保证 18+14pt、提取表 18pt
- 数据字号：正常/乐观表 18pt、保证/提取表 14pt
- 隔行填充 FAFAFA；保证vs非保证表不隔行
"""
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

RED = RGBColor(0xC0, 0x00, 0x00)      # 主红（表头/产品名/本币列/关键数字）
DARK = RGBColor(0x22, 0x22, 0x22)      # 正文
GRAY = RGBColor(0x88, 0x88, 0x88)      # 说明 / 汇率
SUB = RGBColor(0x33, 0x33, 0x33)       # 缴费描述普通文字
INK0 = RGBColor(0x0D, 0x0D, 0x0D)      # 提取页标题
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xFA, 0xFA, 0xFA)     # 隔行填充


def pick_rows(data_rows):
    last_year = data_rows[-1]["year"] if data_rows else None
    has_gap = any(r.get("is_age_marker") and r["year"] != last_year for r in data_rows)
    if has_gap:
        out = [r for r in data_rows if (r.get("is_age_marker") and r["year"] > 30)
               or (not r.get("is_age_marker") and r["year"] <= 30 and r["year"] % 5 == 0)]
    else:
        out = [r for r in data_rows if r["year"] % 5 == 0]
    last = data_rows[-1] if data_rows else None
    if last and not any(r["year"] == last["year"] for r in out):
        out.append(last)
    seen = set(); res = []
    for r in sorted(out, key=lambda x: x["year"]):
        if r["year"] in seen: continue
        seen.add(r["year"]); res.append(r)
    return res


def _wan(v):
    """美元金额缩写：>=1亿显示亿，否则万；整亿/整万去小数，其余保留1位小数。"""
    if v >= 1e8:
        t = int(v // 1e7 + 1e-9) / 10
        return f"{int(t)} 亿" if t == int(t) else f"{t:.1f} 亿"
    t = int(v // 1000 + 1e-9) / 10
    return f"{int(t)} 万" if t == int(t) else f"{t:.1f} 万"


def _rmb(v, rate):
    """人民币/港币换算，显示规则与美元一致。"""
    total = v * rate
    return _wan(total)


def _cl(data):
    return data.get("curLabel", "人民币")


def _style_cell(cell, text, bold=False, color=DARK, size=14, align=PP_ALIGN.CENTER, fill=None):
    cell.text = ""
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "微软雅黑"
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


def _box(slide, text, l, t, w, size=14, color=DARK, bold=False, align=PP_ALIGN.LEFT, h=0.5):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    p = box.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "微软雅黑"
    return box


def _richbox(slide, l, t, w, size, segs, align=PP_ALIGN.LEFT, h=0.5):
    """多 run 文本框，segs = [(text, color, bold), ...]。用于同一句内部分红/加粗。"""
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    p = box.text_frame.paragraphs[0]
    p.alignment = align
    for txt, color, bold in segs:
        run = p.add_run()
        run.text = txt
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = color
        run.font.name = "微软雅黑"
    return box


def _add_line(slide, l, t, w, h=0.02, color=RED):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    shp.line.fill.background()


def _footer(slide, rate=6.8, cl="人民币"):
    _box(slide, f"美元汇率：{rate}（仅供参考）", 5.1, 11.35, 2.8, size=9, color=GRAY, align=PP_ALIGN.RIGHT)


def _build_table(slide, top_in, headers, data_rows, col_w, row_h=0.44,
                 two_row_header=False, hdr_size=11, hdr2_size=None,
                 data_size=10, zebra=True, red_cols=()):
    """通用表格。red_cols 指定本币列索引（数据用红色），zebra 控制隔行填充。"""
    n_cols = len(headers[0])
    n_rows = len(data_rows) + (2 if two_row_header else 1)
    height = row_h * n_rows
    tbl_shape = slide.shapes.add_table(n_rows, n_cols, Inches(0.52), Inches(top_in),
                                       Inches(7.22), Inches(height))
    tbl = tbl_shape.table
    for i, w in enumerate(col_w):
        tbl.columns[i].width = Inches(w)
    for r in range(n_rows):
        tbl.rows[r].height = Inches(row_h)

    if two_row_header:
        for j, h in enumerate(headers[0]):
            _style_cell(tbl.cell(0, j), h or "", bold=True, color=WHITE, size=hdr_size, fill=RED)
        tbl.cell(0, 0).merge(tbl.cell(1, 0))
        tbl.cell(0, 1).merge(tbl.cell(0, 2))
        tbl.cell(0, 3).merge(tbl.cell(0, 4))
        for j, h in enumerate(headers[1]):
            _style_cell(tbl.cell(1, j), h, bold=True, color=WHITE, size=hdr2_size or hdr_size, fill=RED)
        start = 2
    else:
        for j, h in enumerate(headers[0]):
            _style_cell(tbl.cell(0, j), h, bold=True, color=WHITE, size=hdr_size, fill=RED)
        start = 1

    for i, row in enumerate(data_rows):
        fill = LIGHT if (zebra and i % 2 == 0) else None
        for j, val in enumerate(row):
            color = RED if j in red_cols else DARK
            _style_cell(tbl.cell(start + i, j), val, size=data_size, color=color, fill=fill)
    return tbl_shape


def _new_prs():
    prs = Presentation()
    prs.slide_width = Inches(8.26)
    prs.slide_height = Inches(11.69)
    return prs, prs.slide_layouts[6]


def _use_year_label(data_rows):
    """前25年里非年龄标记行>=10个 → 按年份显示，否则按岁数显示"""
    yr_set = {r["year"] for r in data_rows if not r.get("is_age_marker")}
    return sum(1 for y in range(1, 26) if y in yr_set) >= 10


def _ylab(r, last, use_year=True):
    if r.get("is_age_marker"):
        # 年龄标记行：若年份可由年龄推算（year 存在），统一显示"第X年"，与其余行一致
        if r.get("year"):
            return f"第 {r['year']} 年"
        return f"{r['age']}岁"
    return f"第 {r['year']} 年"


def _premium_desc(b, cur):
    """Slide1 缴费描述富文本段：普通部分 SUB，金额/年数部分红加粗。"""
    pt = b.get("premium_type") or ""
    annual = b.get("annual_premium") or 0
    total = b.get("total_premium") or annual
    years = b.get("pay_years") or 0
    if ("趸" in pt) or (years and years <= 1):
        return [("趸缴 ", SUB, False), (_wan(annual), RED, True), (" " + cur, SUB, False)]
    return [("年缴 ", SUB, False), (_wan(annual), RED, True), (f" {cur}，缴 ", SUB, False),
            (str(years), RED, True), (" 年，合共 ", SUB, False),
            (_wan(total), RED, True), (" " + cur, SUB, False)]


def add_base_slides(prs, blank, data, rate=6.8):
    """投入方案 + 正常/乐观/保证 三张基础表。"""
    b = data["basic"]
    cl = _cl(data)
    cur = b.get("currency", "美元")
    rows = pick_rows(data["rows"])
    last = rows[-1] if rows else None
    uy = _use_year_label(data["rows"])
    ylab = lambda r: _ylab(r, last, uy)
    age = b.get("anb_age", "")

    # === Slide 1：投入方案 + 正常预期退保价值 ===
    slide = prs.slides.add_slide(blank)
    _box(slide, data["product_name"], 0.71, 0.15, 6.85, size=24, bold=True, color=RED, align=PP_ALIGN.CENTER)
    _richbox(slide, 0.71, 0.69, 6.85, 16, _premium_desc(b, cur), align=PP_ALIGN.CENTER)
    _box(slide, f"{age}岁｜保单货币：{cur}", 0.71, 1.14, 6.85, size=12, color=GRAY, align=PP_ALIGN.CENTER)
    _add_line(slide, 0.52, 1.54, 7.08)
    _box(slide, "退保价值", 0.52, 1.60, 7.22, size=20, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    normal_rows = [[ylab(r), _wan(r["surrender"]), _rmb(r["surrender"], rate)] for r in rows]
    _build_table(slide, 2.05, [["保单年度", cur, f"{cl}（约）"]], normal_rows,
                 col_w=[2.62, 2.3, 2.3], row_h=0.42, hdr_size=24, data_size=18, red_cols=[2])
    _box(slide, "剩余资金可一次性取出，或传承给下一代", 0.52, 10.95, 7.22, size=16, color=RED, align=PP_ALIGN.CENTER)
    _footer(slide, rate, cl)

    # === Slide 2：乐观预期 ===
    has_opt = bool(data.get("optimistic"))
    def opt_of(r):
        o = data["optimistic"].get(r["year"]) or data["optimistic"].get(str(r["year"])) or {}
        return o.get("optimistic", r["surrender"])
    if has_opt:
        slide = prs.slides.add_slide(blank)
        _box(slide, "退保价值", 3.43, 0.20, 1.44, size=22, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        _box(slide, "（投资回报比现时预期每年上升约 1.4%）", 0.52, 0.72, 7.22, size=16, color=DARK, align=PP_ALIGN.CENTER)
        opt_rows = [[ylab(r), _wan(opt_of(r)), _rmb(opt_of(r), rate)] for r in rows]
        _build_table(slide, 1.14, [["保单年度", cur, f"{cl}（约）"]], opt_rows,
                     col_w=[2.62, 2.3, 2.3], row_h=0.45, hdr_size=24, data_size=18, red_cols=[2])
        _box(slide, "长远受每年 6.5% 客户总内部回报率规限，乐观与正常情景趋于一致",
             0.52, 11.0, 7.22, size=16, color=RED, align=PP_ALIGN.CENTER)
        _box(slide, "乐观预期", 6.93, 10.76, 0.97, size=12, color=INK0, align=PP_ALIGN.LEFT)
        _footer(slide, rate, cl)

    # === Slide 3：保证 vs 非保证 ===
    if has_opt:
        slide = prs.slides.add_slide(blank)
        _box(slide, "保证收益 vs 非保证收益", 0.0, 0.20, 8.26, size=22, bold=True, color=DARK, align=PP_ALIGN.CENTER)
        _box(slide, "保证现金价值（无论市场表现如何，这部分金额确定可获得）",
             0.52, 0.85, 7.22, size=16, color=GRAY, align=PP_ALIGN.CENTER)
        g_rows = [[ylab(r), _wan(r["guaranteed"]), _rmb(r["guaranteed"], rate),
                   _wan(r["bonus"]), _rmb(r["bonus"], rate)] for r in rows]
        _build_table(slide, 1.25,
                     [["保单年度", "保证收益", "", "非保证收益", ""],
                      ["", cur, f"{cl}（约）", cur, f"{cl}（约）"]],
                     g_rows, col_w=[1.34, 1.41, 1.24, 1.52, 1.71], row_h=0.44,
                     two_row_header=True, hdr_size=18, hdr2_size=14, data_size=14,
                     zebra=False, red_cols=[2, 4])
        _box(slide, "保证收益为保险公司承诺的最低现金价值，不包含非保证红利与终期分红",
             0.43, 11.0, 7.38, size=16, color=GRAY, align=PP_ALIGN.CENTER)
        _footer(slide, rate, cl)


def add_withdrawal_slide(prs, blank, data, rate=6.8, label=None):
    """提取方案页。label 用于多方案序号（一/二/三…）。"""
    w = data.get("withdrawal")
    if not w or not w.get("start_year"):
        return
    cl = _cl(data)
    cur = data["basic"].get("currency", "美元")
    slide = prs.slides.add_slide(blank)
    wrows = [x for x in w["rows"] if x["withdrawal"] > 0]
    n_years = len(wrows)
    wlast = wrows[-1]
    annual = w["annual_withdrawal"]
    total = wlast["accumulated"]

    _box(slide, "建议提取方案" + (label or ""), 3.05, 0.30, 2.17, size=20, bold=True, color=INK0, align=PP_ALIGN.CENTER)
    _box(slide, f"从第 {w['start_year']} 年开始提取，共提取 {n_years} 年", 2.72, 0.80, 2.81, size=12, color=DARK, align=PP_ALIGN.CENTER)
    _richbox(slide, 2.54, 1.15, 3.18, 20,
             [("每年领取 ", RED, True), (f"{annual:,.0f} ", RED, True), (cur, RED, True)],
             align=PP_ALIGN.CENTER)
    _box(slide, f"（每年约 {_rmb(annual, rate)}{cl}，每月约 {_rmb(annual/12, rate)}{cl}）",
         1.92, 1.55, 4.13, size=12, color=DARK, align=PP_ALIGN.CENTER)
    _richbox(slide, 1.32, 1.9, 5.63, 16,
             [(f"{n_years} 年总共领取 {_wan(total)}{cur}（约 {_rmb(total, rate)}{cl}）", DARK, True)],
             align=PP_ALIGN.CENTER)
    _box(slide, "提取后账户剩余的总现金价值", 2.21, 2.40, 3.85, size=20, bold=True, color=DARK, align=PP_ALIGN.CENTER)
    _add_line(slide, 0.59, 2.33, 7.08)
    key = [x for x in wrows if x["year"] == w["start_year"] or x["year"] % 5 == 0 or x == wrows[-1]]
    rem_rows = [[f"第 {x['year']} 年",
                 _wan(x["surrender"]), _rmb(x["surrender"], rate)] for x in key]
    _build_table(slide, 2.85, [["保单年度", cur + "（剩余价值）", f"{cl}（约）"]], rem_rows,
                 col_w=[2.62, 2.3, 2.3], row_h=0.40, hdr_size=18, data_size=14, red_cols=[2])
    _richbox(slide, 0.59, 11.0, 7.15, 16,
             [("领完一辈子，账户还剩近 ", DARK, False),
              (_wan(wlast["surrender"]) + " ", RED, False),
              ("继续传承给下一代", DARK, False)],
             align=PP_ALIGN.CENTER)
    _footer(slide, rate, cl)


def _save(prs):
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf


def build_pptx(data, rate=6.8):
    """单份：基础表 + 提取页。"""
    prs, blank = _new_prs()
    add_base_slides(prs, blank, data, rate)
    add_withdrawal_slide(prs, blank, data, rate)
    return _save(prs)


def build_combined_pptx(base, variants, rate=6.8):
    """同一投入方案多提取方案：基础表一次，每个提取方案各一页。"""
    cn = "一二三四五六七八九十"
    prs, blank = _new_prs()
    add_base_slides(prs, blank, base, rate)
    for i, v in enumerate(variants):
        label = cn[i] if i < len(cn) else str(i + 1)
        add_withdrawal_slide(prs, blank, v, rate, label=label)
    return _save(prs)
