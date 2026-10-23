# -*- coding: utf-8 -*-
"""把解析结果生成为 .pptx（仿保诚参考PPT版式：A4竖版、红表头、表格铺满、双行分组表头）。
支持单份与多提取方案组合：基础表(投入/正常/乐观/保证)只出一次，每个提取方案各出一页。"""
import io
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

RED = RGBColor(0xE6, 0x00, 0x12)
DARK = RGBColor(0x22, 0x22, 0x22)
GRAY = RGBColor(0x88, 0x88, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xFA, 0xFA, 0xFA)

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
    return f"{int(v//1000 + 1e-9)/10:.1f} 万"


def _rmb(v, rate):
    return f"约 {int(v*rate//10000 + 1e-9)} 万"


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


def _box(slide, text, l, t, w, size=14, color=DARK, bold=False, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(0.5))
    p = box.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "微软雅黑"
    return box


def _footer(slide, rate=6.8, cl="人民币"):
    _box(slide, f"美元汇率：{rate}（仅供参考）", 5.1, 11.35, 2.8, size=9, color=GRAY, align=PP_ALIGN.RIGHT)


def _build_table(slide, top_in, headers, data_rows, col_w, row_h=0.48, two_row_header=False):
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
            _style_cell(tbl.cell(0, j), h or "", bold=True, color=WHITE, size=11, fill=RED)
        tbl.cell(0, 0).merge(tbl.cell(1, 0))
        tbl.cell(0, 1).merge(tbl.cell(0, 2))
        tbl.cell(0, 3).merge(tbl.cell(0, 4))
        for j, h in enumerate(headers[1]):
            _style_cell(tbl.cell(1, j), h, bold=True, color=WHITE, size=10, fill=RED)
        start = 2
    else:
        for j, h in enumerate(headers[0]):
            _style_cell(tbl.cell(0, j), h, bold=True, color=WHITE, size=11, fill=RED)
        start = 1

    for i, row in enumerate(data_rows):
        fill = LIGHT if (i % 2 == 0) else None
        for j, val in enumerate(row):
            _style_cell(tbl.cell(start + i, j), val, size=10, fill=fill)
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
        return f"{r['age']}岁"
    return f"第 {r['year']} 年"


def add_base_slides(prs, blank, data, rate=6.8):
    """投入方案 + 正常/乐观/保证 三张基础表。"""
    b = data["basic"]
    cl = _cl(data)
    cur = b.get("currency", "美元")
    rows = pick_rows(data["rows"])
    last = rows[-1] if rows else None
    uy = _use_year_label(data["rows"])
    ylab = lambda r: _ylab(r, last, uy)

    # 投入 + 正常预期
    slide = prs.slides.add_slide(blank)
    _box(slide, data["product_name"], 0.71, 0.30, 6.85, size=22, bold=True)
    _box(slide, f"趸交 {b.get('total_premium',0)/10000:.1f} 万 {cur}", 0.71, 0.80, 6.85, size=14)
    _box(slide, f"{b.get('anb_age','')}岁{b.get('gender','')}性｜保单货币：{cur}", 0.71, 1.18, 6.85, size=11, color=GRAY)
    _box(slide, "退保价值", 0.74, 1.60, 6.85, size=16, bold=True)
    normal_rows = [[ylab(r), _wan(r["surrender"]), _rmb(r["surrender"], rate)] for r in rows]
    _build_table(slide, 2.05, [["保单年度", cur, f"{cl}（约）"]], normal_rows,
                 col_w=[2.62, 2.3, 2.3], row_h=0.48)
    _box(slide, "剩余资金可一次性取出，或传承给下一代", 0.71, 10.95, 6.85, size=10, color=GRAY)
    _footer(slide, rate, cl)

    # 乐观预期
    has_opt = bool(data.get("optimistic"))
    def opt_of(r):
        o = data["optimistic"].get(r["year"]) or data["optimistic"].get(str(r["year"])) or {}
        return o.get("optimistic", r["surrender"])
    if has_opt:
        slide = prs.slides.add_slide(blank)
        _box(slide, "退保价值", 0.72, 0.20, 6.85, size=22, bold=True)
        _box(slide, "（投资回报比现时预期每年上升约 1.4%）", 0.72, 0.72, 6.85, size=11, color=GRAY)
        opt_rows = [[ylab(r), _wan(opt_of(r)), _rmb(opt_of(r), rate)] for r in rows]
        _build_table(slide, 1.25, [["保单年度", cur, f"{cl}（约）"]], opt_rows,
                     col_w=[2.62, 2.3, 2.3], row_h=0.48)
        _box(slide, "长远受每年 6.5% 客户总内部回报率规限，乐观与正常情景趋于一致", 0.67, 10.9, 6.85, size=10, color=GRAY)
        _footer(slide, rate, cl)

    # 保证 vs 非保证
    if has_opt:
        slide = prs.slides.add_slide(blank)
        _box(slide, "保证收益 vs 非保证收益", 0.0, 0.20, 8.26, size=22, bold=True, align=PP_ALIGN.CENTER)
        _box(slide, "保证现金价值（无论市场表现如何，这部分金额确定可获得）", 0.73, 0.85, 7.0, size=11, color=GRAY)
        g_rows = [[ylab(r), _wan(r["guaranteed"]), _rmb(r["guaranteed"], rate),
                   _wan(r["bonus"]), _rmb(r["bonus"], rate)] for r in rows]
        _build_table(slide, 1.25,
                     [["保单年度", "保证收益", "", "非保证收益", ""],
                      ["", cur, f"{cl}（约）", cur, f"{cl}（约）"]],
                     g_rows, col_w=[1.62, 1.4, 1.4, 1.4, 1.4], row_h=0.46, two_row_header=True)
        _box(slide, "保证收益为保险公司承诺的最低现金价值，不包含非保证红利与终期分红", 0.62, 11.0, 7.0, size=10, color=GRAY)
        _footer(slide, rate, cl)


def add_withdrawal_slide(prs, blank, data, rate=6.8, label=None):
    “””提取方案页。label 用于多方案时区分（如”第6年起提取”）。”””
    w = data.get(“withdrawal”)
    if not w or not w.get(“start_year”):
        return
    cl = _cl(data)
    cur = data[“basic”].get(“currency”, “美元”)
    slide = prs.slides.add_slide(blank)
    wrows = [x for x in w[“rows”] if x[“withdrawal”] > 0]
    n_years = len(wrows)
    wlast = wrows[-1]
    annual = w[“annual_withdrawal”]
    total = wlast[“accumulated”]
    _box(slide, “建议提取方案” + (f”（{label}）” if label else “”), 0.59, 0.30, 7.08, size=18, bold=True)
    _box(slide, f”从第 {w['start_year']} 年开始提取，共提取 {n_years} 年”, 0.59, 0.80, 7.08, size=12)
    _box(slide, f”每年领取 {annual:,.0f} {cur}”, 0.59, 1.15, 7.08, size=16, bold=True)
    _box(slide, f”（每年约 {int(annual*rate//1000 + 1e-9)/10:.1f} 万{cl}，每月约 {int(annual*rate//10000/12*100 + 1e-9)/100:.2f} 万{cl}）”, 0.44, 1.55, 7.08, size=10, color=GRAY)
    _box(slide, f”{n_years} 年总共领取 {int(total//1000 + 1e-9)/10:.1f} 万{cur}（约 {int(total*rate//10000 + 1e-9)} 万{cl}）”, 0.59, 1.90, 7.08, size=13, bold=True)
    _box(slide, “提取后账户剩余的总现金价值”, 0.59, 2.40, 7.08, size=14, bold=True)
    key = [x for x in wrows if x[“year”] == w[“start_year”] or x[“year”] % 5 == 0 or x == wrows[-1]]
    rem_rows = [[f”第 {x['year']} 年” + (f”（{x['age']}岁）” if x == wrows[-1] else “”),
                 _wan(x[“surrender”]), _rmb(x[“surrender”], rate)] for x in key]
    _build_table(slide, 2.85, [[“保单年度”, cur+”（剩余价值）”, f”{cl}（约）”]], rem_rows,
                 col_w=[2.62, 2.3, 2.3], row_h=0.48)
    _box(slide, f”领完一辈子，账户还剩近 {int(wlast['surrender']//1000 + 1e-9)/10:.1f} 万 继续传承给下一代”, 0.59, 11.0, 7.08, size=11, color=GRAY)
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
    """同一投入方案多提取方案：基础表一次，每个 variant 各一页提取。"""
    prs, blank = _new_prs()
    add_base_slides(prs, blank, base, rate)
    for v in variants:
        w = v.get("withdrawal") or {}
        label = v.get("variant_label") or (f"第{w.get('start_year','?')}年起提取" if w.get("start_year") else None)
        add_withdrawal_slide(prs, blank, v, rate, label=label)
    return _save(prs)
