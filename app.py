# -*- coding: utf-8 -*-
"""
计划书助手 - 本地后端
启动：python app.py  然后浏览器打开 http://127.0.0.1:8000
"""
import io
import json
import re
import tempfile
import os

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from typing import List
import uvicorn

import parser
import ppt_export

app = FastAPI(title="计划书助手")

STATIC_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "static")


@app.middleware("http")
async def no_cache(request, call_next):
    # 开发期禁用缓存，避免前端改完浏览器还加载旧版
    response = await call_next(request)
    response.headers["Cache-Control"] = "no-cache, no-store, must-revalidate"
    return response


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/parse")
async def parse_plan(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="请上传 PDF 文件")
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="空文件")
    try:
        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
            tmp.write(data)
            tmp_path = tmp.name
        result = parser.parse_pdf(tmp_path)
        os.unlink(tmp_path)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"解析失败：{e}")
    if result["product_code"] == "UNKNOWN":
        result["warning"] = "未识别产品，仅提取到通用数据；请确认是否为保诚 PRMESP / TRST 计划书。"
    return JSONResponse(result)


@app.post("/api/export-ppt")
async def export_ppt(payload: dict = None):
    try:
        rate = float((payload or {}).get("rate", 6.8))
        buf = ppt_export.build_pptx(payload, rate)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": 'attachment; filename="plan.pptx"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"导出PPT失败：{e}")


@app.post("/api/export-combined-ppt")
async def export_combined(payload: dict = None):
    try:
        rate = float((payload or {}).get("rate", 6.8))
        cl = payload.get("curLabel", "人民币")
        base = payload["base"]
        variants = payload.get("variants", []) or [base]
        base["curLabel"] = cl
        for v in variants:
            v["curLabel"] = cl
        buf = ppt_export.build_combined_pptx(base, variants, rate)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": 'attachment; filename="plan.pptx"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"导出组合PPT失败：{e}")


# ============ 家庭保单规划 ============

@app.post("/api/family-parse")
async def family_parse(files: List[UploadFile] = File(...)):
    """接收多个 PDF，返回每张的基本信息。"""
    results = []
    for file in files:
        if not file.filename.lower().endswith(".pdf"):
            results.append({"filename": file.filename, "error": "非PDF文件"})
            continue
        data = await file.read()
        if not data:
            results.append({"filename": file.filename, "error": "空文件"})
            continue
        try:
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(data)
                tmp_path = tmp.name
            info = parser.parse_basic_fast(tmp_path)
            os.unlink(tmp_path)
            info["filename"] = file.filename
            results.append(info)
        except Exception as e:
            results.append({"filename": file.filename, "error": f"解析失败：{e}"})
    return JSONResponse({"policies": results})


@app.post("/api/renewal-parse")
async def renewal_parse(file: UploadFile = File(...), rate: float = 7.9, dda_day: int = 6, include_monthly: bool = False):
    """接收 AES Premium Position Enquiry 另存的 HTML，返回按客户合并的续保名单。"""
    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="空文件")
    try:
        text = data.decode("utf-8", errors="replace")
        import renewal
        result = renewal.parse_renewal_html(text, rate=rate, dda_day=dda_day, include_monthly=include_monthly)
        return JSONResponse(result)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"解析失败：{e}")


@app.post("/api/family-export-ppt")
async def family_export_ppt(payload: dict = None):
    """接收家庭保单汇总数据，生成 PPT。"""
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt, Emu
        from pptx.dml.color import RGBColor
        from pptx.enum.text import PP_ALIGN

        policies = (payload or {}).get("policies", [])
        usd_hkd = float((payload or {}).get("usd_hkd", 7.8))
        usd_rmb = float((payload or {}).get("usd_rmb", 7.2))

        prs = Presentation()
        prs.slide_width = Inches(13.333)
        prs.slide_height = Inches(7.5)

        RED = RGBColor(0xB5, 0x00, 0x00)
        DARK = RGBColor(0x1F, 0x24, 0x30)
        GRAY = RGBColor(0x6B, 0x72, 0x80)
        WHITE = RGBColor(0xFF, 0xFF, 0xFF)
        LIGHT_BG = RGBColor(0xF4, 0xF5, 0xF7)

        def add_title_slide(title, subtitle=""):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
            bg.fill.solid()
            bg.fill.fore_color.rgb = RED
            bg.line.fill.background()
            txBox = slide.shapes.add_textbox(Inches(1), Inches(2.5), Inches(11), Inches(2))
            tf = txBox.text_frame
            p = tf.paragraphs[0]
            p.text = title
            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = WHITE
            if subtitle:
                p2 = tf.add_paragraph()
                p2.text = subtitle
                p2.font.size = Pt(18)
                p2.font.color.rgb = RGBColor(0xFF, 0xDD, 0xDD)
            return slide

        def add_table_slide(title, headers, rows):
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            # 标题
            txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.6))
            p = txBox.text_frame.paragraphs[0]
            p.text = title
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = DARK
            # 表格
            nrows = len(rows) + 1
            ncols = len(headers)
            left = Inches(0.5)
            top = Inches(1.1)
            width = Inches(12.3)
            height = Inches(0.4 * nrows)
            table_shape = slide.shapes.add_table(nrows, ncols, left, top, width, height)
            table = table_shape.table
            for ci, h in enumerate(headers):
                cell = table.cell(0, ci)
                cell.text = h
                for para in cell.text_frame.paragraphs:
                    para.font.size = Pt(11)
                    para.font.bold = True
                    para.font.color.rgb = WHITE
                cell.fill.solid()
                cell.fill.fore_color.rgb = RED
            for ri, row in enumerate(rows):
                for ci, val in enumerate(row):
                    cell = table.cell(ri + 1, ci)
                    cell.text = str(val) if val is not None else "—"
                    for para in cell.text_frame.paragraphs:
                        para.font.size = Pt(10)
                        para.font.color.rgb = DARK
                    if ri % 2 == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = LIGHT_BG
            return slide

        def fmt_money(v):
            if v is None: return "—"
            return f"${v:,.0f}"

        def fmt_hkd(v):
            if v is None: return "—"
            return f"HK${v*usd_hkd:,.0f}"

        def fmt_rmb(v):
            if v is None: return "—"
            return f"¥{v*usd_rmb:,.0f}"

        # 万为单位
        def fmt_wan(v):
            if v is None: return "—"
            if v >= 10000:
                w = v / 10000
                return f"{w:.0f}万" if w >= 100 else f"{w:.1f}万"
            return f"{v:,.0f}"

        # 保额单元格：$47.5万
        def fmt_sum_usd(v):
            if v is None: return "—"
            return f"${fmt_wan(v)}"

        # HKD保额：HKD 373万
        def fmt_sum_hkd(v):
            if v is None: return "—"
            return f"HKD {fmt_wan(v * usd_hkd)}"

        # 年缴USD：$17,922.5
        def fmt_prem_usd(v):
            if v is None: return "—"
            return f"${v:,.1f}"

        # 年缴HKD：HK$140,871
        def fmt_prem_hkd(v):
            if v is None: return "—"
            return f"HK${v*usd_hkd:,.0f}"

        # Slide 1: 封面
        add_title_slide("家庭保单规划方案", f"共 {len(policies)} 张保单 ｜ USD→HKD {usd_hkd}")

        # 收集所有家庭成员（按年龄排序）
        person_set = set()
        for pol in policies:
            person_set.add(f"{pol.get('age','?')}岁{pol.get('gender','?')}")
        persons = sorted(person_set, key=lambda x: int(x.replace('岁','').replace('男','').replace('女','')))

        # 按产品类型分组
        cats = {}
        for pol in policies:
            cat = pol.get("category", "其他")
            cats.setdefault(cat, []).append(pol)

        # 每个类型一页 PPT：列=家庭成员|产品|总保额|年保费|供款期
        for cat, pols in cats.items():
            # 计算总行数：每人的保单数 + 表头 + 合计行
            total_rows = 1  # 表头
            for person in persons:
                cnt = sum(1 for p in pols if f"{p.get('age','?')}岁{p.get('gender','?')}" == person)
                total_rows += cnt
            total_rows += 1  # 合计行

            ncols = 7
            slide = prs.slides.add_slide(prs.slide_layouts[6])

            # 标题
            txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.6))
            p = txBox.text_frame.paragraphs[0]
            p.text = f"{cat} 对比"
            p.font.size = Pt(24)
            p.font.bold = True
            p.font.color.rgb = RED

            left = Inches(0.3)
            top = Inches(1.1)
            width = Inches(12.7)
            height = Inches(0.5 * total_rows)
            table_shape = slide.shapes.add_table(total_rows, ncols, left, top, width, height)
            table = table_shape.table

            # 列宽
            table.columns[0].width = Inches(1.3)
            table.columns[1].width = Inches(2.5)
            table.columns[2].width = Inches(1.6)
            table.columns[3].width = Inches(1.6)
            table.columns[4].width = Inches(1.7)
            table.columns[5].width = Inches(1.8)
            table.columns[6].width = Inches(1.0)

            # 表头
            headers = ["家庭成员", "产品", "总保额(USD)", "总保额(HKD)", "年保费(USD)", "年保费(HKD)", "供款期"]
            for ci, h in enumerate(headers):
                cell = table.cell(0, ci)
                cell.text = h
                cell.fill.solid()
                cell.fill.fore_color.rgb = RED
                for para in cell.text_frame.paragraphs:
                    para.font.size = Pt(10)
                    para.font.bold = True
                    para.font.color.rgb = WHITE
                    para.alignment = PP_ALIGN.CENTER

            # 数据行
            ri = 1
            for person in persons:
                person_pols = [p for p in pols if f"{p.get('age','?')}岁{p.get('gender','?')}" == person]
                if not person_pols:
                    continue
                for idx, pol in enumerate(person_pols):
                    # 产品
                    table.cell(ri, 1).text = pol.get("product_name", "—")
                    # 总保额USD
                    table.cell(ri, 2).text = fmt_sum_usd(pol.get("sum_assured"))
                    # 总保额HKD
                    table.cell(ri, 3).text = fmt_sum_hkd(pol.get("sum_assured"))
                    # 年保费USD
                    table.cell(ri, 4).text = fmt_prem_usd(pol.get("annual_premium"))
                    # 年保费HKD
                    table.cell(ri, 5).text = fmt_prem_hkd(pol.get("annual_premium"))
                    # 供款期
                    py = pol.get("pay_years")
                    table.cell(ri, 6).text = f"{py}年" if py else "—"

                    for ci in range(7):
                        cell = table.cell(ri, ci)
                        for para in cell.text_frame.paragraphs:
                            para.font.size = Pt(9)
                            para.font.color.rgb = DARK
                            para.alignment = PP_ALIGN.CENTER if ci != 1 else PP_ALIGN.LEFT

                    # 第一行才写家庭成员
                    if idx == 0:
                        table.cell(ri, 0).text = person
                        for para in table.cell(ri, 0).text_frame.paragraphs:
                            para.font.size = Pt(10)
                            para.font.bold = True
                            para.font.color.rgb = DARK
                            para.alignment = PP_ALIGN.CENTER
                        table.cell(ri, 0).fill.solid()
                        table.cell(ri, 0).fill.fore_color.rgb = LIGHT_BG
                    ri += 1

            # 合计行
            total_prem = sum(p.get("annual_premium") or 0 for p in pols)
            table.cell(ri, 0).text = "年缴合计"
            table.cell(ri, 1).text = ""
            table.cell(ri, 2).text = ""
            table.cell(ri, 3).text = ""
            table.cell(ri, 4).text = fmt_prem_usd(total_prem)
            table.cell(ri, 5).text = fmt_prem_hkd(total_prem)
            table.cell(ri, 6).text = ""
            for ci in range(7):
                cell = table.cell(ri, ci)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xFD, 0xEC, 0xEC)
                for para in cell.text_frame.paragraphs:
                    para.font.size = Pt(11)
                    para.font.bold = True
                    para.font.color.rgb = RED

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": 'attachment; filename="family_plan.pptx"'},
        )
    except Exception as e:
        import traceback
        raise HTTPException(status_code=500, detail=f"导出PPT失败：{e}\n{traceback.format_exc()}")


# ============ 核保查询（Underwriting）============
UW_JSON_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "uw_data.json")
_uw_cache = {"mtime": None, "data": None}

# 简繁转换（数据已为简体原文，用户可能输入繁体，统一转简体再匹配）
try:
    from opencc import OpenCC
    _cc = OpenCC("t2s")
except Exception:  # pragma: no cover
    _cc = None


def _norm_uw(s: str) -> str:
    """归一化：去首尾空格、转小写、去空格与括号。"""
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r"[\s\(\)（）\[\]【】「」『』]", "", s)
    return s


def _to_simplified(s: str) -> str:
    if _cc is None:
        return s
    try:
        return _cc.convert(s)
    except Exception:
        return s


def load_uw_data():
    if not os.path.exists(UW_JSON_PATH):
        return None
    mtime = os.path.getmtime(UW_JSON_PATH)
    if _uw_cache["data"] is not None and _uw_cache["mtime"] == mtime:
        return _uw_cache["data"]
    with open(UW_JSON_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)
    _uw_cache["data"] = data
    _uw_cache["mtime"] = mtime
    return data


def _disease_keys(d: dict):
    """返回该疾病用于匹配的归一化键集合。"""
    keys = set()
    cn = _norm_uw(d.get("name_cn", ""))
    if cn:
        keys.add(cn)
    en = _norm_uw(d.get("name_en", ""))
    if en:
        keys.add(en)
    for al in d.get("aliases", []) or []:
        a = _norm_uw(al)
        if a:
            keys.add(a)
    return keys


def _score_disease(d: dict, q_norm: str, q_simp: str):
    """返回 (tier, ) 越小越匹配；不匹配返回 None。"""
    cn = _norm_uw(d.get("name_cn", ""))
    en = _norm_uw(d.get("name_en", ""))
    aliases = [_norm_uw(a) for a in (d.get("aliases", []) or []) if a]
    # 用户查询的两种归一化形式（原文 + 转简体）
    qs = [q_norm, q_simp]
    # ① 精确相等（name_cn / name_en / aliases）
    for q in qs:
        if not q:
            continue
        if q == cn or q == en:
            return 1
        if q in aliases:
            return 1
    # ② 名称以 q 开头（主名 name_cn / name_en）
    for q in qs:
        if not q:
            continue
        if cn.startswith(q) or en.startswith(q):
            return 2
    # ③ q 或名称互为子串（双向 contains，针对主名）
    for q in qs:
        if not q:
            continue
        if q in cn or cn in q:
            return 3
        if en and (q in en or en in q):
            return 3
    # ④ 别名 / 英文模糊匹配（子串）
    for q in qs:
        if not q:
            continue
        for a in aliases:
            if q in a or a in q:
                return 4
        if en and (q in en or en in q):
            return 4
    return None


@app.get("/api/underwriting/search")
def uw_search(q: str = ""):
    data = load_uw_data()
    if data is None:
        raise HTTPException(status_code=503, detail="核保知识库尚未就绪：uw_data.json 不存在，请稍后再试。")
    diseases = data.get("diseases", []) or []
    q_raw = (q or "").strip()
    if not q_raw:
        return JSONResponse({"query": q_raw, "count": 0, "results": [], "category_match": False},
                            status_code=200)
    q_norm = _norm_uw(q_raw)
    q_simp = _norm_uw(_to_simplified(q_raw))

    # 分类精确匹配：q 等于某分类名 → 返回该分类全部条目
    cat_map = {}
    for d in diseases:
        cat = (d.get("category") or "").strip()
        if cat:
            cat_map.setdefault(cat, []).append(d)
    for cat, items in cat_map.items():
        cn = _norm_uw(cat)
        if cn in (q_norm, q_simp) or (q_norm and q_norm in cn) or (q_simp and q_simp in cn):
            return JSONResponse(
                {"query": q_raw, "count": len(items), "results": items,
                 "category_match": True, "category": cat},
                status_code=200)

    # 疾病名匹配
    scored = []
    for d in diseases:
        t = _score_disease(d, q_norm, q_simp)
        if t is not None:
            scored.append((t, d))
    scored.sort(key=lambda x: x[0])
    results = [d for _, d in scored]
    return JSONResponse(
        {"query": q_raw, "count": len(results), "results": results, "category_match": False},
        status_code=200)


@app.get("/api/underwriting/meta")
def uw_meta():
    data = load_uw_data()
    if data is None:
        raise HTTPException(status_code=503, detail="核保知识库尚未就绪：uw_data.json 不存在，请稍后再试。")
    diseases = data.get("diseases", []) or []
    categories = sorted({(d.get("category") or "").strip() for d in diseases if d.get("category")})
    return JSONResponse(
        {"manual": data.get("manual", {}),
         "disease_count": len(diseases),
         "appendices_count": len(data.get("appendices", []) or []),
         "categories": categories},
        status_code=200)


# 静态前端（放最后挂载，避免拦截 API）
app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
