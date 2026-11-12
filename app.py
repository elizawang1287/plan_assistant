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

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from typing import List
import uvicorn

import parser
import ppt_export

app = FastAPI(title="计划书助手")

# 允许跨域（junrong-team.pages.dev 访问）
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # 生产环境可以改成具体域名
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

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


# ===== 团队业绩目标存储 =====
TARGETS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "targets.json")


@app.get("/api/targets")
def get_targets():
    """读取团队业绩目标"""
    try:
        if os.path.exists(TARGETS_FILE):
            with open(TARGETS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception as e:
        pass
    return {}


@app.post("/api/targets")
async def save_targets(payload: dict = None):
    """保存团队业绩目标"""
    try:
        targets = payload or {}
        with open(TARGETS_FILE, "w", encoding="utf-8") as f:
            json.dump(targets, f, ensure_ascii=False, indent=2)
        return {"status": "ok", "count": len(targets)}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"保存失败：{e}")


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
        # 生成文件名：储蓄类=客户+年保费+年数；重疾/医疗类=客户+保额
        b = (payload or {}).get("basic", {}) or {}
        age = b.get("anb_age", "")
        gender = b.get("gender", "")
        prod = (payload or {}).get("product_code", "")
        savings_codes = {"PRMESP", "TRST", "PHOR", "LFRE2", "PLP2", "ENCASH", "EGI", "EGIP"}
        if prod in savings_codes:
            prem = b.get("annual_premium", 0) or 0
            years = b.get("pay_years", "") or ""
            prem_wan = round(prem / 10000, 1) if prem else 0
            fname = f"{age}岁{gender}-年缴{prem_wan}万-{years}年.pptx"
        else:
            sa = b.get("sum_assured", 0) or 0
            years = b.get("pay_years", "") or ""
            sa_wan = round(sa / 10000, 1) if sa else 0
            years_str = f"-{years}年" if years else ""
            fname = f"{age}岁{gender}-保额{sa_wan}万{years_str}.pptx"
        # 文件名编码（中文）
        from urllib.parse import quote
        fname_enc = quote(fname)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{fname_enc}"},
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
        # 生成文件名
        b = base.get("basic", {}) or {}
        age = b.get("anb_age", "")
        gender = b.get("gender", "")
        prod = base.get("product_code", "")
        savings_codes = {"PRMESP", "TRST", "PHOR", "LFRE2", "PLP2", "ENCASH", "EGI", "EGIP"}
        if prod in savings_codes:
            prem = b.get("annual_premium", 0) or 0
            years = b.get("pay_years", "") or ""
            prem_wan = round(prem / 10000, 1) if prem else 0
            fname = f"{age}岁{gender}-年缴{prem_wan}万-{years}年-多提取方案.pptx"
        else:
            sa = b.get("sum_assured", 0) or 0
            years = b.get("pay_years", "") or ""
            sa_wan = round(sa / 10000, 1) if sa else 0
            years_str = f"-{years}年" if years else ""
            fname = f"{age}岁{gender}-保额{sa_wan}万{years_str}-多提取方案.pptx"
        from urllib.parse import quote
        fname_enc = quote(fname)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{fname_enc}"},
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


@app.post("/api/family-detail")
async def family_detail(files: List[UploadFile] = File(...)):
    """完整解析PDF，返回退保价值表数据（用于详细方案对比）。"""
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
            full = parser.parse_pdf(tmp_path)
            os.unlink(tmp_path)
            # 只返回前端需要的数据
            results.append({
                "filename": file.filename,
                "product_code": full.get("product_code"),
                "product_name": full.get("product_name"),
                "basic": full.get("basic", {}),
                "rows": full.get("rows", []),  # 退保价值表
                "optimistic": full.get("optimistic", {}),
            })
        except Exception as e:
            results.append({"filename": file.filename, "error": f"解析失败：{e}"})
    return JSONResponse({"policies": results})


# OCR 懒加载（首次调用时才初始化，避免启动慢）
_ocr_reader = None

def _get_ocr_reader():
    global _ocr_reader
    if _ocr_reader is None:
        import easyocr
        _ocr_reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)
    return _ocr_reader


@app.post("/api/family-parse-html")
async def family_parse_html(files: List[UploadFile] = File(...)):
    """接收多个HTML文件（保诚系统另存的页面），提取纯文本后解析基本信息。"""
    results = []
    for file in files:
        data = await file.read()
        if not data:
            results.append({"filename": file.filename, "error": "空文件"})
            continue
        try:
            text = data.decode("utf-8", errors="replace")
            # 去掉HTML标签，提取纯文本
            import re
            text = re.sub(r"<script[^>]*>.*?</script>", "", text, flags=re.S | re.I)
            text = re.sub(r"<style[^>]*>.*?</style>", "", text, flags=re.S | re.I)
            text = re.sub(r"<[^>]+>", "\n", text)
            text = re.sub(r"\n{3,}", "\n\n", text)
            text = re.sub(r"[ \t]+", " ", text)
            # 判断是保单详情页还是计划书
            if "保單詳情" in text or "保單號碼" in text or "保单详情" in text or "保单号码" in text:
                info = parser.parse_policy_detail_from_text(text)
            else:
                info = parser.parse_basic_from_text(text)
            info["filename"] = file.filename
            results.append(info)
        except Exception as e:
            results.append({"filename": file.filename, "error": f"HTML解析失败：{e}"})
    return JSONResponse({"policies": results})


@app.post("/api/family-parse-image")
async def family_parse_image(files: List[UploadFile] = File(...)):
    """接收多张图片（计划书截图），OCR 后提取基本信息。"""
    results = []
    for file in files:
        data = await file.read()
        if not data:
            results.append({"filename": file.filename, "error": "空文件"})
            continue
        try:
            import io
            from PIL import Image
            img = Image.open(io.BytesIO(data))
            reader = _get_ocr_reader()
            ocr_result = reader.readtext(img, detail=0, paragraph=True)
            text = "\n".join(ocr_result)
            info = parser.parse_basic_from_text(text)
            info["filename"] = file.filename
            info["ocr_text"] = text[:500]  # 前500字留作调试
            results.append(info)
        except Exception as e:
            results.append({"filename": file.filename, "error": f"OCR解析失败：{e}"})
    return JSONResponse({"policies": results})


@app.post("/api/family-review-export-ppt")
async def family_review_export_ppt(
    policies: str = Form(...),
    extract_files: List[UploadFile] = File(default=[]),
    shot_files: List[UploadFile] = File(default=[])
):
    """生成家庭保单检视PPT。
    FormData:
      - policies: JSON字符串，保单列表
      - extract_files: 提取方案文件列表（PDF/图片）
      - shot_files: 财务概览文件列表（PDF/图片）
    返回：pptx 文件下载
    """
    try:
        import json
        import family_review_export
        policies_list = json.loads(policies)
        if not policies_list:
            raise HTTPException(status_code=400, detail="没有保单数据")

        # 保存提取方案文件到临时目录（保留原始文件名，排序要靠文件名里的「第N年提取」）
        import tempfile, os, uuid
        extract_paths = []
        for f in extract_files:
            data = await f.read()
            name = os.path.basename(f.filename or "") or "upload.pdf"
            save_path = os.path.join(tempfile.gettempdir(), f"up_ext_{uuid.uuid4().hex[:8]}_{name}")
            with open(save_path, "wb") as fp:
                fp.write(data)
            extract_paths.append(save_path)

        # 保存财务概览文件到临时目录（同样保留原始文件名）
        shot_paths = []
        for f in shot_files:
            data = await f.read()
            name = os.path.basename(f.filename or "") or "upload.pdf"
            save_path = os.path.join(tempfile.gettempdir(), f"up_shot_{uuid.uuid4().hex[:8]}_{name}")
            with open(save_path, "wb") as fp:
                fp.write(data)
            shot_paths.append(save_path)

        buf = family_review_export.generate_family_review_ppt(
            policies_list,
            extract_paths=extract_paths,
            shot_paths=shot_paths
        )
        filename = "家庭保单检视.pptx"
        from urllib.parse import quote
        filename_encoded = quote(filename)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": f"attachment; filename*=UTF-8''{filename_encoded}"}
        )
    except Exception as e:
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"生成PPT失败：{e}")


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
        currency = (payload or {}).get("currency", "HKD")  # HKD 或 RMB

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
        add_title_slide("方案对比", f"共 {len(policies)} 张保单 ｜ USD→HKD {usd_hkd}")

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

            ncols = 7 if currency == "RMB" else 8
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
            # 彻底删除表格默认样式（隔行变色、首列高亮）
            tbl_pr = table._tbl.tblPr
            for child in list(tbl_pr):
                if 'tableStyleId' in child.tag or 'firstRow' in child.tag or 'bandRow' in child.tag:
                    tbl_pr.remove(child)

            # 列宽
            if currency == "RMB":
                table.columns[0].width = Inches(1.3)
                table.columns[1].width = Inches(2.5)
                table.columns[2].width = Inches(1.7)
                table.columns[3].width = Inches(1.6)
                table.columns[4].width = Inches(1.8)
                table.columns[5].width = Inches(1.8)
                table.columns[6].width = Inches(1.0)
                headers = ["家庭成员", "产品", "总保额(USD)", "前十年赠送(USD)", "年保费(USD)", f"年保费({currency})", "供款期"]
            else:
                table.columns[0].width = Inches(1.2)
                table.columns[1].width = Inches(2.2)
                table.columns[2].width = Inches(1.5)
                table.columns[3].width = Inches(1.5)
                table.columns[4].width = Inches(1.4)
                table.columns[5].width = Inches(1.6)
                table.columns[6].width = Inches(1.6)
                table.columns[7].width = Inches(0.9)
                headers = ["家庭成员", "产品", "总保额(USD)", "总保额(HKD)", "前十年赠送(USD)", "年保费(USD)", "年保费(HKD)", "供款期"]
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
                    
                    if currency == "RMB":
                        # RMB模式：总保额(USD) | 前十年赠送(USD) | 年保费(USD) | 年保费(RMB) | 供款期
                        table.cell(ri, 3).text = fmt_sum_usd(pol.get("first_10y_bonus"))
                        table.cell(ri, 4).text = fmt_prem_usd(pol.get("annual_premium"))
                        table.cell(ri, 5).text = fmt_rmb(pol.get("annual_premium"))
                        py = pol.get("pay_years")
                        table.cell(ri, 6).text = f"{py}年" if py else "—"
                    else:
                        # HKD模式：总保额(USD) | 总保额(HKD) | 前十年赠送(USD) | 年保费(USD) | 年保费(HKD) | 供款期
                        table.cell(ri, 3).text = fmt_sum_hkd(pol.get("sum_assured"))
                        table.cell(ri, 4).text = fmt_sum_usd(pol.get("first_10y_bonus"))
                        table.cell(ri, 5).text = fmt_prem_usd(pol.get("annual_premium"))
                        table.cell(ri, 6).text = fmt_prem_hkd(pol.get("annual_premium"))
                        py = pol.get("pay_years")
                        table.cell(ri, 7).text = f"{py}年" if py else "—"

                    for ci in range(ncols):
                        cell = table.cell(ri, ci)
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = WHITE
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
                    ri += 1

            # 强制所有数据行白底（覆盖表格默认样式）
            for r in range(1, ri):
                for c in range(ncols):
                    cell = table.cell(r, c)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = WHITE

        # ===== 按人汇总页（每个产品一行，带合计行） =====
        person_policies = {}
        person_order = []
        for pol in policies:
            key = f"{pol.get('age','?')}岁{pol.get('gender','?')}"
            if key not in person_policies:
                person_policies[key] = []
                person_order.append(key)
            person_policies[key].append(pol)

        # 计算总行数：表头 + 所有产品行 + 合计行
        total_person_rows = 1 + len(policies) + 1
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.6))
        p = txBox.text_frame.paragraphs[0]
        p.text = "按人汇总"
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = RED

        headers_person = ["家庭成员", "产品", "总保额(USD)", "总保额(HKD)", "年保费(USD)", "年保费(HKD)", "供款期"]
        left = Inches(0.3)
        top = Inches(1.1)
        width = Inches(12.7)
        height = Inches(0.5 * total_person_rows)
        table_shape = slide.shapes.add_table(total_person_rows, 7, left, top, width, height)
        table = table_shape.table
        tbl_pr = table._tbl.tblPr
        for child in list(tbl_pr):
            if 'tableStyleId' in child.tag or 'firstRow' in child.tag or 'bandRow' in child.tag:
                tbl_pr.remove(child)
        table.columns[0].width = Inches(1.3)
        table.columns[1].width = Inches(2.5)
        table.columns[2].width = Inches(1.6)
        table.columns[3].width = Inches(1.6)
        table.columns[4].width = Inches(1.7)
        table.columns[5].width = Inches(1.8)
        table.columns[6].width = Inches(1.0)

        # 表头
        for ci, h in enumerate(headers_person):
            cell = table.cell(0, ci)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = RED
            for para in cell.text_frame.paragraphs:
                para.font.size = Pt(10)
                para.font.bold = True
                para.font.color.rgb = WHITE
                para.alignment = PP_ALIGN.CENTER

        # 数据行：每个产品一行
        ri = 1
        grand_total_usd = 0
        for person in person_order:
            pols = person_policies[person]
            for idx, pol in enumerate(pols):
                if idx == 0:
                    table.cell(ri, 0).text = person
                table.cell(ri, 1).text = pol.get("product_name", "—")
                table.cell(ri, 2).text = fmt_sum_usd(pol.get("sum_assured"))
                table.cell(ri, 3).text = fmt_sum_hkd(pol.get("sum_assured"))
                table.cell(ri, 4).text = fmt_prem_usd(pol.get("annual_premium"))
                table.cell(ri, 5).text = fmt_prem_hkd(pol.get("annual_premium"))
                py = pol.get("pay_years")
                table.cell(ri, 6).text = f"{py}年" if py else "—"
                grand_total_usd += pol.get("annual_premium") or 0

                for ci in range(7):
                    cell = table.cell(ri, ci)
                    cell.fill.solid()
                    cell.fill.fore_color.rgb = WHITE
                    for para in cell.text_frame.paragraphs:
                        para.font.size = Pt(9)
                        para.font.color.rgb = DARK
                        para.alignment = PP_ALIGN.CENTER if ci != 1 else PP_ALIGN.LEFT
                    if idx == 0 and ci == 0:
                        for para in cell.text_frame.paragraphs:
                            para.font.size = Pt(10)
                            para.font.bold = True
                ri += 1

        # 合计行（只汇总年保费）
        grand_total_hkd = grand_total_usd * usd_hkd
        table.cell(ri, 0).text = "年缴合计"
        table.cell(ri, 1).text = ""
        table.cell(ri, 2).text = ""
        table.cell(ri, 3).text = ""
        table.cell(ri, 4).text = fmt_prem_usd(grand_total_usd)
        table.cell(ri, 5).text = f"HK${grand_total_hkd:,.0f}"
        table.cell(ri, 6).text = ""
        for ci in range(7):
            cell = table.cell(ri, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0xFD, 0xEC, 0xEC)
            for para in cell.text_frame.paragraphs:
                para.font.size = Pt(11)
                para.font.bold = True
                para.font.color.rgb = RED
                para.alignment = PP_ALIGN.CENTER

        # ===== 产品介绍页（每个产品一段） =====
        PRODUCT_DESC = {
            "PRMESP": {"title": "PRMESP 世誉财富保险计划", "func": "整付保费，多元货币（美元/港元/人民币），保证现金价值+非保证归原红利+终期分红", "use": "一次性投入一笔钱，长线增值，退休时提取作为养老金，或传承给下一代", "feature": "整付即享，无供款压力；多元货币转换自由；保证回报打底，非保证红利提升收益"},
            "TRST": {"title": "TRST 信守明天多元货币保险计划", "func": "5年缴，多元货币，保证现金价值高+非保证红利，支持灵活提取", "use": "中期储蓄（5年供款），退休前开始每年提取一笔补充收入，剩余继续滚存传承", "feature": "保证现金价值高，回本快；提取灵活，可按需调整；多元货币分散汇率风险"},
            "PHOR": {"title": "PHOR 诚誉财富保险计划", "func": "2年缴，多元货币，短期投入长期回报，保证+非保证红利", "use": "短期资金（2年缴）做长线储蓄，适合手头有一笔钱想快速完成储蓄目标", "feature": "2年即缴清，快速完成供款；长期滚存复利效应明显；适合高净值客户资产配置"},
            "PACE": {"title": "PACE 骏誉财富保险计划", "func": "3年缴，5列退保价值表（保证金额A+归原红利B+特别红利C+总额），多元货币", "use": "3年缴储蓄，平衡供款压力和长期回报，可提取作为退休补充收入", "feature": "3年缴清，时间适中；保证+归原红利+特别红利三重收益；提取灵活"},
            "LFRE2": {"title": "LFRE2 自主未来II", "func": "定期寿险+储蓄成分，供款期满后有现金价值", "use": "既有寿险保障，又有储蓄积累，适合想兼顾保障和理财的客户", "feature": "保障与储蓄一体；供款期满后现金价值持续增长"},
            "PLP2": {"title": "PLP2 美好人寿II", "func": "终身寿险，含储蓄成分，身故赔付+现金价值", "use": "终身保障+财富传承，适合留给家人一笔确定的资产", "feature": "终身保障；现金价值可贷款；传承资产确定性高"},
            "CIM3": {"title": "CIM3 诚保一生危疾保险计划", "func": "终身危疾保障，覆盖多种严重疾病，含首次重疾、多次赔付、严重脑退化及帕金森年金", "use": "确诊危疾时一次性赔付，弥补收入损失、康复费用及高端医疗自付额", "feature": "终身保障；多次赔付；含保费豁免；有现金价值（具备储蓄成分）"},
            "CIE3": {"title": "CIE3 守护健康危疾加护保III", "func": "终身危疾保障，覆盖74种病况，含早期危疾、深切治疗、保费豁免", "use": "确诊危疾时一次性赔付，弥补收入损失、康复费用", "feature": "终身保障；覆盖病况广；含早期危疾保障"},
            "CIP2": {"title": "CIP2 危疾首护保险计划II", "func": "终身危疾保障，覆盖74种病况；首10年额外定期危疾保障；含早期危疾、深切治疗、保费豁免", "use": "确诊危疾时一次性赔付，弥补收入损失、康复费用及高端医疗自付额", "feature": "20年供款终身保障；有现金价值及终期红利（第3年起）；具备储蓄成分"},
            "CIT2": {"title": "CIT2 守护健康危疾定期保II", "func": "定期危疾保障，保障期长（至指定年龄），纯保障无储蓄成分", "use": "人生责任最重的阶段（房贷期、子女教育期）用低保费撬动高保额", "feature": "保费低、保额高；定期保障，到期无现金价值；适合预算有限但需要高保额的客户"},
            "CIT": {"title": "CIT 定期危疾保险计划", "func": "定期危疾保障，纯保障无储蓄成分", "use": "特定阶段高保额危疾保障，补充终身危疾的保额不足", "feature": "保费低、保额高；定期保障；灵活选择保障年期"},
            "MCVIP": {"title": "MCVIP 高端医疗", "func": "终身高端医疗保障，覆盖住院及手术费用，自付额内地$1,250/境外$3,750", "use": "实报实销大额医疗开支，避免疾病导致的财务冲击", "feature": "终身续保；全球保障；高端私立医院直付"},
            "VIP": {"title": "VIP 自愿医保尚宾", "func": "自愿医保（标准计划），住院及手术费用实报实销", "use": "基础住院医疗保障，报销公立/私家医院大额医疗费用", "feature": "自愿医保扣税；保证续保；灵活选择自付额"},
            "CPER": {"title": "CPER 特选危疾治疗保", "func": "危疾治疗费用报销，覆盖指定疾病的治疗开支", "use": "补充危疾险，实报实销高端治疗费用（如靶向药、质子治疗等）", "feature": "实报实销；覆盖高端治疗；与危疾险互补"},
            "ENCASH": {"title": "ENCASH 加护一生住院现金储蓄保险", "func": "住院每日现金赔付+储蓄成分，人寿保障", "use": "住院时每天领取现金，弥补住院期间的收入损失和额外开支", "feature": "住院即赔；有现金价值；储蓄与保障一体"},
            "PLT1": {"title": "PLT1 守护家人定期寿险（1年续保）", "func": "1年续保定期寿险，纯身故/末期疾病保障，无储蓄成分", "use": "家庭经济支柱身故时，提供现金赔付给家人，覆盖房贷、子女教育、生活开支", "feature": "每年续保，保费随年龄递增；保费低、保额高；无储蓄成分"},
            "PLT5": {"title": "PLT5 守护家人定期寿险（5年续保）", "func": "5年续保定期寿险，纯身故/末期疾病保障，无储蓄成分", "use": "家庭经济支柱身故时，提供现金赔付给家人，覆盖房贷、子女教育、生活开支", "feature": "每5年续保保费递增；保费低、保额高；无储蓄成分"},
            "PLT10": {"title": "PLT10 守护家人定期寿险（10年续保）", "func": "10年续保定期寿险，纯身故/末期疾病保障，无储蓄成分", "use": "家庭经济支柱身故时，提供现金赔付给家人，覆盖房贷、子女教育、生活开支", "feature": "每10年续保保费递增；保费低、保额高；无储蓄成分"},
            "PLT20": {"title": "PLT20 守护家人定期寿险（20年续保）", "func": "20年续保定期寿险，纯身故/末期疾病保障，无储蓄成分", "use": "家庭经济支柱身故时，提供现金赔付给家人，覆盖房贷、子女教育、生活开支", "feature": "每20年续保保费递增；保费低、保额高；无储蓄成分"},
            "PLT": {"title": "PLT 定期寿险", "func": "定期寿险，纯身故保障，无储蓄成分", "use": "家庭经济支柱身故时，提供现金赔付给家人", "feature": "保费低、保额高；无储蓄成分"},
        }

        # 去重产品代码
        seen_codes = set()
        for pol in policies:
            code = pol.get("product_code")
            if not code or code in seen_codes:
                continue
            seen_codes.add(code)
            desc = PRODUCT_DESC.get(code)
            if not desc:
                continue

            # 每2个产品一页
            if len(seen_codes) % 2 == 1:
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                txBox = slide.shapes.add_textbox(Inches(0.5), Inches(0.3), Inches(12), Inches(0.6))
                p = txBox.text_frame.paragraphs[0]
                p.text = "产品介绍"
                p.font.size = Pt(24)
                p.font.bold = True
                p.font.color.rgb = RED
                y_offset = 1.2

            # 动态填充本保单信息
            sa = pol.get("sum_assured")
            prem = pol.get("annual_premium")
            pay_y = pol.get("pay_years")
            sa_str = f"保额 ${sa:,.0f}" if sa else "保额 —"
            prem_str = f"年缴 ${prem:,.0f}" if prem else "年缴 —"
            pay_str = f"{pay_y}年供款" if pay_y else "供款期 —"

            # 产品介绍文本框
            txBox = slide.shapes.add_textbox(Inches(0.8), Inches(y_offset), Inches(11.5), Inches(2.5))
            tf = txBox.text_frame
            tf.word_wrap = True
            p1 = tf.paragraphs[0]
            p1.text = desc["title"]
            p1.font.size = Pt(16)
            p1.font.bold = True
            p1.font.color.rgb = RED
            p2 = tf.add_paragraph()
            p2.text = f"• 功能：{desc['func']}"
            p2.font.size = Pt(11)
            p2.font.color.rgb = DARK
            p3 = tf.add_paragraph()
            p3.text = f"• 用途：{desc['use']}"
            p3.font.size = Pt(11)
            p3.font.color.rgb = DARK
            p4 = tf.add_paragraph()
            p4.text = f"• 特点：{desc['feature']}"
            p4.font.size = Pt(11)
            p4.font.color.rgb = DARK
            p5 = tf.add_paragraph()
            p5.text = f"• 本保单：{sa_str}，{prem_str}，{pay_str}"
            p5.font.size = Pt(11)
            p5.font.bold = True
            p5.font.color.rgb = DARK

            y_offset += 2.8

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
