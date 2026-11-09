# -*- coding: utf-8 -*-
"""家庭保单检视PPT生成（竖版A4，仿李晓波模板）。
结构：封面 → 现有保单汇总（按人分表+总保费汇总）→ 产品详情页
"""
import io
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

RED = RGBColor(0xE6, 0x00, 0x12)      # 保诚红
DARK = RGBColor(0x22, 0x22, 0x22)
GRAY = RGBColor(0x88, 0x88, 0x88)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xFA, 0xFA, 0xFA)
LIGHT_ORANGE = RGBColor(0xFF, 0xF3, 0xE0)  # 汇总表浅橙底


def _style_cell(cell, text, bold=False, color=DARK, size=10, align=PP_ALIGN.CENTER, fill=None):
    cell.text = ""
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = cell.text_frame.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = str(text) if text else "—"
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "微软雅黑"
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill


def _remove_table_style(table):
    tbl = table._tbl
    tblPr = tbl.find('{http://schemas.openxmlformats.org/drawingml/2006/main}tblPr')
    if tblPr is not None:
        for child in list(tblPr):
            if child.tag.endswith('tableStyleId'):
                tblPr.remove(child)


def _add_textbox(slide, left, top, width, height, text, size=14, bold=False, color=DARK, align=PP_ALIGN.LEFT):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    run = p.add_run()
    run.text = text
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = "微软雅黑"
    return txBox


def _pdf_page_to_image(pdf_path, page_num, clip=None):
    """把PDF指定页码转成图片（0索引）。
    clip: (x0, y0, x1, y1) 裁剪区域，单位是PDF点（72dpi）。None=整页。
    """
    import fitz
    doc = fitz.open(pdf_path)
    if page_num >= len(doc):
        doc.close()
        return None
    page = doc[page_num]
    if clip:
        rect = fitz.Rect(clip)
        pix = page.get_pixmap(dpi=150, clip=rect)
    else:
        pix = page.get_pixmap(dpi=150)
    img_data = pix.tobytes("png")
    import io
    from PIL import Image
    img = Image.open(io.BytesIO(img_data))
    doc.close()
    return img


def _extract_extraction_plan(pdf_path, product_code=""):
    """从提取方案PDF提取提取方案信息。
    一个PDF = 一个提取方案 = 表格一行
    复用 parser.parse_withdrawal 解析第17页的逐年表格
    """
    import fitz
    import re
    import os
    import sys
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from parser import parse_withdrawal

    doc = fitz.open(pdf_path)
    result = {}

    # 如果没有product_code，自动从PDF检测
    product_names = {
        "TRST": "TRST 信守明天多元货币保险计划",
        "PRMESP": "PRMESP 世誉财富保险计划",
    }
    if not product_code:
        for i in range(min(10, len(doc))):
            text = doc[i].get_text()
            if "TRST" in text or "信守明天" in text:
                product_code = "TRST"
                break
            elif "PRMESP" in text or "世譽" in text or "世誉" in text:
                product_code = "PRMESP"
                break

    result["product_code"] = product_code
    result["product_name"] = product_names.get(product_code, product_code)

    # 提取年龄和性别
    for i in range(min(10, len(doc))):
        text = doc[i].get_text()
        # 年龄：翌年歲（ANB）：39歲
        m = re.search(r'翌年[歲岁][（(]ANB[)）][:：]\s*(\d+)\s*[歲岁]', text)
        if m:
            result["age"] = m.group(1) + "岁"
        # 性别：性別：女（非吸煙者）
        m = re.search(r'性[別别][:：]\s*(男|女)', text)
        if m:
            result["gender"] = m.group(1)
        if result.get("age") and result.get("gender"):
            break

    # 把第12-22页的文本拼起来，让parse_withdrawal自己找位置
    all_text = ""
    for i in range(12, min(22, len(doc))):
        all_text += doc[i].get_text() + "\n"

    # anb_age 暂时传0，后面可以优化
    withdrawal_data = parse_withdrawal(all_text, 0, product_code)
    if withdrawal_data:
        result["start_year"] = withdrawal_data.get("start_year")
        result["annual_extraction"] = withdrawal_data.get("annual_withdrawal")

    # 概览页：提取总额、101岁退保价值
    if product_code == "PRMESP":
        summary_page = 20  # 第21页
    elif product_code == "TRST":
        summary_page = 23  # 第24页
    else:
        summary_page = None
        for i in range(len(doc)):
            text = doc[i].get_text()
            if "補充說明摘要之概覽" in text and "提取" in text:
                summary_page = i
                break

    if summary_page is not None and summary_page < len(doc):
        text = doc[summary_page].get_text()
        # 提取总额
        m = re.search(r'提取之總額\s*([\d,]+)\s*([\d,]+)\s*([\d,]+)', text)
        if m:
            result["total_extraction"] = float(m.group(3).replace(',', ''))
        # 101岁退保价值
        m = re.search(r'於101歲.*?退保價值總額.*?([\d,]+)\s*([\d,]+)\s*([\d,]+)', text, re.S)
        if m:
            result["surrender_value_101"] = float(m.group(3).replace(',', ''))

    doc.close()
    return result


def _pdf_to_images(pdf_path, max_pages=5):
    """把PDF前N页转成图片列表（PIL Image）。"""
    import fitz
    images = []
    doc = fitz.open(pdf_path)
    for i in range(min(len(doc), max_pages)):
        page = doc[i]
        pix = page.get_pixmap(dpi=150)
        img_data = pix.tobytes("png")
        import io
        from PIL import Image
        img = Image.open(io.BytesIO(img_data))
        images.append(img)
    doc.close()
    return images


def _html_to_image(html_path):
    """用subprocess调用单独的Python脚本把HTML截图成图片（避免asyncio loop问题）。"""
    import subprocess
    import sys
    import tempfile
    import os
    from PIL import Image
    
    # 生成临时PNG文件路径
    tmp_png = os.path.join(tempfile.gettempdir(), f"shot_{os.getpid()}_{hash(html_path)}.png")
    
    # 把路径里的反斜杠换成正斜杠，避免转义问题
    html_path_unix = html_path.replace('\\', '/')
    tmp_png_unix = tmp_png.replace('\\', '/')
    
    # 写一个临时脚本
    script = f'''
from playwright.sync_api import sync_playwright
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page(viewport={{'width': 800, 'height': 1200}})
    page.goto("file:///{html_path_unix}")
    page.wait_for_load_state("networkidle")
    page.screenshot(path="{tmp_png_unix}", full_page=True)
    browser.close()
'''
    script_file = os.path.join(tempfile.gettempdir(), f"shot_script_{os.getpid()}.py")
    with open(script_file, 'w', encoding='utf-8') as f:
        f.write(script)
    
    # 执行脚本
    result = subprocess.run([sys.executable, script_file], capture_output=True, text=True)
    
    # 清理临时脚本文件
    os.unlink(script_file)
    
    if result.returncode != 0:
        raise Exception(f"HTML截图失败: {result.stderr}")
    
    img = Image.open(tmp_png)
    img.load()  # 确保图片完全加载
    os.unlink(tmp_png)
    return img


def _load_images_from_paths(paths):
    """从文件路径列表加载图片，PDF转成图片，HTML截图，图片直接加载。"""
    import os
    from PIL import Image
    all_images = []
    for path in paths:
        ext = os.path.splitext(path)[1].lower()
        if ext == '.pdf':
            try:
                imgs = _pdf_to_images(path)
                all_images.extend(imgs)
            except Exception as e:
                print(f"PDF转图片失败 {path}: {e}")
        elif ext in ('.png', '.jpg', '.jpeg', '.gif', '.bmp'):
            try:
                img = Image.open(path)
                all_images.append(img)
            except Exception as e:
                print(f"加载图片失败 {path}: {e}")
        elif ext in ('.html', '.htm'):
            try:
                img = _html_to_image(path)
                all_images.append(img)
            except Exception as e:
                print(f"HTML截图失败 {path}: {e}")
    return all_images


# 产品特征词映射：财务概览文件里通常带产品名，用有区分度的特征词识别，避免"保險計劃"这类通用词误配
_PRODUCT_KEYWORDS = {
    "PRMESP": ["世譽", "世誉", "PRMESP"],
    "TRST": ["信守明天", "TRST"],
    "CIM": ["誠保一生", "诚保一生", "智危", "CIM"],
    "CIE": ["危疾首護", "危疾首护", "CIE"],
    "PHOR": ["誠譽", "诚誉", "PHOR"],
    "LFRE": ["雋珀", "隽珀", "LFRE"],
    "PLP": ["美好人生", "PLP"],
    "PLT": ["守護家人", "守护家人", "PLT"],
    "PLJ": ["PLJ"],
    "MCVIP": ["MCVIP", "醫療", "医疗"],
}

# 通用词：不能用于区分产品，匹配前先剔除
_GENERIC_WORDS = ["保險計劃", "保险计划", "保險", "保险", "計劃", "计划", "保障", "財富", "财富", "多元貨幣", "多元货币", "危疾"]


def _norm_text(s):
    """归一化文本：全角转半角，去掉空格/横线/标点，只留数字字母中文，大写。"""
    import re
    if not s:
        return ""
    s = str(s)
    res = []
    for ch in s:
        code = ord(ch)
        if code == 0x3000:
            res.append(" ")
        elif 0xFF01 <= code <= 0xFF5E:
            res.append(chr(code - 0xFEE0))
        else:
            res.append(ch)
    s = "".join(res)
    return re.sub(r"[^0-9a-zA-Z\u4e00-\u9fff]", "", s).upper()


def _product_keywords(policy):
    """按产品代码/产品名给出有区分度的特征词列表。"""
    code = (policy.get("product_code") or "").upper()
    name = policy.get("product_name") or ""
    # 先按产品代码前缀找（CIM3/CIM1 都归 CIM）
    for prefix, kws in _PRODUCT_KEYWORDS.items():
        if prefix in code:
            return kws
    # 代码认不出，从产品名里剔除通用词后取剩下的特征片段
    kws = []
    cleaned = name
    for gw in _GENERIC_WORDS:
        cleaned = cleaned.replace(gw, "")
    cleaned = cleaned.strip("【】「」（）() ")
    if len(cleaned) >= 3:
        kws.append(cleaned[:6])
    return kws


def _shot_clip_rect(page):
    """财务概览页裁剪区域：去掉顶部日期/保单号标题行、左侧导航栏、底部备注与页脚，只留主体表格。
    这个 PDF 的表格线是细矩形不是线对象，所以用文本块定位：
    - 上边界：'保單號碼' 所在行
    - 下边界：最后一行的 '已繳總保費'（含金额，排除下方以 * 开头的注释）
    - 左右：该区域内 x>105 的内容范围（105 以左是导航栏）
    """
    rect = page.rect
    blocks = page.get_text('blocks')

    top_y = None
    for kw in ('保單號碼', '保单号码'):
        hits = page.search_for(kw)
        if hits:
            top_y = hits[0].y0 - 4
            break
    if top_y is None:
        top_y = rect.height * 0.095

    bottom_y = None
    for b in blocks:
        txt = b[4].replace('\xa0', ' ').strip()
        if txt.startswith('已繳總保費') or txt.startswith('已繳費保費') or txt.startswith('已缴总保费'):
            if b[3] > top_y:
                bottom_y = max(bottom_y or 0, b[3])
    if bottom_y is None:
        inner = [b for b in blocks if b[1] > top_y and b[3] < rect.height * 0.9]
        bottom_y = max((b[3] for b in inner), default=rect.height * 0.7)
    bottom_y += 8

    inner = [b for b in blocks if b[1] >= top_y - 2 and b[3] <= bottom_y and b[0] > 105]
    if inner:
        left_x = min(b[0] for b in inner) - 4
        right_x = max(b[2] for b in inner) + 4
    else:
        left_x, right_x = 108.0, rect.width - 18

    return (max(rect.x0, left_x), max(rect.y0, top_y), min(rect.x1, right_x), min(rect.y1, bottom_y))


def _shot_pdf_to_image(path):
    """财务概览 PDF 只取第 1 页，并按主体表格区域裁剪。"""
    import fitz
    from PIL import Image
    doc = fitz.open(path)
    try:
        if len(doc) == 0:
            return None
        page = doc[0]
        clip = _shot_clip_rect(page)
        if clip and clip[3] - clip[1] > 50 and clip[2] - clip[0] > 100:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=fitz.Rect(*clip))
        else:
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        return Image.open(io.BytesIO(pix.tobytes('png')))
    finally:
        doc.close()


def _build_shot_candidates(shot_paths):
    """把财务概览文件逐一转成候选：{path, text, image, used}。
    关键：每个文件单独配对一张图，不用全局图片列表下标，避免 PDF 多页时错位。
    """
    import os
    candidates = []
    for path in shot_paths:
        ext = os.path.splitext(path)[1].lower()
        text = ""
        image = None
        try:
            if ext == '.pdf':
                import fitz
                doc = fitz.open(path)
                text = doc[0].get_text() if len(doc) > 0 else ""
                doc.close()
                image = _shot_pdf_to_image(path)
            elif ext in ('.png', '.jpg', '.jpeg', '.gif', '.bmp'):
                from PIL import Image
                image = Image.open(path)
            elif ext in ('.html', '.htm'):
                import re
                raw = ""
                for enc in ('utf-8', 'gbk'):
                    try:
                        with open(path, 'r', encoding=enc, errors='ignore') as f:
                            raw = f.read()
                        break
                    except Exception:
                        continue
                # 剥掉标签和实体，拿到可匹配的纯文本
                text = re.sub(r'<[^>]+>', ' ', raw)
                text = text.replace('&nbsp;', ' ').replace('&amp;', '&')
                image = _html_to_image(path)
        except Exception as e:
            print(f"财务概览加载失败 {path}: {e}")
            continue
        if image is None:
            continue
        candidates.append({"path": path, "text": text, "image": image, "used": False})
    return candidates


def _score_shot_match(policy, text, keywords):
    """给一份财务概览对某个保单的打分。"""
    text_n = _norm_text(text)
    score = 0
    reasons = []
    billing = _norm_text(policy.get("billing_number", ""))
    if len(billing) >= 5 and billing in text_n:
        score += 100
        reasons.append("保单号")
    for kw in keywords:
        kw_n = _norm_text(kw)
        if kw_n and kw_n in text_n:
            score += 40
            reasons.append(f"产品词[{kw}]")
            break
    name = policy.get("insured_name", "")
    if name and _norm_text(name) in text_n:
        score += 30
        reasons.append("受保人")
    return score, reasons


def _assign_shots(policies, shot_candidates):
    """先整体分配财务概览，再生成PPT：保证一对一、不重复、不缺页。"""
    import os
    assign = {}
    if not shot_candidates:
        return assign
    for p in policies:
        kw = _product_keywords(p)
        best = None
        best_score = 0
        best_reason = ""
        for c in shot_candidates:
            if c["used"]:
                continue
            score, reasons = _score_shot_match(p, c["text"], kw)
            if score > best_score:
                best_score = score
                best = c
                best_reason = "+".join(reasons)
        # 至少要命中一项身份证据（名字30分/产品词40分/保单号100分）
        if best is not None and best_score >= 30:
            best["used"] = True
            assign[id(p)] = best
            print(f"  财务概览匹配({best_reason}, 分数{best_score}): {os.path.basename(best['path'])} -> {p.get('product_name', '')}")
    # 兜底：没匹配上的产品，按顺序分配剩下的财务概览，避免缺页
    remaining = [c for c in shot_candidates if not c["used"]]
    for p in policies:
        if id(p) in assign:
            continue
        if not remaining:
            break
        c = remaining.pop(0)
        c["used"] = True
        assign[id(p)] = c
        print(f"  财务概览按顺序兜底: {os.path.basename(c['path'])} -> {p.get('product_name', '')}")
    return assign


def _cn_num_to_int(s):
    """把 '二' / '十五' / '二十' / '11' 这类写法转成整数，认不出返回 None。"""
    if not s:
        return None
    if s.isdigit():
        return int(s)
    cn = {'零': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6, '七': 7, '八': 8, '九': 9}
    if '十' in s:
        parts = s.split('十')
        tens = cn.get(parts[0], 1) if parts[0] else 1
        ones = cn.get(parts[1], 0) if len(parts) > 1 and parts[1] else 0
        return tens * 10 + ones
    return cn.get(s)


def _extract_year_value(path, text=""):
    """提取「第N年提取」里的 N（支持中文数字），认不出返回 9999（排到最后）。"""
    import re
    import os
    name = os.path.basename(path)
    num = r'([0-9一二三四五六七八九十]+)'
    m = re.search(r'第\s*' + num + r'\s*年', name) or re.search(num + r'\s*年', name)
    if not m and text:
        m = re.search(r'第\s*' + num + r'\s*(?:個保單年度|年)', text)
    if not m:
        return 9999
    val = _cn_num_to_int(m.group(1))
    return val if val is not None else 9999


def _extract_insured_name(text):
    import re
    m = re.search(r"受保人\s*\n?\s*[A-Z ]*([\u4e00-\u9fff]{2,5})\s*\(", text)
    return m.group(1) if m else ""


def _order_extract_paths(policies, extract_paths):
    """提取方案排序：先按产品分组（组顺序与产品详情页一致），组内按「第N年提取」从小到大。
    认不出属于哪个产品的，排在最后并保持上传顺序。
    """
    import os

    groups = {}
    others = []
    for idx, path in enumerate(extract_paths):
        text = ""
        code = ""
        pname = ""
        insured = ""
        year = None
        try:
            import fitz
            from parser import parse_basic_from_text, parse_withdrawal
            doc = fitz.open(path)
            head = ""
            for i in range(min(5, len(doc))):
                head += doc[i].get_text() + "\n"
            basic = parse_basic_from_text(head)
            code = (basic.get("product_code") or "").upper()
            pname = basic.get("product_name") or ""
            insured = _extract_insured_name(head)
            # 内容优先：从逐年表格里解析「第N年开始提取」
            body = ""
            for i in range(12, min(22, len(doc))):
                body += doc[i].get_text() + "\n"
            doc.close()
            if "TRST" in head or "信守明天" in head:
                code = "TRST"
            elif "PRMESP" in head or "世譽" in head or "世誉" in head:
                code = "PRMESP"
            if body and code:
                try:
                    wd = parse_withdrawal(body, 0, code)
                    if wd:
                        year = wd.get("start_year")
                except Exception:
                    year = None
        except Exception as e:
            print(f"提取方案排序解析失败 {os.path.basename(path)}: {e}")

        # 内容解析不到再退回文件名
        if year is None:
            fn_year = _extract_year_value(path, text)
            year = None if fn_year == 9999 else fn_year
        if year is None:
            year = 9999

        best_i = None
        best_score = 0
        for i, p in enumerate(policies):
            score = 0
            pc = (p.get("product_code") or "").upper()
            pn = p.get("product_name") or ""
            if code and pc:
                if code == pc or code in pc or pc in code:
                    score += 50
            if pname and pn and len(pname) >= 4 and len(pn) >= 4:
                if pname[:6] in pn or pn[:6] in pname:
                    score += 30
            pi = p.get("insured_name") or ""
            if insured and pi and (insured in pi or pi in insured):
                score += 30
            if score > best_score:
                best_score = score
                best_i = i

        item = (year, idx, path)
        if best_i is None:
            others.append(item)
        else:
            groups.setdefault(best_i, []).append(item)

    ordered = []
    for i in range(len(policies)):
        if i in groups:
            for _, _, p in sorted(groups[i]):
                ordered.append(p)
    for _, _, p in sorted(others):
        ordered.append(p)
    return ordered


def generate_family_review_ppt(policies, client_name="", advisor_name="武安孟", extract_paths=None, shot_paths=None):
    """生成家庭保单检视PPT（竖版A4）。
    extract_paths: 提取方案文件路径列表（PDF/图片）
    shot_paths: 财务概览文件路径列表（PDF/图片）
    """
    extract_paths = extract_paths or []
    shot_paths = shot_paths or []
    prs = Presentation()
    prs.slide_width = Inches(8.27)   # A4竖版宽
    prs.slide_height = Inches(11.69)  # A4竖版高

    SW = prs.slide_width
    SH = prs.slide_height

    # ===== Slide 1: 封面 =====
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 红色顶部条
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Inches(0.5))
    shape.fill.solid()
    shape.fill.fore_color.rgb = RED
    shape.line.fill.background()
    # 标题
    _add_textbox(slide, Inches(0.5), Inches(3.5), Inches(7.27), Inches(1.5),
                 f"{client_name}家庭保单检视表", size=32, bold=True, color=RED, align=PP_ALIGN.CENTER)
    # 副标题
    _add_textbox(slide, Inches(0.5), Inches(5), Inches(7.27), Inches(0.8),
                 "用心聆听，更知你心", size=16, color=GRAY, align=PP_ALIGN.CENTER)
    # 顾问和日期
    _add_textbox(slide, Inches(0.5), Inches(10), Inches(3.5), Inches(0.5),
                 f"理财顾问：{advisor_name}", size=12, color=DARK)
    _add_textbox(slide, Inches(4.27), Inches(10), Inches(3.5), Inches(0.5),
                 "日  期：2026年11月", size=12, color=DARK, align=PP_ALIGN.RIGHT)

    # ===== Slide 2: 现有保单汇总 =====
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    # 标题
    _add_textbox(slide, Inches(0.4), Inches(0.3), Inches(7.5), Inches(0.6),
                 "1、现有保单汇总", size=20, bold=True, color=DARK)

    # 按受保人分组
    groups = {}
    for p in policies:
        name = p.get("insured_name") or "未知"
        if name not in groups:
            groups[name] = []
        groups[name].append(p)

    current_top = Inches(1.0)

    # ===== 按受保人汇总年缴保费的小表格 =====
    # 先算每个人的年缴保费合计
    summary_data = []
    total_premium = 0
    for name, plist in groups.items():
        person_total = sum(p.get("annual_premium") or 0 for p in plist)
        summary_data.append((name, person_total))
        total_premium += person_total

    # 汇总表（2列：受保人、年缴保费(美元)）
    rows = len(summary_data) + 2  # +表头+总计行
    cols = 2
    tbl_w = Inches(3.5)
    tbl_shape = slide.shapes.add_table(rows, cols, Inches(0.4), current_top, tbl_w, Inches(0.35 * rows))
    table = tbl_shape.table
    _remove_table_style(table)
    table.columns[0].width = Inches(2)
    table.columns[1].width = Inches(1.5)

    # 表头
    _style_cell(table.cell(0, 0), "受保人", bold=True, color=WHITE, size=11, fill=RGBColor(0xFF, 0xC0, 0x00))
    _style_cell(table.cell(0, 1), "年缴保费(美元)", bold=True, color=WHITE, size=11, fill=RGBColor(0xFF, 0xC0, 0x00))

    # 数据行
    for i, (name, prem) in enumerate(summary_data):
        row_idx = i + 1
        fill = LIGHT_ORANGE if i % 2 == 0 else None
        _style_cell(table.cell(row_idx, 0), name, size=10, fill=fill)
        _style_cell(table.cell(row_idx, 1), f"${prem:,.2f}", size=10, fill=fill)

    # 总计行
    last_row = rows - 1
    _style_cell(table.cell(last_row, 0), "总保费", bold=True, color=RED, size=11, fill=LIGHT_ORANGE)
    _style_cell(table.cell(last_row, 1), f"${total_premium:,.2f}", bold=True, color=RED, size=11, fill=LIGHT_ORANGE)

    current_top += Inches(0.35 * rows + 0.3)

    # ===== 每个人的明细表格 =====
    col_widths = [Inches(0.8), Inches(1.8), Inches(1.0), Inches(0.9), Inches(0.8), Inches(0.9), Inches(0.7), Inches(0.8)]
    headers = ["类型", "产品", "缴费编号", "保额", "前十年增", "年缴保费", "供款期", "剩余缴费"]

    for name, plist in groups.items():
        # 人名标题
        _add_textbox(slide, Inches(0.4), current_top, Inches(7.5), Inches(0.4),
                     f"{name}保单", size=13, bold=True, color=RED)
        current_top += Inches(0.45)

        # 表格
        rows = len(plist) + 1
        cols = len(headers)
        total_w = sum(col_widths, Emu(0))
        table_shape = slide.shapes.add_table(rows, cols, Inches(0.4), current_top,
                                             total_w, Inches(0.3 * rows))
        table = table_shape.table
        _remove_table_style(table)

        for i, w in enumerate(col_widths):
            table.columns[i].width = w

        # 表头
        for i, h in enumerate(headers):
            _style_cell(table.cell(0, i), h, bold=True, color=WHITE, size=9, fill=RED)

        # 数据行
        for r, p in enumerate(plist):
            row_idx = r + 1
            cat = p.get("category", "")
            cat_short = "储蓄" if "储蓄" in cat else ("危疾" if "危疾" in cat else ("医疗" if "医疗" in cat else "其他"))
            sa = f"${p['sum_assured']:,.0f}" if p.get("sum_assured") else "—"
            bonus = f"${p['first_10y_bonus']:,.0f}" if p.get("first_10y_bonus") else "—"
            prem = f"${p['annual_premium']:,.0f}" if p.get("annual_premium") else "—"
            pay = f"{p['pay_years']}年" if p.get("pay_years") else "—"
            remain = f"{p['remaining_pay_years']}年" if p.get("remaining_pay_years") and p['remaining_pay_years'] > 0 else "缴费期满"

            fill = LIGHT if r % 2 == 1 else None
            _style_cell(table.cell(row_idx, 0), cat_short, size=9, fill=fill)
            _style_cell(table.cell(row_idx, 1), p.get("product_name", "—"), size=9, align=PP_ALIGN.LEFT, fill=fill)
            _style_cell(table.cell(row_idx, 2), p.get("billing_number", "—"), size=8, fill=fill)
            _style_cell(table.cell(row_idx, 3), sa, size=9, fill=fill)
            _style_cell(table.cell(row_idx, 4), bonus, size=9, fill=fill)
            _style_cell(table.cell(row_idx, 5), prem, size=9, fill=fill)
            _style_cell(table.cell(row_idx, 6), pay, size=9, fill=fill)
            _style_cell(table.cell(row_idx, 7), remain, size=9, fill=fill)

        current_top += Inches(0.3 * rows + 0.25)

    # 加载提取方案和财务概览图片
    extract_images = _load_images_from_paths(extract_paths) if extract_paths else []
    shot_candidates = _build_shot_candidates(shot_paths) if shot_paths else []
    # 生成PPT前先把财务概览和产品一一分配好（保单号 > 产品特征词 > 受保人名字）
    shot_assign = _assign_shots(policies, shot_candidates)

    # ===== Slide 3: 储蓄计划提取现金价值汇总 =====
    # 先筛选出储蓄类产品
    savings_policies = [p for p in policies if "储蓄" in (p.get("category", ""))]
    if savings_policies:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # 标题
        _add_textbox(slide, Inches(0.4), Inches(0.3), Inches(7.5), Inches(0.6),
                     "2、储蓄计划 提取现金价值", size=20, bold=True, color=RED, align=PP_ALIGN.CENTER)

        current_top = Inches(1.2)

        # 按受保人分组储蓄产品
        savings_groups = {}
        for p in savings_policies:
            name = p.get("insured_name") or "未知"
            if name not in savings_groups:
                savings_groups[name] = []
            savings_groups[name].append(p)

        # 先解析所有提取方案PDF，按产品代码分组
        extract_by_product = {}  # {product_code: [row1, row2, ...]}
        for pdf_path in extract_paths:
            summary = _extract_extraction_plan(pdf_path)
            product_code = summary.get("product_code", "其他")

            total_ext = summary.get("total_extraction", 0)
            surrender_101 = summary.get("surrender_value_101", 0)
            total = total_ext + surrender_101
            start_year = summary.get("start_year")
            annual_ext = summary.get("annual_extraction")
            row = [
                f"第{start_year}年" if start_year else "—",
                f"${annual_ext:,.0f}" if annual_ext else "—",
                f"${total_ext:,.0f}" if total_ext else "—",
                f"${surrender_101:,.0f}" if surrender_101 else "—",
                f"${total:,.0f}" if total else "—",
            ]
            if product_code not in extract_by_product:
                extract_by_product[product_code] = []
            extract_by_product[product_code].append(row)

        # 按受保人显示
        for name_idx, (name, plist) in enumerate(savings_groups.items()):
            # 小卡片：受保人信息
            p = plist[0]  # 取第一个产品的信息
            card_items = [
                ("受保人", name),
                ("编号编号", p.get("billing_number", "—")),
                ("计划名称", p.get("product_name", "—")),
                ("保额(美金)", f"${p['sum_assured']:,.0f}" if p.get("sum_assured") else "—"),
                ("首期保费日", p.get("first_premium_date", "—")),
                ("缴费年期", f"{p['pay_years']}年" if p.get("pay_years") else "—"),
            ]

            rows = 2
            cols = len(card_items)
            table_shape = slide.shapes.add_table(rows, cols, Inches(0.3), current_top,
                                                 Inches(7.7), Inches(0.6))
            table = table_shape.table
            _remove_table_style(table)
            for i, (label, value) in enumerate(card_items):
                table.columns[i].width = Inches(7.7 / cols)
                _style_cell(table.cell(0, i), label, bold=True, size=9, color=DARK, fill=RGBColor(0xFF, 0xE0, 0xB0))
                _style_cell(table.cell(1, i), value, size=9, align=PP_ALIGN.CENTER, fill=RGBColor(0xFF, 0xF3, 0xE0))

            current_top += Inches(0.7)

            # 不同提取方案现金流情况表格
            _add_textbox(slide, Inches(0.3), current_top, Inches(7.7), Inches(0.4),
                         "不同提取方案，现金流情况", size=12, bold=True, color=DARK, align=PP_ALIGN.CENTER)
            current_top += Inches(0.5)

            # 按产品代码找这个受保人的提取方案
            product_code = p.get("product_code", "")
            extract_rows = extract_by_product.get(product_code, [])
            # 按起提年限从小到大排序
            def extract_year(row):
                try:
                    return int(row[0].replace("第", "").replace("年", ""))
                except:
                    return 999
            extract_rows.sort(key=extract_year)
            if not extract_rows:
                extract_rows = [["—", "—", "—", "—", "—"]]

            extract_headers = ["起提年限", "每年提取金额(USD)", "终身总提取(USD)", "百年后剩余现金价值(USD)", "合计总收益(提取+剩余)"]

            rows = len(extract_rows) + 1
            cols = len(extract_headers)
            table_shape = slide.shapes.add_table(rows, cols, Inches(0.3), current_top,
                                                 Inches(7.7), Inches(0.35 * rows))
            table = table_shape.table
            _remove_table_style(table)

            for i, h in enumerate(extract_headers):
                _style_cell(table.cell(0, i), h, bold=True, size=9, color=WHITE, fill=RED)

            for r, row_data in enumerate(extract_rows):
                fill = LIGHT if r % 2 == 1 else None
                for c, val in enumerate(row_data):
                    _style_cell(table.cell(r+1, c), val, size=9, fill=fill)

            current_top += Inches(0.35 * rows + 0.3)

        # 底部说明
        _add_textbox(slide, Inches(0.3), Inches(11), Inches(7.7), Inches(0.4),
                     "请参考计划书详情", size=10, color=GRAY, align=PP_ALIGN.CENTER)

    # ===== 每个产品的详情页 =====
    # 产品详情页从提取方案PDF里截取第6页
    from parser import parse_basic_from_text
    
    # 先解析所有提取方案PDF的产品信息
    pdf_info_list = []
    for ep in extract_paths:
        try:
            import fitz
            doc_tmp = fitz.open(ep)
            text = ""
            for j in range(min(10, len(doc_tmp))):
                text += doc_tmp[j].get_text() + "\n"
            doc_tmp.close()
            basic = parse_basic_from_text(text)
            
            # 提取受保人名字
            import re
            insured_name = None
            m = re.search(r"受保人\s*\n?\s*[A-Z ]*([\u4e00-\u9fff]{2,5})\s*\(", text)
            if m:
                insured_name = m.group(1)
            
            pdf_info = {
                'path': ep,
                'product_name': basic.get("product_name", ""),
                'age': basic.get("age", ""),
                'gender': basic.get("gender", ""),
                'product_code': basic.get("product_code", ""),
                'insured_name': insured_name,
                'is_savings': "储蓄" in basic.get("category", "") or "PRMESP" in basic.get("product_code", "") or "TRST" in basic.get("product_code", ""),
            }
            pdf_info_list.append(pdf_info)
        except:
            pass
    
    # 第二步：按产品名分组，同产品的保单和PDF都按年龄排序，一一对应
    # 先按产品代码分组保单
    policies_by_code = {}
    for p in policies:
        code = p.get("product_code", "")
        if code not in policies_by_code:
            policies_by_code[code] = []
        policies_by_code[code].append(p)
    
    # 同产品的保单按年龄从小到大排序
    for code in policies_by_code:
        policies_by_code[code].sort(key=lambda x: int(x.get("age") or 0))
    
    # 同产品的PDF按年龄从小到大排序
    pdfs_by_code = {}
    for pdf_info in pdf_info_list:
        code = pdf_info.get("product_code", "")
        if code not in pdfs_by_code:
            pdfs_by_code[code] = []
        pdfs_by_code[code].append(pdf_info)
    
    for code in pdfs_by_code:
        pdfs_by_code[code].sort(key=lambda x: int(x.get("age") or 0))
    
    # 一一对应：同产品的第N个保单对应同产品的第N个PDF
    matched_pairs = {}  # 保单索引 -> PDF路径
    for code in policies_by_code:
        policy_list = policies_by_code[code]
        pdf_list = pdfs_by_code.get(code, [])
        for idx, p in enumerate(policy_list):
            if idx < len(pdf_list):
                matched_pairs[id(p)] = pdf_list[idx]['path']
    
    for i, p in enumerate(policies):
        product_name = p.get("product_name", "")
        product_code = p.get("product_code", "")
        age = p.get("age", "")
        gender = p.get("gender", "")
        insured_name = p.get("insured_name", "")
        
        print(f"=== 产品{i+1}: {product_name} ({product_code}), {age}岁{gender}, 受保人={insured_name} ===")
        for pdf_info in pdf_info_list:
            print(f"  PDF: {pdf_info['product_name']} ({pdf_info['product_code']}), {pdf_info['age']}岁{pdf_info['gender']}, 受保人={pdf_info['insured_name']}")
        
        # 第一步：优先按受保人名字匹配
        file_path = None
        if insured_name:
            for pdf_info in pdf_info_list:
                if pdf_info['insured_name'] and insured_name in pdf_info['insured_name']:
                    file_path = pdf_info['path']
                    print(f"  名字匹配成功: {pdf_info['path']}")
                    break
        
        # 第二步：按产品名分组 + 按年龄排序 + 一一对应
        if not file_path:
            if id(p) in matched_pairs:
                file_path = matched_pairs[id(p)]
                print(f"  分组排序对应成功: {file_path}")
        
        # 第三步：产品+年龄性别匹配
        if not file_path:
            for pdf_info in pdf_info_list:
                # 先匹配产品名
                product_match = False
                if product_code and product_code in pdf_info['product_code']:
                    product_match = True
                elif "信守明天" in product_name and "信守明天" in pdf_info['product_name']:
                    product_match = True
                elif "世譽" in product_name and "世譽" in pdf_info['product_name']:
                    product_match = True
                elif "世誉" in product_name and "世誉" in pdf_info['product_name']:
                    product_match = True
                elif "智危" in product_name or "CIM" in product_name or "誠保一生" in product_name:
                    if "智危" in pdf_info['product_name'] or "CIM" in pdf_info['product_name'] or "誠保一生" in pdf_info['product_name']:
                        product_match = True
                elif product_name and pdf_info['product_name'] and product_name in pdf_info['product_name']:
                    product_match = True
                
                if not product_match:
                    continue
                
                # 再匹配性别（年龄放宽±2岁）
                age_match = True
                if age and pdf_info['age']:
                    try:
                        age_match = abs(int(age) - int(pdf_info['age'])) <= 2
                    except:
                        age_match = True
                if gender and pdf_info['gender']:
                    age_match = age_match and (gender == pdf_info['gender'])
                
                if age_match:
                    file_path = pdf_info['path']
                    print(f"  产品+年龄性别匹配成功: {pdf_info['path']}")
                    break
        
        # 第四步：只按产品名匹配
        if not file_path:
            print(f"  年龄性别匹配失败，只按产品名匹配")
            for pdf_info in pdf_info_list:
                product_match = False
                if product_code and product_code in pdf_info['product_code']:
                    product_match = True
                elif "信守明天" in product_name and "信守明天" in pdf_info['product_name']:
                    product_match = True
                elif "世譽" in product_name and "世譽" in pdf_info['product_name']:
                    product_match = True
                elif "世誉" in product_name and "世誉" in pdf_info['product_name']:
                    product_match = True
                elif "智危" in product_name or "CIM" in product_name or "誠保一生" in product_name:
                    if "智危" in pdf_info['product_name'] or "CIM" in pdf_info['product_name'] or "誠保一生" in pdf_info['product_name']:
                        product_match = True
                if product_match:
                    file_path = pdf_info['path']
                    print(f"  产品名匹配成功: {pdf_info['path']}")
                    break
        
        if not file_path and len(pdf_info_list) > 0:
            print(f"  匹配失败，兜底用第一个: {pdf_info_list[0]['path']}")
            file_path = pdf_info_list[0]['path']
        
        # --- 第1页：基本信息 + PDF第6页（铺满） ---
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        # 标题条
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Inches(0.5))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RED
        shape.line.fill.background()
        _add_textbox(slide, Inches(0.3), Inches(0.05), Inches(7.5), Inches(0.4),
                     f"{p.get('product_name', '未知产品')} - 保单详情", size=14, bold=True, color=WHITE)

        # 基本信息表格（精简版，一行显示更多）
        info_items = [
            ("受保人", p.get("insured_name", "—")),
            ("缴费编号", p.get("billing_number", "—")),
            ("保障额", f"${p['sum_assured']:,.0f}" if p.get("sum_assured") else "—"),
            ("年缴保费", f"${p['annual_premium']:,.2f}" if p.get("annual_premium") else "—"),
            ("供款期", f"{p['pay_years']}年" if p.get("pay_years") else "—"),
            ("剩余缴费", f"{p['remaining_pay_years']}年" if p.get("remaining_pay_years") and p['remaining_pay_years'] > 0 else "缴费期满"),
        ]
        rows = len(info_items)
        table_shape = slide.shapes.add_table(rows, 2, Inches(0.3), Inches(0.7),
                                             Inches(7.6), Inches(0.3 * rows))
        table = table_shape.table
        _remove_table_style(table)
        table.columns[0].width = Inches(1.2)
        table.columns[1].width = Inches(6.4)
        for r, (label, value) in enumerate(info_items):
            fill = LIGHT if r % 2 == 1 else None
            _style_cell(table.cell(r, 0), label, bold=True, size=9, align=PP_ALIGN.RIGHT, fill=fill)
            _style_cell(table.cell(r, 1), value, size=9, align=PP_ALIGN.LEFT, fill=fill)

        # 嵌入PDF第6页截图
        try:
            top = Inches(0.7 + 0.3 * rows + 0.1)
            max_w = Inches(7.8)
            max_h = Inches(7.5)

            if file_path:
                # 第6页，自动检测裁剪区域
                import fitz
                doc = fitz.open(file_path)
                page = doc[5]  # 第6页
                # 找所有长水平线（只找上边界以下的）
                title_rects = page.search_for('基本計劃')
                title_rects.sort(key=lambda r: r.y0)
                top_y = title_rects[0].y0 - 10 if title_rects else 50
                h_lines = []
                for draw in page.get_drawings():
                    for item in draw['items']:
                        if item[0] == 'l':
                            p1 = item[1]
                            p2 = item[2]
                            if abs(p1.y - p2.y) < 2 and p2.x - p1.x > 400 and p1.y > top_y:
                                h_lines.append(p1.y)
                h_lines.sort()
                bottom_y = h_lines[-1] + 5 if h_lines else 680
                clip = (20, top_y, 580, bottom_y)
                doc.close()
                img = _pdf_page_to_image(file_path, 5, clip=clip)
            else:
                img = None

            if img:
                img_buf = io.BytesIO()
                img.save(img_buf, format='PNG')
                img_buf.seek(0)
                img_w, img_h = img.size
                ratio = min(max_w / img_w, max_h / img_h)
                pic_w = int(img_w * ratio)
                pic_h = int(img_h * ratio)
                left = int((SW - pic_w) / 2)
                slide.shapes.add_picture(img_buf, left, top, width=pic_w, height=pic_h)
        except Exception as e:
            print(f"嵌入PDF截图失败: {e}")
        
        # 每个产品详情页后面跟着对应的财务概览页（用预先分配好的结果）
        matched_candidate = shot_assign.get(id(p))
        matched_shot = matched_candidate["image"] if matched_candidate else None

        if matched_shot:
            slide = prs.slides.add_slide(prs.slide_layouts[6])
            # 标题条
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Inches(0.5))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RED
            shape.line.fill.background()
            _add_textbox(slide, Inches(0.3), Inches(0.05), Inches(7.5), Inches(0.4),
                         f"{p.get('product_name', '未知产品')} - 财务概览", size=14, bold=True, color=WHITE)
            
            # 嵌入截图
            img_buf = io.BytesIO()
            matched_shot.save(img_buf, format='PNG')
            img_buf.seek(0)
            img_w, img_h = matched_shot.size
            max_w = Inches(7.8)
            max_h = Inches(10.5)
            ratio = min(max_w / img_w, max_h / img_h)
            pic_w = int(img_w * ratio)
            pic_h = int(img_h * ratio)
            left = int((SW - pic_w) / 2)
            top = Inches(0.7)
            slide.shapes.add_picture(img_buf, left, top, width=pic_w, height=pic_h)

    # ===== 每个提取方案PDF生成一页PPT =====
    # 先按产品分组（顺序同产品详情页），组内按「第N年提取」从小到大
    from parser import parse_basic_from_text
    ordered_extracts = _order_extract_paths(policies, extract_paths)
    if ordered_extracts:
        import os as _os
        print("提取方案顺序: " + " | ".join(_os.path.basename(x) for x in ordered_extracts))
    for extract_idx, extract_file_path in enumerate(ordered_extracts):
        # 先判断是不是储蓄类PDF，重疾类PDF不做提取方案表格
        try:
            import fitz
            doc_tmp = fitz.open(extract_file_path)
            text = ""
            for i in range(min(5, len(doc_tmp))):
                text += doc_tmp[i].get_text() + "\n"
            doc_tmp.close()
            basic = parse_basic_from_text(text)
            product_code = basic.get("product_code", "")
            category = basic.get("category", "")
            # 重疾类PDF跳过，不做提取方案表格
            if "危疾" in category or "重疾" in category or "CIM" in product_code:
                print(f"提取方案{extract_idx+1}是重疾类PDF，跳过提取方案表格")
                continue
        except:
            pass
        
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Inches(0.5))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RED
        shape.line.fill.background()
        
        # 从PDF前几页提取产品名字、年龄、性别
        title_text = f"提取方案 {extract_idx + 1}"
        try:
            import fitz
            import re
            doc_tmp = fitz.open(extract_file_path)
            # 拼前10页文本
            text = ""
            for i in range(min(10, len(doc_tmp))):
                text += doc_tmp[i].get_text() + "\n"
            doc_tmp.close()
            basic = parse_basic_from_text(text)
            product_name = basic.get("product_name", "")
            age = basic.get("age", "")
            gender = basic.get("gender", "")
            
            # 从文件名提取提取年限信息（比如"第二年提取"、"第六年提取"）
            filename = extract_file_path.split('\\')[-1]
            extract_year = None
            m = re.search(r'第(\d+)年提取', filename)
            if m:
                extract_year = f"第{m.group(1)}年提取"
            else:
                m = re.search(r'(\d+)年提取', filename)
                if m:
                    extract_year = f"第{m.group(1)}年提取"
                else:
                    # 从PDF文本里找
                    m = re.search(r'第(\d+)年.*?提取', text)
                    if m:
                        extract_year = f"第{m.group(1)}年提取"
            
            # 找受保人名字
            insured_name = None
            m = re.search(r"受保人\s*\n?\s*[A-Z ]*([\u4e00-\u9fff]{2,5})\s*\(", text)
            if m:
                insured_name = m.group(1)
            
            # 组合标题
            title_parts = []
            if insured_name:
                title_parts.append(insured_name)
            if product_name:
                title_parts.append(product_name)
            if age:
                title_parts.append(f"{age}岁{gender or ''}")
            if extract_year:
                title_parts.append(extract_year)
            if title_parts:
                title_text = " - ".join(title_parts)
        except Exception as e:
            print(f"提取方案{extract_idx+1}标题解析失败: {e}")
        
        _add_textbox(slide, Inches(0.3), Inches(0.05), Inches(7.5), Inches(0.4),
                     title_text, size=14, bold=True, color=WHITE)

        try:
            import fitz
            doc = fitz.open(extract_file_path)
            
            # 找第一个包含"投保時整付保費"或"投保時每年保費"且包含"現金價值提取"或"現金提取"的页面作为开始
            start_page = None
            for i in range(10, min(35, len(doc))):
                text = doc[i].get_text()
                if ('投保時整付保費' in text or '投保時每年保費' in text) and ('現金價值提取' in text or '現金提取' in text):
                    start_page = i
                    break
            
            if start_page is None:
                start_page = 16  # 默认第17页（索引16）
            
            # 从开始页往后找，直到遇到页面开头是新的表格（投保時整付保費或投保時每年保費）
            page_nums = [start_page]
            for i in range(start_page + 1, min(start_page + 5, len(doc))):
                text = doc[i].get_text()
                first_200 = text[:200]
                if '投保時整付保費' in first_200 or '投保時每年保費' in first_200:
                    break
                page_nums.append(i)
            
            print(f"提取方案{extract_idx+1}表格页数: {len(page_nums)}, 页码: {[p+1 for p in page_nums]}")

            # 自动找表格裁剪区域的函数
            def _auto_find_table_clip(page, page_idx=0):
                # 先确定上边界
                if page_idx == 0:
                    # 第一页：从y=205开始
                    top_y = 205
                elif page_idx == 1:
                    # 第二页：找"保單年度"表头关键词，往上留10点边距（保留表头）
                    title_rects = page.search_for('保單年度')
                    if title_rects:
                        top_y = title_rects[0].y0 - 10
                    else:
                        top_y = 80  # 兜底
                else:
                    # 第三页及以后：找"保單年度"表头关键词，往下64点（跳过表头）
                    title_rects = page.search_for('保單年度')
                    if title_rects:
                        top_y = title_rects[0].y0 + 64
                    else:
                        top_y = 155  # 兜底
                # 下边界：找上边界以下的最下面的长水平线
                h_lines = []
                for draw in page.get_drawings():
                    for item in draw['items']:
                        if item[0] == 'l':
                            p1 = item[1]
                            p2 = item[2]
                            if abs(p1.y - p2.y) < 2 and p2.x - p1.x > 400 and p1.y > top_y:
                                h_lines.append(p1.y)
                h_lines.sort()
                if not h_lines:
                    return (30, top_y, 575, 700)  # 兜底
                bottom_y = h_lines[-1] + 5
                return (30, top_y, 575, bottom_y)

            # 每一页表格单独放一页PPT，不再拼接
            for idx, page_num in enumerate(page_nums):
                try:
                    # 自动找表格裁剪区域
                    page = doc[page_num]
                    clip = _auto_find_table_clip(page, page_idx=idx)
                    img = _pdf_page_to_image(extract_file_path, page_num, clip=clip)
                    if img:
                        # 如果不是第一页，新建一页PPT
                        if idx > 0:
                            slide = prs.slides.add_slide(prs.slide_layouts[6])
                            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, SW, Inches(0.5))
                            shape.fill.solid()
                            shape.fill.fore_color.rgb = RED
                            shape.line.fill.background()
                            _add_textbox(slide, Inches(0.3), Inches(0.05), Inches(7.5), Inches(0.4),
                                         f"{title_text} (续)", size=14, bold=True, color=WHITE)
                        
                        img_buf = io.BytesIO()
                        img.save(img_buf, format='PNG')
                        img_buf.seek(0)

                        # 铺满
                        img_w, img_h = img.size
                        max_w = Inches(7.8)
                        max_h = Inches(10.8)
                        ratio = min(max_w / img_w, max_h / img_h)
                        pic_w = int(img_w * ratio)
                        pic_h = int(img_h * ratio)
                        left = int((SW - pic_w) / 2)
                        slide.shapes.add_picture(img_buf, left, Inches(0.6), width=pic_w, height=pic_h)
                except Exception as e:
                    print(f"嵌入第{page_num+1}页失败: {e}")
            doc.close()
        except Exception as e:
            print(f"提取方案{extract_idx+1}生成失败: {e}")

    # 输出到内存
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return buf
