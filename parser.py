# -*- coding: utf-8 -*-
"""
保诚计划书解析器（MVP）
支持产品：
  - PRMESP  世誉财富保险计划（一次整付）
  - TRST    保诚信守明天多元货币保险计划（年缴，多年缴费）

输入：PDF 文件路径
输出：结构化 dict（产品、受保人、保费、逐年退保价值表等）
说明：基于数字版 PDF 文字层提取；若为扫描图片版则文字层为空，需另接 OCR。
"""
import re
import datetime
import fitz  # PyMuPDF

PRODUCT_NAMES = {
    "PRMESP": "世誉财富保险计划",
    "TRST": "保诚信守明天多元货币保险计划",
    "CIM3": "诚保一生危疾保险计划",
    "CIE3": "守护健康危疾加护保 III",
    "CIP2": "危疾首护保险计划 II",
    "ENCASH": "保诚加护一生住院现金储蓄保险",
    "PLP2": "美好人生保险计划 II",
    "LFRE2": "自主未来保险计划 II",
    "PHOR": "诚誉财富保险计划",
    "PACE": "骏誉财富保险计划",
    "CIT2": "守护健康危疾定期保II",
}

# 产品配置表：集中管理所有产品的属性
# render_type: 渲染模板类型（savings=储蓄险, critical=重疾险, term=定期险, medical=医疗险）
# extra_years: 退保价值表额外展示的年份
# breakeven_year: 预期回本年份
PRODUCT_CONFIG = {
    "PRMESP": {
        "keywords": ["PRMESP", "世譽財富", "世誉财富"],
        "category": "储蓄/退休",
        "render_type": "savings",
        "extra_years": {4},
        "breakeven_year": None,
    },
    "TRST": {
        "keywords": ["TRST", "信守明天"],
        "category": "储蓄/退休",
        "render_type": "savings",
        "extra_years": {8},
        "breakeven_year": 8,
    },
    "CIM3": {
        "keywords": ["CIM3", "誠保一生", "诚保一生"],
        "category": "危疾",
        "render_type": "critical",
        "extra_years": {1},
        "breakeven_year": None,
    },
    "CIE3": {
        "keywords": ["CIE3"],
        "category": "危疾",
        "render_type": "critical",
        "extra_years": {1},
        "breakeven_year": None,
    },
    "CIP2": {
        "keywords": ["CIP2", "危疾首護", "危疾首护"],
        "category": "危疾",
        "render_type": "critical",
        "extra_years": {1},
        "breakeven_year": None,
    },
    "PHOR": {
        "keywords": ["PHOR", "誠譽財富", "诚誉财富"],
        "category": "储蓄/退休",
        "render_type": "savings",
        "extra_years": set(),
        "breakeven_year": None,
    },
    "PACE": {
        "keywords": ["PACE", "駿譽財富", "骏誉财富"],
        "category": "储蓄/退休",
        "render_type": "savings",
        "extra_years": set(),
        "breakeven_year": None,
    },
    "LFRE2": {
        "keywords": ["LFRE2", "自主未來", "自主未来"],
        "category": "储蓄/寿险",
        "render_type": "savings",
        "extra_years": set(),
        "breakeven_year": None,
    },
    "PLP2": {
        "keywords": ["PLP2", "美好人生"],
        "category": "储蓄/寿险",
        "render_type": "savings",
        "extra_years": set(),
        "breakeven_year": None,
    },
    "ENCASH": {
        "keywords": ["ENCASH", "加護一生住院現金", "住院現金儲蓄保險"],
        "category": "住院现金",
        "render_type": "medical",
        "extra_years": set(),
        "breakeven_year": None,
    },
    "PMP": {
        "keywords": ["PMP", "醫療加倍保", "医疗加倍保"],
        "category": "医疗",
        "render_type": "medical",
        "extra_years": set(),
        "breakeven_year": None,
    },
    "CIT2": {
        "keywords": ["CIT2", "守護健康危疾定期保II", "守护健康危疾定期保II"],
        "category": "定期危疾",
        "render_type": "term",
        "extra_years": set(),
        "breakeven_year": None,
    },
}

NUM = re.compile(r"^\d{1,3}(,\d{3})+(\.\d+)?$")  # 25,000 / 1,240,550
VAL = re.compile(r"^\d+(\.\d+)?$")                # 任意数字（含 0 / 25,000 由 NUM 另配）
PLAIN_NUM = re.compile(r"^\d+$")                  # 纯整数年份
ANB = re.compile(r"@?ANB\s*(\d+)\s*歲")


def _to_num(s):
    return float(s.replace(",", ""))


def _elapsed_since_eff(eff_year, eff_month, eff_day):
    """按生效日算已缴年数：生效当天计 1 年，每过一个生效周年再加 1（满周年当天为 2 年）。
    取不到完整生效日返回 None。"""
    if not eff_year:
        return None
    today = datetime.date.today()
    try:
        eff = datetime.date(int(eff_year), int(eff_month or 1), int(eff_day or 1))
    except ValueError:
        return None
    diff = today.year - eff.year
    if (today.month, today.day) < (eff.month, eff.day):
        diff -= 1
    return max(1, diff + 1)


def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    pages = [p.get_text() for p in doc]
    doc.close()
    return "\n".join(pages)


def detect_product(text):
    # 用产品配置表驱动，不再逐个产品写
    for code, config in PRODUCT_CONFIG.items():
        for keyword in config["keywords"]:
            if keyword in text:
                return code
    # PLT系列：守护家人定期寿险（特殊处理，因为有多个续保期）
    if "守護家人" in text or "守护家人" in text:
        if "(PLT1)" in text or "（PLT1）" in text:
            return "PLT1"
        if "(PLT5)" in text or "（PLT5）" in text:
            return "PLT5"
        if "(PLT10)" in text or "（PLT10）" in text:
            return "PLT10"
        if "(PLT20)" in text or "（PLT20）" in text:
            return "PLT20"
        return "PLT"
    return "UNKNOWN"


def parse_basic(text):
    info = {}
    m = re.search(r"翌年歲[（(]ANB[)）][:：]\s*(\d+)\s*歲", text)
    info["anb_age"] = int(m.group(1)) if m else None

    m = re.search(r"性別[：:]\s*(男|女)[（(](非?吸煙者)[)）]", text)
    if m:
        info["gender"] = "男" if m.group(1) == "男" else "女"
        info["smoker"] = "非吸烟" if "非" in m.group(2) else "吸烟"
    else:
        info["gender"] = None
        info["smoker"] = None

    m = re.search(r"保單貨幣\s*[:：]\s*([A-Za-z\u4e00-\u9fff]+)", text)
    info["currency"] = m.group(1) if m else None

    m = re.search(r"投保時整付總保費\s*([\d,]+\.?\d*)", text)
    if m:
        info["premium_type"] = "一次整付"
        info["annual_premium"] = _to_num(m.group(1))
        info["pay_years"] = 1
        info["total_premium"] = _to_num(m.group(1))
    else:
        m2 = re.search(r"投保時每年總保費\s*([\d,]+\.?\d*)", text)
        annual = _to_num(m2.group(1)) if m2 else None
        # 供款年期：PDF表格布局下表头和值不在同一行，从"保費供款年期"往后找
        years = None
        m3 = re.search(r"保費供款年期", text)
        if m3:
            tail = text[m3.end():m3.end()+500]
            m4 = re.search(r"(\d+)\s*年", tail)
            if m4 and int(m4.group(1)) <= 50:
                years = int(m4.group(1))
        info["premium_type"] = "年缴"
        info["annual_premium"] = annual
        info["pay_years"] = years
        info["total_premium"] = (annual * years) if (annual and years) else None
    return info


def _following_numbers(lines, j):
    """从 j 起连续收集数字行（含 0 与带逗号数字）。"""
    out = []
    while j < len(lines) and (NUM.match(lines[j]) or VAL.match(lines[j])):
        out.append(_to_num(lines[j]))
        j += 1
    return out, j


def _parse_rows_from_lines(lines, anb_age, total_premium, n_values, col_paid, col_g, col_e, trst_dynamic=False):
    """从已截取的行里解析退保价值行。col_*是列下标。
    trst_dynamic=True时：TRST补充表列序不固定，动态识别：
      vals[0]=paid, 其余4个值中surrender=最大值, guaranteed=第二大值。"""
    raw = []
    i = 0
    n = len(lines)

    def take_values(j):
        vals = []
        k = j
        while k < n and len(vals) < n_values:
            if NUM.match(lines[k]) or (VAL.match(lines[k]) and "." in lines[k]):
                vals.append(_to_num(lines[k]))
                k += 1
            elif VAL.match(lines[k]):
                v = int(lines[k])
                if len(vals) == 0 and 1 <= v <= 101:
                    break
                vals.append(float(v))
                k += 1
            else:
                break
        return vals, k

    def extract_vals(following):
        if trst_dynamic and len(following) >= 5:
            paid = following[0]
            rest = following[1:]
            surrender = max(rest)
            rest2 = [v for v in rest if v != surrender]
            guaranteed = max(rest2) if rest2 else 0
            return paid, guaranteed, surrender
        return following[col_paid], following[col_g], following[col_e]

    while i < n:
        ln = lines[i]
        if PLAIN_NUM.match(ln) and 1 <= int(ln) <= 101:
            following, j = take_values(i + 1)
            need = max(col_paid, col_g, col_e) if not trst_dynamic else 4
            if len(following) > need:
                year = int(ln)
                age = (anb_age + year) if anb_age else None
                _paid, _g, _e = extract_vals(following)
                raw.append({"year": year, "age": age, "paid": _paid,
                            "guaranteed": _g, "bonus": _e - _g, "surrender": _e})
                i = j
                continue
        m = ANB.search(ln)
        if m:
            following, j = take_values(i + 1)
            need = max(col_paid, col_g, col_e) if not trst_dynamic else 4
            if len(following) > need:
                age = int(m.group(1))
                year = (age - anb_age) if anb_age else None
                _paid, _g, _e = extract_vals(following)
                raw.append({"year": year, "age": age, "paid": _paid,
                            "guaranteed": _g, "bonus": _e - _g, "surrender": _e,
                            "is_age_marker": True})
                i = j
                continue
        i += 1

    out = []
    seen = set()
    tol = max(5000, (total_premium or 0) * 1.0)
    for r in raw:
        if r.get("age") and r["age"] > 101:
            continue
        if total_premium and abs(r["paid"] - total_premium) > tol:
            continue
        key = r.get("year")
        if key in seen:
            continue
        seen.add(key)
        out.append(r)
    out.sort(key=lambda r: (r.get("year") or 0))
    return out


def parse_surrender_table(text, anb_age, total_premium, product="PRMESP"):
    """
    解析不做提取的退保价值表。
    PRMESP 4列：paid(0), A保证(1), B红利(2), C退保总额(3)
    TRST 5列：paid(0), A保证(1), B归原红利(2), C终期红利(3), D退保总额(4)
    """
    is_trst = product in ("TRST", "PACE")
    is_plp2 = product == "PLP2"
    is_lfre2 = product == "LFRE2"
    n_cols = 5 if is_plp2 else (4 if is_lfre2 else (5 if is_trst else 4))
    col_e = 4 if is_plp2 else (3 if is_lfre2 else 3)  # PLP2的surrender在第4列（最后一列），LFRE2在第3列（最后一列）

    def detect_trst_col_g(seg):
        """TRST列序：基础表 paid,A,B,surr,C；补充表 paid,B,C,surr,A。
        根据表头中'保證金額'与'累積歸原紅利'出现顺序判断。"""
        text = " ".join(seg[:30]).replace("説", "說")
        i_a = text.find("保證金額")
        i_b = text.find("累積歸原紅利")
        if i_a >= 0 and i_b >= 0:
            return 1 if i_a < i_b else 4
        return 1  # 默认基础表

    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    def find_table(keywords_exclude, keywords_include):
        for idx, ln in enumerate(lines):
            s = ln.replace("説", "說")
            if all(k in s for k in keywords_include) and not any(k in s for k in keywords_exclude):
                return idx
        return None

    # 1) 优先找「基本計劃補充說明 – 退保價值之說明摘要」（完整年度，不提取）
    start = find_table(["現金價值提取", "現金提取", "身故賠償", "身故赔偿"],
                        ["退保價值", "補充說明"])
    if start is not None:
        seg = lines[start:]
        cut = None
        for idx, ln in enumerate(seg):
            s = ln.replace("説", "說")
            if "身故賠償" in s or "現金價值提取" in s or "現金提取" in s:
                cut = idx
                break
        if cut:
            seg = seg[:cut]
        col_g = detect_trst_col_g(seg) if is_trst else 1
        rows = _parse_rows_from_lines(seg, anb_age, total_premium, n_cols,
                                       col_paid=0, col_g=col_g, col_e=col_e)
        if rows:
            return rows

    # 2) 退回基础表「基本計劃 – 退保價值之說明摘要」（有断档）
    start = find_table(["現金價值提取", "現金提取", "補充說明", "补充说明", "身故賠償"],
                        ["退保價值", "說明摘要"])
    if start is not None:
        seg = lines[start:]
    else:
        seg = lines
    cut = None
    for idx, ln in enumerate(seg):
        s = ln.replace("説", "說")
        if "身故賠償" in s or "現金價值提取" in s or "現金提取" in s:
            cut = idx
            break
    if cut:
        seg = seg[:cut]
    col_g = detect_trst_col_g(seg) if is_trst else 1
    return _parse_rows_from_lines(seg, anb_age, total_premium, n_cols,
                                   col_paid=0, col_g=col_g, col_e=col_e)


def parse_plp2_death(text):
    """解析PLP2身故赔偿表，返回 {year: death_total} 字典。"""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    
    # 找身故赔偿表
    start = None
    for idx, ln in enumerate(lines):
        if "身故賠償之説明摘要" in ln and "4." in ln:
            start = idx
            break
    if start is None:
        return {}
    
    # 截取到下一个表
    seg = lines[start:]
    cut = None
    for idx, ln in enumerate(seg[1:], 1):
        if "5." in ln or "6." in ln or "現金價值提取" in ln:
            cut = idx
            break
    if cut:
        seg = seg[:cut]
    
    # 解析每行：年度 + 7个数字
    # 列顺序：年度、缴付保费、基本保额A、额外身故B、总额A+B、归原红利C、特别红利D、总额A+B+C+D
    result = {}
    i = 0
    n = len(seg)
    while i < n:
        ln = seg[i]
        if PLAIN_NUM.match(ln) and 1 <= int(ln) <= 101:
            year = int(ln)
            # 接下来的7个数字
            vals = []
            k = i + 1
            while k < n and len(vals) < 7:
                if NUM.match(seg[k]) or (VAL.match(seg[k]) and "." in seg[k]):
                    vals.append(_to_num(seg[k]))
                    k += 1
                elif VAL.match(seg[k]):
                    v = int(seg[k])
                    if len(vals) == 0 and 1 <= v <= 101:
                        break
                    vals.append(float(v))
                    k += 1
                else:
                    break
            if len(vals) >= 7:
                # 列顺序：缴付保费、基本保额A、额外身故B、总额A+B、归原红利C、总额A+B+C+D、特别红利D
                # 第6个数字（索引5）是总额A+B+C+D（含非保证身故赔偿）
                death_total = vals[5]
                result[year] = death_total
            i = k
            continue
        i += 1
    
    return result


def parse_lfre2_death(text):
    """解析LFRE2身故赔偿表，返回 {year: death_total} 字典。
    LFRE2是5列：年度、缴付保费、保证金额A、非保证终期红利B、身故赔偿总额A+B"""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    
    # 找身故赔偿表
    start = None
    for idx, ln in enumerate(lines):
        if "身故賠償之説明摘要" in ln and "4." in ln:
            start = idx
            break
    if start is None:
        return {}
    
    # 截取到下一个表
    seg = lines[start:]
    cut = None
    for idx, ln in enumerate(seg[1:], 1):
        if "5." in ln or "6." in ln or "現金價值提取" in ln:
            cut = idx
            break
    if cut:
        seg = seg[:cut]
    
    # 解析每行：年度 + 4个数字
    result = {}
    i = 0
    n = len(seg)
    while i < n:
        ln = seg[i]
        if PLAIN_NUM.match(ln) and 1 <= int(ln) <= 101:
            year = int(ln)
            # 接下来的4个数字
            vals = []
            k = i + 1
            while k < n and len(vals) < 4:
                if NUM.match(seg[k]) or (VAL.match(seg[k]) and "." in seg[k]):
                    vals.append(_to_num(seg[k]))
                    k += 1
                elif VAL.match(seg[k]):
                    v = int(seg[k])
                    if len(vals) == 0 and 1 <= v <= 101:
                        break
                    vals.append(float(v))
                    k += 1
                else:
                    break
            if len(vals) >= 4:
                # 列顺序：缴付保费、保证金额A、非保证终期红利B、身故赔偿总额A+B
                # 第4个数字（索引3）是身故赔偿总额
                death_total = vals[3]
                result[year] = death_total
            i = k
            continue
        i += 1
    
    return result


def _irr(annual, pay_years, surrender, yr):
    """IRR：解 Σ(k=0..p-1) annual/(1+r)^k = surrender/(1+r)^yr 求 r（返回小数）。
    无正回报（退保≤已缴保费）或无解时返回 None。"""
    if surrender <= pay_years * annual or yr <= 0:
        return None

    def f(r):
        s = 0.0
        for k in range(pay_years):
            s += annual / (1 + r) ** k
        return s - surrender / (1 + r) ** yr

    lo, flo = 0.0, f(0.0)
    if flo >= 0:
        return None
    hi = 2.0
    while f(hi) < 0 and hi < 50:
        hi *= 2
    if f(hi) < 0:
        return None
    for _ in range(300):
        mid = (lo + hi) / 2
        fm = f(mid)
        if abs(fm) < 1e-9:
            return mid
        if fm * flo < 0:
            hi = mid
        else:
            lo = mid
            flo = fm
    return (lo + hi) / 2


def compute_metrics(row, total_premium, annual_premium=None, pay_years=None):
    out = dict(row)
    yr = row.get("year") or 1
    if total_premium:
        out["multiple"] = round(row["surrender"] / total_premium, 1)
        # 单利/复利：优先用与 junrong-team「单利福利计算器」一致的口径
        #   单利 = (S-T)/T / 加权平均在账年数(Y-(p-1)/2)
        #   复利 = IRR（年初缴费、Y年退保）
        if annual_premium and pay_years and row["surrender"] > total_premium and yr > 0:
            avg_years = yr - (pay_years - 1) / 2.0
            out["simple_pct"] = round((row["surrender"] - total_premium) / total_premium / avg_years * 100, 1)
            irr = _irr(annual_premium, pay_years, row["surrender"], yr)
            out["compound_pct"] = round(irr * 100, 1) if irr is not None else None
        else:
            # 回退旧口径（缺参数或前期无正回报）
            out["simple_pct"] = round((row["surrender"] - total_premium) / total_premium / yr * 100, 1)
            if row["surrender"] > 0 and yr > 0:
                out["compound_pct"] = round(((row["surrender"] / total_premium) ** (1.0 / yr) - 1) * 100, 1)
            else:
                out["compound_pct"] = None
    return out


def parse_cim3(text, basic):
    """解析 CIM3（诚保一生危疾保险计划）。
    返回 basic_extra / rows（含 compute_metrics）/ product_advantage。"""
    basic_extra = {}
    # (CIM3) 后值顺序：投保时每年保费 | 保障年期 | 投保时保额 | 保费供款年期
    m = re.search(r"\(CIM3\)\s*\n([\d,]+\.?\d*)\s*\n(\S+)\s*\n([\d,]+)\s*\n(\d+)\s*年", text)
    if m:
        basic_extra["annual_premium"] = _to_num(m.group(1))
        basic_extra["coverage"] = m.group(2)
        basic_extra["basic_amount"] = _to_num(m.group(3))
        basic_extra["pay_years"] = int(m.group(4))
    # 升级保障 (CIME3) 首10年赠送
    m2 = re.search(r"\(CIME3\)\s*\n([\d,]+)", text)
    basic_extra["gift10"] = _to_num(m2.group(1)) if m2 else None
    ba = basic_extra.get("basic_amount")
    basic_extra["total_amount"] = (ba + (basic_extra.get("gift10") or 0)) if ba else None
    ann = basic_extra.get("annual_premium")
    yrs = basic_extra.get("pay_years")
    basic_extra["total_premium"] = (ann * yrs) if (ann and yrs) else basic.get("total_premium")
    basic_extra["premium_type"] = "年缴"

    rows = parse_cim3_rows(text, basic.get("anb_age"), basic_extra.get("total_premium"),
                           basic_extra.get("annual_premium"), basic_extra.get("pay_years"))

    basic_amount = basic_extra.get("basic_amount") or 0
    annuity_amount = int(basic_amount * 0.06)
    gift10 = basic_extra.get("gift10") or 0
    total_amount = basic_extra.get("total_amount") or (basic_amount + gift10)

    product_advantage = [
        {
            "title": "保障超全，持续一生",
            "items": [
                "生·给宝宝的专属保障：严重湿疹、荨麻疹、肺炎住院津贴等均可获得补贴；终身保障疾病涵盖127种（49项早期严重疾病+56项严重疾病+15种儿童重疾+7种儿童次级疾病），确诊严重疾病获赔100%保额+累计红利，并豁免后续所有保费",
                "病·深切治疗ICU保障：连续入住ICU三日或以上获赔10%保额；严重疾病延伸保障最高累计900%，每笔赔100%保额可多次赔付——癌症2次、心脏病/中风2次、脑退化症/帕金森1次、主要器官相关2次、其他严重疾病2次",
                f"老·给老人的专属保障：确诊严重脑退化症或帕金森病，提供终身领取的年金，每年{annuity_amount:,}美元直至身故",
                f"死·身故额外保障：严重疾病理赔后、86岁之前身故，额外获赔{basic_amount:,}美元人寿延伸保障",
            ],
        },
        {
            "title": "重疾保额逐年增长，抗通胀",
            "items": [
                f"首10年附送升级保障{gift10:,}美元，总保障{total_amount:,}美元（首10年150%保障）",
                "特别红利随保单年期累积，身故/重疾保障逐年增长，抵御通胀",
                "理赔无需发票报销，与内地社保、医保及商业保险不冲突，全球通用",
            ],
        },
        {
            "title": "有病赔钱，无病储蓄",
            "items": [
                "确诊严重疾病即赔100%基本保额+特别红利，并豁免日后所有保费",
                "终身未患重疾，86岁后身故仍可获赔与重疾保额相当的身故保险金，传承子孙",
                "退保可取回保单现金价值（连本带利），资金灵活周转",
            ],
        },
    ]
    return {"basic_extra": basic_extra, "rows": rows, "product_advantage": product_advantage}


def parse_cie3(text, basic):
    """解析 CIE3（守护健康危疾加护保 III）。数据行结构与 CIM3 同构，复用 parse_cim3_rows。"""
    basic_extra = {}
    # (CIE3) 后值顺序：投保时每年保费 | 保障年期 | 投保时保额 | 保费供款年期
    m = re.search(r"\(CIE3\)\s*\n([\d,]+\.?\d*)\s*\n(\S+)\s*\n([\d,]+)\s*\n(\d+)\s*年", text)
    if m:
        basic_extra["annual_premium"] = _to_num(m.group(1))
        basic_extra["coverage"] = m.group(2).rstrip("*")
        basic_extra["basic_amount"] = _to_num(m.group(3))
        basic_extra["pay_years"] = int(m.group(4))
    # 额外危疾加护保 III (CICE3) 首10年赠送
    m2 = re.search(r"\(CICE3\)\s*\n([\d,]+)", text)
    basic_extra["gift10"] = _to_num(m2.group(1)) if m2 else None
    ba = basic_extra.get("basic_amount")
    basic_extra["total_amount"] = (ba + (basic_extra.get("gift10") or 0)) if ba else None
    ann = basic_extra.get("annual_premium")
    yrs = basic_extra.get("pay_years")
    basic_extra["total_premium"] = (ann * yrs) if (ann and yrs) else basic.get("total_premium")
    basic_extra["premium_type"] = "年缴"

    rows = parse_cim3_rows(text, basic.get("anb_age"), basic_extra.get("total_premium"),
                           basic_extra.get("annual_premium"), basic_extra.get("pay_years"))

    product_advantage = [
        {
            "title": "保障超全，117种病况",
            "items": [
                "严重疾病保障：覆盖56种严重疾病，确诊即赔100%基本保额＋累计特别红利，并豁免后续所有保费",
                "早期严重疾病保障：61种早期严重疾病（原位癌、冠状动脉血管成形术等），按当时保额20%~25%赔付，最多索偿3次",
                "深切治疗ICU保障：连续入住ICU三日或以上，获赔当时保额20%（境外为10%）",
            ],
        },
        {
            "title": "多次赔付，最高860%保额",
            "items": [
                "严重疾病延伸保障：癌症额外2次、心脏病/中风额外2次，每次赔付100%保额",
                "癌症治疗额外保障：持续治疗期间按当时保额60%赔付",
                "首10年额外危疾加护保III：赠送基本保额50%，首10年总保障达150%",
            ],
        },
        {
            "title": "有病赔钱，无病储蓄",
            "items": [
                "确诊严重疾病即赔并豁免日后所有保费；早期疾病理赔后豁免下一年保费",
                "终身未患重疾，86岁后身故仍可获赔与重疾保额相当的身故保险金，传承子孙",
                "退保可取回保单现金价值（连本带利），特别红利随年期累积、保额逐年增长抗通胀",
                "家人同行：两位家庭成员一同投保，首10年赠送50%额外人寿保障；并提供安心医国际二次诊疗服务",
            ],
        },
    ]
    return {"basic_extra": basic_extra, "rows": rows, "product_advantage": product_advantage}


def parse_cip2(text, basic):
    """解析 CIP2（危疾首护保险计划II）。数据行结构与 CIM3 同构，复用 parse_cim3_rows。"""
    basic_extra = {}
    # (CIP2) 后值顺序：投保时每年保费 | 保障年期 | 投保时保额 | 保费供款年期
    m = re.search(r"\(CIP2\)\s*\n([\d,]+\.?\d*)\s*\n(\S+)\s*\n([\d,]+)\s*\n(\d+)\s*年", text)
    if m:
        basic_extra["annual_premium"] = _to_num(m.group(1))
        basic_extra["coverage"] = m.group(2).rstrip("*")
        basic_extra["basic_amount"] = _to_num(m.group(3))
        basic_extra["pay_years"] = int(m.group(4))
    # 额外定期危疾首护保II (CIPE2) 首10年赠送
    m2 = re.search(r"\(CIPE2\)\s*\n([\d,]+)", text)
    basic_extra["gift10"] = _to_num(m2.group(1)) if m2 else None
    ba = basic_extra.get("basic_amount")
    basic_extra["total_amount"] = (ba + (basic_extra.get("gift10") or 0)) if ba else None
    ann = basic_extra.get("annual_premium")
    yrs = basic_extra.get("pay_years")
    basic_extra["total_premium"] = (ann * yrs) if (ann and yrs) else basic.get("total_premium")
    basic_extra["premium_type"] = "年缴"

    rows = parse_cim3_rows(text, basic.get("anb_age"), basic_extra.get("total_premium"),
                           basic_extra.get("annual_premium"), basic_extra.get("pay_years"))

    basic_amount = basic_extra.get("basic_amount") or 0
    gift10 = basic_extra.get("gift10") or 0
    total_amount = basic_extra.get("total_amount") or (basic_amount + gift10)

    product_advantage = [
        {
            "title": "产品特点",
            "items": [
                "有病赔钱，无病储蓄：保额逐年递增，抗通胀和医疗费用增长；重疾保障增长远超市场同类产品！退保时可连本带利取出，具备储蓄增值功能",
                "疾病包括共74种，其中56种重大疾病、18种早期疾病，索偿早期严重疾病保障后，享有12个月保费豁免",
                f"重大疾病保障额预期不低于{basic_amount:,}美元。重大疾病赔偿一次；早期疾病最多赔偿3次，每次可赔偿当时保额的25%",
            ],
        },
        {
            "title": "权益与附加",
            "items": [
                "递增保障权益，疏解通胀压力；自选附加选项，加强保障后盾",
                f"首10年赠送保额：额外定期危疾首护保II赠送{gift10:,}美元，首10年总保障达{total_amount:,}美元",
            ],
        },
        {
            "title": "理赔条件",
            "items": [
                "确诊患上保单覆盖的重大疾病即可申请理赔，无需以治疗疾病为理赔前提，不需要发票报销，全球通用，与国内社保、医保、商业保险不冲突",
            ],
        },
        {
            "title": "传承",
            "items": [
                "若终身未出现重大疾病，百年之时会得到与重大疾病累积保额相同的保险金，传承子孙后代",
            ],
        },
    ]
    return {"basic_extra": basic_extra, "rows": rows, "product_advantage": product_advantage}


def parse_cim3_rows(text, anb_age, total_premium, annual_premium=None, pay_years=None):
    """解析 CIM3 退保价值 + 身故/严重疾病保障表。
    优先取「基本計劃補充說明」的完整逐年表；若无该页，回退到「基本計劃説明摘要」的里程碑表。
    每行 8 个数值：[paid, 保证退保A, 红利B, 退保总额, 保证身故, 升级保障, 红利E, 身故总额]。
    返回 compute_metrics 后的行列表。"""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    n = len(lines)
    start = None
    # 优先：完整逐年表（含「補充」的説明摘要）
    for i, ln in enumerate(lines):
        s = ln.replace("説", "說")
        if ("説明摘要" in s or "說明摘要" in s) and "補充" in s:
            start = i
            break
    # 回退：基本计划说明摘要（里程碑表）
    if start is None:
        for i, ln in enumerate(lines):
            s = ln.replace("説", "說")
            if ("説明摘要" in s or "說明摘要" in s) and "補充" not in s:
                start = i
                break
    if start is None:
        return []

    raw = []
    i = start
    while i < n:
        ln = lines[i]
        if "備註" in ln:
            break
        year = None
        m = ANB.search(ln)
        if m:
            age = int(m.group(1))
            year = (age - anb_age) if anb_age else None
        elif PLAIN_NUM.match(ln) and 1 <= int(ln) <= 101:
            year = int(ln)
        if year is not None:
            vals = []
            j = i + 1
            while j < n and len(vals) < 8 and (NUM.match(lines[j]) or VAL.match(lines[j])):
                vals.append(_to_num(lines[j]))
                j += 1
            if len(vals) >= 8:
                age = (anb_age + year) if anb_age else None
                raw.append({
                    "year": year,
                    "age": age,
                    "is_age_marker": bool(ANB.search(ln)),
                    "paid": vals[0],
                    "guaranteed_surr": vals[1],
                    "bonus_surr": vals[2],
                    "surrender": vals[3],
                    "death_total": vals[7],
                })
                i = j
                continue
        i += 1

    seen = set()
    out = []
    for r in raw:
        if r["year"] in seen:
            continue
        seen.add(r["year"])
        out.append(r)
    out.sort(key=lambda r: r["year"])
    return [compute_metrics(r, total_premium, annual_premium, pay_years) for r in out]


def parse_encash_rows(text, anb_age, total_premium, annual_premium=None, pay_years=None):
    """解析 ENCASH（加护一生住院现金储蓄保险）退保价值 + 身故保障表。
    优先取「基本計劃補充説明」的完整逐年表（→ 全部按「第X年」显示）；
    若无该页，回退到「基本計劃説明摘要」的里程碑表（→ 30年前「第X年」、30年后「X岁」）。
    每行 7 个数值：[已缴保费, 保证退保A, 非保证红利B, 退保总额, 保证身故, 身故红利, 身故总额]。
    返回 compute_metrics 后的行列表。"""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    n = len(lines)
    start = None
    # 优先：完整逐年表（含「補充」的説明摘要）
    for i, ln in enumerate(lines):
        s = ln.replace("説", "說")
        if ("説明摘要" in s or "說明摘要" in s) and "補充" in s:
            start = i
            break
    # 回退：基本计划说明摘要（里程碑表）
    if start is None:
        for i, ln in enumerate(lines):
            s = ln.replace("説", "說")
            if ("説明摘要" in s or "說明摘要" in s) and "補充" not in s:
                start = i
                break
    if start is None:
        return []

    raw = []
    i = start
    while i < n:
        ln = lines[i]
        if "備註" in ln:
            break
        year = None
        m = ANB.search(ln)
        if m:
            age = int(m.group(1))
            year = (age - anb_age) if anb_age else None
        elif PLAIN_NUM.match(ln) and 1 <= int(ln) <= 101:
            year = int(ln)
        if year is not None:
            vals = []
            j = i + 1
            while j < n and len(vals) < 7 and (NUM.match(lines[j]) or VAL.match(lines[j])):
                vals.append(_to_num(lines[j]))
                j += 1
            if len(vals) >= 7:
                age = (anb_age + year) if anb_age else None
                raw.append({
                    "year": year,
                    "age": age,
                    "is_age_marker": bool(ANB.search(ln)),
                    "paid": vals[0],
                    "surr_guaranteed": vals[1],
                    "surr_bonus": vals[2],
                    "surrender": vals[3],
                    "death_guaranteed": vals[4],
                    "death_bonus": vals[5],
                    "death_total": vals[6],
                })
                i = j
                continue
        i += 1

    seen = set()
    out = []
    for r in raw:
        if r["year"] in seen:
            continue
        seen.add(r["year"])
        out.append(r)
    out.sort(key=lambda r: r["year"])
    return [compute_metrics(r, total_premium, annual_premium, pay_years) for r in out]


def parse_encash(text, basic):
    """解析 ENCASH（加护一生住院现金储蓄保险）。
    返回 basic_extra / rows（含 compute_metrics）/ product_advantage。"""
    rows = parse_encash_rows(text, basic.get("anb_age"), basic.get("total_premium"),
                             basic.get("annual_premium"), basic.get("pay_years"))
    # 从退保表推导总保费与供款期：缴清后已缴保费恒定的第一个年份即供款期，其值为总保费
    max_paid = max((r.get("paid") or 0) for r in rows) if rows else 0
    pay_years = next((r["year"] for r in rows if (r.get("paid") or 0) >= max_paid), None) if rows else None
    total_premium = max_paid or None
    # 用推导出的正确参数重算单复利
    rows = [compute_metrics(r, total_premium, basic.get("annual_premium"), pay_years) for r in rows]
    basic_extra = {
        "basic_amount": rows[0]["death_guaranteed"] if rows else basic.get("basic_amount"),
        "coverage": "終身",
        "annual_premium": basic.get("annual_premium"),
        "pay_years": pay_years,
        "total_premium": total_premium,
        "premium_type": "年缴",
    }
    product_advantage = [
        {
            "title": "住院现金与长期病况保障",
            "items": [
                "每日住院现金：每日赔偿当时保额的0.1%（即每日250美元），每次住院最多90日",
                "深切治疗住院现金：每日额外赔偿当时保额的0.1%（即额外250美元/日），双倍保障，最多90日",
                "长期病况保障：每次一笔过赔偿当时保额的5%（即12,500美元），合共最高保额的15%（37,500美元）",
                "覆盖严重脑退化症/帕金森病、严重糖尿病、长期住院超过90日，每项最多赔偿一次",
            ],
        },
        {
            "title": "身故保障，高杠杆",
            "items": [
                "身故赔偿：保证赔付100%保额（250,000美元）+累计终期红利面值，锁定高额身故杠杆",
                "首年杠杆高：第一年保费3,105美元即可撬动250,000美元身故保障（约80.5倍）",
                "与现有医疗重疾险不冲突，可叠加使用",
            ],
        },
        {
            "title": "有病赔钱，无病储蓄",
            "items": [
                "住院现金随心使用，住院期间每日现金流进账",
                "分红型保单，现金价值随时间持续增长，退保可取回连本带利",
                "退保价值逐年增长，101岁退保可领6,308,365美元（约4,416万人民币），传承下一代",
            ],
        },
    ]
    return {"basic_extra": basic_extra, "rows": rows, "product_advantage": product_advantage}


def parse_withdrawal(text, anb_age, product, total_premium=None, annual_premium=None, pay_years=None):
    """
    解析「現金價值提取下退保價值」表（即有提取安排的计划书）。
    TRST 每行 7 个数字：[已缴保费, 当年提取, 保证A, 归原红利B, 终期红利C, 退保总额, 名义金额]
    PRMESP 每行 8 个数字：[名义金额, x, 保证A, x, B, 退保总额, x, 当年提取]
    退保总额列与当年提取列位置因产品而异，按产品取位。
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    start_idx = None
    for idx, ln in enumerate(lines):
        if "現金價值提取下退保價值" in ln or "現金提取下退保價值" in ln:
            start_idx = idx
            break
    if start_idx is None:
        return None

    if product in ("TRST", "PACE"):
        n_cols = 7
    elif product == "PHOR":
        n_cols = 6
    else:
        n_cols = 8
    rows = []
    i = start_idx
    n = len(lines)
    while i < n:
        ln = lines[i]
        # 收集够行后，遇到下一个补充说明标题就停
        if rows and ("補充說明摘要" in ln and "退保價值" not in ln):
            break
        year = None
        m = ANB.search(ln)
        if m:
            # @ANB 101歲 → year = age - anb_age
            age = int(m.group(1))
            year = (age - anb_age) if anb_age else None
        elif PLAIN_NUM.match(ln) and 1 <= int(ln) <= 101:
            year = int(ln)
        if year is not None:
            vals = []
            j = i + 1
            while j < n and len(vals) < n_cols and (NUM.match(lines[j]) or VAL.match(lines[j])):
                vals.append(_to_num(lines[j]))
                j += 1
            if len(vals) >= n_cols:
                age = (anb_age + year) if anb_age else None
                if product in ("TRST", "PACE"):
                    withdrawal, surrender = vals[1], vals[5]
                elif product == "PHOR":
                    # PHOR 6列：保费, 名义金额, 红利B, 保证A, 退保总额, 当年提取
                    withdrawal, surrender = vals[5], vals[4]
                else:
                    withdrawal, surrender = vals[-1], vals[5]
                rows.append({"year": year,
                             "age": age,
                             "withdrawal": withdrawal,
                             "surrender": surrender})
                i = j
                continue
        i += 1

    if not rows:
        return None
    acc = 0
    for r in rows:
        acc += r["withdrawal"]
        r["accumulated"] = acc
        # 加单利复利（基于提取后剩余退保价值）
        if total_premium:
            r2 = compute_metrics(r, total_premium, annual_premium, pay_years)
            r["simple_pct"] = r2.get("simple_pct")
            r["compound_pct"] = r2.get("compound_pct")
    start_year = next((r["year"] for r in rows if r["withdrawal"] > 0), None)
    annual = next((r["withdrawal"] for r in rows if r["withdrawal"] > 0), 0)
    return {"rows": rows, "start_year": start_year, "annual_withdrawal": annual}


def parse_optimistic(text, anb_age):
    """
    解析「退保價值 – 不同投資回報下的說明」表（乐观/悲观情景）。
    每行 6 个数字：[缴费, 悲观红利, 保证A, 悲观总额, 乐观红利, 乐观总额(A)+(C)]
    返回 {year: {"age":, "optimistic":, "pessimistic":}} 仅保留里程碑年。
    """
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    start_idx = None
    for idx, ln in enumerate(lines):
        if "不同投資回報" in ln and "退保" in ln:
            start_idx = idx
            break
    if start_idx is None:
        # 退而求其次：找「樂觀情景」附近
        for idx, ln in enumerate(lines):
            if "樂觀情景" in ln:
                start_idx = idx - 5
                break
    if start_idx is None:
        return {}

    out = {}
    i = start_idx
    n = len(lines)
    while i < n:
        ln = lines[i]
        # 遇到身故赔偿表（8列）就停
        if "身故賠償" in ln and "不同投資回報" in ln:
            break
        
        # 处理纯数字年份
        year = None
        age = None
        if PLAIN_NUM.match(ln) and 1 <= int(ln) <= 101:
            year = int(ln)
            age = (anb_age + year) if anb_age else None
        else:
            # 处理@ANB XX歲格式
            m = ANB.search(ln)
            if m:
                age = int(m.group(1))
                year = (age - anb_age) if anb_age else None
        
        if year is not None:
            vals = []
            j = i + 1
            while j < n and len(vals) < 6 and (NUM.match(lines[j]) or VAL.match(lines[j])):
                vals.append(_to_num(lines[j]))
                j += 1
            if len(vals) == 6:
                out[year] = {
                    "age": age,
                    "optimistic": vals[5],     # 乐观总额 (A)+(C)
                    "pessimistic": vals[3],    # 悲观总额
                }
                i = j
                continue
        i += 1
    return out


def parse_pdf(pdf_path):
    text = extract_text(pdf_path)
    product = detect_product(text)
    basic = parse_basic(text)
    if product in ("CIM3", "CIE3", "CIP2"):
        if product == "CIM3":
            mod = parse_cim3(text, basic)
        elif product == "CIE3":
            mod = parse_cie3(text, basic)
        else:
            mod = parse_cip2(text, basic)
        basic.update(mod["basic_extra"])
        rows = mod["rows"]
        product_advantage = mod["product_advantage"]
        withdrawal = None
        optimistic = {}
    elif product == "ENCASH":
        enc = parse_encash(text, basic)
        basic.update(enc["basic_extra"])
        rows = enc["rows"]
        product_advantage = enc["product_advantage"]
        withdrawal = None
        optimistic = {}
    elif product.startswith("PLT"):
        # PLT定期寿险：提取保额
        m = re.search(r"投保時保額\s*\n?\s*(?:美元\s*)?([\d,]+\.?\d*)", text)
        if m:
            basic["basic_amount"] = _to_num(m.group(1))
        
        # 解析续保保费表
        premium_schedule = []
        m_sched = re.search(r"翌年[歲岁]\s*[(（]ANB[)）]", text)
        if m_sched:
            tail = text[m_sched.end():m_sched.end()+5000]
            # 找所有"X歲" + 数字的对（年龄行和保费行交替）
            # 格式：60歲\n1,533.00
            age_matches = list(re.finditer(r"(\d+)\s*[歲岁]\s*\n\s*([\d,]+\.\d{2})", tail))
            for am in age_matches:
                age = int(am.group(1))
                premium = _to_num(am.group(2))
                premium_schedule.append({"age": age, "annual_premium": premium})
        
        # 计算续保周期
        if len(premium_schedule) >= 2:
            first_age = premium_schedule[0]["age"]
            second_age = premium_schedule[1]["age"]
            renew_cycle = second_age - first_age
            basic["renew_cycle"] = renew_cycle
        else:
            basic["renew_cycle"] = 1
        
        # 正确计算总保费（各续保期累加）
        total_premium = 0
        for i, node in enumerate(premium_schedule):
            if i < len(premium_schedule) - 1:
                # 当前节点到下一个节点之间的年数
                years = premium_schedule[i+1]["age"] - node["age"]
                total_premium += node["annual_premium"] * years
            else:
                # 最后一个节点，到保障期结束
                # 保障年期 = pay_years - (current_age - entry_age)
                # 简单处理：最后一个节点只算1年（或按剩余保障期）
                entry_age = basic.get("anb_age", node["age"])
                coverage_years = basic.get("pay_years", 0)
                entry_to_node = node["age"] - entry_age
                remaining_years = coverage_years - entry_to_node
                if remaining_years > 0:
                    total_premium += node["annual_premium"] * remaining_years
                else:
                    total_premium += node["annual_premium"]  # 至少算1年
        
        basic["total_premium"] = total_premium if premium_schedule else basic.get("total_premium")
        basic["premium_schedule"] = premium_schedule
        
        rows = []
        product_advantage = None
        withdrawal = None
        optimistic = {}
    elif product == "CIT2":
        # CIT2守护健康危疾定期保II：和PLT定期寿险结构一样
        m = re.search(r"投保時保額\s*\n?\s*(?:美元\s*)?([\d,]+\.?\d*)", text)
        if m:
            basic["basic_amount"] = _to_num(m.group(1))
        
        # 解析保障年期（定期险是保障多久就供多久）
        m_years = re.search(r"保障年期\s+(.*?)\s*(\d+)\s*年", text, re.DOTALL)
        if m_years:
            basic["pay_years"] = int(m_years.group(2))
        elif not basic.get("pay_years"):
            # 兜底：从"保費供款年期"后面找
            m_years2 = re.search(r"保費供款年期\s+(.*?)\s*(\d+)\s*年", text, re.DOTALL)
            if m_years2:
                basic["pay_years"] = int(m_years2.group(2))
        
        # 解析续保保费表
        premium_schedule = []
        m_sched = re.search(r"翌年[歲岁]\s*[(（]ANB[)）]", text)
        if m_sched:
            tail = text[m_sched.end():m_sched.end()+5000]
            # 找所有"X歲" + 数字的对（年龄行和保费行交替）
            age_matches = list(re.finditer(r"(\d+)\s*[歲岁]\s*\n\s*([\d,]+\.?\d{2})", tail))
            for am in age_matches:
                age = int(am.group(1))
                premium = _to_num(am.group(2))
                premium_schedule.append({"age": age, "annual_premium": premium})
        
        # 计算续保周期
        if len(premium_schedule) >= 2:
            first_age = premium_schedule[0]["age"]
            second_age = premium_schedule[1]["age"]
            renew_cycle = second_age - first_age
            basic["renew_cycle"] = renew_cycle
        else:
            basic["renew_cycle"] = 5
        
        # 正确计算总保费（各续保期累加）
        total_premium = 0
        for i, node in enumerate(premium_schedule):
            if i < len(premium_schedule) - 1:
                years = premium_schedule[i+1]["age"] - node["age"]
                total_premium += node["annual_premium"] * years
            else:
                entry_age = basic.get("anb_age", node["age"])
                coverage_years = basic.get("pay_years", 0)
                entry_to_node = node["age"] - entry_age
                remaining_years = coverage_years - entry_to_node
                if remaining_years > 0:
                    total_premium += node["annual_premium"] * remaining_years
                else:
                    total_premium += node["annual_premium"]
        
        basic["total_premium"] = total_premium if premium_schedule else basic.get("total_premium")
        basic["premium_schedule"] = premium_schedule
        
        rows = []
        product_advantage = None
        withdrawal = None
        optimistic = {}
    else:
        rows = parse_surrender_table(text, basic.get("anb_age"),
                                     basic.get("total_premium"), product=product)
        # PLP2/LFRE2：合并身故赔偿表数据
        if product == "PLP2":
            death_map = parse_plp2_death(text)
            for r in rows:
                if r.get("year") in death_map:
                    r["death_total"] = death_map[r["year"]]
        elif product == "LFRE2":
            death_map = parse_lfre2_death(text)
            for r in rows:
                if r.get("year") in death_map:
                    r["death_total"] = death_map[r["year"]]
        rows = [compute_metrics(r, basic.get("total_premium"),
                                basic.get("annual_premium"), basic.get("pay_years")) for r in rows]
        withdrawal = parse_withdrawal(text, basic.get("anb_age"), product,
                                      basic.get("total_premium"), basic.get("annual_premium"),
                                      basic.get("pay_years"))
        optimistic = parse_optimistic(text, basic.get("anb_age"))
        product_advantage = None
    return {
        "product_code": product,
        "product_name": PRODUCT_NAMES.get(product, product),
        "basic": basic,
        "rows": rows,
        "optimistic": optimistic,
        "withdrawal": withdrawal,
        "product_advantage": product_advantage,
        "text_length": len(text),
    }


# ============ 家庭保单规划：快速基本信息解析 ============

FAMILY_PRODUCT_NAMES = {
    "PRMESP": "PRMESP 世誉财富保险计划",
    "TRST": "TRST 信守明天多元货币",
    "CIM3": "CIM3 诚保一生危疾",
    "CIE3": "CIE3 守护健康危疾加护保III",
    "CIP2": "CIP2 危疾首护保险计划II",
    "ENCASH": "ENCASH 加护一生住院现金",
    "PMP": "PMP 医疗加倍保",
    "PLT20": "PLT20 守护家人定期寿险",
    "PLT10": "PLT10 守护家人定期寿险（10年续保）",
    "PLT5": "PLT5 守护家人定期寿险（5年续保）",
    "PLT1": "PLT1 守护家人定期寿险（1年续保）",
    "PLT": "PLT 定期寿险",
    "CIT2": "CIT2 守护健康危疾定期保II",
    "CIT": "CIT 定期危疾",
    "CPER": "CPER 特选危疾治疗保",
    "MCVIP": "MCVIP",
    "VIP": "VIP 自愿医保尚宾",
    "LFRE2": "LFRE2 自主未来II",
    "PLP2": "PLP2 美好人寿II",
    "PHOR": "PHOR 诚誉财富保险计划",
    "PACE": "PACE 骏誉财富保险计划",
}

FAMILY_PRODUCT_CATEGORY = {
    "PRMESP": "储蓄/退休",
    "TRST": "储蓄/退休",
    "PHOR": "储蓄/退休",
    "PACE": "储蓄/退休",
    "CIM3": "危疾",
    "CIE3": "危疾",
    "CIP2": "危疾",
    "CIT2": "定期危疾",
    "CIT": "定期危疾",
    "ENCASH": "住院现金",
    "PLT20": "定期寿险",
    "PLT10": "定期寿险",
    "PLT5": "定期寿险",
    "PLT1": "定期寿险",
    "PLT": "定期寿险",
    "CPER": "医疗",
    "MCVIP": "医疗",
    "VIP": "医疗",
    "PMP": "医疗",
    "LFRE2": "储蓄/寿险",
    "PLP2": "储蓄/寿险",
}


def detect_product_family(text):
    """家庭规划用：更宽松的产品代码识别。"""
    for code, _ in FAMILY_PRODUCT_NAMES.items():
        if code in text:
            return code
    # 关键词兜底
    if "世譽財富" in text or "世誉财富" in text: return "PRMESP"
    if "信守明天" in text or "信信守" in text: return "TRST"
    if "誠保一生" in text or "诚保一生" in text: return "CIM3"
    if "守護健康" in text or "守护健康" in text:
        if "加護保" in text or "加护保" in text: return "CIE3"
        return "CIT2"
    if "危疾首護" in text or "危疾首护" in text: return "CIP2"
    if "住院現金" in text or "住院现金" in text: return "ENCASH"
    if "守護家人" in text or "守护家人" in text:
        # 按续保期区分 PLT1 / PLT5 / PLT10 / PLT20
        if "(PLT1)" in text or "（PLT1）" in text: return "PLT1"
        if "(PLT5)" in text or "（PLT5）" in text: return "PLT5"
        if "(PLT10)" in text or "（PLT10）" in text: return "PLT10"
        if "(PLT20)" in text or "（PLT20）" in text: return "PLT20"
        return "PLT20"  # 默认
    # 先按括号里的产品代码区分 VIP / MCVIP
    if "(PMP)" in text or "醫療加倍保" in text or "医疗加倍保" in text:
        return "PMP"
    if "(MCVIP)" in text or "高端醫療自由行" in text or "高端医疗自由行" in text:
        return "MCVIP"
    if "(VIP)" in text or "自願醫保尚賓" in text or "自愿医保尚宾" in text:
        return "VIP"
    if "尚賓" in text or "尚宾" in text: return "MCVIP"  # 关键词兜底
    if "自主未來" in text or "自主未来" in text: return "LFRE2"
    if "美好人壽" in text or "美好人寿" in text: return "PLP2"
    if "誠譽財富" in text or "诚誉财富" in text: return "PHOR"
    if "駿譽財富" in text or "骏誉财富" in text: return "PACE"
    return "UNKNOWN"


def parse_basic_from_text(text):
    """从文本中提取保单基本信息（PDF和图片OCR共用）。
    返回 {product_code, product_name, category, age, gender, currency,
           sum_assured, annual_premium, pay_years, total_premium, ...}
    """
    product = detect_product_family(text)
    product_name = FAMILY_PRODUCT_NAMES.get(product, product)

    # MCVIP 区分计划一/计划二
    if product == "MCVIP":
        if "計劃一" in text or "计划一" in text:
            product_name += "（计划一）"
        elif "計劃二" in text or "计划二" in text:
            product_name += "（计划二）"

    info = {
        "product_code": product,
        "product_name": product_name,
        "category": FAMILY_PRODUCT_CATEGORY.get(product, "其他"),
        "age": None, "gender": None, "currency": None,
        "sum_assured": None, "annual_premium": None,
        "pay_years": None, "pay_term": None, "total_premium": None,
        "last_surrender_value": None,
        "rider_name": None,
        "rider_premium": None,
    }

    # 年龄（繁简都支持）
    m = re.search(r"翌年[歲岁][（(]ANB[)）][:：]\s*(\d+)\s*[歲岁]", text)
    if not m:
        m = re.search(r"ANB\s*[:：]?\s*(\d+)\s*[歲岁]", text)
    info["age"] = int(m.group(1)) if m else None

    # 性别（繁简都支持）
    m = re.search(r"性[別别][：:]\s*(男|女)", text)
    info["gender"] = m.group(1) if m else None

    # 货币（繁简都支持）
    m = re.search(r"保[單单][貨货][幣币]\s*[:：]\s*([A-Za-z\u4e00-\u9fff]+)", text)
    if m:
        cur = m.group(1)
        info["currency"] = "HKD" if "港" in cur else ("USD" if "美" in cur else cur)

    # 年缴保费（最可靠：找"投保時每年總保費"或"投保時整付總保費"后面的数字）
    annual_premium = None
    for pat in [
        r"投保[時时]每年[總总][保保][費费]\s*\n?([\d,]+\.?\d*)",
        r"投保[時时]整付[總总][保保][費费]\s*\n?([\d,]+\.?\d*)",
        r"投保[時时]每年[保保][費费]\s*\n?([\d,]+\.?\d*)",
        r"投保[時时]整付[保保][費费]\s*\n?([\d,]+\.?\d*)",
    ]:
        m = re.search(pat, text)
        if m:
            annual_premium = _to_num(m.group(1))
            break
    info["annual_premium"] = annual_premium

    # 纯储蓄险产品（保额不适用）：PRMESP
    # TRST/PHOR/PACE等储蓄险也从PDF提取总保费作为"总保额"显示
    savings_no_sum = {"PRMESP"}
    # 定期寿险/终身寿险产品（PLT系列 + PLP2）：表格布局特殊，需单独提取保额
    term_life_products = {"PLT1", "PLT5", "PLT10", "PLT20", "PLT", "PLP2"}
    # 重疾类产品（CIM3/CIE3/CIP2等）：产品代码后多行数字（年缴/保障年期/保额/供款期）
    critical_illness_products = {"CIM3", "CIE3", "CIP2", "CIT2", "CIT", "CPER"}
    if product in savings_no_sum:
        sum_assured = None
    elif product in term_life_products or product in critical_illness_products:
        # 定期寿险/重疾类：从产品代码括号往后找所有数字，排除年缴和<100的，剩下的就是保额
        sum_assured = None
        all_products = list(term_life_products | critical_illness_products)
        m = re.search(r"\(\s*(" + "|".join(all_products) + r")\s*\)", text)
        if m:
            # 产品名后面的500字符
            tail = text[m.end():m.end()+500]
            numbers = re.findall(r"([\d,]+\.?\d*)", tail)
            for num_str in numbers:
                val = _to_num(num_str)
                if not val:
                    continue
                # 排除年缴
                if annual_premium and abs(val - annual_premium) < 1:
                    continue
                # 排除供款期/保障年期（< 100的数字）
                if val < 100:
                    continue
                # 剩下的大数字就是保额
                if val > 10000:
                    sum_assured = val
                    break
    else:
        # 保额：找产品名后面的数字，排除年缴和供款期
        sum_assured = None
        m = re.search(r"\(((" + "|".join(FAMILY_PRODUCT_NAMES.keys()) + r"))\)\s*\n?([\d,]+\.?\d*)", text)
        if m:
            info["product_code"] = m.group(1)
            info["product_name"] = FAMILY_PRODUCT_NAMES.get(m.group(1), m.group(1))
            info["category"] = FAMILY_PRODUCT_CATEGORY.get(m.group(1), "其他")
            # 产品名后面的所有数字
            tail = text[m.end():m.end()+500]
            numbers = re.findall(r"([\d,]+\.?\d*)", tail)
            for num_str in numbers:
                val = _to_num(num_str)
                if not val:
                    continue
                # 排除年缴（如果年缴找到了，跳过和它一样的数字）
                if annual_premium and abs(val - annual_premium) < 1:
                    continue
                # 排除供款期（X 年）
                if val < 100:  # 供款期一般 < 100
                    continue
                # 剩下的大数字就是保额
                if val > 50000:
                    sum_assured = val
                    break
    info["sum_assured"] = sum_assured

    # 前十年赠送保额（升级保障 CIMEx / 首十年额外保障）
    bonus_10y = None
    for pat in [
        r"[附附]有升[級级]保障\s*\([A-Za-z0-9]+\)\s*\n?\s*([\d,]+\.?\d*)",
        r"升[級级]保障\s*\n?\s*([\d,]+\.?\d*)",
        r"(?:首十年|首10年|前十年)[^\d\n]{0,24}\s*\n?\s*([\d,]+\.?\d*)",
        # CIM3/CIMEx 格式：(CIME3)\n50,000 / (CICE3)\n75,000
        r"\(CI[MC]E\d*\)\s*\n?\s*([\d,]+\.?\d*)",
        # CIP2 首10年额外保障
        r"首10年[^\d\n]{0,30}\n?\s*([\d,]+\.?\d*)",
    ]:
        m = re.search(pat, text)
        if m:
            val = _to_num(m.group(1))
            if val and val >= 1000:
                bonus_10y = val
                break
    info["first_10y_bonus"] = bonus_10y

    # 供款年期（繁简都支持）
    # PDF文本里"保費供款年期"是表头，值在后面几行，不是同一行
    pay_years = None
    pay_term = None
    for line_pat in [r"保[費费]供款年期", r"保费供款年期"]:
        m = re.search(line_pat, text)
        if m:
            # 从表头位置往后取500字符
            tail = text[m.end():m.end()+500]
            # 只找"数字 年"的模式（避免把年缴保费的数字当供款期）
            m2 = re.search(r"(\d+)\s*[年年]", tail)
            if m2:
                val = int(m2.group(1))
                if val <= 50:  # 供款期一般 <= 50年
                    pay_years = val
                    break
    # 兜底：一次整付产品（PRMESP）供款期是1年
    if not pay_years and product in ("PRMESP",):
        pay_years = 1
    # 终身缴交的医疗险（医疗类且未取到数字缴费年期）：供款期按终身，不计算总保费
    medical_codes = {"PMP", "MCVIP", "VIP", "CPER"}
    if not pay_years and product in medical_codes:
        pay_term = "终身"
    info["pay_years"] = pay_years
    info["pay_term"] = pay_term

    # 总保费：仅数字供款期用 年保费×供款年数；终身缴交/无固定供款期留空，不退化成年保费
    if info["annual_premium"] and isinstance(pay_years, int):
        info["total_premium"] = info["annual_premium"] * pay_years
    else:
        info["total_premium"] = None

    # 储蓄险产品：提取最后剩余价值
    savings_products = {"TRST", "PRMESP", "LFRE2", "PLP2"}
    if product in savings_products:
        try:
            rows = parse_surrender_table(text, info.get("age"), info.get("total_premium"), product=product)
            if rows:
                last_row = rows[-1]
                info["last_surrender_value"] = last_row.get("guaranteed_cv") or last_row.get("total_cv")
        except Exception:
            pass

    # 检测附加保障 VFPR（灵活自主医保）——放在最后，避免被保额解析覆盖
    if "VFPR" in text or "靈活自主醫保" in text or "灵活自主医保" in text:
        info["rider_name"] = "VFPR灵活自主医保"
        m = re.search(r"\(VFPR\)[^\d]*?([\d,]+\.?\d*)", text, re.S)
        if m:
            info["rider_premium"] = _to_num(m.group(1))
        info["product_name"] += "（含VFPR）"

    return info


def parse_policy_detail_from_text(text):
    """解析保诚系统保单详情页HTML提取的文本。
    和计划书PDF格式完全不同，单独一套正则。
    """
    info = {
        "product_code": None,
        "product_name": None,
        "category": "其他",
        "policy_number": None,
        "billing_number": None,
        "insured_name": None,
        "age": None,
        "gender": None,
        "currency": None,
        "sum_assured": None,
        "annual_premium": None,
        "pay_years": None,
        "pay_term": None,
        "total_premium": None,
        "remaining_pay_years": None,
        "first_10y_bonus": None,
        "status": None,
    }

    # 保单号码
    m = re.search(r"保[單单][號号][碼码]\s*\n?\s*(\d{6,})", text)
    if m:
        info["policy_number"] = m.group(1)

    # 缴费编号（去掉BN前缀）
    m = re.search(r"繳費編號\s*\n?\s*BN(\d+)", text)
    if m:
        info["billing_number"] = m.group(1)

    # 计划名称 + 产品代码：「誠保一生」危疾保險計劃 (CIM3)
    # 格式：計劃名稱 → 换行 → 产品名 → 换行 → (CIM3)
    m = re.search(r"計劃名[稱称][\s\S]{0,20}?([^\n(]+?)\s*\n?\s*&nbsp;\s*\(([A-Z0-9]+)\)", text)
    if m:
        info["product_name"] = m.group(1).strip()
        info["product_code"] = m.group(2)
    else:
        # 兜底：从括号里找产品代码
        m = re.search(r"\(([A-Z]{2,}\d*)\)", text)
        if m:
            info["product_code"] = m.group(1)
            info["product_name"] = FAMILY_PRODUCT_NAMES.get(info["product_code"], info["product_code"])

    # 产品分类
    if not info["product_code"]:
        # 兜底关键词：未提取到括号代码时按计划名识别
        if "醫療加倍保" in text or "医疗加倍保" in text:
            info["product_code"] = "PMP"
            info["product_name"] = FAMILY_PRODUCT_NAMES.get("PMP", "PMP 医疗加倍保")
    if info["product_code"]:
        info["category"] = FAMILY_PRODUCT_CATEGORY.get(info["product_code"], "其他")

    # 受保人姓名：受保人 → YAN LAURA沈如清 → 提取中文名字
    # 先把 &nbsp; 替换成空格
    clean_text = text.replace("&nbsp;", " ")
    m = re.search(r"受保人\s*\n?\s*[A-Z ]*([\u4e00-\u9fff]{2,5})\s*\(", clean_text)
    if m:
        info["insured_name"] = m.group(1)

    # 性别
    m = re.search(r"性[別别]\s*\n?\s*(F|M)", text)
    if m:
        info["gender"] = "女" if m.group(1) == "F" else "男"

    # 出生日期
    birth_year, birth_month, birth_day = None, None, None
    m = re.search(r"出生日期\s*\n?\s*(\d{2})/(\d{2})/(\d{4})", text)
    if m:
        birth_day, birth_month, birth_year = int(m.group(1)), int(m.group(2)), int(m.group(3))

    # 保单生效日期（批核日期）
    eff_year, eff_month, eff_day = None, None, None
    m = re.search(r"保[單单]生效日期\s*\n?\s*(\d{2})/(\d{2})/(\d{4})", text)
    if m:
        eff_day, eff_month, eff_year = int(m.group(1)), int(m.group(2)), int(m.group(3))

    # 用出生日期和保单生效日期算年龄
    if birth_year and eff_year:
        age = eff_year - birth_year
        if eff_month < birth_month or (eff_month == birth_month and eff_day < birth_day):
            age -= 1
        info["age"] = age
    elif birth_year:
        # 兜底：用当前日期算
        ref_year, ref_month = 2026, 11
        age = ref_year - birth_year
        if birth_month > ref_month:
            age -= 1
        info["age"] = age

    # 保单货币
    m = re.search(r"保[單单][貨货][幣币]\s*\n?\s*([A-Z]{3})", text)
    if m:
        info["currency"] = m.group(1)

    # 保单状态
    m = re.search(r"狀[況况]\s*\n?\s*([A-Z ]+)", text)
    if m:
        info["status"] = m.group(1).strip()

    # 保单生效日期 → 算已缴年数
    effective_year = None
    m = re.search(r"保[單单]生效日期\s*\n?\s*\d{2}/\d{2}/(\d{4})", text)
    if m:
        effective_year = int(m.group(1))

    # 缴费模式：ANNUAL / SINGLE
    pay_mode = None
    m = re.search(r"繳費模式\s*\n?\s*(\w+)", text)
    if m:
        pay_mode = m.group(1)

    # ===== 在「保障資料」表格里提取 =====
    # 先定位「保障資料」后面的区域
    benefit_start = text.find("保障資料")
    if benefit_start >= 0:
        benefit_text = text[benefit_start:benefit_start+3000]
    else:
        benefit_text = text

    # 供款期：在保障资料表格数据行里，状态(INFORCE/FULLY PAID UP) → 首期保费日 → 供款期
    pay_years = None
    m = re.search(r"(INFORCE|FULLY PAID UP)\s*\n?\s*\d{2}/\d{2}/\d{4}\s*\n?\s*(\d+)", benefit_text)
    if m:
        val = int(m.group(2))
        if val <= 50:
            pay_years = val
    # 兜底：一次缴清
    if not pay_years and pay_mode == "SINGLE":
        pay_years = 1
    # 终身缴交的医疗险：未取到数字供款期，按终身处理
    medical_codes = {"PMP", "MCVIP", "VIP", "CPER"}
    pay_term = "终身" if (not pay_years and info["product_code"] in medical_codes) else None
    info["pay_years"] = pay_years
    info["pay_term"] = pay_term

    # ===== 在「遞增保障權益選擇詳情」表格里提取 =====
    # 这个表格里才有保障额和每期保费
    detail_start = text.find("遞增保障權益選擇詳情")
    if detail_start >= 0:
        detail_text = text[detail_start:detail_start+1000]
    else:
        detail_text = benefit_text

    # 找所有数字（带逗号的金额格式）
    numbers = re.findall(r"([\d,]+\.\d{2})", detail_text)
    numbers = [_to_num(n) for n in numbers if _to_num(n) and _to_num(n) > 0]

    # 保障额：第一个大数字（>10000）
    sum_assured = None
    for n in numbers:
        if n > 10000:
            sum_assured = n
            break
    info["sum_assured"] = sum_assured

    # 年缴保费：第二个数字（如果第一个是保障额，第二个就是保费）
    annual_premium = None
    if len(numbers) >= 2:
        # 跳过保障额那个大数字，找后面的
        for i, n in enumerate(numbers):
            if sum_assured and n == sum_assured:
                # 找下一个数字
                if i + 1 < len(numbers):
                    annual_premium = numbers[i + 1]
                    break
                else:
                    annual_premium = numbers[i]
                    break
    # 兜底：在「保費資料」区域找产品名后面的保费
    if not annual_premium and info["product_code"]:
        premium_start = text.find("保費資料")
        if premium_start >= 0:
            premium_text = text[premium_start:premium_start+1000]
            pattern = r"\(" + re.escape(info["product_code"]) + r"\)[^\d]*?([\d,]+\.\d{2})"
            m = re.search(pattern, premium_text)
            if m:
                annual_premium = _to_num(m.group(1))
    info["annual_premium"] = annual_premium

    # 总保费
    if annual_premium and isinstance(pay_years, int):
        info["total_premium"] = annual_premium * pay_years
    else:
        info["total_premium"] = None

    # 剩余缴费年期 = 总供款期 - 已缴年数（生效当天计 1 年，每过一个周年 +1）
    if pay_years and eff_year:
        elapsed_years = _elapsed_since_eff(eff_year, eff_month, eff_day)
        if elapsed_years is not None:
            info["remaining_pay_years"] = max(0, pay_years - elapsed_years)

    return info


def parse_basic_fast(pdf_path):
    """快速解析：只提取前6页基本信息，不解析退保价值表。
    返回 {product_code, product_name, category, age, gender, currency,
           sum_assured, annual_premium, pay_years, total_premium, filename}
    """
    doc = fitz.open(pdf_path)
    # 扫描前6页（封面+空白页+摘要页通常在前6页内）
    n = min(6, len(doc))
    text = "\n".join(doc[i].get_text() for i in range(n))
    doc.close()
    return parse_basic_from_text(text)


if __name__ == "__main__":
    import sys, json
    r = parse_pdf(sys.argv[1])
    print(json.dumps(r, ensure_ascii=False, indent=2))
