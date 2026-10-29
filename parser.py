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
import fitz  # PyMuPDF

PRODUCT_NAMES = {
    "PRMESP": "世誉财富保险计划",
    "TRST": "保诚信守明天多元货币保险计划",
    "CIM3": "诚保一生危疾保险计划",
}

NUM = re.compile(r"^\d{1,3}(,\d{3})+(\.\d+)?$")  # 25,000 / 1,240,550
VAL = re.compile(r"^\d+(\.\d+)?$")                # 任意数字（含 0 / 25,000 由 NUM 另配）
PLAIN_NUM = re.compile(r"^\d+$")                  # 纯整数年份
ANB = re.compile(r"@?ANB\s*(\d+)\s*歲")


def _to_num(s):
    return float(s.replace(",", ""))


def extract_text(pdf_path):
    doc = fitz.open(pdf_path)
    pages = [p.get_text() for p in doc]
    doc.close()
    return "\n".join(pages)


def detect_product(text):
    if "PRMESP" in text or "世譽財富" in text:
        return "PRMESP"
    if "TRST" in text or "信守明天" in text:
        return "TRST"
    if "CIM3" in text or "誠保一生" in text or "誠保一生" in text or "危疾保險計劃" in text:
        return "CIM3"
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
        m3 = re.search(r"保費供款年期\s*(\d+)\s*年", text)
        annual = _to_num(m2.group(1)) if m2 else None
        years = int(m3.group(1)) if m3 else None
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
    is_trst = product == "TRST"
    n_cols = 5 if is_trst else 4

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
                                       col_paid=0, col_g=col_g, col_e=3)
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
                                   col_paid=0, col_g=col_g, col_e=3)


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

    product_advantage = [
        {
            "title": "保障超全，持续一生",
            "items": [
                "生·给宝宝的专属保障：严重湿疹、荨麻疹、肺炎住院津贴等均可获得补贴；终身保障疾病涵盖127种（49项早期严重疾病+56项严重疾病+15种儿童重疾+7种儿童次级疾病），确诊严重疾病获赔100%保额+累计红利，并豁免后续所有保费",
                "病·深切治疗ICU保障：连续入住ICU三日或以上获赔10%保额；严重疾病延伸保障最高累计900%，每笔赔100%保额可多次赔付——癌症2次、心脏病/中风2次、脑退化症/帕金森1次、主要器官相关2次、其他严重疾病2次",
                "老·给老人的专属保障：确诊严重脑退化症或帕金森病，提供终身领取的年金，每年6,000美元直至身故",
                "死·身故额外保障：严重疾病理赔后、86岁之前身故，额外获赔100,000美元人寿延伸保障",
            ],
        },
        {
            "title": "重疾保额逐年增长，抗通胀",
            "items": [
                "首10年附送升级保障50,000美元，总保障150,000美元（首10年150%保障）",
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


def parse_withdrawal(text, anb_age, product):
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

    n_cols = 7 if product == "TRST" else 8
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
                if product == "TRST":
                    withdrawal, surrender = vals[1], vals[5]
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
        if PLAIN_NUM.match(ln) and 1 <= int(ln) <= 101:
            vals = []
            j = i + 1
            while j < n and len(vals) < 6 and (NUM.match(lines[j]) or VAL.match(lines[j])):
                vals.append(_to_num(lines[j]))
                j += 1
            if len(vals) == 6:
                year = int(ln)
                out[year] = {
                    "age": (anb_age + year) if anb_age else None,
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
    if product == "CIM3":
        cim3 = parse_cim3(text, basic)
        basic.update(cim3["basic_extra"])
        rows = cim3["rows"]
        product_advantage = cim3["product_advantage"]
        withdrawal = None
        optimistic = {}
    else:
        rows = parse_surrender_table(text, basic.get("anb_age"),
                                     basic.get("total_premium"), product=product)
        rows = [compute_metrics(r, basic.get("total_premium"),
                                basic.get("annual_premium"), basic.get("pay_years")) for r in rows]
        withdrawal = parse_withdrawal(text, basic.get("anb_age"), product)
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


if __name__ == "__main__":
    import sys, json
    r = parse_pdf(sys.argv[1])
    print(json.dumps(r, ensure_ascii=False, indent=2))
