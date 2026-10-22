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


def compute_metrics(row, total_premium):
    out = dict(row)
    yr = row.get("year") or 1
    if total_premium:
        out["multiple"] = round(row["surrender"] / total_premium, 1)
        out["simple_pct"] = round((row["surrender"] - total_premium) / total_premium / yr * 100, 1)
        if row["surrender"] > 0 and yr > 0:
            out["compound_pct"] = round(((row["surrender"] / total_premium) ** (1.0 / yr) - 1) * 100, 1)
        else:
            out["compound_pct"] = None
    return out


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
    rows = parse_surrender_table(text, basic.get("anb_age"),
                                 basic.get("total_premium"), product=product)
    rows = [compute_metrics(r, basic.get("total_premium")) for r in rows]
    withdrawal = parse_withdrawal(text, basic.get("anb_age"), product)
    optimistic = parse_optimistic(text, basic.get("anb_age"))
    return {
        "product_code": product,
        "product_name": PRODUCT_NAMES.get(product, product),
        "basic": basic,
        "rows": rows,
        "optimistic": optimistic,
        "withdrawal": withdrawal,
        "text_length": len(text),
    }


if __name__ == "__main__":
    import sys, json
    r = parse_pdf(sys.argv[1])
    print(json.dumps(r, ensure_ascii=False, indent=2))
