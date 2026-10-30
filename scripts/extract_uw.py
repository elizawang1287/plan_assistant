# -*- coding: utf-8 -*-
"""完整提取：sections + 矩阵，生成 uw_data.json。"""
import fitz, json, re, os

SRC = r"C:\Users\user\Documents\working\PowerPoint Presentation.pdf"
OUT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT1 = os.path.join(OUT_ROOT, "uw_data.json")
OUT2 = os.path.join(OUT_ROOT, "static", "uw_data.json")
NAV = "回到疾病索引"
TITLE_RE = re.compile(r'^([^\n（）()]+)[（(]([^（）()]*[A-Za-z][^（）()]*)[）)]$')
SECTION_KEYS = ["簡介", "病徵", "治療", "考慮因素", "基本核保要求"]
FOOT = ("回到疾病索引", "版本", "本指引", "以上為", "僅限內部", "Confidential")

doc = fitz.open(SRC)
PAGETXT = [doc[i].get_text() for i in range(doc.page_count)]
PAGEW = [doc[i].get_text("words") for i in range(doc.page_count)]


def norm(s):
    s = (s or "").strip()
    return s.replace("／", "/").replace("卵", "卵").replace("　", " ")


def norm_ns(s):
    return norm(s).replace(" ", "")


def clean_en(s):
    s = (s or "").strip()
    s = s.strip("（）()")
    return re.sub(r'\s+', ' ', s).strip()


TITLE_ALIAS = {
    "宮頸上皮內瘤樣病變／宮頸原位癌／人類乳頭狀瘤病毒感染": "宮頸上皮內瘤樣病變",
    "前列腺特異抗原上升": "前列腺特異性抗原上升",
    "柏金遜症家族病史": "帕金森病家族病史",
}


def parse_overview():
    txt = PAGETXT[6]
    cat, cur, pending_seq, whitelist_ns = {}, None, False, set()
    for raw in txt.splitlines():
        l = raw.strip()
        if "疾病總覽" in l:
            continue
        if "回到疾病索引" in l or "核保指引" in l:
            break
        if l in ("Confidential",) or re.match(r'^[0-9]+$', l):
            continue
        if re.match(r'^[0-9]+\.\s*.+', l):
            dn = re.sub(r'^[0-9]+\.\s*', '', l).strip()
            if cur is not None and dn:
                cat.setdefault(cur, []).append(dn); whitelist_ns.add(norm_ns(dn))
            pending_seq = False
            continue
        if re.match(r'^[0-9]+\.?\s*$', l):
            pending_seq = True
            continue
        if pending_seq:
            pending_seq = False
            if cur is not None and l:
                cat.setdefault(cur, []).append(l); whitelist_ns.add(norm_ns(l))
            continue
        if l and "（" not in l and "(" not in l:
            cur = l
    if "乳房疾病疾病" in cat:
        cat["乳房疾病"] = cat.pop("乳房疾病疾病")
    return cat, whitelist_ns


def collect_part1(whitelist_ns):
    alias_ns = {norm_ns(k): v for k, v in TITLE_ALIAS.items()}
    entries, cur = [], None
    for i, txt in enumerate(PAGETXT):
        if i in (6, 7) or NAV not in txt:  # 仅跳过疾病总览页(PAGE7/8)
            continue
        lines = [l.strip() for l in txt.splitlines() if l.strip()]
        hit = None
        for l in lines:
            if l.startswith(("▪", "•", "-", "・")):
                continue
            mm = TITLE_RE.match(l)
            if not mm:
                continue
            cn_ns = norm_ns(mm.group(1))
            if cn_ns in whitelist_ns:
                hit = (mm.group(1).strip(), mm.group(1).strip(), mm.group(2).strip()); break
            if cn_ns in alias_ns:
                hit = (alias_ns[cn_ns], mm.group(1).strip(), mm.group(2).strip()); break
        if hit is None:
            for k, l in enumerate(lines):
                if l.startswith(("▪", "•", "-", "・")) or "（" in l or "(" in l:
                    continue
                if l.endswith(("：", ":")):
                    continue
                nn = norm_ns(l)
                matched = None
                if nn in whitelist_ns:
                    matched = nn
                elif nn in alias_ns:
                    matched = alias_ns[nn]
                else:
                    for wl in sorted(whitelist_ns, key=len, reverse=True):
                        if nn.startswith(wl):
                            matched = wl; break
                if matched is None:
                    continue
                en = ""
                for k2 in range(k + 1, min(k + 4, len(lines))):
                    l2 = lines[k2]; m2 = TITLE_RE.match(l2)
                    if m2:
                        en = m2.group(2).strip(); break
                    if re.search(r'[A-Za-z]', l2) and not re.search(r'[\u4e00-\u9fff]', l2):
                        en = l2.strip(); break
                if not en:
                    continue
                hit = (matched, l, en); break
        if hit is None:
            continue
        std, disp, en = hit
        if cur is None:
            cur = {"name_cn": std, "disp": disp, "name_en": en, "start": i + 1, "end": i + 1}
        elif norm_ns(cur["name_cn"]) == norm_ns(std):
            cur["end"] = i + 1
            if not cur["name_en"]:
                cur["name_en"] = en
        else:
            entries.append(cur); cur = {"name_cn": std, "disp": disp, "name_en": en, "start": i + 1, "end": i + 1}
    if cur:
        entries.append(cur)
    seen, out = set(), []
    for e in entries:
        if norm_ns(e["name_cn"]) in seen:
            continue
        seen.add(norm_ns(e["name_cn"])); out.append(e)
    return out


def collect_part2():
    out = []
    skip = re.compile(r'(Confidential|回到疾病索引|健康保障核保考慮因素|本指引|^version|^\d+$|^[*＊*])')
    for i in range(193, min(213, doc.page_count)):
        txt = PAGETXT[i]
        if "健康保障核保" not in txt:
            continue
        lines = [l.strip() for l in txt.splitlines()
                 if l.strip() and not skip.search(l.strip())]
        name_cn, name_en = "", ""
        k = 0
        while k < len(lines):
            if re.search(r'[A-Za-z]', lines[k]) and not re.search(r'[\u4e00-\u9fff]', lines[k]):
                en = []
                j = k
                while j < len(lines) and re.search(r'[A-Za-z]', lines[j]) and not re.search(r'[\u4e00-\u9fff]', lines[j]):
                    en.append(lines[j])
                    j += 1
                name_en = " ".join(en).strip()
                p = k - 1
                cn = []
                while p >= 0 and re.search(r'[\u4e00-\u9fff]', lines[p]):
                    if DEC_RE.search(lines[p]) or re.search(r'[：:]', lines[p]):
                        break
                    cn.insert(0, lines[p])
                    p -= 1
                name_cn = "".join(cn).strip()
                if i + 1 == 204:  # 该页疾病名上方为矩阵决定行（无冒号），特例
                    name_cn = "聽力喪失"
                break
            k += 1
        out.append({"name_cn": name_cn, "name_en": name_en, "page": i + 1, "part": "II"})
    return out


def is_foot_line(l, whitelist_ns):
    if not l:
        return True
    if any(f in l for f in FOOT):
        return True
    if re.match(r'^\d+$', l) or l.startswith("version"):
        return True
    if "核保指引" in l:
        return True
    if l in ("人壽", "早期危疾", "危疾", "醫療", "癌症", "早期危疾 / 危疾", "早期危疾/危疾"):
        return True
    mm = TITLE_RE.match(l)
    if mm and norm_ns(mm.group(1)) in whitelist_ns:
        return True
    return False


def extract_sections(entry, part, whitelist_ns):
    if part == "II":
        return extract_part2_sections(entry, whitelist_ns)
    pages = list(range(entry["start"] - 1, entry["end"]))
    if norm_ns(entry["name_cn"]) == norm_ns("2019冠狀病毒病"):
        pages += [70, 71]
    sec, cur, buf = {}, None, []
    for pi in sorted(set(pages)):
        if pi < 0 or pi >= doc.page_count:
            continue
        for l in PAGETXT[pi].splitlines():
            l = l.strip()
            if is_foot_line(l, whitelist_ns):
                continue
            if l in SECTION_KEYS:
                if cur is not None:
                    sec[cur] = "\n".join(buf).strip()
                cur, buf = l, []
                continue
            buf.append(l)
    if cur is not None:
        sec[cur] = "\n".join(buf).strip()
    return sec


def extract_part2_sections(entry, whitelist_ns):
    """Part II：簡介正文在「簡介」标记上方、矩阵之下（疾病名英文块之下）。"""
    pi = entry["page"] - 1
    lines = [l.strip() for l in PAGETXT[pi].splitlines() if l.strip()]
    intro_idx = next((k for k, l in enumerate(lines) if l == "簡介"), None)
    sec = {}
    if intro_idx is not None:
        body = []
        j = intro_idx - 1
        while j >= 0:
            l = lines[j]
            if "回到疾病索引" in l:
                j -= 1
                continue
            if DEC_RE.search(l) or re.search(r'[：:]', l):
                j -= 1
                continue  # 跳过矩阵行
            if re.search(r'[\u4e00-\u9fff]', l):
                body.insert(0, l)
            else:
                break  # 疾病名英文块之上，停止
            j -= 1
        if body:
            sec["簡介"] = "\n".join(body).strip()
    return sec


HDR_NAMES = ("人壽", "早期危疾", "危疾", "醫療", "癌症", "早期危疾 / 危疾", "早期危疾/危疾")


DEC_RE = re.compile(r'標準|加價|拒絕|擱置|不保|受保|承保|事項|延期|至\s*(?:拒絕|擱置|不保|受保|承保|事項)')


def block_matrix(words, gap=25):
    """按条件列聚类成块（y差>gap开新块），决定词按其 y 所在 [块min, 下块min) 区间归属。"""
    conds = sorted([w for w in words if not DEC_RE.search(w[4])], key=lambda w: (w[1], w[0]))
    blocks = []
    for w in conds:
        if blocks and w[1] - blocks[-1]["max_y"] > gap:
            blocks.append({"min_y": w[1], "max_y": w[1], "words": [w]})
        elif blocks:
            blocks[-1]["words"].append(w)
            blocks[-1]["max_y"] = max(blocks[-1]["max_y"], w[1])
        else:
            blocks.append({"min_y": w[1], "max_y": w[1], "words": [w]})
    decs = sorted([w for w in words if DEC_RE.search(w[4])], key=lambda w: (w[1], w[0]))
    for d in decs:
        for idx, b in enumerate(blocks):
            nmin = blocks[idx + 1]["min_y"] if idx + 1 < len(blocks) else 10 ** 9
            if b["min_y"] <= d[1] < nmin:
                b.setdefault("decs", []).append(d)
                break
    rows = []
    for b in blocks:
        cond = "".join(w[4] for w in sorted(b["words"], key=lambda v: (v[1], v[0])))
        dec = "".join(w[4] for w in sorted(b.get("decs", []), key=lambda v: (v[1], v[0])))
        dec = re.sub(r'[（(][^）)]*[）)]', '', dec)  # 去决定列括号注释
        dec = re.sub(r'(?:version\w*|\d+)\s*$', '', dec).strip()  # 去尾部页码/version
        rows.append({"condition": cond.strip(), "decision": dec.strip()})
    return rows


def page_matrix_part1(i, whitelist_ns):
    """Part I 单页矩阵：返回 (columns, rows)。"""
    if "核保指引" not in PAGETXT[i]:
        return None, []
    words = PAGEW[i]
    hdr_y = None
    cands = [w for w in words if w[4] == "人壽"]
    if cands:
        hdr_y = min(w[1] for w in cands)
    if hdr_y is None:
        return None, []
    low_y = 9999
    for w in words:
        if any(f in w[4] for f in ("本指引", "以上為", "版本")) and w[1] > hdr_y:
            low_y = min(low_y, w[1])
    if low_y == 9999:
        low_y = max(w[1] for w in words)
    mwords = [w for w in words if hdr_y + 2 < w[1] < low_y
              and not (re.match(r'^\d+$', w[4]) and w[0] > 850)]
    rows = block_matrix(mwords)
    cols = []
    for w in words:
        if abs(w[1] - hdr_y) < 4 and w[4] in HDR_NAMES:
            cols.append((w[0], w[4]))
    cols = [t for _, t in sorted(cols)]
    return cols, rows


def page_matrix_part2(i):
    """Part II 单页矩阵：双栏（条件 | 决定），columns=None。"""
    words = PAGEW[i]
    mwords = [w for w in words if w[1] < 700 and w[0] >= 440 and not re.match(r'^\d+$', w[4])
              and not any(f in w[4] for f in
                          ("回到", "版本", "核保部", "本指引", "以上", "Confidential",
                           "簡介", "健康保障核保考慮因素"))]
    rows = block_matrix(mwords)
    return None, rows


def extract_matrix(entry, part, whitelist_ns):
    if part == "II":
        cols, rows = page_matrix_part2(entry["page"] - 1)
        return cols, rows
    cols = None
    rows = []
    pages = list(range(entry["start"] - 1, entry["end"] + 1))
    if norm_ns(entry["name_cn"]) == norm_ns("2019冠狀病毒病"):
        pages += [70, 71]
    for pi in sorted(set(pages)):
        if pi < 0 or pi >= doc.page_count:
            continue
        c, r = page_matrix_part1(pi, whitelist_ns)
        if r:
            if cols is None:
                cols = c
            rows.extend(r)
    return cols, rows


def build_disease(e, part, cat, whitelist_ns):
    sec = extract_sections(e, part, whitelist_ns)
    cols, rows = extract_matrix(e, part, whitelist_ns)
    cat_name = ""
    if part == "I":
        for c, names in cat.items():
            if any(norm_ns(e["name_cn"]) == norm_ns(n) or norm_ns(e["name_cn"]) in norm_ns(n) for n in names):
                cat_name = c
                break
    else:
        cat_name = "健康保障核保專屬指引"
    pages = [e["start"], e["end"]] if part == "I" else [e["page"]]
    notes = []
    aliases = []
    if norm_ns(e["name_cn"]) == norm_ns("乙型肝炎"):
        aliases = ["乙肝", "乙型肝", "B型肝炎"]
    if norm_ns(e["name_cn"]) in ("二型糖尿病家族史", "糖尿病", "乙型肝炎", "肺結節", "甲狀腺結節"):
        notes.append("该条目核保矩阵较复杂，提取结果仅供参考，请以手册原文为准。")
    return {
        "name_cn": e["name_cn"],
        "name_en": clean_en(e.get("name_en", "")),
        "aliases": aliases,
        "category": cat_name,
        "part": part,
        "pages": pages,
        "revision": "",
        "sections": sec,
        "matrix": {"columns": cols, "rows": rows},
        "notes": notes,
    }


cat, whitelist_ns = parse_overview()
p1 = collect_part1(whitelist_ns)
p2 = collect_part2()
diseases = [build_disease(e, "I", cat, whitelist_ns) for e in p1]
diseases += [build_disease(e, "II", cat, whitelist_ns) for e in p2]
data = {
    "manual": {"version": "202103", "version_date": "2026-10-07",
               "title": "醫療核保指引手冊", "part2_title": "健康保障核保專屬指引"},
    "diseases": diseases,
    "appendices": [
        {"title": "附件二 常見直接拒絕／延期受保原因", "pages": [217, 222]},
        {"title": "附件三 BCIM3 簡易投保小貼士", "pages": [223, 231]},
    ],
}
with open(OUT1, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=1)
os.makedirs(os.path.dirname(OUT2), exist_ok=True)
with open(OUT2, "w", encoding="utf-8") as f:
    json.dump(data, f, ensure_ascii=False, indent=1)
print("written diseases:", len(diseases), "P1:", len(p1), "P2:", len(p2))
# 抽查
for e in p1:
    if e["name_cn"] in ("高血壓", "乙型肝炎", "糖尿病"):
        cols, rows = extract_matrix(e, "I", whitelist_ns)
        print("\nMATRIX", e["name_cn"], "cols=", cols)
        for r in rows[:6]:
            print("  ", r["condition"][:24], "=>", r["decision"][:24])
for e in p2[:2]:
    cols, rows = extract_matrix(e, "II", whitelist_ns)
    print("\nP2 MATRIX", e["name_cn"], "rows=", len(rows))
    for r in rows[:5]:
        print("  ", r["condition"][:24], "=>", r["decision"][:24])
