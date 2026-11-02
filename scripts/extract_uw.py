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
    HDR_SET = {"人壽", "早期危疾", "危疾", "醫療", "癌症", "癌病",
               "早期危疾 / 危疾", "早期危疾/危疾"}
    for pi in sorted(set(pages)):
        if pi < 0 or pi >= doc.page_count:
            continue
        for l in PAGETXT[pi].splitlines():
            l = l.strip()
            if l == "核保指引" or l in HDR_SET:
                # 矩阵区起始行（标题 或 数据列表头行）：finalize 当前 section
                if cur is not None:
                    sec[cur] = "\n".join(buf).strip()
                cur, buf = None, []
                continue
            if l in SECTION_KEYS:
                if cur is not None:
                    sec[cur] = "\n".join(buf).strip()
                cur, buf = l, []
                continue
            if cur is None:
                continue
            if is_foot_line(l, whitelist_ns):
                continue
            buf.append(l)
    if cur is not None:
        sec[cur] = "\n".join(buf).strip()
    _recover_basic_uw(sec)
    return sec


def _recover_basic_uw(sec):
    """基本核保要求为空时，把紧邻前一节尾部的编号检查列表块移入基本核保要求。
    仅触发于标题后置、列表被错收进前一节的页（前期糖尿病/糖尿病）。"""
    if sec.get("基本核保要求", "").strip():
        return
    keys = list(sec.keys())
    if "基本核保要求" not in keys:
        return
    idx = keys.index("基本核保要求")
    if idx == 0:
        return
    prev = keys[idx - 1]
    lines = sec[prev].split("\n")
    cut = None
    for i, l in enumerate(lines):
        s = l.strip()
        if re.match(r'^\d+\.?\s*$', s) or re.match(r'^\d+\.\S', s):
            cut = i
            break
    if cut is None:
        return
    head = "\n".join(lines[:cut]).strip()
    numblock = "\n".join(lines[cut:]).strip()
    if not numblock:
        return
    sec[prev] = head
    sec["基本核保要求"] = numblock


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


def _fill_kind(f):
    """表头填充色分类：red=数据列, grey=条件列, blue=整宽横幅(排除)。"""
    if not f:
        return "none"
    r, g, b = f[0], f[1], f[2]
    if max(r, g, b) - min(r, g, b) < 0.08:
        return "grey" if r < 0.9 else "white"
    if b - r > 0.15:
        return "blue"
    if r > g + 0.15 and r > b + 0.15:
        return "red"
    return "other"


def norm_colname(s):
    """归一化表头列名：去括号注解、去尾部脚注、压缩 '/' 与空白。"""
    s = (s or "").strip()
    s = re.sub(r'[（(][^）)]*[）)]', '', s)
    s = s.strip()
    s = re.sub(r'[#¹²³*＊]+$', '', s)
    s = re.sub(r'\s*/\s*', '/', s)
    s = re.sub(r'\s+', '', s)
    return s


def _collect_header_rects(page):
    """返回 (red数据列矩形[(x0,y0,x1,y1)], grey条件列矩形)。无表头返回空。"""
    red, grey = [], []
    for dr in page.get_drawings():
        if dr["type"] not in ("f", "fs"):
            continue
        fill = dr.get("fill")
        if not fill:
            continue
        r = dr["rect"]
        if r.width < 30 or r.height < 20:
            continue
        kind = _fill_kind(fill)
        if kind == "red":
            red.append((r.x0, r.y0, r.x1, r.y1))
        elif kind == "grey":
            grey.append((r.x0, r.y0, r.x1, r.y1))
    red.sort(key=lambda v: v[0])
    grey.sort(key=lambda v: v[0])
    return red, grey


def _words_in_rect(words, rect, tol=2.0):
    x0, y0, x1, y1 = rect
    out = []
    for w in words:
        wx0, wy0, wx1, wy1 = w[0], w[1], w[2], w[3]
        if wx0 >= x0 - tol and wx1 <= x1 + tol and wy0 >= y0 - tol and wy1 <= y1 + tol:
            out.append(w)
    out.sort(key=lambda v: (round(v[1]), v[0]))
    return "".join(w[4] for w in out)


def _cell_text(words):
    """单元格文字：y-chain(gap≤16) 成簇；簇内若相邻 x0 间隙>40（两竖列）则整体按 x
    升序（列优先），否则按短 y 间隙(≤8)分行、行内 x 升序。行间/句间空格。"""
    if not words:
        return ""
    ws = sorted(words, key=lambda v: v[1])
    clusters = [[ws[0]]]
    for w in ws[1:]:
        if w[1] - clusters[-1][-1][1] <= 16:
            clusters[-1].append(w)
        else:
            clusters.append([w])
    out = []
    for cl in clusters:
        xs = sorted(w[0] for w in cl)
        biggap = any(xs[k + 1] - xs[k] > 40 for k in range(len(xs) - 1))
        if biggap:
            cl2 = sorted(cl, key=lambda v: v[0])
            out.append("".join(w[4] for w in cl2))
            continue
        lines = [[cl[0]]]
        for w in cl[1:]:
            if w[1] - lines[-1][0][1] <= 8:
                lines[-1].append(w)
            else:
                lines.append([w])
        for ln in lines:
            ln.sort(key=lambda v: v[0])
            out.append("".join(w[4] for w in ln))
    return "\n".join(p for p in out if p).strip()


def _cond_text(words):
    """条件列：y 基线行还原词序（行内 x 升序、行间空格），再按右括号「)」边界
    一疾病名单元一行。括号配对，英文名内部空格保留。"""
    if not words:
        return ""
    ws = sorted(words, key=lambda v: v[1])
    rows = [[ws[0]]]
    for w in ws[1:]:
        if w[1] - rows[-1][0][1] <= 10:
            rows[-1].append(w)
        else:
            rows.append([w])
    def _join_row(r):
        r = sorted(r, key=lambda v: v[0])
        parts = [w[4] for w in r]
        out = parts[0] if parts else ""
        for prev, nxt in zip(parts, parts[1:]):
            prev_lat = bool(re.search(r'[A-Za-z0-9)$]', prev[-1:]))
            next_lat = bool(re.search(r'[A-Za-z0-9(]', nxt[:1]))
            prev_cjk = bool(re.search(r'[\u4e00-\u9fff]', prev[-1:]))
            next_cjk = bool(re.search(r'[\u4e00-\u9fff]', nxt[:1]))
            if prev_cjk and next_cjk:
                out += nxt
            elif prev_lat and next_lat:
                out += " " + nxt
            elif prev_cjk and next_lat:
                out += nxt
            else:
                out += " " + nxt
        return out
    line = ""
    for r in rows:
        line += _join_row(r)
    units, buf = [], ""
    for ch in line:
        buf += ch
        if ch == ")":
            buf = buf.strip()
            if buf:
                units.append(buf)
            buf = ""
    buf = buf.strip()
    if buf:
        units.append(buf)
    return "\n".join(units)


def page_matrix_part1(i, whitelist_ns):
    """Part I 单页矩阵：表头矩形+网格线精确分列。返回 (columns, rows)。"""
    page = doc[i]
    red_raw, grey = _collect_header_rects(page)
    words = PAGEW[i]

    # 修复4a：数据列矩形文字必须含产品关键词，否则不算数据列（页69分类索引页被跳过）
    red = []
    for rc in red_raw:
        name = norm_colname(_words_in_rect(words, rc))
        if any(k in name for k in ("人壽", "早期危疾", "危疾", "醫療", "癌症", "癌病")):
            red.append(rc)
    if not red:
        return None, []
    columns = [norm_colname(_words_in_rect(words, rc)) for rc in red]
    ncol = len(red)
    cond_x0 = min(rc[0] for rc in red)
    hdr_bottom = max(rc[3] for rc in red)

    # 网格线
    h_lines, v_lines = [], []
    for dr in page.get_drawings():
        if dr["type"] not in ("s", "l"):
            continue
        r = dr["rect"]
        if r.width > 5 and r.height < 3:
            h_lines.append((r.x0, r.y0, r.x1))
        elif r.height > 5 and r.width < 3:
            v_lines.append((r.x0, r.y0, r.y1))

    right_x = max(rc[2] for rc in red)
    table_w = right_x - cond_x0
    left_edge = min([v[0] for v in v_lines] + [cond_x0])

    # 行边界 y：同 y 水平线段合并覆盖度≥表宽25%，且在表头下方
    ybuckets = {}
    for x0, y, x1 in h_lines:
        ybuckets.setdefault(round(y, 0), []).append((x0, x1))
    row_seps = [hdr_bottom]
    for key, segs in ybuckets.items():
        if key <= hdr_bottom + 3:
            continue
        cover = sum(min(x1, right_x) - max(x0, cond_x0)
                    for x0, x1 in segs if x1 > cond_x0 and x0 < right_x)
        if cover >= 0.25 * table_w:
            row_seps.append(key)
    data_bottom = max([s[1] for s in h_lines] + [hdr_bottom])
    edges = sorted(set(row_seps + [hdr_bottom, data_bottom]))
    bands = [(a, b) for a, b in zip(edges[:-1], edges[1:]) if b - a >= 4]

    # 修复1：词按垂直中心单归属到行带
    def band_of(w):
        cy = (w[1] + w[3]) / 2
        if cy <= bands[0][0]:
            return 0
        for bi, (a, b) in enumerate(bands):
            if a < cy <= b:
                return bi
        return len(bands) - 1

    # 条件子列区间
    grey_page = len(grey) >= 2
    cond_subs = []  # [(x0,x1)] 条件区子列
    if grey_page:
        grey_sorted = sorted(grey, key=lambda v: v[0])
        main_int = (grey_sorted[0][0], grey_sorted[0][2])
        sec_int = (grey_sorted[1][0], grey_sorted[-1][2])
        cond_subs = [(g[0], g[2]) for g in grey_sorted]
    else:
        cond_vx = sorted([v[0] for v in v_lines
                          if left_edge - 3 < v[0] < cond_x0 - 3])
        bounds = [left_edge] + cond_vx + [cond_x0]
        cond_subs = [(bounds[k], bounds[k + 1]) for k in range(len(bounds) - 1)]

    def cond_sub_of(w):
        cx = (w[0] + w[2]) / 2
        for si, (a, b) in enumerate(cond_subs):
            if a - 2 <= cx < b + 2:
                return si
        return 0

    # 每带：条件词（按子列）、数据词（按列中心）
    band_data = []
    for bi, (by, ty) in enumerate(bands):
        cond_w = [[] for _ in cond_subs]
        data_w = [[] for _ in range(ncol)]
        for w in words:
            cy = (w[1] + w[3]) / 2
            if cy <= hdr_bottom + 1 or cy > data_bottom:
                continue
            if band_of(w) != bi:
                continue
            cx = (w[0] + w[2]) / 2
            if cx < cond_x0:
                cond_w[cond_sub_of(w)].append(w)
            else:
                ci = min(range(ncol), key=lambda k: abs(cx - (red[k][0] + red[k][2]) / 2))
                data_w[ci].append(w)
        band_data.append({"cond": cond_w, "data": data_w, "by": by, "ty": ty})

    def hline_covers_col(y, ci):
        cx0, cx1 = red[ci][0], red[ci][2]
        for x0, ly, x1 in h_lines:
            if abs(ly - y) > 1.5:
                continue
            if (min(x1, cx1) - max(x0, cx0)) > 0.5 * (cx1 - cx0):
                return True
        return False

    def vline_present_col(xb, by, ty):
        for vx, vy0, vy1 in v_lines:
            if abs(vx - xb) > 2.5:
                continue
            if max(0.0, min(vy1, ty) - max(vy0, by)) >= 3.0:
                return True
        return False

    def build_cells(bds, by, ty):
        """合并组按[by,ty]区间垂直分隔线判定；数据词按组合并。返回(cells, merged_groups)。"""
        parent = list(range(ncol))

        def find(x):
            while parent[x] != x:
                parent[x] = parent[parent[x]]
                x = parent[x]
            return x
        for ci in range(ncol - 1):
            if not vline_present_col(red[ci][2], by, ty):
                parent[find(ci)] = find(ci + 1)
        groups = {}
        for ci in range(ncol):
            groups.setdefault(find(ci), []).append(ci)
        # 汇总数据词
        all_data = [w for bd in bds for ci in range(ncol) for w in bd["data"][ci]]
        gw = {g: [] for g in groups}
        for w in all_data:
            cx = (w[0] + w[2]) / 2
            gid = None
            for g, members in groups.items():
                gx0, gx1 = red[members[0]][0], red[members[-1]][2]
                if gx0 - 2 <= cx < gx1 + 2:
                    gid = g
                    break
            if gid is None:
                ci = min(range(ncol), key=lambda k: abs(cx - (red[k][0] + red[k][2]) / 2))
                gid = find(ci)
            gw[gid].append(w)
        cells, merged_groups = {}, []
        for g, members in groups.items():
            txt = _cell_text(gw.get(g, []))
            if not txt:
                continue
            if len(members) >= 2:
                merged_groups.append([columns[m] for m in members])
            cells[columns[members[0]]] = txt
        return cells, merged_groups

    rows = []

    if grey_page:
        # 灰色主/次条件列页（肺结节）视觉行重构 F1/F2/F3
        # F1: 主/次词按 x 中心严格分类（无容差）
        band_ms = []
        for bd in band_data:
            mw, sw = [], []
            for sub in bd["cond"]:
                for w in sub:
                    cx = (w[0] + w[2]) / 2
                    (mw if main_int[0] <= cx < main_int[1] else sw).append(w)
            band_ms.append((mw, sw))

        # 分组：以次列词为锚；无次列词的带并入上一组；
        # F3：相邻两个含次列词的带，次列词 y0 差≤16 并入同一组
        groups = []
        for bi in range(len(band_data)):
            mw, sw = band_ms[bi]
            if not groups:
                groups.append([bi]); continue
            if not sw:
                groups[-1].append(bi); continue
            prev_sec_y = None
            for p in groups[-1]:
                if band_ms[p][1]:
                    prev_sec_y = min(w[1] for w in band_ms[p][1])
            this_sec_y = min(w[1] for w in sw)
            if prev_sec_y is not None and this_sec_y - prev_sec_y <= 16:
                groups[-1].append(bi)
            else:
                groups.append([bi])

        row_main, row_sec = [], []
        for grp in groups:
            mw = [w for bi in grp for w in band_ms[bi][0]]
            # 次列按带分别组句再按 y 序空格连接，避免跨带 x 混排
            sec_parts = [_cell_text(band_ms[bi][1]) for bi in grp if band_ms[bi][1]]
            row_main.append(_cell_text(mw))
            row_sec.append(" ".join(sec_parts))

        # F2：主列值双向最近传播（空值行取最近有值行，平局取上方）
        filled = list(row_main)
        for gi in range(len(filled)):
            if filled[gi]:
                continue
            best, bestd = None, 10 ** 6
            for gj in range(len(row_main)):
                if row_main[gj]:
                    d = abs(gj - gi)
                    if d < bestd or (d == bestd and gj < gi):
                        bestd, best = d, gj
            filled[gi] = row_main[best] if best is not None else ""

        for gi, grp in enumerate(groups):
            bds = [band_data[bi] for bi in grp]
            sec = row_sec[gi]
            cond = f"結節大小{filled[gi]}；年期{sec}".strip("；").strip()
            cells, merged_groups = build_cells(bds, bds[0]["by"], bds[-1]["ty"])
            row = {"condition": cond, "cells": cells}
            if merged_groups:
                row["merged"] = merged_groups
            if cond or cells:
                rows.append(row)
    else:
        # 普通页 / 双条件列页（页70）：每带一行
        pre_cells = []
        pre_merged = []
        pre_cond_sub = []
        for bd in band_data:
            cells, merged_groups = build_cells([bd], bd["by"], bd["ty"])
            pre_cells.append(cells)
            pre_merged.append(merged_groups)
            allw = [w for sub in bd["cond"] for w in sub]
            n_open = sum(w[4].count("(") for w in allw)
            if not grey_page and n_open >= 2:
                # 多「名(英)」单元格：整格收集一次重组（R1 类竖排疾病名）
                pre_cond_sub.append([_cond_text(allw)])
            else:
                pre_cond_sub.append([_cond_text(s) for s in bd["cond"]])

        def _region_covered(y, x0, x1):
            for hx0, ly, hx1 in h_lines:
                if abs(ly - y) > 1.5:
                    continue
                if (min(hx1, x1) - max(hx0, x0)) >= 0.5 * (x1 - x0):
                    return True
            return False

        def _union(a, b):
            if not a:
                return b or ""
            if not b:
                return a
            if a == b:
                return a
            if a in b:
                return b
            if b in a:
                return a
            return a + " " + b

        # R2：跨带垂直合并（双向并集）
        if len(cond_subs) >= 2:
            for bi in range(1, len(band_data)):
                y_bound = band_data[bi]["by"]
                for ci in range(ncol):
                    if not _region_covered(y_bound, red[ci][0], red[ci][2]):
                        k = columns[ci]
                        u = _union(pre_cells[bi - 1].get(k, ""), pre_cells[bi].get(k, ""))
                        pre_cells[bi - 1][k] = u
                        pre_cells[bi][k] = u
                for si in range(min(len(cond_subs), len(pre_cond_sub[bi]),
                                    len(pre_cond_sub[bi - 1]))):
                    if not grey_page:
                        continue  # 非灰页条件子列不跨带并集（避免中英分隔线误合并）
                    if not _region_covered(y_bound, cond_subs[si][0], cond_subs[si][1]):
                        u = _union(pre_cond_sub[bi - 1][si], pre_cond_sub[bi][si])
                        pre_cond_sub[bi - 1][si] = u
                        pre_cond_sub[bi][si] = u

        for bi, bd in enumerate(band_data):
            cond = " ".join(filter(None, pre_cond_sub[bi])).strip()
            row = {"condition": cond, "cells": pre_cells[bi]}
            if pre_merged[bi]:
                row["merged"] = pre_merged[bi]
            rows.append(row)

    rows = [r for r in rows if r["condition"] or r["cells"]]
    return columns, rows


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
    cols, rows = [], []
    pages = list(range(entry["start"] - 1, entry["end"] + 1))
    if norm_ns(entry["name_cn"]) == norm_ns("2019冠狀病毒病"):
        pages += [70, 71]
    col_union = []
    for pi in sorted(set(pages)):
        if pi < 0 or pi >= doc.page_count:
            continue
        c, r = page_matrix_part1(pi, whitelist_ns)
        if r:
            for cn in c:
                if cn and cn not in col_union:
                    col_union.append(cn)
            rows.extend(r)
    return (col_union or None), rows


def _norm(s):
    return re.sub(r"\s+", "", s or "")


def _page_is_matrix(pno):
    p = doc[pno]
    lines = [_norm(w[4]) for w in p.get_text("words")]
    def has(s):
        return any(s in ln for ln in lines)
    return has("人壽") and has("早期危疾") and (has("醫療") or has("癌症") or has("癌病"))


def extract_section_tables(entry, part):
    """提取文字页内的内嵌小表格（非核保指引矩阵），按 y 上方最近 SECTION 标题归位。
    返回 {section名: [table,...]}；table={cols, rows, merged?}。"""
    pages = list(range(entry["start"] - 1, entry["end"])) if part == "I" else []
    if norm_ns(entry["name_cn"]) == norm_ns("2019冠狀病毒病"):
        pages += [70, 71]
    result = {}
    region_words = set()
    if not pages:
        return result, region_words
    for pi in sorted(set(pages)):
        if pi < 0 or pi >= doc.page_count:
            continue
        if _page_is_matrix(pi):
            continue  # 核保指引矩阵页，跳过
        page = doc[pi]
        words = page.get_text("words")
        hl, vl = [], []
        for dr in page.get_drawings():
            r = dr["rect"]
            if dr["type"] not in ("s", "l"):
                continue
            if r.width > 40 and r.height < 3:
                hl.append(round(r.y0, 1))
            elif r.height > 20 and r.width < 3:
                vl.append((round(r.x0, 1), round(r.y0, 1), round(r.y1, 1)))
        if len(hl) < 3 or len(vl) < 2:
            continue
        hy0, hy1 = min(hl), max(hl)
        vspan = sorted(set(v[0] for v in vl if v[1] <= hy0 + 4 and v[2] >= hy1 - 4))
        if len(vspan) < 3:
            continue  # 至少 2 列，避免单列假表
        rows_y = sorted(set(hl))
        col_x = vspan
        # 归属 section：本页 y 小于表顶的最近 SECTION_KEY
        owner = None
        best_y = -1
        for w in words:
            if w[4] in SECTION_KEYS and w[1] < hy0 + 2 and w[1] > best_y:
                best_y = w[1]; owner = w[4]
        if not owner:
            owner = "基本核保要求"  # 页内无 section 标题时（该表即基本核保要求本身）
        # 填格
        table_rows = []
        for ri in range(len(rows_y) - 1):
            ry0, ry1 = rows_y[ri], rows_y[ri + 1]
            rowcells = []
            for ci in range(len(col_x) - 1):
                cx0, cx1 = col_x[ci], col_x[ci + 1]
                cw = [w for w in words
                      if cx0 - 2 <= (w[0] + w[2]) / 2 < cx1 + 2
                      and ry0 - 1 <= (w[1] + w[3]) / 2 < ry1 + 1]
                cw.sort(key=lambda v: (v[1], v[0]))
                for w in cw:
                    region_words.add(_norm(w[4]))
                # 行内 y 聚类
                txt = _cell_text(cw)
                txt = txt.strip().lstrip("▪•").strip()
                rowcells.append(txt)
            table_rows.append(rowcells)
        if len(table_rows) < 2:
            continue
        cols = table_rows[0]
        data_rows = table_rows[1:]
        if not cols[0].strip():
            cols[0] = "定义"  # 首列维度表头
        region_words.add(_norm("定義"))
        result.setdefault(owner, []).append({"cols": cols, "rows": data_rows})
    return result, region_words


def build_disease(e, part, cat, whitelist_ns):
    sec = extract_sections(e, part, whitelist_ns)
    sectabs, region_words = extract_section_tables(e, part)
    # 从所有 section 字符串中移出表格行文本（表格词可能落在任一 section）
    tokens = set(region_words)
    for tabs in sectabs.values():
        for t in tabs:
            for c in t["cols"]:
                for ln in c.split("\n"):
                    s = re.sub(r"\s+", "", ln)
                    if s: tokens.add(s)
            for row in t["rows"]:
                for c in row:
                    for ln in c.split("\n"):
                        s = re.sub(r"\s+", "", ln)
                        if s: tokens.add(s)
    title_norm = _norm(e.get("name_cn", ""))
    title_trad = {"帕金森病家族病史": "柏金遜症家族病史"}.get(e.get("name_cn", ""), "")
    title_trad_norm = _norm(title_trad)
    for sname in list(sec.keys()):
        kept = []
        for ln in sec[sname].split("\n"):
            n = _norm(ln.strip())
            if not n:
                continue
            if n in tokens:
                continue
            if title_norm and n == title_norm:
                continue
            if title_trad_norm and n == title_trad_norm:
                continue  # PDF 繁体疾病名独立行
            if re.match(r'^（.*[A-Za-z]{3,}.*）$', ln.strip()):
                continue  # 英文病名括号行
            kept.append(ln)
        # 相邻重复行去重
        dedup = []
        for ln in kept:
            if dedup and _norm(dedup[-1]) == _norm(ln):
                continue
            dedup.append(ln)
        sec[sname] = "\n".join(dedup).strip()
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
    if norm_ns(e["name_cn"]) == norm_ns("肺結節"):
        notes.append("产品适用分类索引（手册页76）未收录进矩阵，各保障类别的适用产品范围以手册各分页表头及原文为准。")
    if norm_ns(e["name_cn"]) == norm_ns("睡眠窒息症"):
        notes.append("该矩阵较复杂，提取结果仅供参考，请以手册原文为准。")
    return {
        "name_cn": e["name_cn"],
        "name_en": clean_en(e.get("name_en", "")),
        "aliases": aliases,
        "category": cat_name,
        "part": part,
        "pages": pages,
        "revision": "",
        "sections": sec,
        "section_tables": sectabs,
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
    if e["name_cn"] in ("高血壓", "甲狀腺結節", "乙型肝炎", "肺結節", "二型糖尿病家族史", "糖尿病"):
        cols, rows = extract_matrix(e, "I", whitelist_ns)
        print("\nMATRIX", e["name_cn"], "rows=", len(rows), "cols=", cols)
        for r in rows[:12]:
            print("   條件:", r["condition"][:40])
            for k, v in r.get("cells", {}).items():
                print("      ", k, "=", v[:50])
for e in p2[:2]:
    cols, rows = extract_matrix(e, "II", whitelist_ns)
    print("\nP2 MATRIX", e["name_cn"], "rows=", len(rows))
    for r in rows[:5]:
        print("  ", r["condition"][:24], "=>", r["decision"][:24])
