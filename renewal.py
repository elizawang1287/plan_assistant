# -*- coding: utf-8 -*-
"""
续保提醒模块：解析 AES «Premium Position Enquiry» 页面另存的 HTML，
按客户合并保费、换算港币，生成催缴名单。

整合自 read_premium_position.py + merge_premium_by_client.py，
改为函数式接口供 FastAPI 调用。
"""

import collections
import html as htmllib
import os
import re
from html.parser import HTMLParser

# 繁体转简体
try:
    from opencc import OpenCC
    _T2S = OpenCC("t2s")
except Exception:
    _T2S = None


def t2s(s):
    if not s or _T2S is None or not isinstance(s, str):
        return s
    return _T2S.convert(s)


# ---------------------------------------------------------------- JS 解析

FIELDS = [
    "branchCd", "servStaffNo", "polNo", "billNo", "name", "clientCd",
    "role", "status", "phone1", "phone2", "ddaDt", "ddaReason", "plan",
    "pMode", "pMethod", "overdue", "prodType", "overPrem", "ptd",
    "modalPrem", "pdaBalance", "modalPremLevy",
]

MONEY_FIELDS = ["overPrem", "modalPrem", "pdaBalance", "modalPremLevy"]

PREM_POS_RE = re.compile(r"data\[(\d+)\]\s*=\s*new\s+premPos\((.*?)\)\s*;", re.S)


def clean(s):
    return " ".join((s or "").split())


def unescape(s):
    return htmllib.unescape(s or "")


def split_args(argstr):
    args, buf, in_quote, esc = [], [], False, False
    for ch in argstr:
        if in_quote:
            if esc:
                buf.append(ch); esc = False
            elif ch == "\\":
                esc = True
            elif ch == '"':
                in_quote = False
            else:
                buf.append(ch)
        else:
            if ch == '"':
                in_quote = True
            elif ch == ",":
                args.append("".join(buf).strip()); buf = []
            else:
                buf.append(ch)
    tail = "".join(buf).strip()
    if tail or args:
        args.append(tail)
    return args


def parse_js(text):
    records = []
    for m in PREM_POS_RE.finditer(text):
        idx = int(m.group(1))
        args = split_args(m.group(2))
        args = [unescape(a).strip().strip('"') for a in args]
        if len(args) < len(FIELDS):
            args += [""] * (len(FIELDS) - len(args))
        rec = {"row": idx}
        for name, val in zip(FIELDS, args):
            rec[name] = val
        records.append(rec)
    records.sort(key=lambda r: r["row"])
    return records


# ---------------------------------------------------------------- 表格解析（交叉校验用）

class RowTableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows = {}
        self.cur = None
        self.buf = []
        self.last_row = None

    def handle_starttag(self, tag, attrs):
        d = dict(attrs)
        tid = d.get("id", "") or ""
        if tag == "td" and tid.startswith("row"):
            head = tid[3:]
            num, _, field = head.partition("_")
            if num.isdigit():
                self.cur = (int(num), field)
                self.buf = []
                self.last_row = int(num)

    def handle_endtag(self, tag):
        if tag == "td" and self.cur:
            num, field = self.cur
            self.rows.setdefault(num, {})[field] = clean("".join(self.buf))
            self.cur = None

    def handle_data(self, data):
        if self.cur:
            self.buf.append(data)


def parse_table(text):
    p = RowTableParser()
    p.feed(text)
    return p.rows


def num(v):
    s = (v or "").replace(",", "").replace("$", "").strip()
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


# ---------------------------------------------------------------- 核心：合并生成名单

def currency_and_hkd(rec, rate):
    """返回 (币种, 港币整数, 港币精确, 原币金额)。"""
    over = num(rec.get("overPrem"))
    modal = num(rec.get("modalPrem"))
    if modal is not None and modal > 0 and over is not None:
        if abs(over / modal - 1.0) < 0.01:
            exact = modal; cur = "HKD"
        else:
            exact = modal * rate; cur = "USD"
        return cur, int(round(exact)), round(exact, 2), modal
    exact = over or 0.0
    return "HKD", int(round(exact)), round(exact, 2), modal or 0.0


def display_name(full):
    cn = "".join(ch for ch in full if "\u4e00" <= ch <= "\u9fff" or ch == "·")
    name = cn if cn else re.sub(r"\s+", " ", full).strip()
    return t2s(name)


def bill_id(rec):
    b = (rec.get("billNo") or "").strip().upper()
    if b.startswith("BN"):
        b = b[2:]
    return b or (rec.get("polNo") or "").strip()


def money(v, unit=True):
    s = format(int(round(v)), ",")
    return ("港币" + s) if unit else s


def _period(recs):
    for r in recs:
        ptd = (r.get("ptd") or "").strip()
        mt = re.match(r"(\d{2})/(\d{2})/(\d{4})", ptd)
        if mt:
            return int(mt.group(2)), int(mt.group(3))
    return None, None


def build_note(recs):
    m, y = _period(recs)
    if not m:
        return ""
    import calendar
    last = calendar.monthrange(y, m)[1]
    return "%d月续保需%d月%d日前交" % (m, m, last)


def build_dda_note(recs, day=6):
    m, y = _period(recs)
    if not m:
        return ""
    return "%d月续保%d月%d号扣款" % (m, m, day)


def parse_renewal_html(html_text, rate=7.9, dda_day=6, include_monthly=False):
    """主入口：解析 HTML，返回结构化结果。

    返回: {
      "total_records": int,
      "client_count": int,
      "skipped_monthly": int,
      "total_hkd": int,
      "total_hkd_exact": float,
      "note": str,          # 自助缴费收尾语
      "dda_note": str,      # 自动转账收尾语
      "clients": [ {
          "name": str, "clientCd": str, "phone": str,
          "total_hkd": int, "total_hkd_exact": float,
          "policies": [ {
              "plan": str, "bill_id": str, "polNo": str,
              "currency": str, "hkd": int, "hkd_exact": float,
              "orig": float, "pMode": str, "pMethod": str,
              "is_auto": bool, "is_pda_zero": bool,
              "ptd": str, "overdue": str, "ddaReason": str,
          } ],
      } ],
      "text_list": str,     # 纯文字催缴名单
    }
    """
    recs = parse_js(html_text)
    if not recs:
        table = parse_table(html_text)
        recs = [dict(v, row=k) for k, v in table.items()]

    note = build_note(recs)
    dda_note = build_dda_note(recs, dda_day)

    # 剔除月缴
    skipped = []
    if not include_monthly:
        skipped = [r for r in recs if (r.get("pMode") or "").strip().upper() == "MONTHLY"]
        recs = [r for r in recs if (r.get("pMode") or "").strip().upper() != "MONTHLY"]

    # 逐保单算港币
    for r in recs:
        cur, hkd, hkd_exact, orig = currency_and_hkd(r, rate)
        r["_cur"], r["_hkd"], r["_hkd_exact"], r["_orig"] = cur, hkd, hkd_exact, orig
        r["_name"] = display_name(r.get("name", ""))
        r["_id"] = bill_id(r)

    # 按客户合并
    groups = collections.OrderedDict()
    for r in recs:
        key = r.get("clientCd") or r.get("polNo")
        g = groups.setdefault(key, {
            "clientCd": key, "name": r["_name"],
            "policies": [], "phone": r.get("phone1", ""), "dda": [],
        })
        g["policies"].append(r)
        if r.get("ddaReason"):
            g["dda"].append(r.get("ddaReason"))

    def is_auto(p):
        return "AUTOPAY" in (p.get("pMethod") or "").upper()

    def amount_label(p):
        tags = []
        if is_auto(p):
            tags.append("自动转账")
        if p["_hkd"] == 0:
            tags.append("PDA余额已覆盖本期保费")
        return ("（%s）" % "，".join(tags)) if tags else ""

    # 生成纯文字名单
    lines = []
    clients_out = []
    for key, g in groups.items():
        pols = g["policies"]
        total = sum(p["_hkd"] for p in pols)
        exact = round(sum(p["_hkd_exact"] for p in pols), 2)

        policy_list = []
        for p in pols:
            policy_list.append({
                "plan": p.get("plan", ""),
                "bill_id": p["_id"],
                "polNo": p.get("polNo", ""),
                "currency": p["_cur"],
                "hkd": p["_hkd"],
                "hkd_exact": p["_hkd_exact"],
                "orig": p["_orig"],
                "pMode": p.get("pMode", ""),
                "pMethod": p.get("pMethod", ""),
                "is_auto": is_auto(p),
                "is_pda_zero": p["_hkd"] == 0,
                "ptd": p.get("ptd", ""),
                "overdue": p.get("overdue", ""),
                "ddaReason": p.get("ddaReason", ""),
            })

        clients_out.append({
            "name": g["name"],
            "clientCd": g["clientCd"],
            "phone": g["phone"],
            "total_hkd": total,
            "total_hkd_exact": exact,
            "policy_count": len(pols),
            "policies": policy_list,
        })

        # 文字名单
        if len(pols) == 1:
            p = pols[0]
            lines.append("%s，%s，%s%s，%s" % (
                g["name"], p["_id"], money(p["_hkd"]), amount_label(p),
                dda_note if is_auto(p) else note))
        else:
            for p in pols:
                tail = "，%s" % dda_note if is_auto(p) else ""
                lines.append("%s，%s，%s%s%s" % (
                    g["name"], p["_id"], money(p["_hkd"]), amount_label(p), tail))
            all_auto = all(is_auto(p) for p in pols)
            if all_auto:
                combo = "（自动转账）"
            elif any(is_auto(p) for p in pols):
                auto_sum = sum(p["_hkd"] for p in pols if is_auto(p))
                combo = "（含自动转账港币%s）" % format(int(auto_sum), ",")
            else:
                combo = ""
            lines.append("%d份单总保费，%s%s，%s" % (
                len(pols), money(total), combo,
                dda_note if all_auto else note))
        lines.append("")

    total_all = sum(c["total_hkd"] for c in clients_out)
    exact_all = round(sum(c["total_hkd_exact"] for c in clients_out), 2)

    header = "本月续保应缴清单（汇率 %.1f，单位：港币，金额取整；%s；编号=缴费编号，即账单号去掉BN前缀）\n\n" % (
        rate, "含年缴+月缴" if include_monthly else "仅年缴保单，已剔除月缴")

    return {
        "total_records": len(recs),
        "client_count": len(clients_out),
        "skipped_monthly": len(skipped),
        "total_hkd": total_all,
        "total_hkd_exact": exact_all,
        "note": note,
        "dda_note": dda_note,
        "clients": clients_out,
        "text_list": header + "\n".join(lines).rstrip() + "\n",
    }
