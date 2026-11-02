"""把核保库 uw_data.json 从繁体转为简体（含 dict key，如 sections 的「簡介」→「简介」）。
备份繁体原版到 uw_data_trad_backup.json。运行：python scripts/_uw_t2s.py"""
import json, os, shutil
from opencc import OpenCC

cc = OpenCC("t2s")

# opencc 漏转的异体字/字符级全局映射（递归覆盖所有文本）
CHAR_MAP = {"穏": "稳", "於": "于", "著": "着"}


def char_map(v):
    if isinstance(v, str):
        for a, b in CHAR_MAP.items():
            v = v.replace(a, b)
        return v
    if isinstance(v, list):
        return [char_map(x) for x in v]
    if isinstance(v, dict):
        return {k: char_map(x) for k, x in v.items()}
    return v

def conv(v):
    if isinstance(v, str):
        return cc.convert(v)
    if isinstance(v, list):
        return [conv(x) for x in v]
    if isinstance(v, dict):
        out = {}
        for k, x in v.items():
            nk = cc.convert(k) if isinstance(k, str) else k
            out[nk] = conv(x)
        return out
    return v


# 机械规范化：按疾病名 + 行索引写死目标（condition / cond_group），不动 cells/columns
NORM = {
    "甲状腺功能减退症": {
        0: {"condition": "有足够治疗和病情受控(如无高血压或心血管异常)\n≤2年\n>2年"},
    },
    "宫颈上皮内瘤样病变": {
        0: {"condition": "CIN I\n诊断后跟进之子宫颈抹片报告正常"},
        1: {"condition": "CIN II\n完全切除；\n病理报告证实无浸润；\n手术后至少6个月的跟进之子宫颈抹片报告正常"},
        2: {"condition": "CIN III\n人类乳头状瘤病毒(HPV)阳性"},
    },
    "肠化生": {
        0: {"condition": "轻微至中度肠化生1\n有非典型增生"},
        1: {"condition": "轻微至中度肠化生1\n没有非典型增生及没有幽门螺旋杆菌感染(或曾经受感染但已完成相关治疗)"},
        2: {"condition": "严重肠化生1"},
    },
    "乳房良性疾病": {
        2: {"condition": "良性乳腺结节(Benign breast nodule)/\n乳房纤维腺瘤(Fibroadenoma)"},
        6: {"condition": "乳房纤维腺瘤(Fibroadenoma)\n单纯性纤维腺瘤(无非典型性质)"},
        7: {"condition": "乳房纤维腺瘤(Fibroadenoma)\n非单纯性纤维腺瘤"},
    },
    "2019 冠状病毒病": {
        6: {"condition": "加护病房(ICU), 重症治疗病房(ITU)\n所有症状已经痊愈，无并发症，"},
    },
    "睡眠窒息症": {
        0: {"cond_group": "未有治疗", "condition": "轻度"},
        1: {"cond_group": "未有治疗", "condition": "中度"},
        2: {"cond_group": "未有治疗", "condition": "严重"},
        3: {"cond_group": "已接受治疗\n或正在治疗\n且进展良好", "condition": "初次诊断；轻度"},
        4: {"cond_group": "已接受治疗\n或正在治疗\n且进展良好", "condition": "初次诊断；中度"},
        5: {"cond_group": "已接受治疗\n或正在治疗\n且进展良好", "condition": "初次诊断；严重"},
    },
    "三尖瓣反流": {
        0: {"condition": "微量(Trivial)三尖瓣反流；没有其他心脏、心瓣、心血管问题"},
        1: {"condition": "轻度(Mild)三尖瓣反流；没有其他心脏、心瓣、心血管问题"},
        2: {"condition": "中度(Moderate)三尖瓣反流；没有其他心脏、心瓣、心血管问题"},
        3: {"condition": "严重(Severe)三尖瓣反流曾经接受心瓣手术"},
    },
    "二尖瓣反流": {
        0: {"condition": "微量(Trivial)二尖瓣反流；没有其他心脏、心瓣、心血管问题"},
        1: {"condition": "轻度(Mild)二尖瓣反流；没有其他心脏、心瓣、心血管问题"},
        2: {"condition": "中度(Moderate)至严重(Severe)二尖瓣反流曾经接受心瓣手术"},
    },
}

# 肺结节：合并断行碎片 + 主列/次列首「；」后插 \n（保留 \n> 真子行）
def fix_nodule(c):
    if not c:
        return c
    c = c.replace("及\n稳定", "及稳定")
    c = c.replace("注示\n#)", "注示#)")
    c = c.replace("注示\n#", "注示#")
    if "；" in c:
        a, b = c.split("；", 1)
        c = a + "；\n" + b
    return c


ROW_SPLIT = {
    "三尖瓣反流": {0: 2, 1: 2},
}


def apply_norm(data):
    for a in data.get("diseases", []):
        nm = a.get("name_cn")
        if not a.get("matrix"):
            continue
        rows = a["matrix"]["rows"]
        if nm in ROW_SPLIT:
            new_rows = []
            for ri, r in enumerate(rows):
                if ri in ROW_SPLIT[nm]:
                    # 复制行（cells 继承），condition 稍后由 NORM 按最终行号写
                    for _ in range(ROW_SPLIT[nm][ri]):
                        new_rows.append({"condition": r.get("condition", ""),
                                         "cells": dict(r.get("cells", {}))})
                else:
                    new_rows.append(r)
            a["matrix"]["rows"] = new_rows
            rows = new_rows
        if nm in NORM:
            for idx, fields in NORM[nm].items():
                if idx < len(rows):
                    for k, v in fields.items():
                        rows[idx][k] = v
        if nm == "肺结节":
            for r in rows:
                r["condition"] = fix_nodule(r.get("condition", ""))
    return data


base = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
paths = [
    os.path.join(base, "uw_data.json"),
    os.path.join(base, "static", "uw_data.json"),
]
for p in paths:
    if not os.path.exists(p):
        print("skip missing:", p); continue
    bak = os.path.join(base, "uw_data_trad_backup.json")
    if p.endswith("uw_data.json") and not os.path.exists(bak):
        shutil.copy2(p, bak)
        print("backup ->", bak)
    with open(p, "r", encoding="utf-8") as f:
        data = json.load(f)
    data = conv(data)
    data = char_map(data)
    data = apply_norm(data)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    # 复核
    with open(p, "r", encoding="utf-8") as f:
        chk = json.load(f)
    d0 = chk["diseases"][0]
    print("OK", os.path.basename(p), "diseases=", len(chk.get("diseases", [])),
          "first name_cn=", d0.get("name_cn"), "| first sec keys=", list((d0.get("sections") or {}).keys()))
