"""把核保库 uw_data.json 从繁体转为简体（含 dict key，如 sections 的「簡介」→「简介」）。
备份繁体原版到 uw_data_trad_backup.json。运行：python scripts/_uw_t2s.py"""
import json, os, shutil
from opencc import OpenCC

cc = OpenCC("t2s")

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
    with open(p, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    # 复核
    with open(p, "r", encoding="utf-8") as f:
        chk = json.load(f)
    d0 = chk["diseases"][0]
    print("OK", os.path.basename(p), "diseases=", len(chk.get("diseases", [])),
          "first name_cn=", d0.get("name_cn"), "| first sec keys=", list((d0.get("sections") or {}).keys()))
