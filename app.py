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

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles
import uvicorn

import parser
import ppt_export

app = FastAPI(title="计划书助手")

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
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": 'attachment; filename="plan.pptx"'},
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
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": 'attachment; filename="plan.pptx"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"导出组合PPT失败：{e}")


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
