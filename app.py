# -*- coding: utf-8 -*-
"""
计划书助手 - 本地后端
启动：python app.py  然后浏览器打开 http://127.0.0.1:8000
"""
import io
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
        base = payload["base"]
        variants = payload.get("variants", []) or [base]
        buf = ppt_export.build_combined_pptx(base, variants, rate)
        return StreamingResponse(
            buf,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            headers={"Content-Disposition": 'attachment; filename="plan.pptx"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"导出组合PPT失败：{e}")


# 静态前端（放最后挂载，避免拦截 API）
app.mount("/", StaticFiles(directory=STATIC_DIR, html=True), name="static")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
