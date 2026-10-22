# 保诚计划书助手

上传保诚数字版储蓄险 PDF → 自动解析受保人/保费/逐年退保价值/提取安排 → 生成保诚红长图和 PPT。

## 支持产品
- **PRMESP** 世誉财富保险计划（一次整付）
- **TRST** 保诚信守明天多元货币保险计划（年缴）

## 本地运行

```bash
pip install -r requirements.txt
python app.py
```

浏览器打开 http://127.0.0.1:8000

## 使用
1. 拖入或选择计划书 PDF（数字版，非扫描件）；
2. 点「开始解析」出长图；
3. 可下载长图、下载 PPT、复制摘要发客户。

## 部署到服务器

```bash
pip install -r requirements.txt
uvicorn app:app --host 0.0.0.0 --port 8000
```

或用 gunicorn + uvicorn worker：
```bash
pip install gunicorn
gunicorn app:app -w 2 -k uvicorn.workers.UvicornWorker -b 0.0.0.0:8000
```

## 文件结构
- `app.py` — FastAPI 后端（上传 /api/parse，静态页面）
- `parser.py` — PDF 文字层解析规则（PRMESP / TRST 双产品列映射）
- `ppt_export.py` — PPT 生成
- `static/index.html` — 前端页面与长图渲染

## 解析规则要点
- 退保价值取「基本計劃–退保價值之說明摘要」（不做提取的情景）
- 优先取「基本計劃補充說明」完整年度表，无则退回基础断档表
- TRST 列顺序在不同表中文字提取位置不同，通过表头自动识别
- 数字直接舍去（不四舍五入），美元列显示「X.X万」
