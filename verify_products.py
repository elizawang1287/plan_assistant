#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
本地验证脚本：跑各产品PDF解析，对比基准结果
用法：
  python verify_products.py baseline   # 生成基准结果
  python verify_products.py verify     # 对比基准，检查是否有变化
"""
import sys, os, json, hashlib
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from parser import parse_pdf

# 各产品代表性测试PDF
TEST_PDFS = [
    ("CIM3", r"C:\Users\user\Documents\working\working\家庭规划书\1993年男\M_34_NS_CIM3_20_100000_20261103.pdf"),
    ("CIE3", r"C:\Users\user\Documents\working\working\家庭规划书\53岁女\M_54_NS_CIE3_15_250000_20261102.pdf"),
    ("CIP2", r"C:\Users\user\Documents\working\working\家庭规划书\1993年男\M_34_NS_CIP2_20_100000_20261103.pdf"),
    ("TRST", r"C:\Users\user\Documents\working\working\家庭规划书\F_33_NS_TRST_5_31000_20261106.pdf"),
    ("ENCASH", r"C:\Users\user\Documents\working\working\家庭规划书\53岁女\M_54_NS_ENCASH_15_250000_20261102.pdf"),
    ("PRMESP", r"C:\Users\user\Documents\working\working\家庭规划书\家庭2\M_5_NS_PRMESP_1_30000_11000_20260930.pdf"),
    ("PLP2", r"C:\Users\user\Documents\working\working\家庭规划书\53岁女\F_53_NS_PLP2_15_250000_rider_20261102.pdf"),
    ("LFRE2", r"C:\Users\user\Documents\working\working\家庭规划书\53岁女\F_53_NS_LFRE2_10_250000_20261102.pdf"),
    ("PLT5", r"C:\Users\user\Documents\working\working\家庭规划书\李总\M_50_NS_PLT5_36_300000_20261106.pdf"),
]

BASELINE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "verify_baseline.json")

def run_all():
    results = {}
    for name, path in TEST_PDFS:
        if not os.path.exists(path):
            print(f"[SKIP] {name}: 文件不存在 {path}")
            continue
        try:
            data = parse_pdf(path)
            # 提取关键字段做对比
            summary = {
                "product_code": data.get("product_code"),
                "product_name": data.get("product_name"),
                "basic": data.get("basic"),
                "rows_count": len(data.get("rows", [])),
                "rows_first3": data.get("rows", [])[:3],
                "rows_last1": data.get("rows", [])[-1:] if data.get("rows") else [],
                "warning": data.get("warning"),
            }
            results[name] = summary
            print(f"[OK] {name}: {data.get('product_name')} | rows={len(data.get('rows', []))} | 年保费={data.get('basic',{}).get('annual_premium')}")
        except Exception as e:
            results[name] = {"error": str(e)}
            print(f"[ERR] {name}: {e}")
    return results

def make_hash(obj):
    s = json.dumps(obj, ensure_ascii=False, sort_keys=True, default=str)
    return hashlib.md5(s.encode("utf-8")).hexdigest()

if __name__ == "__main__":
    mode = sys.argv[1] if len(sys.argv) > 1 else "verify"
    results = run_all()

    if mode == "baseline":
        with open(BASELINE_FILE, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2, default=str)
        print(f"\n基准已保存到 {BASELINE_FILE}")
    else:
        if not os.path.exists(BASELINE_FILE):
            print("没有基准文件，请先运行: python verify_products.py baseline")
            sys.exit(1)
        with open(BASELINE_FILE, "r", encoding="utf-8") as f:
            baseline = json.load(f)

        diff_count = 0
        for name, current in results.items():
            old = baseline.get(name, {})
            if make_hash(current) != make_hash(old):
                diff_count += 1
                print(f"\n[DIFF] {name}:")
                # 找出具体哪个字段变了
                for k in set(list(current.keys()) + list(old.keys())):
                    if make_hash(current.get(k)) != make_hash(old.get(k)):
                        print(f"  - {k}:")
                        print(f"    旧: {json.dumps(old.get(k), ensure_ascii=False, default=str)[:200]}")
                        print(f"    新: {json.dumps(current.get(k), ensure_ascii=False, default=str)[:200]}")
        if diff_count == 0:
            print("\n✅ 所有产品解析结果与基准一致，没有变化")
        else:
            print(f"\n❌ 有 {diff_count} 个产品的结果发生了变化，请检查！")
            sys.exit(2)
