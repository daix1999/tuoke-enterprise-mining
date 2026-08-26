#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
一键跑完整拓客流程：搜索 → 补电话 → 写回 → 分表导出

用法:
  python run_full.py --db <ledger.db> --outdir <输出目录> \
      --identity "联想商用电脑" --customer-type "渠道商/经销商" --regcap "100￥500"

省 token 设计：内部串起 4 步，只打印关键摘要，名单明细只在 Excel 里。
"""
import argparse
import json
import os
import sqlite3
import subprocess
import sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
MINE = os.path.join(SCRIPT_DIR, "riskbird_random_mining.py")
BATCH = os.path.join(SCRIPT_DIR, "riskbird_batch.py")


def run(cmd, label):
    print(f"▶ {label}")
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
    out = (r.stdout or "").strip()
    # 只打印最后 8 行摘要，避免长输出进 token
    lines = [l for l in out.splitlines() if l.strip()]
    print("\n".join(lines[-8:]))
    if r.returncode != 0:
        print((r.stderr or "").strip(), file=sys.stderr)
        sys.exit(r.returncode)
    return out


def main():
    ap = argparse.ArgumentParser(description="一键拓客完整流程")
    ap.add_argument("--db", required=True)
    ap.add_argument("--outdir", required=True)
    ap.add_argument("--identity", default=None)
    ap.add_argument("--customer-type", default=None)
    ap.add_argument("--scope", default=None)
    ap.add_argument("--region", default=None)
    ap.add_argument("--regcap", default=None)
    ap.add_argument("--esdate", default=None)
    ap.add_argument("--contact", default=None)
    ap.add_argument("--has-bid-win", action="store_true")
    ap.add_argument("--has-patent", action="store_true")
    ap.add_argument("--has-soft-copyright", action="store_true")
    ap.add_argument("--count", type=int, default=50)
    args = ap.parse_args()

    py = sys.executable
    raw_dir = os.path.dirname(os.path.abspath(args.db))

    # 1. 搜索 + 入库 + 导出（无电话版）
    cmd = [py, MINE, "mine", "--db", args.db, "--count", str(args.count), "--outdir", args.outdir]
    for opt, attr in [("identity", "identity"), ("customer-type", "customer_type"),
                      ("scope", "scope"), ("region", "region"), ("regcap", "regcap"),
                      ("esdate", "esdate"), ("contact", "contact")]:
        v = getattr(args, attr)
        if v:
            cmd += [f"--{opt}", v]
    for flag in ["has_bid_win", "has_patent", "has_soft_copyright"]:
        if getattr(args, flag):
            cmd.append(f"--{flag.replace('_', '-')}")
    run(cmd, "步骤1/4 搜索")

    # 2. 补电话
    db = sqlite3.connect(args.db)
    names = [r[0] for r in db.execute(
        "SELECT company_name FROM companies WHERE (phone IS NULL OR phone='') AND (email IS NULL OR email='')")]
    db.close()
    if names:
        names_file = os.path.join(raw_dir, "raw", "_pending.json")
        os.makedirs(os.path.dirname(names_file), exist_ok=True)
        json.dump(names, open(names_file, "w", encoding="utf-8"), ensure_ascii=False)
        result_xlsx = os.path.join(raw_dir, "_补电话结果.xlsx")
        result_json = os.path.join(raw_dir, "_补电话结果_联系方式.json")
        run([py, BATCH, names_file, "--out", result_xlsx, "--json-out", result_json],
            f"步骤2/4 补电话 {len(names)} 家")

        # 3. 写回电话
        db = sqlite3.connect(args.db)
        data = json.load(open(result_json, encoding="utf-8"))
        n_phone = 0
        for x in data:
            if x.get("company") and not x.get("error"):
                db.execute("UPDATE companies SET phone=?, email=? WHERE company_name=?",
                           (x.get("phone", "") or "", x.get("email", "") or "", x["company"]))
                if x.get("phone"):
                    n_phone += 1
        db.commit()
        db.close()
        print(f"   写回电话 {n_phone} 家")
    else:
        print("✅ 无需补电话")

    # 4. 分表导出（含电话）
    run([py, MINE, "export", "--db", args.db, "--outdir", args.outdir], "步骤4/4 分表导出")

    # 最终统计
    db = sqlite3.connect(args.db)
    total = db.execute("SELECT COUNT(*) FROM companies").fetchone()[0]
    has_phone = db.execute("SELECT COUNT(*) FROM companies WHERE phone IS NOT NULL AND phone!=''").fetchone()[0]
    db.close()
    print(f"🎉 完成 | 累计 {total} 家 | 有电话 {has_phone} 家")


if __name__ == "__main__":
    main()
