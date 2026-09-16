#!/usr/bin/env python3
"""Summarize a债转文件 workbook."""
from __future__ import annotations

import argparse
import json
import sys
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common import money, read_debt_workbook  # noqa: E402


def main() -> int:
    ap = argparse.ArgumentParser(description="读取债转文件并输出各表笔数与金额")
    ap.add_argument("excel", help="债转文件.xlsx 路径")
    ap.add_argument("--json", action="store_true", help="JSON 输出")
    args = ap.parse_args()

    data = read_debt_workbook(args.excel)
    summary = []
    all_orders = set()
    total_rows = 0
    total_amt = Decimal("0")
    for name, rows in data.items():
        amt = sum((money(r[7]) for r in rows), Decimal("0"))
        orders = {r[1] for r in rows}
        all_orders |= orders
        total_rows += len(rows)
        total_amt += amt
        seq_ok = all(r[0] == str(i) for i, r in enumerate(rows, 1))
        summary.append(
            {
                "sheet": name,
                "rows": len(rows),
                "unique_orders": len(orders),
                "amount": f"{amt:.2f}",
                "seq_ok": seq_ok,
            }
        )

    out = {
        "file": str(Path(args.excel).resolve()),
        "sheets": summary,
        "total_rows": total_rows,
        "unique_orders": len(all_orders),
        "sum_of_sheet_amounts": f"{total_amt:.2f}",
        "note": "两段转让订单会在两张表各计一次，unique_orders 为去重后笔数",
    }
    if args.json:
        print(json.dumps(out, ensure_ascii=False, indent=2))
    else:
        print(f"文件: {out['file']}")
        print(f"工作表: {len(summary)} | 明细行: {total_rows} | 唯一订单: {len(all_orders)}")
        for row in summary:
            flag = "" if row["seq_ok"] else " [序号异常]"
            print(
                f"  {row['sheet']}: {row['rows']}笔 / {row['amount']}元{flag}"
            )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
