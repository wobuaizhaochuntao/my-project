#!/usr/bin/env python3
"""Generate一转/二转/三转/补充协议 from a JSON config + 债转文件."""
from __future__ import annotations

import argparse
import json
import shutil
import sys
from decimal import Decimal
from pathlib import Path

ROOT = Path(__file__).resolve().parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from common import (  # noqa: E402
    cn_money,
    fill_detail_table,
    find_docx,
    fmt_bracket_date,
    fmt_cn_date,
    money,
    read_debt_workbook,
    replace_doc_text,
    ymd_compact,
)
from docx import Document  # noqa: E402


def load_config(path: Path) -> dict:
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def party(cfg: dict, code: str) -> dict:
    p = cfg["parties"][code]
    required = ["short", "name"]
    for k in required:
        if k not in p:
            raise KeyError(f"parties.{code} 缺少 {k}")
    return p


def build_first(cfg: dict, data: dict, out_first: Path) -> list[Path]:
    written = []
    tpl = Path(cfg["templates"]["first"])
    first_date = cfg["dates"]["first"]
    cn = fmt_cn_date(first_date)
    compact = ymd_compact(first_date)
    for item in cfg.get("first_transfers", []):
        sheet = item["sheet"]
        recipient_code = item["recipient"]
        rec = party(cfg, recipient_code)
        rows = data[sheet]
        total = sum((money(r[7]) for r in rows), Decimal("0")).quantize(Decimal("0.01"))
        doc = Document(str(tpl))
        # Template is typically三农信->锋泰; replace recipient fields.
        old_rec = cfg["templates"].get("first_template_recipient", {})
        replacements = [
            (old_rec.get("name", "北京锋泰科技有限公司"), rec["name"]),
            (
                old_rec.get("address", "北京市朝阳区东三环北路19号楼中青大厦1908室"),
                rec.get("address", old_rec.get("address", "北京市朝阳区东三环北路19号楼中青大厦1908室")),
            ),
            (old_rec.get("account", "10242000000483660"), rec.get("account", old_rec.get("account", ""))),
            (cfg["templates"].get("first_template_date_cn", "2026年08月13日"), cn),
            (cfg["templates"].get("first_template_date_compact", "20260813"), compact),
        ]
        if rec.get("bank") and old_rec.get("bank"):
            replacements.append((old_rec["bank"], rec["bank"]))
        amount_text = f"{cn_money(total)}( ￥{total:,.2f}元)"
        regexes = [
            (
                r"(借款债权及附属权益总额为).*?(,共计)\d+(笔,后续如新增)",
                lambda m: m.group(1) + amount_text + m.group(2) + str(len(rows)) + m.group(3),
            ),
        ]
        replace_doc_text(doc, replacements, regexes)
        fill_detail_table(doc, rows, six_cols=True)
        out = out_first / f"【上海法久&{rec['short']}】一转协议-{compact}.docx"
        doc.save(str(out))
        written.append(out)
    return written


def build_second_series(cfg: dict, data: dict, out_second: Path) -> list[Path]:
    written = []
    batch_label = cfg["batch_label"]  # e.g. 第12批
    old_batch = cfg["templates"].get("old_batch_label", "第11批")
    second_date = cfg["dates"]["second"]
    cn = fmt_cn_date(second_date)
    compact = ymd_compact(second_date)
    bracket = fmt_bracket_date(second_date)
    old_cn = cfg["templates"].get("second_template_date_cn", "2026年08月17日")
    old_compact = cfg["templates"].get("second_template_date_compact", "20260817")
    old_bracket = cfg["templates"].get(
        "second_template_date_bracket", "【2026】年【08】月【17】日"
    )
    old_dest = cfg["templates"]["second_template_recipient"]
    unified_address = cfg.get("second_contact_address")

    for item in cfg["second_groups"]:
        sheet = item["sheet"]
        transferor_code = item["transferor"]
        dest_code = item["recipient"]
        tpl_folder = Path(item["template_folder"])
        transferor = party(cfg, transferor_code)
        dest = party(cfg, dest_code)
        rows = data[sheet]
        target_folder = out_second / sheet
        target_folder.mkdir(parents=True, exist_ok=True)

        for kind, suffix in (("二转协议", ""), ("三转协议", "-2"), ("补充协议", "-1")):
            src = find_docx(tpl_folder, kind)
            doc = Document(str(src))
            replacements = [
                (old_batch, batch_label),
                (old_compact, compact),
                (old_cn, cn),
                (old_bracket, bracket),
                (old_dest["name"], dest["name"]),
                (old_dest["address"], dest.get("address", old_dest["address"])),
                (old_dest["bank"], dest.get("bank", old_dest["bank"])),
                (old_dest["account"], dest.get("account", old_dest["account"])),
                (old_dest["code"], dest_code),
            ]
            # Optional transferor swap when template transferor differs (e.g. 锋泰 template -> 焕然)
            old_tr = item.get("template_transferor")
            if old_tr:
                ot = party(cfg, old_tr) if old_tr in cfg["parties"] else cfg["templates"].get("parties_extra", {}).get(old_tr, {})
                # Prefer explicit override map on the item
                swap = item.get("transferor_replacements", [])
                for a, b in swap:
                    replacements.append((a, b))
                if not swap and ot:
                    if ot.get("name") and transferor.get("name"):
                        replacements.append((ot["name"], transferor["name"]))
                    if ot.get("address") and transferor.get("address"):
                        replacements.append((ot["address"], transferor["address"]))
                    if ot.get("account") and transferor.get("account"):
                        replacements.append((ot["account"], transferor["account"]))
                    if ot.get("bank") and transferor.get("bank"):
                        replacements.append((ot["bank"], transferor["bank"]))
                    if old_tr != transferor_code:
                        replacements.append((f"{old_tr}-", f"{transferor_code}-"))

            regexes = [
                (r"共计【\d+】笔", f"共计【{len(rows)}】笔"),
                (r"共【\d+】笔", f"共【{len(rows)}】笔"),
            ]
            replace_doc_text(doc, replacements, regexes)

            if unified_address and kind == "二转协议":
                # After SPV address replace, force contact address if requested.
                replace_doc_text(
                    doc,
                    [
                        (dest.get("address", ""), unified_address),
                        (old_dest.get("address", ""), unified_address),
                    ],
                )

            if kind in ("二转协议", "三转协议"):
                fill_detail_table(doc, rows, six_cols=False)

            filename = f"【法久&{transferor['short']}&{batch_label}】{kind}-{compact}.docx"
            out = target_folder / filename
            doc.save(str(out))
            written.append(out)
    return written


def main() -> int:
    ap = argparse.ArgumentParser(description="按配置生成债转协议 Word")
    ap.add_argument("--config", required=True, help="JSON 配置文件路径")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划，不写文件")
    args = ap.parse_args()

    cfg = load_config(Path(args.config))
    excel = Path(cfg["debt_excel"])
    data = read_debt_workbook(excel)

    missing = []
    for item in cfg.get("first_transfers", []):
        if item["sheet"] not in data:
            missing.append(item["sheet"])
    for item in cfg.get("second_groups", []):
        if item["sheet"] not in data:
            missing.append(item["sheet"])
    if missing:
        raise SystemExit(f"债转文件缺少工作表: {missing}")

    root = Path(cfg["output_root"])
    batch_n = cfg["batch_number"]
    first_mmdd = ymd_compact(cfg["dates"]["first"])[4:]
    second_mmdd = ymd_compact(cfg["dates"]["second"])[4:]
    out_first = root / f"法久第{batch_n}批一转协议{first_mmdd}"
    out_second = root / f"法久第{batch_n}批二转协议{second_mmdd}"

    plan = {
        "debt_excel": str(excel),
        "first_docs": len(cfg.get("first_transfers", [])),
        "second_groups": len(cfg.get("second_groups", [])),
        "second_docs": len(cfg.get("second_groups", [])) * 3,
        "out_first": str(out_first),
        "out_second": str(out_second),
        "sheets": {k: len(v) for k, v in data.items()},
    }
    print(json.dumps(plan, ensure_ascii=False, indent=2))
    if args.dry_run:
        return 0

    if out_first.exists():
        shutil.rmtree(out_first)
    if out_second.exists():
        shutil.rmtree(out_second)
    out_first.mkdir(parents=True)
    out_second.mkdir(parents=True)

    written = []
    written += build_first(cfg, data, out_first)
    written += build_second_series(cfg, data, out_second)
    print(json.dumps({"written": len(written), "files": [str(p) for p in written]}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
