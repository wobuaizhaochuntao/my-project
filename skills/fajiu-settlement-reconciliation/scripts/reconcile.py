#!/usr/bin/env python3
"""Reconcile Fajiu settlement applications against posting-success details."""

from __future__ import annotations

import argparse
import json
import os
import re
from collections import Counter, defaultdict
from copy import copy
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill


SOURCE_HEADERS = {
    "order": "订单号",
    "amount": "还款总金额",
    "review_date": "repaid_date",
}
POSTING_HEADERS = {
    "order": "借款单号",
    "amount": "还款金额",
    "posting_date": "复核日期",
}


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--settlement", type=Path)
    parser.add_argument("--posting", type=Path)
    parser.add_argument("--settlement-output", type=Path)
    parser.add_argument("--posting-output", type=Path)
    parser.add_argument(
        "--settlement-cols",
        default="",
        help="覆盖结算表列名，如 order=借款单号,amount=还款金额,review_date=复核日期",
    )
    parser.add_argument(
        "--posting-cols",
        default="",
        help="覆盖入账表列名，如 order=订单号,amount=还款总金额,posting_date=repaid_date",
    )
    parser.add_argument(
        "--date-order",
        choices=["posting_first", "app_first", "none"],
        default="posting_first",
        help="正常时间口径：posting_first=入账日期≤结算日期（默认，先复核后确认）；"
        "app_first=结算日期≤入账日期；none=不限制",
    )
    return parser.parse_args()


def apply_col_overrides(spec: str, target: dict[str, str]) -> None:
    if not spec:
        return
    for pair in spec.split(","):
        key, sep, value = pair.partition("=")
        key = key.strip()
        if not sep or key not in target or not value.strip():
            raise ValueError(
                f"无效的列覆盖 {pair!r}，可用键：{', '.join(target)}"
            )
        target[key] = value.strip()


def order_id(value) -> str | None:
    if value is None:
        return None
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip() or None


def cents(value) -> int | None:
    if value is None:
        return None
    return int((Decimal(str(value)) * 100).quantize(Decimal("1")))


def as_date(value) -> date | None:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    return datetime.strptime(str(value)[:10], "%Y-%m-%d").date()


def header_indexes(values) -> dict[str, int]:
    return {value: index for index, value in enumerate(values) if value is not None}


def find_sheet(workbook, required: dict[str, str], path: Path):
    active = workbook.active
    sheets = [active] + [s for s in workbook.worksheets if s.title != active.title]
    for sheet in sheets:
        first = next(sheet.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if first is None:
            continue
        indexes = header_indexes(first)
        if all(header in indexes for header in required.values()):
            return sheet, indexes
    raise ValueError(
        f"{path.name} 中未找到包含 {', '.join(required.values())} 表头的工作表"
    )


def read_settlement(path: Path):
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet, indexes = find_sheet(workbook, SOURCE_HEADERS, path)

    records = []
    incomplete = []
    for row_number, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
        values = {
            "row": row_number,
            "order": order_id(row[indexes[SOURCE_HEADERS["order"]]]),
            "amount": cents(row[indexes[SOURCE_HEADERS["amount"]]]),
            "date": as_date(row[indexes[SOURCE_HEADERS["review_date"]]]),
        }
        if all(value is None for value in (values["order"], values["amount"], values["date"])):
            continue
        if values["order"] is None or values["amount"] is None or values["date"] is None:
            incomplete.append(row_number)
        else:
            records.append(values)
    if incomplete:
        raise ValueError(f"结算表存在字段不完整的数据行：{incomplete[:20]}")
    return records, sheet.title


def read_postings(path: Path):
    workbook = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sheet, indexes = find_sheet(workbook, POSTING_HEADERS, path)
    records = []
    incomplete = []
    for row_number, row in enumerate(sheet.iter_rows(min_row=2, values_only=True), 2):
        values = {
            "row": row_number,
            "order": order_id(row[indexes[POSTING_HEADERS["order"]]]),
            "amount": cents(row[indexes[POSTING_HEADERS["amount"]]]),
            "date": as_date(row[indexes[POSTING_HEADERS["posting_date"]]]),
        }
        if all(value is None for value in (values["order"], values["amount"], values["date"])):
            continue
        if values["order"] is None or values["amount"] is None or values["date"] is None:
            incomplete.append(row_number)
        else:
            records.append(values)
    if incomplete:
        raise ValueError(f"入账表存在字段不完整的数据行：{incomplete[:20]}")
    return records, sheet.title


def best_assignment(applications, postings, date_order: str):
    """Choose disjoint exact-sum subsets for postings.

    date_order: "posting_first" 入账日期不晚于结算日期（先复核后确认）；
    "app_first" 结算日期不晚于入账日期；"none" 不限制。
    """
    count = len(applications)
    if count > 18:
        raise ValueError(
            f"借款单号 {applications[0]['order']} 有 {count} 条结算记录，"
            "超过安全组合上限18条，请人工拆分后重试"
        )

    options = []
    for posting in postings:
        candidates = []
        for mask in range(1, 1 << count):
            subset = [applications[i] for i in range(count) if mask >> i & 1]
            if sum(item["amount"] for item in subset) != posting["amount"]:
                continue
            if date_order == "app_first" and not all(
                item["date"] <= posting["date"] for item in subset
            ):
                continue
            if date_order == "posting_first" and not all(
                posting["date"] <= item["date"] for item in subset
            ):
                continue
            lag = sum(abs((posting["date"] - item["date"]).days) for item in subset)
            candidates.append((mask, lag))
        options.append(candidates)

    best_score = (-1, -1, -(10**18))
    best = []

    def search(posting_index, used_mask, chosen, matched_postings, matched_apps, lag):
        nonlocal best_score, best
        if posting_index == len(postings):
            score = (matched_postings, matched_apps, -lag)
            if score > best_score:
                best_score = score
                best = chosen[:]
            return

        search(
            posting_index + 1,
            used_mask,
            chosen,
            matched_postings,
            matched_apps,
            lag,
        )
        for mask, cost in options[posting_index]:
            if mask & used_mask:
                continue
            search(
                posting_index + 1,
                used_mask | mask,
                chosen + [(posting_index, mask)],
                matched_postings + 1,
                matched_apps + bin(mask).count("1"),
                lag + cost,
            )

    search(0, 0, [], 0, 0, 0)
    return best


def reconcile(applications, postings, date_order: str):
    applications_by_order = defaultdict(list)
    postings_by_order = defaultdict(list)
    for item in applications:
        applications_by_order[item["order"]].append(item)
    for item in postings:
        postings_by_order[item["order"]].append(item)

    source_result = {}
    posting_result = {}

    for order, raw_applications in applications_by_order.items():
        apps = sorted(raw_applications, key=lambda item: (item["date"], item["row"]))
        posts = sorted(
            postings_by_order.get(order, []),
            key=lambda item: (item["date"], item["row"]),
        )
        used_apps = set()
        used_posts = set()

        for posting_index, mask in best_assignment(apps, posts, date_order):
            members = [i for i in range(len(apps)) if mask >> i & 1]
            used_apps.update(members)
            used_posts.add(posting_index)
            posting = posts[posting_index]
            posting_result[posting["row"]] = {
                "kind": "normal",
                "source_rows": [apps[i]["row"] for i in members],
                "amount": sum(apps[i]["amount"] for i in members),
            }
            for i in members:
                source_result[apps[i]["row"]] = {
                    "status": "入账成功",
                    "posting": posting,
                    "group_size": len(members),
                }

        remaining_apps = [
            item for i, item in enumerate(apps) if i not in used_apps
        ]
        remaining_posts = [
            item for i, item in enumerate(posts) if i not in used_posts
        ]
        anomaly_used_apps = set()
        anomaly_used_posts = set()

        for posting_index, mask in best_assignment(
            remaining_apps, remaining_posts, "none"
        ):
            members = [
                i for i in range(len(remaining_apps)) if mask >> i & 1
            ]
            anomaly_used_apps.update(members)
            anomaly_used_posts.add(posting_index)
            posting = remaining_posts[posting_index]
            posting_result[posting["row"]] = {
                "kind": "anomaly",
                "source_rows": [remaining_apps[i]["row"] for i in members],
                "amount": sum(remaining_apps[i]["amount"] for i in members),
            }
            for i in members:
                source_result[remaining_apps[i]["row"]] = {
                    "status": "时间异常待复核",
                    "posting": posting,
                    "group_size": len(members),
                }

        unmatched_posts = [
            item
            for i, item in enumerate(remaining_posts)
            if i not in anomaly_used_posts
        ]
        for i, application in enumerate(remaining_apps):
            if i in anomaly_used_apps:
                continue
            if order not in postings_by_order:
                status = "未找到入账成功记录"
            elif not unmatched_posts:
                status = "重复申请/无剩余成功记录"
            else:
                status = "金额不一致"
            source_result[application["row"]] = {
                "status": status,
                "posting": None,
                "group_size": 0,
            }

    return source_result, posting_result


def style_header(cell, source_cell) -> None:
    if source_cell.has_style:
        cell._style = copy(source_cell._style)
    cell.font = Font(
        name=cell.font.name,
        size=cell.font.sz,
        bold=True,
        color=cell.font.color,
    )
    cell.alignment = Alignment(horizontal="center", vertical="center")


def fmt_dates(items) -> str:
    return ",".join(sorted({str(item["date"]) for item in items}))


def fmt_amounts(items) -> str:
    return "+".join(f"{item['amount'] / 100:.2f}" for item in items)


def period_label(path: Path) -> str:
    match = re.search(r"(\d{1,2})月结算", path.stem)
    return f"{match.group(1)}月结算" if match else "结算表"


def write_settlement(
    source: Path,
    output: Path,
    sheet_title: str,
    applications,
    postings,
    source_result,
) -> None:
    workbook = openpyxl.load_workbook(source)
    sheet = workbook[sheet_title]
    nonempty_headers = [
        column
        for column in range(1, sheet.max_column + 1)
        if sheet.cell(1, column).value is not None
    ]
    last_header = max(nonempty_headers)
    names = [
        "核对结果",
        "入账成功日期",
        "匹配还款金额",
        "时间差（天）",
        "合并匹配组",
        "合并后金额",
        "对应入账日期",
        "对应入账金额",
        "复核口径",
    ]
    existing = {
        sheet.cell(1, column).value: column
        for column in range(1, sheet.max_column + 1)
        if sheet.cell(1, column).value is not None
    }
    columns = {}
    next_column = last_header + 1
    for name in names:
        if name in existing:
            columns[name] = existing[name]
        else:
            columns[name] = next_column
            next_column += 1
        cell = sheet.cell(1, columns[name], name)
        style_header(cell, sheet.cell(1, last_header))

    fills = {
        "入账成功": "C6EFCE",
        "时间异常待复核": "FFEB9C",
        "重复申请/无剩余成功记录": "D9EAD3",
        "金额不一致": "F4CCCC",
        "未找到入账成功记录": "E7E6E6",
    }
    by_row = {item["row"]: item for item in applications}
    posts_by_order = defaultdict(list)
    for item in postings:
        posts_by_order[item["order"]].append(item)

    for row_number, result in source_result.items():
        status = result["status"]
        sheet.cell(row_number, columns["核对结果"], status)
        sheet.cell(row_number, columns["核对结果"]).fill = PatternFill(
            "solid", fgColor=fills[status]
        )
        for name in names[1:-1]:
            sheet.cell(row_number, columns[name], None)
        sheet.cell(
            row_number,
            columns["复核口径"],
            "订单号 + 还款总金额合计（精确到分） + 时间一对一"
            "（入账复核日期 ≤ 结算确认日期）",
        )

        posting = result["posting"]
        application = by_row[row_number]
        if posting is None:
            order_posts = posts_by_order.get(application["order"], [])
            if order_posts:
                sheet.cell(
                    row_number, columns["对应入账日期"], fmt_dates(order_posts)
                )
                sheet.cell(
                    row_number,
                    columns["对应入账金额"],
                    fmt_amounts(order_posts),
                )
            continue
        sheet.cell(row_number, columns["入账成功日期"], posting["date"])
        sheet.cell(row_number, columns["入账成功日期"]).number_format = "yyyy-mm-dd"
        sheet.cell(
            row_number, columns["匹配还款金额"], application["amount"] / 100
        )
        sheet.cell(row_number, columns["匹配还款金额"]).number_format = "0.00"
        sheet.cell(
            row_number,
            columns["时间差（天）"],
            (posting["date"] - application["date"]).days,
        )
        sheet.cell(
            row_number,
            columns["合并匹配组"],
            f"{application['order']}-{posting['date']}-L{posting['row']}",
        )
        sheet.cell(row_number, columns["合并后金额"], posting["amount"] / 100)
        sheet.cell(row_number, columns["合并后金额"]).number_format = "0.00"
        sheet.cell(row_number, columns["对应入账日期"], str(posting["date"]))
        sheet.cell(row_number, columns["对应入账金额"], posting["amount"] / 100)
        sheet.cell(row_number, columns["对应入账金额"]).number_format = "0.00"

    for name, column in columns.items():
        if name == "复核口径":
            width = 44
        elif name in ("核对结果", "合并匹配组"):
            width = 28
        else:
            width = 16
        sheet.column_dimensions[openpyxl.utils.get_column_letter(column)].width = width
    sheet.auto_filter.ref = (
        f"A1:{openpyxl.utils.get_column_letter(max(columns.values()))}{sheet.max_row}"
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)


def write_postings(
    source: Path,
    output: Path,
    posting_sheet_name: str,
    applications,
    postings,
    posting_result,
    label: str,
) -> None:
    workbook = openpyxl.load_workbook(source)
    sheet = workbook[posting_sheet_name]
    last_header = max(
        column
        for column in range(1, sheet.max_column + 1)
        if sheet.cell(1, column).value is not None
    )
    names = [
        f"{label}体现情况",
        f"对应{label}行",
        "对应还款金额合计",
        f"对应{label}日期",
        f"对应{label}金额",
        "时间核对",
    ]
    existing = {
        sheet.cell(1, column).value: column
        for column in range(1, sheet.max_column + 1)
        if sheet.cell(1, column).value is not None
    }
    columns = {}
    next_column = last_header + 1
    for name in names:
        if name in existing:
            columns[name] = existing[name]
        else:
            columns[name] = next_column
            next_column += 1
        cell = sheet.cell(1, columns[name], name)
        style_header(cell, sheet.cell(1, last_header))

    statuses = {
        "normal": f"已在{label}体现",
        "anomaly": "已体现-时间异常待复核",
        "missing_no_order": f"未在{label}体现（无此订单）",
        "missing_amount": f"未在{label}体现（金额不一致）",
    }
    fills = {
        statuses["normal"]: "C6EFCE",
        statuses["anomaly"]: "FFEB9C",
        statuses["missing_no_order"]: "F4CCCC",
        statuses["missing_amount"]: "FCE5CD",
    }
    posting_rows = {item["row"] for item in postings}
    posting_by_row = {item["row"]: item for item in postings}
    app_by_row = {item["row"]: item for item in applications}
    apps_by_order = defaultdict(list)
    for item in applications:
        apps_by_order[item["order"]].append(item)

    for row_number in posting_rows:
        for name in names:
            sheet.cell(row_number, columns[name], None)
        result = posting_result.get(row_number)
        order = posting_by_row[row_number]["order"]
        if result:
            status = statuses[result["kind"]]
            members = [app_by_row[row] for row in result["source_rows"]]
            sheet.cell(
                row_number,
                columns[f"对应{label}行"],
                ",".join(map(str, result["source_rows"])),
            )
            sheet.cell(
                row_number,
                columns["对应还款金额合计"],
                result["amount"] / 100,
            )
            sheet.cell(row_number, columns["对应还款金额合计"]).number_format = "0.00"
            sheet.cell(row_number, columns[f"对应{label}日期"], fmt_dates(members))
            sheet.cell(row_number, columns[f"对应{label}金额"], fmt_amounts(members))
            sheet.cell(
                row_number,
                columns["时间核对"],
                "正常" if result["kind"] == "normal" else "异常",
            )
        else:
            order_apps = apps_by_order.get(order, [])
            if order_apps:
                status = statuses["missing_amount"]
                sheet.cell(
                    row_number, columns[f"对应{label}日期"], fmt_dates(order_apps)
                )
                sheet.cell(
                    row_number, columns[f"对应{label}金额"], fmt_amounts(order_apps)
                )
            else:
                status = statuses["missing_no_order"]
        sheet.cell(row_number, columns[f"{label}体现情况"], status)
        sheet.cell(row_number, columns[f"{label}体现情况"]).fill = PatternFill(
            "solid", fgColor=fills[status]
        )

    for name, column in columns.items():
        sheet.column_dimensions[openpyxl.utils.get_column_letter(column)].width = (
            28 if "情况" in name or "行" in name else 20
        )
    sheet.auto_filter.ref = (
        f"A1:{openpyxl.utils.get_column_letter(max(columns.values()))}{sheet.max_row}"
    )
    output.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output)


def classify_posting_kinds(postings, posting_result, settlement_orders):
    kinds = []
    for item in postings:
        kind = posting_result.get(item["row"], {}).get("kind")
        if kind is None:
            kind = (
                "missing_amount"
                if item["order"] in settlement_orders
                else "missing_no_order"
            )
        kinds.append(kind)
    return kinds


def summarize(applications, postings, source_result, posting_result):
    source_counts = Counter(item["status"] for item in source_result.values())
    settlement_orders = {item["order"] for item in applications}
    posting_counts = Counter(
        classify_posting_kinds(postings, posting_result, settlement_orders)
    )
    application_by_row = {item["row"]: item for item in applications}
    posting_by_row = {item["row"]: item for item in postings}

    normal_source_amount = sum(
        application_by_row[row]["amount"]
        for row, result in source_result.items()
        if result["status"] == "入账成功"
    )
    normal_posting_amount = sum(
        posting_by_row[row]["amount"]
        for row, result in posting_result.items()
        if result["kind"] == "normal"
    )
    anomaly_source_amount = sum(
        application_by_row[row]["amount"]
        for row, result in source_result.items()
        if result["status"] == "时间异常待复核"
    )
    anomaly_posting_amount = sum(
        posting_by_row[row]["amount"]
        for row, result in posting_result.items()
        if result["kind"] == "anomaly"
    )
    missing_posting_amount = sum(
        item["amount"] for item in postings if item["row"] not in posting_result
    )

    balanced = (
        normal_source_amount == normal_posting_amount
        and anomaly_source_amount == anomaly_posting_amount
        and sum(source_counts.values()) == len(applications)
        and sum(posting_counts.values()) == len(postings)
    )
    return {
        "balanced": balanced,
        "settlement_rows": len(applications),
        "posting_rows": len(postings),
        "settlement_counts": dict(source_counts),
        "posting_counts": {
            "已正常体现": posting_counts["normal"],
            "时间异常待复核": posting_counts["anomaly"],
            "未体现-无此订单": posting_counts["missing_no_order"],
            "未体现-金额不一致": posting_counts["missing_amount"],
        },
        "normal_matched_amount": f"{normal_source_amount / 100:.2f}",
        "anomaly_matched_amount": f"{anomaly_source_amount / 100:.2f}",
        "missing_posting_amount": f"{missing_posting_amount / 100:.2f}",
    }


def output_name(path: Path, tag: str) -> str:
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return f"{path.stem}_{tag}_{stamp}.xlsx"


def main() -> None:
    args = parse_args()
    settlement_value = args.settlement or os.environ.get("FAJIU_SETTLEMENT")
    posting_value = args.posting or os.environ.get("FAJIU_POSTING")
    if not settlement_value or not posting_value:
        raise ValueError(
            "请通过--settlement/--posting或"
            "FAJIU_SETTLEMENT/FAJIU_POSTING指定两个输入文件"
        )
    settlement = Path(settlement_value).resolve()
    posting = Path(posting_value).resolve()
    if not settlement.exists():
        raise FileNotFoundError(settlement)
    if not posting.exists():
        raise FileNotFoundError(posting)

    label = period_label(settlement)
    settlement_output = (
        args.settlement_output.resolve()
        if args.settlement_output
        else settlement.with_name(output_name(settlement, "已核对"))
    )
    posting_output = (
        args.posting_output.resolve()
        if args.posting_output
        else posting.with_name(output_name(posting, f"{label}体现标记"))
    )

    apply_col_overrides(args.settlement_cols, SOURCE_HEADERS)
    apply_col_overrides(args.posting_cols, POSTING_HEADERS)

    applications, settlement_sheet = read_settlement(settlement)
    postings, posting_sheet_name = read_postings(posting)
    source_result, posting_result = reconcile(
        applications, postings, args.date_order
    )
    write_settlement(
        settlement,
        settlement_output,
        settlement_sheet,
        applications,
        postings,
        source_result,
    )
    write_postings(
        posting,
        posting_output,
        posting_sheet_name,
        applications,
        postings,
        posting_result,
        label,
    )
    summary = summarize(
        applications,
        postings,
        source_result,
        posting_result,
    )
    summary["settlement_output"] = str(settlement_output)
    summary["posting_output"] = str(posting_output)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    if not summary["balanced"]:
        raise RuntimeError("核对输出未通过金额或行数平衡校验")


if __name__ == "__main__":
    main()
