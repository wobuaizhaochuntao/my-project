"""Shared helpers for Fajiu debt-transfer SMS generation."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Any

SHORT_PARTIES: list[tuple[str, str]] = [
    ("黑龙江三农信融资担保有限公司", "SNX"),
    ("新疆应瑞融资担保有限公司", "YR"),
    ("北京锋泰科技有限公司", "FT"),
    ("北京焕然数字科技有限公司", "HR"),
    ("上海序普洪商务咨询有限公司", "序普洪"),
    ("上海序嘉洪商务咨询有限公司", "序嘉洪"),
    ("上海序洪商务咨询有限公司", "序洪"),
    ("上海启心程商务咨询有限公司", "启心程"),
    ("上海洲序商务咨询有限公司", "洲序"),
    ("杭州灵机一动商务咨询有限公司", "灵机一动"),
    ("上海臻行远商务咨询有限公司", "臻行远"),
    ("杭州君商喆辉商务咨询有限公司", "君商"),
    ("上海竣洪商务咨询有限公司", "竣洪"),
    ("杭州万妙汇新商务咨询有限公司", "万妙汇新"),
    ("上海洪均商务咨询有限公司", "洪均"),
]

DEFAULT_FUND = "众邦银行"

# FJ upload template columns (1-based)
COL_ORDER_ID = 1
COL_CURRENT_HOLDER = 7  # 当前债权归属主体
COL_SECOND_TRANSFEREE = 13  # 二次债转债权受让方
COL_THIRD_TRANSFEROR = 24  # 第三次债转债权出让方
COL_FUND = 31  # 资金方名称


def clean_party(name: Any) -> str:
    """Strip trailing numeric IDs like 3660 / 9768 from party names."""
    if name is None:
        return ""
    text = str(name).strip()
    return re.sub(r"\d+$", "", text).strip()


def short_of(name: Any) -> str:
    full = clean_party(name)
    if not full:
        raise ValueError("空主体名称")
    for party, short in SHORT_PARTIES:
        if full == party or full.startswith(party):
            return short
    raise ValueError(f"未知主体，请补 SHORT_PARTIES: {full!r}")


def to_order_id(value: Any) -> int | str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return int(text) if text.isdigit() else text


def normalize_oid_key(value: Any) -> str:
    oid = to_order_id(value)
    return str(oid)


def output_filename(from_party: str, to_party: str, date_mmdd: str) -> str:
    return f"FJ-债转短信-{short_of(from_party)}转{short_of(to_party)}-{date_mmdd}.xlsx"


def find_default_template(search_roots: list[Path]) -> Path | None:
    """Prefer a previously Excel-SaveAs'd FJ SMS file (has sharedStrings)."""
    candidates: list[Path] = []
    for root in search_roots:
        if not root.exists():
            continue
        for path in root.rglob("FJ-债转短信-*.xlsx"):
            if path.name.startswith("~$"):
                continue
            candidates.append(path)
    if not candidates:
        return None
    # Prefer newer files; larger ones after SaveAs are usually safer.
    candidates.sort(key=lambda p: (p.stat().st_mtime, p.stat().st_size), reverse=True)
    return candidates[0]


def excel_saveas(path: Path) -> None:
    """Rewrite xlsx via Excel COM so Fajiu format checks pass."""
    import pythoncom
    import win32com.client

    path = path.resolve()
    tmp = path.with_name("_tmp_excel_save.xlsx")
    if tmp.exists():
        tmp.unlink()

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    app.ScreenUpdating = False
    try:
        workbook = app.Workbooks.Open(str(path))
        workbook.SaveAs(str(tmp), FileFormat=51)
        workbook.Close(SaveChanges=False)
    finally:
        app.Quit()

    tmp.replace(path)
