"""Create an Excel QA report for the Slide QA Copilot MVP."""
from __future__ import annotations

import io
import json
import re
from collections import Counter
from typing import Any

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

DARK = PatternFill("solid", fgColor="1F2937")
TEAL = PatternFill("solid", fgColor="0F766E")
WHITE = Font(color="FFFFFF", bold=True)
WRAP = Alignment(vertical="top", wrap_text=True)
THIN = Side(style="thin", color="E5E7EB")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Excel worksheets are XML files. Source pages and model verbatims can contain
# invisible control characters that are valid in Python strings but illegal in
# Excel XML. Remove only those forbidden characters before writing cells.
ILLEGAL_EXCEL_CHARACTERS_RE = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F\uFFFE\uFFFF]")
EXCEL_CELL_TEXT_LIMIT = 32767


def _clean(value: Any) -> Any:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(value, (dict, list, tuple, set)):
        value = json.dumps(value, ensure_ascii=False, default=str)

    if isinstance(value, str):
        value = ILLEGAL_EXCEL_CHARACTERS_RE.sub("", value)
        return value[:EXCEL_CELL_TEXT_LIMIT]
    return value


def _write_dataframe(ws, df: pd.DataFrame, widths: dict[str, int] | None = None) -> None:
    for col_idx, col in enumerate(df.columns, start=1):
        c = ws.cell(row=1, column=col_idx, value=col)
        c.fill = DARK
        c.font = WHITE
        c.border = BORDER
        c.alignment = WRAP
    for r_idx, row in enumerate(df.itertuples(index=False), start=2):
        for c_idx, value in enumerate(row, start=1):
            c = ws.cell(row=r_idx, column=c_idx, value=_clean(value))
            c.border = BORDER
            c.alignment = WRAP
            if isinstance(value, str) and value.startswith("http"):
                c.hyperlink = value
                c.font = Font(color="1D4ED8", underline="single")
    for idx, col in enumerate(df.columns, start=1):
        width = (widths or {}).get(col, 22)
        ws.column_dimensions[get_column_letter(idx)].width = width
    ws.freeze_panes = "A2"
    if len(df.columns):
        ws.auto_filter.ref = f"A1:{get_column_letter(len(df.columns))}{max(1, len(df) + 1)}"


def build_report(results: list[dict]) -> bytes:
    claims_rows: list[dict] = []
    source_rows: list[dict] = []
    usage_rows: list[dict] = []

    for result in results:
        claims_rows.append({k: v for k, v in result.items() if k not in {"sources", "usage"}})
        for source in result.get("sources", []) or []:
            source_rows.append({
                "Claim ID": result.get("Claim ID", ""),
                "Slide": result.get("Slide", ""),
                "Slide Title": result.get("Slide Title", ""),
                **source,
            })
        usage = result.get("usage") or {}
        usage_rows.append({
            "Claim ID": result.get("Claim ID", ""),
            "Slide": result.get("Slide", ""),
            **usage,
        })

    claims_df = pd.DataFrame(claims_rows)
    sources_df = pd.DataFrame(source_rows)
    usage_df = pd.DataFrame(usage_rows)

    wb = Workbook()
    summary = wb.active
    summary.title = "QA Summary"
    summary["A1"] = "Slide QA Summary"
    summary["A1"].font = Font(size=16, bold=True)

    status_counts = Counter(claims_df.get("Status", pd.Series(dtype=str)).tolist())
    metrics = [
        ("Claims checked", len(claims_df)),
        ("Confirmed", status_counts.get("✅ Confirmed", 0)),
        ("Partial", status_counts.get("⚠️ Partial", 0)),
        ("Incorrect", status_counts.get("❌ Incorrect", 0)),
        ("Not Found", status_counts.get("❓ Not Found", 0)),
        ("Inaccessible", status_counts.get("🔒 Inaccessible", 0)),
        ("No Reference", status_counts.get("❓ No Reference", 0)),
    ]
    for c_idx, name in enumerate(["Metric", "Value"], start=1):
        c = summary.cell(row=3, column=c_idx, value=name)
        c.fill = TEAL
        c.font = WHITE
        c.border = BORDER
    for r_idx, (name, value) in enumerate(metrics, start=4):
        summary.cell(r_idx, 1, name).border = BORDER
        summary.cell(r_idx, 2, value).border = BORDER

    total_cost = sum(float(v or 0) for v in usage_df.get("estimated_cost_usd", pd.Series(dtype=float)).tolist()) if not usage_df.empty else 0
    total_tokens = sum(int(v or 0) for v in usage_df.get("total_tokens", pd.Series(dtype=int)).tolist()) if not usage_df.empty else 0
    summary["A13"] = "Usage estimate"
    summary["A13"].font = Font(bold=True, size=12)
    summary["A14"] = "Total tokens reported"
    summary["B14"] = total_tokens
    summary["A15"] = "Estimated model cost (USD)"
    summary["B15"] = round(total_cost, 6)
    summary["A16"] = "Note"
    summary["B16"] = "Estimate only; verify against the provider billing console."
    summary.column_dimensions["A"].width = 34
    summary.column_dimensions["B"].width = 58

    ws_claims = wb.create_sheet("QA Results")
    _write_dataframe(ws_claims, claims_df, {
        "Claim": 55, "Verdict": 60, "Supported Claims": 50,
        "Unsupported Claims": 50, "Unverified Claims": 50, "Direct Verbatims": 60,
        "Evidence by Source": 60, "Suggested Fix": 60, "URLs": 55,
    })

    ws_sources = wb.create_sheet("QA Sources")
    _write_dataframe(ws_sources, sources_df, {"url": 65, "Slide Title": 45, "raw_status": 32})

    ws_usage = wb.create_sheet("API Usage")
    _write_dataframe(ws_usage, usage_df)

    output = io.BytesIO()
    wb.save(output)
    return output.getvalue()
