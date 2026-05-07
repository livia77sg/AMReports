"""Core processor: Track360 raw report → filled-out conversion templates."""

from __future__ import annotations

import io
import re
import shutil
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

import config


@dataclass
class ProcessResult:
    """Summary of a processing run."""
    output_files: dict[str, bytes] = field(default_factory=dict)  # filename -> content
    report_date: str = ""
    total_signups: int = 0
    total_ftds: int = 0
    per_output_counts: dict[str, dict] = field(default_factory=dict)  # key -> {signups, ftd_lines}
    warnings: list[str] = field(default_factory=list)
    unmapped_brands: list[str] = field(default_factory=list)


def _normalize_brand_fallback(raw: str) -> str:
    """Fallback rule when brand isn't in brand_map.csv:
    lowercase, strip common geo/category suffixes, remove dashes/spaces.
    Example: 'Betfred-UK-Casino' -> 'betfred'.
    """
    if not raw:
        return ""
    s = raw.strip()
    # Strip trailing region/category fragments
    s = re.sub(r"-(UK|GB|US|EU|DE|IT|BR|AU|NL)-?(Casino|Bingo|Sport|Sports)?$", "", s, flags=re.IGNORECASE)
    s = re.sub(r"-(Casino|Bingo|Sport|Sports)-(UK|GB|US|EU|DE|IT|BR|AU|NL)$", "", s, flags=re.IGNORECASE)
    s = s.lower().replace("-", "").replace(" ", "")
    return s


def _detect_category_fallback(raw: str) -> str:
    """Fallback category detection from raw brand name."""
    s = (raw or "").lower()
    if "bingo" in s:
        return "bingo"
    if "casino" in s:
        return "casino"
    return "sport"  # default


def load_brand_map(csv_path: Path) -> dict[str, tuple[str, str]]:
    """Load brand_map.csv → {report_name: (clean_name, category)}.
    Lookup is case-insensitive."""
    if not csv_path.exists():
        return {}
    df = pd.read_csv(csv_path)
    return {
        str(row["report_brand_name"]).strip().lower(): (
            str(row["clean_name"]).strip(),
            str(row["category"]).strip().lower(),
        )
        for _, row in df.iterrows()
    }


def resolve_brand(raw: str, brand_map: dict[str, tuple[str, str]]) -> tuple[str, str, bool]:
    """Return (clean_name, category, was_mapped).
    Falls back to heuristics if not in map."""
    if pd.isna(raw):
        return ("", "sport", False)
    key = str(raw).strip().lower()
    if key in brand_map:
        clean, cat = brand_map[key]
        return (clean, cat, True)
    return (_normalize_brand_fallback(str(raw)), _detect_category_fallback(str(raw)), False)


def _detect_report_date(df: pd.DataFrame) -> tuple[str, list[str]]:
    """Pick the most common date across Signup/FTD date columns.
    Returns (YYYY-MM-DD string, warnings)."""
    warnings = []
    dates = []
    for col in (config.COL_SIGNUP_DATE, config.COL_FTD_DATE):
        if col in df.columns:
            dates.extend(pd.to_datetime(df[col], errors="coerce").dropna().dt.date.tolist())
    if not dates:
        # Fallback: today (shouldn't happen with real reports)
        warnings.append("Nuk u gjet asnjë datë në raport — u përdor data e sotme.")
        return (datetime.now().strftime("%Y-%m-%d"), warnings)
    counter = Counter(dates)
    most_common, count = counter.most_common(1)[0]
    if len(counter) > 1:
        others = ", ".join(str(d) for d, _ in counter.most_common()[1:5])
        warnings.append(
            f"Raporti përmban data të ndryshme. U përdor {most_common} ({count} rreshta). "
            f"Datat e tjera: {others}."
        )
    return (most_common.strftime("%Y-%m-%d"), warnings)


def _build_rows_for_output(
    df: pd.DataFrame,
    output_cfg: dict,
    report_date: str,
    brand_map: dict,
    unmapped_brands: set[str],
) -> list[list]:
    """Build the rows (Cid, ConversionName, Time, Value, Currency) to append
    to the template for one output (sport_google, bingo_bing, etc.).
    """
    site_ids = output_cfg["site_ids"]
    types = config.CHANNEL_TYPES[output_cfg["channel"]]

    sub = df[
        df[config.COL_SITE_ID].isin(site_ids)
        & df[config.COL_TYPE].astype(str).str.lower().isin([t.lower() for t in types])
    ].copy()

    if sub.empty:
        return []

    # Strip Bing prefix from Cid
    if output_cfg["channel"] == "bing":
        sub[config.COL_CID] = sub[config.COL_CID].astype(str).apply(
            lambda c: c[len(config.BING_CID_PREFIX):] if c.startswith(config.BING_CID_PREFIX) else c
        )

    collapse_casino = output_cfg["key"] in config.COLLAPSE_CASINO_OUTPUTS
    signup_dt = f"{report_date} {config.SIGNUP_TIME}"
    ftd_dt = f"{report_date} {config.FTD_TIME}"
    rows: list[list] = []

    # === SIGNUP rows ===
    signup_df = sub[sub[config.COL_SIGNUPS].fillna(0) > 0]
    for _, r in signup_df.iterrows():
        cid = r[config.COL_CID]
        if pd.isna(cid) or str(cid).strip().lower() in ("na", "nan", ""):
            continue
        # If a row has signups > 1, emit one row per signup (matches manual workflow
        # where the AM clicks into the pivot and pulls one Cid per signup)
        n = int(r[config.COL_SIGNUPS])
        for _ in range(n):
            rows.append([str(cid), output_cfg["signup_label"], signup_dt, 1, config.CURRENCY])

    # === FTD rows ===
    ftd_df = sub[sub[config.COL_FTDS].fillna(0) > 0]
    for _, r in ftd_df.iterrows():
        cid = r[config.COL_CID]
        if pd.isna(cid) or str(cid).strip().lower() in ("na", "nan", ""):
            continue
        n_ftds = int(r[config.COL_FTDS])
        revenue = r[config.COL_REVENUES]
        revenue_int = int(round(float(revenue))) if pd.notna(revenue) else 0
        # Per-FTD revenue split: if a row has multiple FTDs, divide and round
        per_ftd_rev = int(round(revenue_int / n_ftds)) if n_ftds > 0 else revenue_int

        clean, category, mapped = resolve_brand(r[config.COL_BRAND], brand_map)
        if not mapped and pd.notna(r[config.COL_BRAND]):
            unmapped_brands.add(str(r[config.COL_BRAND]))

        for _ in range(n_ftds):
            if category == "casino" and collapse_casino:
                # Single row labeled as the output's casino_label
                rows.append([str(cid), output_cfg["casino_label"], ftd_dt, per_ftd_rev, config.CURRENCY])
            else:
                # Two-row treatment: per-brand row + generic ftds row
                rows.append([str(cid), f"{output_cfg['brand_prefix']}{clean}", ftd_dt, per_ftd_rev, config.CURRENCY])
                rows.append([str(cid), output_cfg["ftd_label"], ftd_dt, per_ftd_rev, config.CURRENCY])

    return rows


def _find_header_row(ws) -> int:
    """Locate the row that starts with 'Microsoft Click ID' or 'Google Click ID'.
    Data rows go directly below it. Returns 1-indexed row number."""
    for row_idx in range(1, min(ws.max_row + 1, 20)):
        cell = ws.cell(row=row_idx, column=1).value
        if cell and isinstance(cell, str) and "Click ID" in cell:
            return row_idx
    raise ValueError("Could not find header row in template.")


def _write_template(template_path: Path, rows: list[list], output_path: Path):
    """Open template, clear existing example data rows, write new rows, save."""
    wb = load_workbook(template_path)
    ws = wb.active
    header_row = _find_header_row(ws)
    first_data_row = header_row + 1

    # Clear existing example data rows (templates ship with sample data)
    last_existing = ws.max_row
    for r in range(first_data_row, last_existing + 1):
        for c in range(1, 6):
            ws.cell(row=r, column=c).value = None

    # Write new rows
    for i, row_data in enumerate(rows):
        for c, val in enumerate(row_data, start=1):
            ws.cell(row=first_data_row + i, column=c).value = val

    wb.save(output_path)


def process_report(
    report_bytes: bytes,
    templates_dir: Path,
    brand_map_path: Path,
) -> ProcessResult:
    """Main entry point. Read raw Track360 report bytes, return ProcessResult
    with all output files in memory."""
    result = ProcessResult()
    df = pd.read_excel(io.BytesIO(report_bytes))

    # Validate columns
    missing = [c for c in config.REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"Mungojnë kolonat e nevojshme në raport: {missing}")

    # Drop Cid == "na" (case-insensitive)
    cid_str = df[config.COL_CID].astype(str).str.strip().str.lower()
    pre = len(df)
    df = df[~cid_str.isin(["na", "nan", ""])].copy()
    skipped = pre - len(df)
    if skipped:
        result.warnings.append(f"U përjashtuan {skipped} rreshta me Cid = 'na'.")

    # Detect report date
    report_date, date_warnings = _detect_report_date(df)
    result.report_date = report_date
    result.warnings.extend(date_warnings)

    # Load brand map
    brand_map = load_brand_map(brand_map_path)
    unmapped: set[str] = set()

    # Process each output
    import tempfile
    with tempfile.TemporaryDirectory() as tmpdir:
        tmp = Path(tmpdir)
        for output_cfg in config.OUTPUTS:
            template_path = templates_dir / output_cfg["template"]
            if not template_path.exists():
                result.warnings.append(f"Mungon template: {template_path.name}")
                continue

            rows = _build_rows_for_output(df, output_cfg, report_date, brand_map, unmapped)

            # Count for summary
            sig_count = sum(1 for r in rows if "signup" in r[1].lower())
            ftd_count = len(rows) - sig_count
            result.per_output_counts[output_cfg["key"]] = {
                "signups": sig_count,
                "ftd_rows": ftd_count,
            }
            result.total_signups += sig_count
            # Each non-collapsed FTD writes 2 rows; collapsed writes 1 — but we just
            # report the row count to the user, which is what they care about.
            result.total_ftds += ftd_count

            # Always produce the file even if empty (so she sees that nothing came through)
            out_filename = f"{output_cfg['output_prefix']}_{report_date}.xlsx"
            out_path = tmp / out_filename
            _write_template(template_path, rows, out_path)
            result.output_files[out_filename] = out_path.read_bytes()

    result.unmapped_brands = sorted(unmapped)
    if unmapped:
        result.warnings.append(
            f"{len(unmapped)} brand pa hartë (u përdor rregulli i parazgjedhur): "
            f"{', '.join(sorted(unmapped))}. "
            f"Shtoji në brand_map.csv për kontroll të plotë."
        )

    return result
