r"""
Batch forecast service.

Reads (Part Number, Tier 1) pairs from an uploaded Excel, runs the
forecast engine for each unique combination, and writes a result
workbook with one sheet per forecast month (12 sheets, or 13 when the
current month is included) plus a Summary.

COLUMN LAYOUT (per monthly sheet)
──────────────────────────────────
Part Number | Tier 1 | Weight (lbs) | Base Price ($) |
Quarter (Current) | Quarter (Previous) |
MC_Q ($/lb) | MC_Q-1 ($/lb) |
PPI_Q | PPI_Q-1 | PPI Factor |
CNG_Q ($/lb) | CNG_Q-1 ($/lb) |
AMS_Q ($/lb) | AMS_Q-1 ($/lb) | AMS Delta ($/lb) |
DF_c | Predicted Price ($) | Predicted Price (without deadband) ($) | Total AOP
"""

import io
import logging
from typing import Optional

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from app.models.response import ForecastResponse
from app.services.forecast_engine import ForecastEngine

# ======================================================================
# NEW - AOP
# This is the new feature as per thapase/srikanth requirements on 10/20/2026
# Importing dependencies for calling AOP calculation methods
# Importing dependencies for ppi and lme prediction
from app.services.forecast_engine import ForecastEngine, compute_total_aop   # NEW - AOP
from app.data.excel_store import _load_aop_lookup                            # NEW - AOP
# For ppi forecast
from app.services.ppi_forecast import refresh_ppi_forecast            # NEW - PPI
from app.services.lme_forecast import refresh_lme_forecast            # NEW - LME
# ======================================================================

logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────────────────────────────────────
# STYLING
# ─────────────────────────────────────────────────────────────────────────────

_THIN        = Side(style="thin", color="BFBFBF")
_BORDER      = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
_ALT_FILL    = PatternFill("solid", fgColor="D6E4F0")
_WHITE_FILL  = PatternFill("solid", fgColor="FFFFFF")
_WARN_FILL   = PatternFill("solid", fgColor="FCE4D6")
_HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
_DATA_FONT   = Font(name="Arial", size=10)
_WARN_FONT   = Font(name="Arial", size=10, color="C00000")
_CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
_LEFT        = Alignment(horizontal="left",   vertical="center")


# ─────────────────────────────────────────────────────────────────────────────
# COLUMN DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────
# ── MODIFICATION ────────────────────────────────────────
# In this adde rounded off values for alloy prices, ppi and unit price index in the output excel sheet
# This is the new feature as per thapase/srikanth requirements on 09/30/2026
def _build_columns() -> list[tuple[str, str, callable]]:
    def _ctx(mf): return mf.quarter_context
    return [
        ("Part Number",          "@",             lambda pn, t1, fr, mf: pn),
        ("Tier 1",               "@",             lambda pn, t1, fr, mf: t1),
        ("Weight\n(lbs)",        "0.00",          lambda pn, t1, fr, mf: fr.pwt_lbs),
        ("Base Price\n($)",      "$#,##0.00",   lambda pn, t1, fr, mf: round(mf.base_price_used, 2)),
        ("Quarter\n(Current)",   "@",             lambda pn, t1, fr, mf: _ctx(mf).quarter_label),
        ("Quarter\n(Previous)",  "@",             lambda pn, t1, fr, mf: _ctx(mf).prev_quarter_label),
        ("MC_Q\n($/lb)",         "$0.00",     lambda pn, t1, fr, mf: round(_ctx(mf).mc_q, 2)),
        ("MC_Q-1\n($/lb)",       "$0.00",     lambda pn, t1, fr, mf: round(_ctx(mf).mc_q_1, 2)),
        ("PPI_Q",                "0.0",         lambda pn, t1, fr, mf: round(_ctx(mf).ppi_q, 1)),
        ("PPI_Q-1",              "0.0",         lambda pn, t1, fr, mf: round(_ctx(mf).ppi_q_1, 1)),
        ("PPI Factor",           "0.0",      lambda pn, t1, fr, mf: round(_ctx(mf).ppi_factor, 1)),
        ("CNG_Q\n($/lb)",        "General",       lambda pn, t1, fr, mf: _ctx(mf).cng_q),
        ("CNG_Q-1\n($/lb)",      "General",       lambda pn, t1, fr, mf: _ctx(mf).cng_q_1),
        ("AMS_Q\n($/lb)",        "$0.00",     lambda pn, t1, fr, mf: round(_ctx(mf).ams_q, 2)),
        ("AMS_Q-1\n($/lb)",      "$0.00",     lambda pn, t1, fr, mf: round(_ctx(mf).ams_q_1, 2)),
        ("AMS Delta\n($/lb)",    "$0.00",     lambda pn, t1, fr, mf: round(_ctx(mf).ams_delta, 2)),
        ("DF_c",                 "General",          lambda pn, t1, fr, mf: mf.df_c),
        # ("Predicted Price\n($)", "$#,##0.00",   lambda pn, t1, fr, mf: round(mf.predicted_price, 2)),
        # ("Predicted Price\n(without deadband)", "$#,##0.00", lambda pn, t1, fr, mf: round(mf.predicted_price_without_deadband, 2)),  # NEW
        # ── MODIFICATION ────────────────────────────────────────
        ("Predicted Price", "$#,##0.00", lambda pn, t1, fr, mf: round(mf.predicted_price_without_deadband, 2)),  # NEW
        # Changing the column name here from Predicted Price to Final Unit Price (with deadband)
        ("Final Unit Price (with deadband)\n($)", "$#,##0.00",   lambda pn, t1, fr, mf: round(mf.predicted_price, 2)),

    ]


COLUMNS = _build_columns()
N_COLS  = len(COLUMNS)


# ─────────────────────────────────────────────────────────────────────────────
# CELL HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _write_header(ws, col_idx: int, label: str) -> None:
    cell = ws.cell(row=1, column=col_idx, value=label)
    cell.font = _HEADER_FONT; cell.fill = _HEADER_FILL
    cell.alignment = _CENTER; cell.border = _BORDER


def _write_data(ws, row: int, col: int, value, fmt: str, fill, left=False) -> None:
    cell = ws.cell(row=row, column=col, value=value)
    cell.font = _DATA_FONT; cell.fill = fill; cell.border = _BORDER
    cell.number_format = fmt
    cell.alignment = _LEFT if left else _CENTER


def _write_error_row(ws, row: int, part_number: str, tier_1: str, error_msg: str, fill) -> None:
    for col, val in [(1, part_number), (2, tier_1)]:
        c = ws.cell(row=row, column=col, value=val)
        c.font = _WARN_FONT; c.fill = _WARN_FILL
        c.border = _BORDER; c.alignment = _LEFT

    err = ws.cell(row=row, column=3, value=f"ERROR: {error_msg}")
    err.font = _WARN_FONT; err.fill = _WARN_FILL
    err.border = _BORDER; err.alignment = _LEFT
    ws.merge_cells(start_row=row, start_column=3, end_row=row, end_column=N_COLS)


def _set_col_widths(ws) -> None:
    widths = [22, 22, 10, 14, 14, 14, 13, 13, 10, 10, 13, 13, 13, 13, 13, 13, 8, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 36


# ─────────────────────────────────────────────────────────────────────────────
# INPUT READER
# ─────────────────────────────────────────────────────────────────────────────

def read_parts_from_upload(file_bytes: bytes) -> list[tuple[str, str]]:
    """
    Read (Part Number, Tier 1) pairs from the uploaded Excel file.

    Expects columns named 'Part Number' and 'Tier 1' (case-insensitive).
    Returns a deduplicated list of (part_number, tier_1) tuples,
    preserving input order.
    """
    df = pd.read_excel(
        io.BytesIO(file_bytes),
        engine="openpyxl",
        dtype=str,
        keep_default_na=False,
        na_values=[""],
    )

    cols_lower = {c.strip().lower(): c for c in df.columns}

    if "part number" not in cols_lower:
        raise ValueError(
            f"Input Excel must have a 'Part Number' column. "
            f"Found: {list(df.columns)}"
        )
    if "tier 1" not in cols_lower:
        raise ValueError(
            f"Input Excel must have a 'Tier 1' column. "
            f"Found: {list(df.columns)}"
        )

    pn_col = cols_lower["part number"]
    t1_col = cols_lower["tier 1"]

    seen: set[tuple[str, str]] = set()
    pairs: list[tuple[str, str]] = []
    for _, row in df.iterrows():
        pn = str(row[pn_col]).strip()
        t1 = str(row[t1_col]).strip()
        if pn and t1 and (pn, t1) not in seen:
            seen.add((pn, t1))
            pairs.append((pn, t1))

    logger.info("Read %d unique (part, tier_1) pairs from uploaded file", len(pairs))
    return pairs


# CNG column name aliases (case-insensitive, matched against stripped headers).
_CNG_Q_ALIASES  = ("cng_q", "cng q", "cng_q current", "cng current", "cng - current qtr")
_CNG_Q1_ALIASES = ("cng_q-1", "cng_q1", "cng q-1", "cng_q_1", "cng previous", "cng - previous qtr")


def read_parts_with_cng_from_upload(file_bytes: bytes) -> list[dict]:
    """
    Read (Part Number, Tier 1, CNG_Q, CNG_Q-1) rows from the uploaded Excel.

    The CNG values are taken straight from the file — the caller no longer
    supplies them. Expects columns 'Part Number', 'Tier 1', 'CNG_Q' and
    'CNG_Q-1' (case-insensitive; a few common spellings are accepted).
    Returns a deduplicated list of dicts, preserving input order:
        {"part_number", "tier_1", "cng_q", "cng_q_1"}
    """
    df = pd.read_excel(
        io.BytesIO(file_bytes),
        engine="openpyxl",
        dtype=str,
        keep_default_na=False,
        na_values=[""],
    )

    cols_lower = {c.strip().lower(): c for c in df.columns}

    if "part number" not in cols_lower:
        raise ValueError(
            f"Input Excel must have a 'Part Number' column. Found: {list(df.columns)}"
        )
    if "tier 1" not in cols_lower:
        raise ValueError(
            f"Input Excel must have a 'Tier 1' column. Found: {list(df.columns)}"
        )

    def _find(aliases: tuple[str, ...]) -> Optional[str]:
        for a in aliases:
            if a in cols_lower:
                return cols_lower[a]
        return None

    cng_q_col  = _find(_CNG_Q_ALIASES)
    cng_q1_col = _find(_CNG_Q1_ALIASES)
    if cng_q_col is None or cng_q1_col is None:
        raise ValueError(
            "Input Excel must have 'CNG_Q' and 'CNG_Q-1' columns "
            f"(CNG is now read from the file). Found: {list(df.columns)}"
        )

    pn_col = cols_lower["part number"]
    t1_col = cols_lower["tier 1"]

    seen: set[tuple[str, str]] = set()
    rows: list[dict] = []
    for _, row in df.iterrows():
        pn = str(row[pn_col]).strip()
        t1 = str(row[t1_col]).strip()
        if not (pn and t1) or (pn, t1) in seen:
            continue
        try:
            cng_q   = float(row[cng_q_col])
            cng_q_1 = float(row[cng_q1_col])
        except (ValueError, TypeError):
            raise ValueError(
                f"Non-numeric CNG value for part '{pn}' / '{t1}'. "
                "CNG_Q and CNG_Q-1 must be numbers."
            )
        seen.add((pn, t1))
        rows.append({"part_number": pn, "tier_1": t1, "cng_q": cng_q, "cng_q_1": cng_q_1})

    logger.info("Read %d unique (part, tier_1, cng) rows from uploaded file", len(rows))
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def _run_forecasts_uniform(
    part_tier_pairs: list[tuple[str, str]],
    engine: ForecastEngine,
    cng_q: float,
    cng_q_1: float,
    include_current_month: bool,
) -> dict[tuple[str, str], "ForecastResponse | Exception"]:
    """Run the engine for every pair using ONE shared CNG_Q / CNG_Q-1."""
    # ======================================================================
    # NEW - PPI
    # calling for ppi forecast
    # This is the new feature as per thapase/srikanth requirements on 10/21/2026
    refresh_ppi_forecast()                          # NEW - PPI: latest predicted PPI before forecasting
    refresh_lme_forecast()                          # NEW - LME: latest predicted LME price before forecasting
    # ======================================================================

    results: dict[tuple[str, str], ForecastResponse | Exception] = {}
    for pn, t1 in part_tier_pairs:
        try:
            results[(pn, t1)] = engine.forecast(
                part_number=pn, tier_1=t1, cng_q=cng_q, cng_q_1=cng_q_1,
                include_current_month=include_current_month,
            )
            logger.debug("Forecast OK: %s / %s", pn, t1)
        except Exception as exc:
            results[(pn, t1)] = exc
            logger.warning("Forecast failed for %s / %s: %s", pn, t1, exc)
    return results


def _run_forecasts_per_row(
    part_rows: list[dict],
    engine: ForecastEngine,
    include_current_month: bool,
) -> tuple[list[tuple[str, str]], dict[tuple[str, str], "ForecastResponse | Exception"]]:
    """Run the engine per row, each with its OWN CNG_Q / CNG_Q-1 from the file."""
    results: dict[tuple[str, str], ForecastResponse | Exception] = {}
    # ======================================================================
    # NEW - PPI
    # calling for ppi forecast
    # This is the new feature as per thapase/srikanth requirements on 10/21/2026
    refresh_ppi_forecast()                          # NEW - PPI: latest predicted PPI before forecasting
    refresh_lme_forecast()                          # NEW - LME: latest predicted LME price before forecasting
    # ======================================================================
    part_tier_pairs: list[tuple[str, str]] = []
    for row in part_rows:
        key = (row["part_number"], row["tier_1"])
        part_tier_pairs.append(key)
        try:
            results[key] = engine.forecast(
                part_number=key[0], tier_1=key[1],
                cng_q=row["cng_q"], cng_q_1=row["cng_q_1"],
                include_current_month=include_current_month,
            )
            logger.debug("Forecast OK: %s / %s (CNG_Q=%s)", key[0], key[1], row["cng_q"])
        except Exception as exc:
            results[key] = exc
            logger.warning("Forecast failed for %s / %s: %s", key[0], key[1], exc)
    return part_tier_pairs, results


def build_forecast_workbook(
    part_tier_pairs: list[tuple[str, str]],
    engine: ForecastEngine,
    cng_q: float,
    cng_q_1: float,
    include_current_month: bool = False,
) -> bytes:
    """
    Run forecasts for all (part_number, tier_1) pairs with a single shared
    CNG pair and build the output workbook. Kept for backward compatibility.
    """
    results = _run_forecasts_uniform(
        part_tier_pairs, engine, cng_q, cng_q_1, include_current_month
    )

    # ======================================================================
    # NEW - AOP
    return _build_workbook_from_results(part_tier_pairs, results, _load_aop_lookup()) # CHANGED - AOP 
    # ======================================================================


def build_forecast_workbook_from_rows(
    part_rows: list[dict],
    engine: ForecastEngine,
    include_current_month: bool = False,
) -> bytes:
    """Run forecasts using per-row CNG (from the file) and build the workbook."""
    part_tier_pairs, results = _run_forecasts_per_row(
        part_rows, engine, include_current_month
    )
    # ======================================================================
    # NEW - AOP
    return _build_workbook_from_results(part_tier_pairs, results, _load_aop_lookup())   # CHANGED - AOP)
    # ======================================================================


def build_preview_from_rows(
    part_rows: list[dict],
    engine: ForecastEngine,
    include_current_month: bool = False,
) -> dict:
    """
    Run forecasts using per-row CNG and return a JSON-serialisable preview
    (Summary + one entry per monthly sheet) for on-screen display.
    """
    part_tier_pairs, results = _run_forecasts_per_row(
        part_rows, engine, include_current_month
    )

    # NEW - AOP, same source the Excel workbook uses
    aop_lookup = _load_aop_lookup()   

    month_labels: list[tuple[str, str]] = []
    for r in results.values():
        if isinstance(r, ForecastResponse):
            month_labels = [(f.year_month, f.month_label) for f in r.forecasts]
            break
    if not month_labels:
        first_error = next(iter(results.values()), None)
        raise ValueError(
            f"All parts failed forecasting — nothing to preview. First error: {first_error}"
        )

    n_ok = sum(1 for r in results.values() if isinstance(r, ForecastResponse))
    n_failed = len(results) - n_ok

    summary: list[dict] = []
    for pn, t1 in part_tier_pairs:
        r = results[(pn, t1)]
        if isinstance(r, Exception):
            summary.append({"part_number": pn, "tier_1": t1, "error": str(r)})
            continue
        monthly = []
        for ym, _ in month_labels:
            mf = next((f for f in r.forecasts if f.year_month == ym), None)
            monthly.append(mf.predicted_price if mf else None)
        ctx0 = r.forecasts[0].quarter_context
        summary.append({
            "part_number": pn, "tier_1": t1,
            "weight_lbs": r.pwt_lbs, "base_price": r.base_price,
            "cng_q": ctx0.cng_q, "cng_q_1": ctx0.cng_q_1,
            "monthly": monthly, "error": None,
        })

    # ── MODIFICATION ────────────────────────────────────────
    # ── Base Price Summary (on-screen) ─────────────────────────────────
    # Mirrors the "Base Price Summary" Excel sheet: base price used (P_current) per month
    base_price_summary: list[dict] = []
    for pn, t1 in part_tier_pairs:
        r = results[(pn, t1)]
        if isinstance(r, Exception):
            base_price_summary.append({"part_number": pn, "tier_1": t1, "error": str(r)})
            continue
        monthly_base = []
        for ym, _ in month_labels:
            mf = next((f for f in r.forecasts if f.year_month == ym), None)
            monthly_base.append(round(mf.base_price_used, 2) if mf else None)
        base_price_summary.append({
            "part_number": pn, "tier_1": t1,
            "weight_lbs": r.pwt_lbs,
            "monthly": monthly_base,          # same order as "months"
            "error": None,
        })

    # NEW - AOP ===================================
    # Total AOP Summary (on-screen)
    # Mirrors the "Total AOP Summary" Excel sheet: Total AOP ($) per part per month
    aop_summary: list[dict] = []
    aop_grand_total = 0.0
    if aop_lookup:
        for pn, t1 in part_tier_pairs:
            r = results[(pn, t1)]
            if isinstance(r, Exception):
                aop_summary.append({"part_number": pn, "tier_1": t1, "error": str(r)})
                continue
            _pn, _t1 = pn.strip().lower(), t1.strip().lower()
            monthly_aop = []
            part_total = 0.0
            has_value = False
            for ym, _ in month_labels:
                mf = next((f for f in r.forecasts if f.year_month == ym), None)
                month_total = None
                if mf is not None:
                    aop_qty = aop_lookup.get((_pn, _t1, ym))
                    if aop_qty is None:
                        aop_qty = aop_lookup.get((_pn, "", ym))
                    month_total = compute_total_aop(mf.predicted_price, aop_qty) # Modified
                    #month_total = compute_total_aop(mf.base_price_used, aop_qty)
                    if month_total is not None:
                        part_total += month_total
                        has_value = True
                monthly_aop.append(round(month_total, 2) if month_total is not None else None)
            if has_value:
                aop_grand_total += part_total
            aop_summary.append({
                "part_number": pn, "tier_1": t1,
                "monthly": monthly_aop,
                "total_all_months": round(part_total, 2) if has_value else None,
                "error": None,
            })
    # ==============================================

    monthly_sheets: list[dict] = []
    for ym, label in month_labels:
        rows: list[dict] = []
        for pn, t1 in part_tier_pairs:
            r = results[(pn, t1)]
            if isinstance(r, Exception):
                rows.append({"part_number": pn, "tier_1": t1, "error": str(r)})
                continue
            mf = next((f for f in r.forecasts if f.year_month == ym), None)
            if mf is None:
                rows.append({"part_number": pn, "tier_1": t1, "error": f"No data for {ym}"})
                continue
            q = mf.quarter_context
            # NEW - AOP ===================================
            total_aop = None
            if aop_lookup:
                _pn, _t1 = pn.strip().lower(), t1.strip().lower()
                aop_qty = aop_lookup.get((_pn, _t1, ym))
                if aop_qty is None:
                    aop_qty = aop_lookup.get((_pn, "", ym))
                total_aop = compute_total_aop(mf.predicted_price, aop_qty)
            # ==============================================
            rows.append({
                "part_number": pn, "tier_1": t1,
                "weight_lbs": r.pwt_lbs, "base_price_used": mf.base_price_used,
                "quarter": q.quarter_label,
                "mc_q": q.mc_q, "mc_q_1": q.mc_q_1,
                "ppi_q": q.ppi_q, "ppi_q_1": q.ppi_q_1, "ppi_factor": q.ppi_factor,
                "cng_q": q.cng_q, "cng_q_1": q.cng_q_1,
                "ams_q": q.ams_q, "ams_q_1": q.ams_q_1, "ams_delta": q.ams_delta,
                "predicted_price": mf.predicted_price, 
                "predicted_price_without_deadband": mf.predicted_price_without_deadband,
                # NEW - AOP ===================================
                "total_aop": round(total_aop, 2) if total_aop is not None else None,   # NEW - AOP
                # NEW - AOP ===================================              
                "error": None,

            })
        monthly_sheets.append({"year_month": ym, "label": label, "rows": rows})
    print(f"AOP GRAND TOTAL VALUE: {round(aop_grand_total, 2)}")
    return {
        "part_count": len(part_tier_pairs),
        "ok": n_ok,
        "failed": n_failed,
        "months": [{"year_month": ym, "month_label": label} for ym, label in month_labels],
        "summary": summary,
        "base_price_summary": base_price_summary,
        # NEW - AOP ===================================
        "aop_summary": aop_summary,                 # NEW - AOP
        "aop_grand_total": round(aop_grand_total, 2) if aop_lookup else None,  # NEW - AOP
        # =============================================
        "monthly_sheets": monthly_sheets,
    }


def _build_workbook_from_results(
    part_tier_pairs: list[tuple[str, str]],
    results: dict[tuple[str, str], "ForecastResponse | Exception"],
    aop_lookup: Optional[dict] = None,     # NEW - AOP (new optional parameter,
) -> bytes:
    """
    Build the formatted .xlsx workbook from already-computed results.

    include_current_month affects the number of monthly sheets via the
    supplied results. Failed rows show an ERROR message instead of stopping
    the batch. Returns raw bytes of the generated .xlsx workbook.
    """
    # ── Step 2: determine month labels from first successful result ───────
    month_labels: list[tuple[str, str]] = []
    for r in results.values():
        if isinstance(r, ForecastResponse):
            month_labels = [(f.year_month, f.month_label) for f in r.forecasts]
            break

    if not month_labels:
        first_error = next(iter(results.values()), None)
        raise ValueError(
            "All parts failed forecasting — cannot generate output workbook. "
            f"First error: {first_error}"
        )

    # ── Step 3: build workbook ────────────────────────────────────────────
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    for year_month, month_label in month_labels:
        ws = wb.create_sheet(title=month_label[:31])
        ws.freeze_panes = "A2"

        for ci, (label, _, _) in enumerate(COLUMNS, 1):
            _write_header(ws, ci, label)

        # ======================================================================
        # NEW - AOP
        # To save in Excel File
        if aop_lookup:  # NEW - AOP (a) header
            _write_header(ws, N_COLS + 1, "Total AOP\n($)")
        # ======================================================================


        for ri, (pn, t1) in enumerate(part_tier_pairs, 2):
            fill   = _ALT_FILL if ri % 2 == 0 else _WHITE_FILL
            result = results[(pn, t1)]

            if isinstance(result, Exception):
                _write_error_row(ws, ri, pn, t1, str(result), fill)
                continue

            mf = next((f for f in result.forecasts if f.year_month == year_month), None)
            if mf is None:
                _write_error_row(ws, ri, pn, t1, f"No forecast data for {year_month}", fill)
                continue

            for ci, (_, fmt, extractor) in enumerate(COLUMNS, 1):
                left = ci in (1, 2)
                try:
                    value = extractor(pn, t1, result, mf)
                except Exception as ex:
                    value = f"ERR: {ex}"
                _write_data(ws, ri, ci, value, fmt, fill, left=left)
            
            # ======================================================================
            # NEW - AOP
            # To Save in Excel File
            # This is the new feature as per thapase/srikanth requirements on 10/20/2026
            if aop_lookup:                                   # NEW - AOP (b) value 
                #aop_qty = aop_lookup.get((pn, year_month))
                # NEW
                _pn, _t1 = pn.strip().lower(), t1.strip().lower()
                aop_qty = aop_lookup.get((_pn, _t1, year_month))      # tier-specific AOP first
                if aop_qty is None:
                    aop_qty = aop_lookup.get((_pn, '', year_month))   # else item-level AOP (no tier)
                # _write_data(ws, ri, N_COLS + 1,
                #             compute_total_aop(mf.predicted_price, aop_qty),
                #             "$#,##0.00", fill)
                # NEW-AOP(qty*base_price_summary_sheet)
                # This is modified as per the thapase requirement on 10/26/2026
                _write_data(ws, ri, N_COLS + 1, 
                            compute_total_aop(round(mf.base_price_used, 2), aop_qty), 
                            "$#,##0.00", fill)
            # ======================================================================


        _set_col_widths(ws)

        # ======================================================================
        # NEW - AOP
        if aop_lookup:                        
            ws.column_dimensions[get_column_letter(N_COLS + 1)].width = 18
        # ======================================================================


    # ── Step 4: summary sheet ─────────────────────────────────────────────
    _build_summary_sheet(wb, part_tier_pairs, results, month_labels)


    # ── MODIFICATION ────────────────────────────────────────
    # This creates a new sheet "Base Price Summary" that shows the base price used for each month
    # This is the new feature as per thapase/srikanth requirements on 09/30/2026
    _build_base_price_summary_sheet(wb, part_tier_pairs, results, month_labels)

    # NEW - AOP 
    # sheet "Total AOP Summary" (part x month)
    # This is the new feature as per thapase/srikanth requirements on 10/22/2026
    if aop_lookup:
        _build_total_aop_summary_sheet(wb, part_tier_pairs, results, month_labels, aop_lookup)




    # ── Step 5: serialise ─────────────────────────────────────────────────
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def _build_summary_sheet(
    wb, part_tier_pairs, results, month_labels
) -> None:
    ws = wb.create_sheet(title="Summary", index=0)
    ws.freeze_panes = "E2"

    fixed   = ["Part Number", "Tier 1", "Weight (lbs)", "Base Price ($)"]
    monthly = [label for _, label in month_labels]

    for ci, h in enumerate(fixed + monthly, 1):
        _write_header(ws, ci, h)
    ws.row_dimensions[1].height = 36

    for ri, (pn, t1) in enumerate(part_tier_pairs, 2):
        fill   = _ALT_FILL if ri % 2 == 0 else _WHITE_FILL
        result = results[(pn, t1)]

        if isinstance(result, Exception):
            _write_error_row(ws, ri, pn, t1, str(result), fill)
            continue

        _write_data(ws, ri, 1, pn,               "@",            fill, left=True)
        _write_data(ws, ri, 2, t1,               "@",            fill, left=True)
        _write_data(ws, ri, 3, result.pwt_lbs,   "0.00",         fill)
        _write_data(ws, ri, 4, round(result.base_price, 2), "$#,##0.00", fill) # # change to "$#,##0.00" and result.base_price to round(result.base_price,2)

        for mi, (year_month, _) in enumerate(month_labels):
            mf  = next((f for f in result.forecasts if f.year_month == year_month), None)
            val = round(mf.predicted_price, 2) if mf else "N/A"   ##### mf.predicted_price to round(mf.predicted_price, 2)
            _write_data(ws, ri, 5 + mi, val, "$#,##0.00", fill)

    # Column widths
    for col, width in zip("ABCD", [22, 22, 12, 16]):
        ws.column_dimensions[col].width = width
    for i in range(len(month_labels)):
        ws.column_dimensions[get_column_letter(5 + i)].width = 18



# ── MODIFICATION ────────────────────────────────────────
# This creates a new sheet "Base Price Summary" that shows the base price used for each month
# This is the new feature as per thapase/srikanth requirements on 09/30/2026
def _build_base_price_summary_sheet(
    wb, part_tier_pairs, results, month_labels
) -> None:
    ws = wb.create_sheet(title="Base Price Summary", index=1)   # right after Summary (index=0)
    ws.freeze_panes = "D2"                                       # was E2, shifted by 1

    fixed   = ["Part Number", "Tier 1", "Weight (lbs)"]           # "Base Price ($)" removed
    monthly = [label for _, label in month_labels]

    for ci, h in enumerate(fixed + monthly, 1):
        _write_header(ws, ci, h)
    ws.row_dimensions[1].height = 36

    for ri, (pn, t1) in enumerate(part_tier_pairs, 2):
        fill   = _ALT_FILL if ri % 2 == 0 else _WHITE_FILL
        result = results[(pn, t1)]

        if isinstance(result, Exception):
            _write_error_row(ws, ri, pn, t1, str(result), fill)
            continue

        _write_data(ws, ri, 1, pn,             "@",    fill, left=True)
        _write_data(ws, ri, 2, t1,             "@",    fill, left=True)
        _write_data(ws, ri, 3, result.pwt_lbs, "0.00", fill)
        # Base Price ($) column removed — month columns now start at 4

        for mi, (year_month, _) in enumerate(month_labels):
            mf  = next((f for f in result.forecasts if f.year_month == year_month), None)
            val = round(mf.base_price_used, 2) if mf else "N/A" ### mf.base_price_used to round(mf.base_price_used, 2)
            _write_data(ws, ri, 4 + mi, val, "$#,##0.00", fill)   # was 5 + mi ## change to "$#,##0.00"

    for col, width in zip("ABC", [22, 22, 12]):                     # was "ABCD", 4 widths
        ws.column_dimensions[col].width = width
    for i in range(len(month_labels)):
        ws.column_dimensions[get_column_letter(4 + i)].width = 18   # was 5 + i


# ======================================================================
# NEW - AOP
# Creates the sheet "Total AOP Summary": one row per (Part Number, Tier 1), one column
# per forecast month. Each month column shows THAT month's own Total AOP ($)
# (AOP qty x predicted price = same value as column T of the month tab) - NOT cumulative.
# The LAST column "Total AOP (All Months)" = sum of all months of that part.
# The "Grand Total" row at the bottom adds up ONLY the "Total AOP (All Months)" column.
# This is the new feature as per thapase/srikanth requirements on 10/22/2026
def _build_total_aop_summary_sheet(
    wb, part_tier_pairs, results, month_labels, aop_lookup
) -> None:
    ws = wb.create_sheet(title="Total AOP Summary", index=2)   # right after Summary and Base Price Summary
    ws.freeze_panes = "C2"                                      # Part Number and Tier 1 stay visible
 
    n_months = len(month_labels)
    fixed    = ["Part Number", "Tier 1"]
    monthly  = [label for _, label in month_labels]
    last_col = 3 + n_months                                     # column of "Total AOP (All Months)"
 
    for ci, h in enumerate(fixed + monthly + ["Total AOP\n(All Months)"], 1):
        _write_header(ws, ci, h)
    ws.row_dimensions[1].height = 36
 
    grand_total = 0.0                        # Grand Total of the "Total AOP (All Months)" column
 
    for ri, (pn, t1) in enumerate(part_tier_pairs, 2):
        fill   = _ALT_FILL if ri % 2 == 0 else _WHITE_FILL
        result = results[(pn, t1)]
 
        if isinstance(result, Exception):
            _write_error_row(ws, ri, pn, t1, str(result), fill)
            continue
 
        _write_data(ws, ri, 1, pn, "@", fill, left=True)
        _write_data(ws, ri, 2, t1, "@", fill, left=True)
 
        _pn, _t1 = pn.strip().lower(), t1.strip().lower()
        part_total = 0.0       # sum of all months of this part
        has_value = False      # stays False if no month has an AOP qty (total cell stays blank)
        for mi, (year_month, _) in enumerate(month_labels):
            month_total = None
            mf = next((f for f in result.forecasts if f.year_month == year_month), None)
            if mf is not None:
                aop_qty = aop_lookup.get((_pn, _t1, year_month))          # tier-specific AOP first
                if aop_qty is None:
                    aop_qty = aop_lookup.get((_pn, "", year_month))       # else item-level AOP (no tier)
                #month_total = compute_total_aop(mf.predicted_price, aop_qty)   # None if no qty
                # NEW - AOP (qty*base_price_summary_sheet)
                # This is modified as per the thapase requirement on 10/26/2026
                month_total = compute_total_aop(round(mf.base_price_used, 2), aop_qty)   # None if no qty
                if month_total is not None:
                    part_total += month_total
                    has_value = True
            # this month's own Total AOP (blank if the month has no AOP qty)
            _write_data(ws, ri, 3 + mi, round(month_total, 2) if month_total is not None else None,
                        "$#,##0.00", fill)
 
        # last column: sum of all months for this part
        _write_data(ws, ri, last_col, round(part_total, 2) if has_value else None, "$#,##0.00", fill)
        if has_value:
            grand_total += part_total
 
    # Grand Total row - bold, below the last part row (only the "Total AOP (All Months)" column)
    tr = len(part_tier_pairs) + 2
    tfill = PatternFill("solid", fgColor="FFE699")
    _write_data(ws, tr, 1, "Grand Total", "@", tfill, left=True)
    _write_data(ws, tr, 2, "", "@", tfill, left=True)
    for mi in range(n_months):                      # month columns: no grand total
        _write_data(ws, tr, 3 + mi, None, "$#,##0.00", tfill)
    _write_data(ws, tr, last_col, round(grand_total, 2), "$#,##0.00", tfill)   # only the last column
    for ci in range(1, last_col + 1):
        ws.cell(row=tr, column=ci).font = Font(name="Arial", size=10, bold=True)
 
    for col, width in zip("AB", [22, 22]):
        ws.column_dimensions[col].width = width
    for i in range(n_months):
        ws.column_dimensions[get_column_letter(3 + i)].width = 18
    ws.column_dimensions[get_column_letter(last_col)].width = 20
# ======================================================================
