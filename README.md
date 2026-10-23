# Aluminium Price Forecast API

A FastAPI service that forecasts aluminium part prices for the next 12 months. Each month's price is built from LME, Midwest premium, PPI and CNG inputs. The base price changes only once per quarter, in the last month of the quarter, and small changes are held steady by a ±2% deadband.

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Forecasting Logic](#2-forecasting-logic)
3. [Repository Structure](#3-repository-structure)
4. [File-by-File Guide](#4-file-by-file-guide)
5. [Data Layer](#5-data-layer)
6. [API Endpoints](#6-api-endpoints)
7. [Getting Started](#7-getting-started)
8. [How to Swap a Data Source](#8-how-to-swap-a-data-source)
9. [Excel File Format](#9-excel-file-format)
10. [Frontend Integration](#10-frontend-integration)
11. [Known Gaps and Legacy Files](#11-known-gaps-and-legacy-files)

---

## 1. Project Overview

The service predicts the price of an aluminium part for the **next 12 calendar months** (or 13 when the current month is included). The current month comes from the **UTC system clock**.

**Key concepts**
- **Part Number + Tier 1** form the unique key for a part. The same part number can appear under several Tier 1 suppliers (e.g. "Kadon Aerospace", "QCC, LLC", "NA") with different weights and prices.
- **PWt** is the part weight in lbs, fixed per Part Number + Tier 1.
- **Base price** is always read from the Parts sheet of `aluminium_data.xlsx`. It cannot be overridden through the API.
- **Market data** (LME, Midwest, PPI) is read from `aluminium_data.xlsx` and cached in memory.
- **CNG_Q and CNG_Q-1** are **required request inputs** ($/lb). They are not stored in Excel, and the same two values are used for every forecast month.
- **Every month is calculated independently** from its own market data, using a base price chosen by the quarterly base-price rule ([§2.3](#23-base-price-quarterly-rule)). A predicted price is not simply rolled into the next month.

---

## 2. Forecasting Logic

### 2.1 Core formula (applied to every forecast month)

```
raw_price = P_current
            + [ (AMS_Q − AMS_Q-1) × PWt ]
            + [ PPI_Factor × P_current ]
```

| Variable | Definition |
|---|---|
| `MC_Q` | `(LME + Midwest) + MC_OFFSET` for the **month being predicted** |
| `MC_Q-1` | Same expression for the previous quarter (see 2.2) |
| `AMS_Q` | `(MC_Q × DF_c) + CNG_Q` |
| `AMS_Q-1` | `(MC_Q-1 × DF_c) + CNG_Q-1` |
| `PPI_Q` | PPI index of the **month being predicted** |
| `PPI_Q-1` | PPI of the previous quarter (see 2.2) |
| `PPI_Factor` | `(PPI_Q − PPI_Q-1) / PPI_Q-1` |
| `CNG_Q`, `CNG_Q-1` | Supplied in the request |
| `DF_c` | `1.44` (`DF_C` in config) |
| `MC_OFFSET` | `0.558883` (config) |
| `PWt` | Part weight in lbs |
| `P_current` | Base price for the month being predicted, chosen by the quarterly rule (see 2.3) |

> **MC formula:** the current code uses `MC = (LME + Midwest) + 0.558883`. The older `× 0.91 + 1.25` form is **not used**. `MC_MULTIPLIER` still exists in `config.py` but the engine ignores it. The previous version is (`× MC_MULTIPLIER + MC_OFFSET`). So this logic is updated and this is from thapase/srikanth suggestions on 9/30/2026

So This MC_OFFSET is the newly added feature in this version. And this is took from thapase/srikanth requirements on 9/30/2026   ######### CHANGE

### 2.2 Previous-quarter values (`_Q-1`)

The previous quarter is the calendar quarter before the one containing the predicted month. Both `MC_Q-1` and `PPI_Q-1` follow the same "actual vs projected" rule:

- If the **last month of the previous quarter** is covered by actuals, use the value of **that month**.
- Otherwise use the **average of the three months** of the previous quarter. For MC this averages MC, LME and Midwest.

A month counts as "actual" when `last_month ≤ current_month − PPI_ACTUAL_LAG_MONTHS`. The default lag is **1** in `config.py`, so sitting in October 2026 actuals run through September 2026.

### 2.3 Base price rule

Month on Month Sheet:
It is observed that base price is changing every 50 days, below is the formula to be updated .
Base price = if it is last month of the quarter, then consider first month final price of the same quarter else consider base price of previous month


**Rule.** It depends on the position of the forecast month inside its calendar quarter (Jan–Mar, Apr–Jun, Jul–Sep, Oct–Dec):

**How the active base moves.** The engine keeps two values while it loops over the months:

- `active_base` starts as the price from the Parts sheet.
- `pending_first_month_price` holds the final price of the current quarter's 1st month.

```python
position = _month_position_in_quarter(forecast_month)       # 1, 2 or 3
if position == 3 and pending_first_month_price is not None:
    current_price = pending_first_month_price               # 3rd month: this quarter's 1st-month final price
else:
    current_price = active_base                             # 1st and 2nd month: active base
...
if position == 1:
    pending_first_month_price = predicted_price             # remember the 1st month's final price
if position == 3:
    active_base = current_price                             # carry it into the next quarter's 1st and 2nd months
    pending_first_month_price = None
```

If the 1st month's calculated price is within ±2% of its base, the deadband holds the price, so its final price equals the base and the base does not change.

So This Base price rule is the newly added feature in this version. And this is took from thapase/srikanth requirements on 10/11/2026   ######### CHANGE

### 2.4 Deadband (±2%)

So This Deadband logic is the newly added feature in this version. And this is took from thapase/srikanth requirements on 9/30/2026   ######### CHANGE

After calculating `raw_price`:

```
if |raw_price − P_current| / P_current ≤ DEADBAND_PCT (0.02):
    predicted_price = P_current        # change too small, price is held
else:
    predicted_price = raw_price
```

The engine computes and returns both `predicted_price` (after the deadband, the "final price") and `predicted_price_without_deadband` (the raw value). The deadband output is what is carried forward as a base price through the quarterly rule (§2.3).

### 2.5 Forecast window and price month

`include_current_month` (`"YES"` / `"NO"`, default `"NO"`) decides where the forecast starts. The Parts sheet price column header must name the month the prices are for, e.g. `Current Price ($) August 2026`.

| `include_current_month` | Price month needed | First forecast month | Months |
|---|---|---|---|
| `"NO"` (default) | **This** month | Next month | 12 |
| `"YES"` | **Last** month | This month | 13 |

If the header month does not match, the request fails with a message saying which month the data holds and which is needed.

### 2.6 Worked example (sitting in Sep 2026, `include_current_month = "NO"`)

The forecast covers Oct 2026 → Sep 2027. `F(month)` means the final price (after the deadband) of that month, and "Excel" is the price from the Parts sheet.

| Month | Position in quarter | `P_current` used | Carried forward as a base? |
|---|---|---|---|
| Oct 2026 | 1st | Excel | Yes: `F(Oct 2026)` is the base for Dec 2026, Jan 2027 and Feb 2027 |
| Nov 2026 | 2nd | Excel | No |
| Dec 2026 | 3rd (last) | `F(Oct 2026)` | No |
| Jan 2027 | 1st | `F(Oct 2026)` | Yes: `F(Jan 2027)` is the base for Mar, Apr and May 2027 |
| Feb 2027 | 2nd | `F(Oct 2026)` | No |
| Mar 2027 | 3rd (last) | `F(Jan 2027)` | No |
| Apr 2027 | 1st | `F(Jan 2027)` | Yes: `F(Apr 2027)` is the base for Jun, Jul and Aug 2027 |
| May 2027 | 2nd | `F(Jan 2027)` | No |
| Jun 2027 | 3rd (last) | `F(Apr 2027)` | No |
| Jul 2027 | 1st | `F(Apr 2027)` | Yes: `F(Jul 2027)` is the base for Sep 2027 onwards |
| Aug 2027 | 2nd | `F(Apr 2027)` | No |
| Sep 2027 | 3rd (last) | `F(Jul 2027)` | No |

Each month computes its own `MC_Q`, `PPI_Q`, `AMS_Q` and `PPI_Factor`, then applies the core formula and the deadband on its own `P_current`. The **Base Price Summary** sheet in the batch output shows the `P_current` column above for every part.

### 2.7 Price-change breakdown

The single-part endpoint also returns a decomposition of `raw_price − P_current` into four exact factors:

```
lme_effect     = (lme_q − lme_q-1) × DF_c × PWt
midwest_effect = (midwest_q − midwest_q-1) × DF_c × PWt
cng_effect     = (cng_q − cng_q-1) × PWt
ppi_effect     = ppi_factor × P_current
```

They sum to `total_change` (`sum_check` should be 0). The breakdown uses the price **before** the deadband so `sum_check` stays 0. It is not produced by the batch endpoint.

### 2.8 Rounding of decimal values

**Principle:** every calculation runs at full floating-point precision. Rounding is applied only when a value is placed in the API response or in the Excel output. Rounded values are never fed back into later calculations, so the deadband check and the base price carried forward (§2.3, §2.4) always use the unrounded numbers.

**API response (JSON), decimals kept:**

| Field | Decimals |
|---|---|
| `predicted_price`, `final_unit_price_with_deadband`, `base_price_used` | 2 |
| `mc_q`, `lme_q`, `midwest_q`, `cng_q`, `ams_q` and their `_q-1` versions | 2 |
| `ams_delta` | 2 |
| `ppi_q`, `ppi_q-1` | 1 |
| `ppi_factor` | 1 |
| `pwt_lbs`, `df_c`, top-level `base_price` | Not rounded (returned as stored) |

**Batch Excel output, decimals kept:**

| Column / sheet | Decimals | Cell format |
|---|---|---|
| Base Price, Predicted Price, Predicted Price (without deadband) | 2 | `$#,##0.00` |
| `MC_Q`, `MC_Q-1`, `AMS_Q`, `AMS_Q-1`, `AMS Delta` | 2 | `$0.00` |
| `PPI_Q`, `PPI_Q-1`, `PPI Factor` | 1 | `0.0` |
| `CNG_Q`, `CNG_Q-1`, `DF_c` | Not rounded | `General` |
| Weight | Not rounded (shown with 2 decimals) | `0.00` |
| Summary and Base Price Summary sheets (all prices) | 2 | `$#,##0.00` |

**Example** (values from a real response): `mc_q` 3.050797 shows as `$3.05`, `ams_delta` −0.027636 as `-$0.03`, and `ppi_factor` 0.00199722 as `0.0` in the Excel sheet. The API keeps all of those decimals.

So for alloy used rounding off 2 decimal places, for ppi rounding off 1 decimal places and for unit price used rounding off 2 decimal places.
---

## 3. Repository Structure

The code imports modules as `app.core...`, `app.data...`, etc. The files below are therefore expected in this layout. In the uploaded archive they are all in one flat folder, so they need to be placed into these folders before the server will start.

```
project/
├── main.py                      # FastAPI entry point
├── aluminium_data.xlsx          # Master data (Parts, LME, Midwest, PPI sheets)
├── input.xlsx                   # Sample batch input (Part Number, Tier 1)
├── requirements.txt
├── approach.txt                 # Original problem statement (older logic, see §11)
│
├── app/
│   ├── api/v1/
│   │   ├── router.py                   # Registers endpoint routers
│   │   └── endpoints/
│   │       ├── forecast_excel.py       # Single-part + /parts + /reload
│   │       ├── forecast_batch.py       # Batch upload/download
│   │       └── forecast.py             # LEGACY, not registered (§11)
│   ├── core/
│   │   ├── config.py                   # Settings and constants
│   │   └── logging_config.py
│   ├── data/
│   │   ├── base.py                     # Abstract repository interfaces
│   │   ├── excel_store.py              # Active data source
│   │   └── hardcoded_store.py          # Legacy in-code data (tests only)
│   ├── models/
│   │   ├── request.py                  # Endpoints Requests
│   │   └── response.py                 # Endpoints Responses
│   └── services/
│       ├── forecast_engine.py          # Current formula logic
│       ├── forecast_engine_v1.py       # Previous version (× 0.91 + offset), reference only
│       └── batch_forecast.py           # Batch orchestration + output workbook
│
└── tests/
    └── test_forecast_engine.py         # For testing
```

> Every folder inside `app/` needs an empty `__init__.py`, otherwise you will get a `ModuleNotFoundError`.

---

## 4. File-by-File Guide

### `main.py`
Creates the FastAPI app, configures logging, adds CORS (`allow_origins=["*"]`, tighten for production), registers a global exception handler returning HTTP 500 JSON, mounts `api_v1_router`, and exposes `GET /health`, which returns `{"status": "ok", "version": ...}`. The startup log prints `DF_c` and the horizon.

```bash
uvicorn main:app --reload --port 8000
```

### `app/core/config.py`
`pydantic-settings` class. Every value can be overridden by an environment variable or a `.env` file.

| Setting | Default | Meaning |
|---|---|---|
| `DF_C` | `1.44` | Conversion factor in AMS |
| `MC_MULTIPLIER` | `0.91` | **Unused by the current engine** |
| `MC_OFFSET` | `0.558883` | Added to `(LME + Midwest)` to get MC |
| `PPI_ACTUAL_LAG_MONTHS` | `1` | Months before the current month that actuals end. Decides actual vs average for `MC_Q-1` and `PPI_Q-1` |
| `FORECAST_HORIZON_MONTHS` | `12` | Months forecast (13 when the current month is included) |
| `DEADBAND_PCT` | `0.02` | ±2% band within which the base price is retained |
| `PRICE_LAG_MONTHS` | `2` | **No longer used by the engine.** The 2-month lag was replaced by the quarterly base-price rule (§2.3). Safe to remove |
| `APP_NAME`, `APP_VERSION`, `DEBUG` | — | App metadata and debug logging |

### `app/core/logging_config.py`
`configure_logging(debug)` sets one log format (`time | level | module:line | message`) and quiets `uvicorn.access`. Other modules just use `logging.getLogger(__name__)`.

### `app/data/base.py`
Abstract contracts the engine depends on:
- **`MarketDataRepository`**: `get_lme()`, `get_midwest_premium()`, `get_ppi()`. Each takes `YYYY-MM` and returns `float | None`.
- **`PartRepository`**: `get_part_weight(part_number, tier_1)`, `get_base_price(part_number, tier_1, year_month)`, plus an optional `get_price_month()` used to give clearer price-month errors.

### `app/data/excel_store.py` (active data source)
Loads four sheets, `Parts`, `LME`, `Midwest` and `PPI`, with `lru_cache`.
- Everything is read as text with `keep_default_na=False`, so the supplier name `"NA"` is not turned into NaN.
- `ExcelMarketDataRepository` looks up month → value for LME, Midwest and PPI.
- `ExcelPartRepository` matches rows on (Part Number, Tier 1), reading columns **by position** (0 part, 1 tier, 2 weight, 3 price).
- `_parse_price_month()` reads the month from the price column header (`Aug 2026`, `August 2026`, `2026-08`, `08/2026`, or a date cell).
- `get_base_price()` returns a price **only if** the requested month equals the header month.
- `get_all_parts()` feeds `GET /parts`. `_invalidate_cache()` backs `POST /reload`.
- The file path defaults to `aluminium_data.xlsx` two levels above `app/data/`; override it with `ALUMINIUM_EXCEL_PATH`.

### `app/data/hardcoded_store.py`
Original in-code data set (July 2025 onwards) kept for tests and as a pattern for new adapters. Not used by active endpoints.

### `app/services/forecast_engine.py` (core logic)
`ForecastEngine(market_repo, part_repo).forecast(part_number, tier_1, cng_q, cng_q_1, include_breakdown=False, include_current_month=False)`

Steps:
1. Read the current UTC month and decide the price month (this month, or last month when `include_current_month` is true).
2. Look up `PWt` (error "Part not found" if missing), then the base price (error naming the expected month if the header month differs).
3. Loop over 12 months (13 if the current month is included), starting the month after the price month. For each month:
   1. Pick `P_current` using the quarterly base-price rule (§2.3).
   2. Build the `QuarterContext` (MC, PPI, AMS and previous-quarter values).
   3. Compute `raw_price`, apply the deadband, and optionally the factor breakdown.
   4. Store the result, flagged `is_data_projected = (step ≥ 1)`.
   5. Update the quarter state: after a 1st month, remember its final price; after a 3rd month, make the base used in that month the new active base.
4. Return a `ForecastResponse`.

Helper functions: `_quarter_of_month`, `_quarter_months`, `_last_month_of_quarter`, `_prev_quarter`, `_month_position_in_quarter`, `_shift_month`, `_compute_mc`, `_ppi_prev_quarter`, `_compute_quarter_context`, `_compute_price_breakdown`, `_price_month_error`.

### `app/services/batch_forecast.py`
- **`read_parts_from_upload(file_bytes)`**: finds `Part Number` and `Tier 1` columns (case-insensitive), de-duplicates while keeping order, returns `(part, tier)` tuples.
- **`build_forecast_workbook(pairs, engine, cng_q, cng_q_1, include_current_month)`**: runs the engine for each pair. A failing part gets a red **ERROR** row rather than stopping the batch. An error is raised only if every part fails.

Output workbook sheets:
1. **Summary**: predicted price per part per month.
2. **Base Price Summary**: the `P_current` used per part per month. And this sheet is updated from thapase/srikanth srequirements.
3. **One sheet per forecast month** (12, or 13 with the current month), with these columns:

`Part Number, Tier 1, Weight, Base Price, Quarter (Current), Quarter (Previous), MC_Q, MC_Q-1, PPI_Q, PPI_Q-1, PPI Factor, CNG_Q, CNG_Q-1, AMS_Q, AMS_Q-1, AMS Delta, DF_c, Predicted Price, Predicted Price (without deadband)`

### `app/models/request.py`
`ForecastRequest`:

```json
{
  "part_number": "09-0052-003",
  "tier_1": "NA",
  "cng_q": 0.97,
  "cng_q-1": 0.95,
  "include_current_month": "NO"
}
```

`cng_q` and `cng_q-1` (alias for `cng_q_1`) are **required**. `include_current_month` accepts `YES` / `NO` in any case.

### `app/models/response.py`
- **`QuarterContext`**: `quarter_label`, `mc_q`, `lme_q`, `midwest_q`, `ppi_q`, `cng_q`, `ams_q`, the previous-quarter equivalents (serialised as `mc_q-1`, `lme_q-1`, etc.), `ppi_factor`, `ams_delta`.
- **`PriceFactorBreakdown`** / **`PriceChangeBreakdown`**: dollar and % effect of LME, Midwest, CNG and PPI, plus `sum_check`.
- **`MonthForecast`**: `year_month`, `month_label`, `predicted_price`, `predicted_price_without_deadband`, `base_price_used`, `pwt`, `df_c`, `quarter_context`, `price_change_breakdown` (single-part endpoint only), `is_data_projected`.
- **`ForecastResponse`**: `part_number`, `tier_1`, `pwt_lbs`, `base_year_month`, `base_price`, `forecast_generated_at`, `forecasts`.
- **`ErrorResponse`**: `error`, `detail`.

> `predicted_price_without_deadband` is a **required** field on `MonthForecast`. Every place that builds a `MonthForecast` (currently only `forecast_engine.py`) must pass it, otherwise validation fails.

### `app/api/v1/router.py`
Mounts the two active routers under `/api/v1`: `forecast_excel_router` and `forecast_batch_router`.

### `app/api/v1/endpoints/forecast_excel.py`
| Method | Path | Purpose |
|---|---|---|
| `GET` | `/api/v1/forecast-excel/parts` | List all parts with weight, price and price month |
| `POST` | `/api/v1/forecast-excel/reload` | Clear the Excel cache |
| `POST` | `/api/v1/forecast-excel` | Single-part forecast (breakdown included) |

Errors: `404` for "Part not found", `422` for other value errors (missing price or market data, price-month mismatch), `500` for the Excel file missing or an unexpected failure. Repositories are wired in with `Depends`.

### `app/api/v1/endpoints/forecast_batch.py`
`POST /api/v1/forecast-batch` (multipart form):
- `file`: `.xlsx` with `Part Number` and `Tier 1`
- `cng_q`, `cng_q-1`: required numbers
- `include_current_month`: `YES` / `NO` (default `NO`)

Returns a `.xlsx` download named `forecast_<YYYY-MM-DD>.xlsx` with `Content-Disposition` and an `X-Parts-Count` header. `400` is returned for a non-.xlsx file, a bad `include_current_month`, missing columns, an empty list, or when all parts fail.

---

## 5. Data Layer

### `aluminium_data.xlsx`: 4 sheets

| Sheet | Columns | Notes |
|---|---|---|
| **Parts** | Part Number, Tier 1, Weight (lbs), Current Price ($) *with month in header* | Key = Part Number + Tier 1 |
| **LME** | Month (`YYYY-MM`), value | $/lb |
| **Midwest** | Month (`YYYY-MM`), value | $/lb |
| **PPI** | Month (`YYYY-MM`), value | Index value, not dollars |

After editing the workbook, restart the server or call `POST /api/v1/forecast-excel/reload`.

**Coverage needed:** each forecast month needs LME, Midwest and PPI for itself, and for the previous quarter (the last month, or all three months when that quarter is not yet covered by actuals). Cover roughly two quarters back to 12 months ahead.

---

## 6. API Endpoints

Base URL: `http://localhost:8000`. Swagger UI at `/docs`, ReDoc at `/redoc`.

| Method | Endpoint | Input | Output |
|---|---|---|---|
| `GET` | `/health` | none | `{"status":"ok","version":"1.0.0"}` |
| `GET` | `/api/v1/forecast-excel/parts` | none | Parts list |
| `POST` | `/api/v1/forecast-excel/reload` | none | Cache cleared |
| `POST` | `/api/v1/forecast-excel` | JSON body | JSON forecast |
| `POST` | `/api/v1/forecast-batch` | multipart form | `.xlsx` download |

### Single-part response (abbreviated)
```json
{
  "part_number": "09-0052-003",
  "tier_1": "NA",
  "pwt_lbs": 3.6,
  "base_year_month": "2026-09",
  "base_price": 47.09,
  "forecast_generated_at": "2026-09-10T08:00:00+00:00",
  "forecasts": [
    {
      "year_month": "2026-10",
      "month_label": "October 2026",
      "predicted_price": 0.0,
      "predicted_price_without_deadband": 0.0,
      "base_price_used": 47.09,
      "pwt": 3.6,
      "df_c": 1.44,
      "quarter_context": {
        "quarter_label": "Q4-2026",
        "mc_q": 0.0, "lme_q": 0.0, "midwest_q": 0.0, "ppi_q": 0.0,
        "cng_q": 0.97, "ams_q": 0.0,
        "prev_quarter_label": "Q3-2026",
        "mc_q-1": 0.0, "lme_q-1": 0.0, "midwest_q-1": 0.0, "ppi_q-1": 0.0,
        "cng_q-1": 0.95, "ams_q-1": 0.0,
        "ppi_factor": 0.0, "ams_delta": 0.0
      },
      "price_change_breakdown": { "...": "lme/midwest/cng/ppi effects + sum_check" },
      "is_data_projected": false
    }
  ]
}
```
Values shown as `0.0` are placeholders. Real numbers depend on the data in the workbook.

---

## 7. Getting Started

### Prerequisites
- Python 3.12 recommended
- `aluminium_data.xlsx` next to `main.py` (or set `ALUMINIUM_EXCEL_PATH`)

### Install
```bash
python -m venv env
env\Scripts\activate          # Windows
source env/bin/activate       # Mac / Linux
pip install -r requirements.txt
```

### `requirements.txt` (as in the repo)
```
fastapi>=0.115.0
uvicorn[standard]>=0.30.6
pydantic>=2.11.0
pydantic-settings>=2.7.0
python-dateutil>=2.9.0
openpyxl>=3.1.0
pandas>=2.0.0
python-multipart>=0.0.9
joblib>=1.5.3
scikit-learn>=1.9.0
```
`joblib` and `scikit-learn` are not imported by any current code (left over from a removed ML endpoint) and can be dropped.

### Run / test
```bash
uvicorn main:app --reload --port 8000
pytest tests/ -v          # see §11: tests need updating
```

### Optional: custom Excel path
```bash
export ALUMINIUM_EXCEL_PATH=/full/path/to/aluminium_data.xlsx      # Mac / Linux
set ALUMINIUM_EXCEL_PATH=C:\full\path\to\aluminium_data.xlsx       # Windows
```

---

## 8. How to Swap a Data Source

The engine depends only on the abstract classes in `base.py`.

**Live market data (BLS, LME feed):** create a class inheriting `MarketDataRepository` that implements `get_lme`, `get_midwest_premium` and `get_ppi`, then change the factory in the endpoint files:

```python
def get_market_repo() -> MarketDataRepository:
    return LiveApiMarketDataRepository()
```

**ERP for parts:** implement `PartRepository` (`get_part_weight`, `get_base_price`, optionally `get_price_month`) and change `get_part_repo()`.

Note the factories are duplicated in `forecast_excel.py` and `forecast_batch.py`, so change both.

---

## 9. Excel File Format

### Parts sheet

| Part Number | Tier 1 | Weight (lbs) | Current Price ($) August 2026 |
|---|---|---|---|
| 09-0052-003 | NA | 3.6 | 47.09 |
| 2190-1015 | NA | 7 | 89.97 |

- **Part Number + Tier 1 = unique key**. `NA` means no specific supplier.
- Matching is exact apart from surrounding whitespace.
- The price column header **must name its month**. Update the prices and the header together.
  - Header month = last month → use `include_current_month: "YES"`.
  - Header month = this month → use `include_current_month: "NO"`.
- Price cells that contain currency text such as `$1,233.07` must be stored as plain numbers, because values are cast with `float()`.

### LME, Midwest, PPI sheets

| Month (YYYY-MM) | Value |
|---|---|
| 2026-07 | 1.4751 |

CNG is **not** stored in Excel.

> The sample `aluminium_data.xlsx` and `input.xlsx` in the archive are tab-separated text files, not real Excel workbooks. The sample `aluminium_data.xlsx` also contains only the Parts data. Re-save them as genuine `.xlsx` files with all four sheets before running the service.

---

## 10. Frontend Integration

### Single-part forecast (Axios)
```javascript
const { data } = await axios.post('/api/v1/forecast-excel', {
  part_number: '09-0052-003',
  tier_1: 'NA',
  cng_q: 0.97,
  'cng_q-1': 0.95,
  include_current_month: 'NO'
});

// data.forecasts[0].predicted_price
// data.forecasts[0].predicted_price_without_deadband
// data.forecasts[0].base_price_used
// data.forecasts[0].price_change_breakdown
```

### Batch forecast (Axios)
```javascript
const form = new FormData();
form.append('file', excelFile);
form.append('cng_q', '0.97');
form.append('cng_q-1', '0.95');
form.append('include_current_month', 'NO');

const response = await axios.post('/api/v1/forecast-batch', form, {
  responseType: 'blob'
});

const url = URL.createObjectURL(response.data);
const a = document.createElement('a');
a.href = url;
a.download = 'forecast_result.xlsx';
a.click();
URL.revokeObjectURL(url);
```

### Batch input format
Columns `Part Number` and `Tier 1` (case-insensitive). Duplicate pairs are de-duplicated silently, and blank rows are skipped.

---

## 11. Known Gaps and Legacy Files

- **`approach.txt`** describes an **earlier design**: the price is the same for every month of a quarter, CNG is fixed at 0.97, `MC = (LME + Midwest) × 0.91 + 1.25`, and `PPI_Q` is a quarter average. The code now differs on all of these (see §2): the base price moves only in the last month of each quarter (and each month still gets its own calculated price), CNG comes from the request, MC has no multiplier, and `PPI_Q` is the predicted month's value.
- **`forecast.py`** is a legacy endpoint that is **not registered** in the router. It still passes `current_price` and uses the hardcoded repositories, so it would fail against the current request model and engine signature.
- **`test_forecast_engine.py`** targets the older engine: it calls `forecast()` without CNG arguments, expects forecasts starting March 2026, and assumes chained prices. It needs updating for the quarterly base-price rule, deadband, new MC formula and CNG inputs.
- **Docstrings** in `response.py` and `router.py` still mention the `× 0.91 + 1.25` MC formula and an ML endpoint. These comments are out of date.
- **`forecast_engine.py`** still has leftover lines from the old 2-month lag: commented-out `price_lag` / `calculated_prices` code, `CHANGED` / `NEW` markers, and a module docstring (ITERATION LOGIC) that still describes the lag. `PRICE_LAG_MONTHS` in `config.py` is no longer read. These can be cleaned up.
- **`excel_store.py`** reads columns by position, so keep the Parts column order unchanged.

---


*For questions about the formula or data sources, refer to the quarterly pricing agreement documentation or contact the pricing team.*
