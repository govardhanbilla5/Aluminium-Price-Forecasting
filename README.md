# Aluminium Price Forecast API

A production-grade FastAPI service that forecasts aluminium part prices for the next 12 months using a quarterly formula driven by LME prices, Midwest premium, PPI index, and CNG costs.

---

## Table of Contents

1. [Project Overview](#1-project-overview)
2. [Forecasting Formula](#2-forecasting-formula)
3. [Repository Structure](#3-repository-structure)
4. [File-by-File Guide](#4-file-by-file-guide)
5. [Data Layer](#5-data-layer)
6. [API Endpoints](#6-api-endpoints)
7. [Getting Started](#7-getting-started)
8. [How to Swap a Data Source](#8-how-to-swap-a-data-source)
9. [Excel File Format](#9-excel-file-format)
10. [Frontend Integration](#10-frontend-integration)

---

## 1. Project Overview

This service predicts the price of aluminium parts for the **next 12 calendar months**, starting from the current month. Prices are forecast **quarter by quarter** — the same quarterly market data (LME, Midwest, PPI) applies to all three months within a quarter, and each month's predicted price is used as the base for the next month (chained forecasting).

**Key concepts:**
- **Part Number + Tier 1** together form the unique identifier for a part. The same part number can appear under multiple Tier 1 suppliers (e.g. "Kadon Aerospace", "Point Precision Inc.", "NA") with different weights and prices.
- **PWt (Part Weight in lbs)** is a fixed property of each part + supplier combination.
- **Current price** is always read from `aluminium_data.xlsx` — it cannot be overridden via the API.
- All market data (LME, Midwest premium, PPI) is stored in `aluminium_data.xlsx` (CNG comes from the request) and loaded once into memory at startup (cached).

---

## 2. Forecasting Formula

The core formula predicts the price at the **last month of each quarter**, then chains forward:

```
P_next = P_current
         + [ (AMS_Q − AMS_Q-1) × PWt ]
         + [ PPI_Factor × P_current ]
```

Where:

| Variable | Definition |
|---|---|
| `AMS_Q` | `(MC_Q × DF_c) + CNG_Q` — Alloy Metal + Gas cost, current quarter |
| `AMS_Q-1` | `(MC_Q-1 × DF_c) + CNG_Q-1` — same for previous quarter |
| `MC_Q` | `(LME + Midwest) × 0.91 + 1.25` for the **month being predicted**, in $/lb |
| `MC_Q-1` | `(LME + Midwest) × 0.91 + 1.25` for the **last month** of the previous quarter, in $/lb |
| `PPI_Q` | PPI index value of the **month being predicted** |
| `PPI_Q-1` | PPI of the **last month** of the previous quarter when that whole quarter is covered by published actuals; otherwise the **average** of its three months (projected values). PPI actuals end `PPI_ACTUAL_LAG_MONTHS` (2) months before the current month — sitting in September, actuals run through July. |
| `PPI_Factor` | `(PPI_Q − PPI_Q-1) / PPI_Q-1` |
| `CNG_Q` | Supplied in the request (`cng_q`), same for every month ($/lb) |
| `CNG_Q-1` | Supplied in the request (`cng_q-1`), same for every month ($/lb) |
| `DF_c` | Fixed constant = **1.44** |
| `PWt` | Part weight in lbs (fixed per Part Number + Tier 1) |

The 0.91 and 1.25 constants are `MC_MULTIPLIER` and `MC_OFFSET` in `app/core/config.py`.

**Iteration logic:** Starting from the current month (auto-detected from system clock), the engine forecasts the first month of the next quarter, then the second, and so on for 12 months. Each predicted price becomes the `P_current` for the next month.

**`include_current_month`** (optional, `"YES"` / `"NO"`, default `"NO"`): with `"YES"` the forecast starts at the **current** month instead of next month and covers 13 months (e.g. sitting in Sep 2026 → Sep 2026 … Sep 2027). 

**Price month:** the header of the Parts sheet price column must name the month its prices are for, e.g. `Current Price ($) Aug 2026` (also accepted: `August 2026`, `2026-08`, `08/2026`). The first `P_current` is always the price of the month just before the first forecast month:

| `include_current_month` | Price month needed (sitting in Sep 2026) | Forecast months |
|---|---|---|
| `"NO"` (default) | Sep 2026 (this month) | Oct 2026 → Sep 2027 (12) |
| `"YES"` | Aug 2026 (last month) | Sep 2026 → Sep 2027 (13) |

If the header's month doesn't match, the request fails with a message saying which month the data holds and which month is needed.

**Example (sitting in Aug 2026):**
- Step 1 → Forecast Sep 2026 using Q3-2026 data, base = Aug 2026 actual price
- Step 2 → Forecast Oct 2026 using Q4-2026 data, base = Sep 2026 predicted price
- Step 3 → Forecast Nov 2026 using Q4-2026 data, base = Oct 2026 predicted price
- ... and so on through Aug 2027

---

## 3. Repository Structure

```
files/
│
├── main.py                             # FastAPI app entry point — run this
├── aluminium_data.xlsx                 # Master data file (parts + market data)
├── requirements.txt                    # Python dependencies
│
├── app/
│   ├── api/
│   │   └── v1/
│   │       ├── router.py               # Registers all endpoint routers
│   │       └── endpoints/
│   │           ├── forecast_excel.py   # Single-part forecast endpoint
│   │           └── forecast_batch.py   # Batch forecast endpoint (file upload)
│   │
│   ├── core/
│   │   ├── config.py                   # App settings and constants (DF_c, horizon)
│   │   └── logging_config.py           # Structured logging setup
│   │
│   ├── data/
│   │   ├── base.py                     # Abstract interfaces for data sources
│   │   ├── excel_store.py              # Reads aluminium_data.xlsx (active source)
│   │   └── hardcoded_store.py          # Legacy hardcoded data (not in active endpoints)
│   │
│   ├── models/
│   │   ├── request.py                  # Pydantic models for API request bodies
│   │   └── response.py                 # Pydantic models for API responses
│   │
│   └── services/
│       ├── forecast_engine.py          # Core forecasting logic (formula lives here)
│       └── batch_forecast.py           # Batch orchestration + Excel output builder
│
└── tests/
    └── test_forecast_engine.py         # Unit tests
```

> Every folder inside `app/` must contain an empty `__init__.py` file or Python will not recognise it as a package and you will get a `ModuleNotFoundError`.

---

## 4. File-by-File Guide

### `main.py`
The application entry point. Creates the FastAPI app, registers CORS middleware, attaches all routers, and defines the `/health` endpoint. **Start reading here** when first exploring the codebase.

```bash
uvicorn main:app --reload --port 8000
```

---

### `app/core/config.py`
Central place for all tunable constants. Uses `pydantic-settings` so any value can be overridden via an environment variable without changing code.

| Setting | Default | Meaning |
|---|---|---|
| `DF_C` | `1.44` | Density/conversion factor in the AMS formula |
| `MC_MULTIPLIER` | `0.91` | MC = (LME + Midwest) × MC_MULTIPLIER + MC_OFFSET |
| `MC_OFFSET` | `1.25` | Added after the multiplier in the MC formula |
| `PPI_ACTUAL_LAG_MONTHS` | `2` | How many months before the current month the published PPI actuals end |
| `FORECAST_HORIZON_MONTHS` | `12` | Number of months to forecast |

If you ever need to change `DF_c`, change it here — not inside the formula code.

---

### `app/core/logging_config.py`
Sets up a single structured log format used across the whole app:
```
2026-08-25 10:30:00 | INFO     | app.services.forecast_engine:285 | Starting forecast...
```
Called once at startup inside `main.py`. Every other module just does:
```python
import logging
logger = logging.getLogger(__name__)
```

---

### `app/data/base.py` ⭐ Read this first
Defines the **abstract interfaces (contracts)** that all business logic depends on. There are two abstract classes:

- **`MarketDataRepository`** — must implement `get_lme()`, `get_midwest_premium()`, `get_ppi()`. Each takes a `YYYY-MM` string and returns `float | None`.
- **`PartRepository`** — must implement `get_part_weight(part_number, tier_1)` and `get_base_price(part_number, tier_1, year_month)`.

**The forecast engine only ever imports these abstractions — never any concrete implementation.** This is the key architectural decision that makes the data source swappable. See [How to Swap a Data Source](#8-how-to-swap-a-data-source).

---

### `app/data/excel_store.py` ⭐ Active data source
Reads `aluminium_data.xlsx` and implements both abstract interfaces.

- **`ExcelMarketDataRepository`** — serves LME, Midwest, PPI values from their respective sheets.
- **`ExcelPartRepository`** — looks up parts using **(Part Number + Tier 1)** as the composite key.
- **`get_all_parts()`** — returns every row from the Parts sheet; used by the `/parts` listing endpoint.
- **Caching** — the workbook is loaded once via `@lru_cache`. To reload after editing the Excel without restarting the server, call `POST /api/v1/forecast-excel/reload`.
- **Important detail:** `keep_default_na=False` is passed to pandas so the string `"NA"` (a valid Tier 1 value) is not silently converted to `NaN`.

---

### `app/data/hardcoded_store.py`
The original hardcoded data store kept for reference. Not connected to any active endpoint. Useful as a pattern reference if you need to write a new data source adapter or run tests without file I/O.

---

### `app/services/forecast_engine.py` ⭐ Core business logic
Contains `ForecastEngine` — a stateless class injected with a `MarketDataRepository` and a `PartRepository` at construction time.

**Main public method:** `forecast(part_number: str, tier_1: str) -> ForecastResponse`

What it does, step by step:
1. Detects the current month from the system clock (`YYYY-MM`)
2. Looks up `PWt` (weight in lbs) for the part + tier_1 combination
3. Looks up the current base price from the data store
4. For each of the next 12 months:
   - Determines the calendar quarter
   - Builds a `QuarterContext` (MC_Q, AMS_Q, PPI_Q, CNG_Q, etc.) for that month — MC_Q and PPI_Q come from the predicted month itself, so the context is computed per month
   - Applies the formula to get `predicted_price`
   - Rolls `predicted_price` forward as `P_current` for the next step
5. Returns a `ForecastResponse` with all 12 `MonthForecast` objects, each containing every intermediate variable used in the formula

**Quarter helper functions** (at the top of the file):

| Function | What it does |
|---|---|
| `_quarter_of_month(month)` | Returns 1–4 for a given month number |
| `_quarter_months(year, quarter)` | Returns the three `YYYY-MM` keys for a quarter |
| `_last_month_of_quarter(year, quarter)` | Returns the last month of a quarter — used for the previous-quarter (`_Q-1`) lookups |
| `_prev_quarter(year, quarter)` | Handles year rollover (Q1 → previous year Q4) |

---

### `app/services/batch_forecast.py`
Orchestrates multi-part forecasting and builds the output Excel workbook.

Two main functions:

- **`read_parts_from_upload(file_bytes)`** — reads the uploaded input Excel, finds `Part Number` and `Tier 1` columns (case-insensitive match), deduplicates while preserving order, and returns a list of `(part_number, tier_1)` tuples.

- **`build_forecast_workbook(part_tier_pairs, engine)`** — runs the engine for every pair, then builds a 13-sheet workbook (1 Summary + 12 monthly sheets). Parts that fail (unknown part, missing price, etc.) produce a red **ERROR row** instead of stopping the entire batch.

**Output column order in each monthly sheet:**

`Part Number → Tier 1 → Weight → Base Price Used → Quarter (Current) → Quarter (Previous) → MC_Q → MC_Q-1 → PPI_Q → PPI_Q-1 → PPI Factor → CNG_Q → CNG_Q-1 → AMS_Q → AMS_Q-1 → AMS Delta → DF_c → Predicted Price`

---

### `app/models/request.py`
Pydantic model for the single-part endpoint body. Validates incoming JSON before it reaches the engine.

```json
{
  "part_number": "09-0052-003",
  "tier_1": "Kadon Aerospace",
  "cng_q": 0.97,
  "cng_q-1": 0.95,
  "include_current_month": "NO"
}
```

`cng_q` and `cng_q-1` are required ($/lb, `cng_q_1` is also accepted) and are used unchanged for every forecast month.

`include_current_month` is optional (`"YES"` / `"NO"`, case-insensitive, default `"NO"`). The batch endpoint accepts it as a multipart form field alongside the file.

---

### `app/models/response.py`
Three nested Pydantic models that define the API response shape:

- **`QuarterContext`** — all 10 intermediate quarterly variables (MC_Q, AMS_Q, PPI_Factor, CNG_Q, AMS_delta, and their previous-quarter counterparts)
- **`MonthForecast`** — one month's result: `predicted_price`, `base_price_used`, `pwt`, `df_c`, and a nested `QuarterContext`
- **`ForecastResponse`** — top-level: `part_number`, `tier_1`, `base_price`, `pwt_lbs`, and a list of 12 `MonthForecast` objects

---

### `app/api/v1/router.py`
Registers the two active endpoint routers under the `/api/v1` prefix:

```python
api_v1_router.include_router(forecast_excel_router)   # /api/v1/forecast-excel
api_v1_router.include_router(forecast_batch_router)   # /api/v1/forecast-batch
```

To add a new endpoint group, create a file in `endpoints/`, define a FastAPI `APIRouter`, and add `include_router()` here.

---

### `app/api/v1/endpoints/forecast_excel.py`
Three routes under `/api/v1/forecast-excel`:

| Method | Path | Purpose |
|---|---|---|
| `GET` | `/parts` | List all Part Number + Tier 1 combinations from Excel with weight and price |
| `POST` | `/reload` | Clear the in-memory Excel cache so edits are picked up without restart |
| `POST` | `` (root) | Single-part forecast — returns JSON with full 12-month breakdown |

Uses **FastAPI dependency injection** (`Depends`) to wire repositories into the engine. To swap data sources, only the two factory functions at the top of this file (`get_market_repo` and `get_part_repo`) need to change.

---

### `app/api/v1/endpoints/forecast_batch.py`
Single route: `POST /api/v1/forecast-batch`

- Accepts `multipart/form-data` file upload (an `.xlsx` with `Part Number` + `Tier 1` columns)
- Returns a binary `.xlsx` download (13-sheet workbook)
- Sets `Content-Disposition: attachment` header so browsers and Axios trigger a file download automatically
- Includes `X-Parts-Count` response header for the frontend to know how many parts were processed

---

### `tests/test_forecast_engine.py`
Unit tests organised into four classes:

| Class | What it covers |
|---|---|
| `TestQuarterHelpers` | All quarter arithmetic utility functions |
| `TestForecastEngineHappyPath` | 12 months returned, correct chaining, manual formula verification, quarter caching |
| `TestForecastEngineErrors` | Unknown part, bad month format, missing market data |
| `TestHardcodedRepositories` | Data layer read behaviour |

```bash
pytest tests/ -v
```

---

## 5. Data Layer

### `aluminium_data.xlsx` — 5 sheets

| Sheet | Columns | Notes |
|---|---|---|
| **Parts** | Part Number, Tier 1, Weight (lbs), Current Price ($) | Unique key = Part Number + Tier 1 |
| **LME** | Month (YYYY-MM), LME Price ($/lb) | Green = actual, Yellow = projected |
| **Midwest** | Month (YYYY-MM), Midwest Premium ($/lb) | Same colour coding |
| **PPI** | Month (YYYY-MM), PPI Index (dimensionless) | BLS index value, not a dollar amount |

**After editing the Excel**, either restart the server or call `POST /api/v1/forecast-excel/reload` to pick up changes without restarting.

**Data coverage required:** The Excel must contain data from at least 2 quarters back through at least 3 quarters forward from today, so the engine can look up previous-quarter values and project 12 months ahead.

---

## 6. API Endpoints

Base URL: `http://localhost:8000`

| Method | Endpoint | Input | Output |
|---|---|---|---|
| `GET` | `/health` | — | `{"status": "ok"}` |
| `GET` | `/api/v1/forecast-excel/parts` | — | JSON list of all parts |
| `POST` | `/api/v1/forecast-excel/reload` | — | Clears Excel cache |
| `POST` | `/api/v1/forecast-excel` | JSON body | JSON 12-month forecast |
| `POST` | `/api/v1/forecast-batch` | `.xlsx` file upload | `.xlsx` file download |

Interactive docs (Swagger UI): **`http://localhost:8000/docs`**
Alternative docs (ReDoc): `http://localhost:8000/redoc`

### Single-part response (abbreviated)
```json
{
  "part_number": "09-0052-003",
  "tier_1": "NA",
  "pwt_lbs": 3.6,
  "base_year_month": "2026-08",
  "base_price": 47.09,
  "forecasts": [
    {
      "year_month": "2026-09",
      "month_label": "September 2026",
      "predicted_price": 57.9267,
      "base_price_used": 47.09,
      "quarter_context": {
        "quarter_label": "Q3-2026",
        "mc_q": 2.5601,
        "ppi_q": 297.637,
        "cng_q": 0.95,
        "ams_q": 4.6365,
        "prev_quarter_label": "Q2-2026",
        "mc_q-1": 1.3298,
        "ppi_q-1": 276.33,
        "cng_q-1": 0.72,
        "ams_q-1": 2.6350,
        "ppi_factor": 0.07710708,
        "ams_delta": 2.0016
      }
    }
    // ... 11 more months
  ]
}
```

---

## 7. Getting Started

### Prerequisites
- Python 3.12 recommended (3.14 may have wheel build issues with some packages)
- `aluminium_data.xlsx` placed in the same folder as `main.py`

### Installation

```bash
cd files/
python -m venv env

# Activate virtual environment
env\Scripts\activate        # Windows
source env/bin/activate     # Mac / Linux

pip install -r requirements.txt
```

### Full `requirements.txt`
```
fastapi==0.115.0
uvicorn[standard]==0.30.6
pydantic==2.9.2
pydantic-settings==2.5.2
python-dateutil==2.9.0
openpyxl>=3.1.0
pandas>=2.0.0
python-multipart>=0.0.9
```

### Run the server
```bash
uvicorn main:app --reload --port 8000
```

### Run the tests
```bash
pytest tests/ -v
```

### Environment variable (optional)
If `aluminium_data.xlsx` is not next to `main.py`, point to it with:
```bash
# Windows
set ALUMINIUM_EXCEL_PATH=C:\full\path\to\aluminium_data.xlsx

# Mac / Linux
export ALUMINIUM_EXCEL_PATH=/full/path/to/aluminium_data.xlsx
```

---

## 8. How to Swap a Data Source

The codebase is built so the **data source can change without touching any business logic**.

### To replace Excel with a live API (e.g. BLS for PPI, LME futures feed)

1. Create `app/data/api_store.py`
2. Write a class that inherits from `MarketDataRepository` (`app/data/base.py`) and implements all four methods
3. In `app/api/v1/endpoints/forecast_excel.py`, update the factory function:

```python
# Before
def get_market_repo() -> MarketDataRepository:
    return ExcelMarketDataRepository()

# After
def get_market_repo() -> MarketDataRepository:
    return LiveApiMarketDataRepository()   # your new class
```

Nothing else changes — the engine, models, and all other endpoints are unaffected.

### To replace the Parts sheet with an ERP system (SAP, Oracle, etc.)

Same pattern — implement `PartRepository` from `base.py` and swap the `get_part_repo()` factory function.

---

## 9. Excel File Format

### Parts sheet — edit this to add or update parts

| Part Number | Tier 1 | Weight (lbs) | Current Price ($) Aug 2026 |
|---|---|---|---|
| 09-0052-003 | NA | 3.6 | 47.09 |
| 09-0052-003 | Kadon Aerospace | 3.6 | 49.50 |
| 2190-1015 | NA | 7.0 | 89.97 |

- **Part Number + Tier 1 = unique key.** The same part number can repeat under different suppliers.
- **Tier 1 = "NA"** means no specific supplier (default value).
- **Current Price** = the actual known price for one month. The column header **must name that month** (e.g. `Current Price ($) Aug 2026`, `August 2026`, `2026-08` or `08/2026`) — the API reads it to decide which months it can forecast. Update the prices *and* the header month together.
  - Header month = last month → use `include_current_month: "YES"` (forecast starts this month).
  - Header month = this month → use `include_current_month: "NO"` (forecast starts next month).

### Market data sheets (LME, Midwest, PPI)

CNG is no longer read from Excel — `cng_q` and `cng_q-1` come from the request.

| Month (YYYY-MM) | Value |
|---|---|
| 2026-07 | 1.4751 |
| 2026-08 | 1.4724 |
| 2026-09 | 1.4751 |

- **Green rows** = historical actuals
- **Yellow rows** = projected / forecasted values (will be replaced by live data in future)
- The engine needs rows from at least **2 quarters back** through **3 quarters forward** from today

---

## 10. Frontend Integration

### Single-part forecast (Axios)
```javascript
const response = await axios.post('/api/v1/forecast-excel', {
  part_number: '09-0052-003',
  tier_1: 'NA'
});

const { forecasts, base_price, pwt_lbs } = response.data;
// forecasts[0].month_label      → "September 2026"
// forecasts[0].predicted_price  → 57.9267
// forecasts[0].base_price_used  → 47.09  (actual for month 1, chained for rest)
```

### Batch forecast — upload input, download result (Axios)
```javascript
const form = new FormData();
form.append('file', excelFile);    // .xlsx with Part Number + Tier 1 columns

const response = await axios.post('/api/v1/forecast-batch', form, {
  responseType: 'blob'             // ← must be 'blob' to receive the binary file
});

// Trigger download in the browser
const url = URL.createObjectURL(response.data);
const a = document.createElement('a');
a.href = url;
a.download = 'forecast_result.xlsx';
a.click();
URL.revokeObjectURL(url);
```

### Input Excel format for the batch endpoint

The uploaded file must have these two column headers (matching is case-insensitive):

| Part Number | Tier 1 |
|---|---|
| 09-0052-003 | NA |
| 2190-1015 | NA |
| 09-0052-003 | Kadon Aerospace |

Duplicate rows (same Part Number + Tier 1 combination) are silently deduplicated.

---

*For questions about the formula or data sources, refer to the quarterly pricing agreement documentation or contact the pricing team.*
