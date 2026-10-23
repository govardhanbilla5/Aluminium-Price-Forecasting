"""
Aluminium Price Forecast Engine.
FORMULA REFERENCE
-----------------
P_next = P_current
         + [(AMS_Q - AMS_Q-1) x PWt]
         + (PPI_Factor x P_current)
Where:
  AMS_Q     = (MC_Q x DF_c) + CNG_Q
  AMS_Q-1   = (MC_Q-1 x DF_c) + CNG_Q-1
  ### MC_Q      = (LME + MWP of the month being PREDICTED) + 0.558883       # CHANGED (offset 0.558883 in config)
  ### MC_Q-1    = (LME + MWP of the LAST month of the previous quarter) + 0.558883  # CHANGED (offset 0.558883 in config)
              when that quarter is covered by actuals, otherwise the       # CHANGED
              AVERAGE of MC of its 3 months
  PPI_Factor = (PPI_Q - PPI_Q-1) / PPI_Q-1
      where PPI_Q   = PPI of the month being PREDICTED
            PPI_Q-1 = PPI of the LAST month of the previous quarter
                      when that whole quarter is covered by published
                      actuals (actuals lag the current month by
                     PPI_ACTUAL_LAG_MONTHS), otherwise the AVERAGE
                      of its 3 (projected) months
  CNG_Q     = supplied in the request, constant for every month
  CNG_Q-1   = supplied in the request, constant for every month
  DF_c      = 1.44 (constant)
  PWt       = part weight in lbs (per part number)
ITERATION LOGIC
---------------
P_current (base price) of a month is NOT the previous predicted price any more:
  - first PRICE_LAG_MONTHS (2) months  -> the original base price
  - later months                       -> predicted price of (month - 2)
    (a price calculated in Oct becomes effective / base in Dec, 50-day lag)
DEADBAND: if the calculated price is within +/-2% of the base price,
the base price is retained (no change).
"""

# Import all the required modules and classes
import logging
from calendar import monthrange
from datetime import date, datetime, timezone
from typing import Optional, NamedTuple
from app.core.config import settings
from app.data.base import MarketDataRepository, PartRepository
from app.models.response import PriceFactorBreakdown, PriceChangeBreakdown
from app.models.response import (
    ForecastResponse,
    MonthForecast,
    QuarterContext,
)



logger = logging.getLogger(__name__)


# ---------------------------------------------------------------------------
# QUARTER HELPERS
# Calendar quarters: Q1 = Jan-Mar, Q2 = Apr-Jun, Q3 = Jul-Sep, Q4 = Oct-Dec.
# Months are always handled as 'YYYY-MM' strings, which sort chronologically.
# ---------------------------------------------------------------------------
def _quarter_of_month(month: int) -> int:
    """Return 1-4 for a given month number (1-12)."""
    return (month - 1) // 3 + 1

def _quarter_months(year: int, quarter: int) -> list[str]:
    """Return the three 'YYYY-MM' keys that make up a calendar quarter."""
    first_month = (quarter - 1) * 3 + 1
    return [f"{year}-{m:02d}" for m in range(first_month, first_month + 3)]

def _last_month_of_quarter(year: int, quarter: int) -> str:
    """Return the 'YYYY-MM' key of the last month of a quarter."""
    last_month = quarter * 3
    return f"{year}-{last_month:02d}"

def _prev_quarter(year: int, quarter: int) -> tuple[int, int]:
    """Return (year, quarter) of the preceding quarter."""
    if quarter == 1:
        return year - 1, 4
    return year, quarter - 1

def _quarter_label(year: int, quarter: int) -> str:
    """Human-readable quarter name used in the response, e.g. 'Q4-2026'."""
    return f"Q{quarter}-{year}"


# ── NEW MODIFICATION ────────────────────────────────────────
def _month_position_in_quarter(month: int) -> int:
    """
    Return 1, 2 or 3 for a given calendar month (1-12), marking its
    position inside its own quarter.
    Quarter 1 (Jan, Feb, Mar)  -> 1, 2, 3
    Quarter 2 (Apr, May, Jun)  -> 1, 2, 3
    Quarter 3 (Jul, Aug, Sep)  -> 1, 2, 3
    Quarter 4 (Oct, Nov, Dec)  -> 1, 2, 3
    Position 3 always means "last month of the quarter".
    """
    return ((month - 1) % 3) + 1

# ---------------------------------------------------------------------------
# DATA FETCH HELPERS  (raise ValueError if data is unavailable)
# ---------------------------------------------------------------------------
def _require(value: Optional[float], description: str) -> float:
    """
    Repositories return None when a value is missing. The formula cannot run
    without it, so turn None into a clear ValueError (the API maps it to 422).
    """
    if value is None:
        raise ValueError(f"Required data unavailable: {description}")
    return value


class _MCResult(NamedTuple):
    """MC for one month plus the LME / Midwest parts it was built from."""
    mc: float       # (LME + Midwest)  + 0.558883   - used in formula
    lme: float      # LME for the month              - used in breakdown
    midwest: float  # Midwest for the month          - used in breakdown


def _compute_mc(
    year_month: str,
    market_repo: MarketDataRepository,
) -> _MCResult:
    """
    MC = (LME + Midwest premium) + MC_OFFSET for a single
    'YYYY-MM' month (0.558883 by default, see config).
    LME and Midwest are also returned so the price-change breakdown can
    attribute the change to each of them separately.
    """
    lme = _require(market_repo.get_lme(year_month), f"LME for {year_month}")
    mwp = _require(
       market_repo.get_midwest_premium(year_month), f"Midwest premium for {year_month}"
    )
    # ── MODIFICATION ────────────────────────────────────────
    # Old formula (no longer used): (lme + mwp) * settings.MC_MULTIPLIER + settings.MC_OFFSET
    mc = (lme + mwp) + settings.MC_OFFSET
    logger.debug("MC_%s: LME=%.4f  MWP=%.4f  MC=%.4f", year_month, lme, mwp, mc)
    return _MCResult(mc=mc, lme=lme, midwest=mwp)



def _ppi_prev_quarter(
    market_repo: MarketDataRepository,
    prev_year: int,
    prev_quarter: int,
    current_month: str,
) -> tuple[float, str]:
    """
    PPI_Q-1 for the previous quarter, plus a short description of how it was
    obtained (for logging). When the whole previous quarter is covered by
    actuals we use the ACTUAL of its last month; otherwise the quarter is
    (partly) projected, so we average all three months instead.
    """
    months = _quarter_months(prev_year, prev_quarter)
    last_month = months[-1]
    # 'YYYY-MM' strings sort chronologically, so plain comparison works.
    if last_month <= _shift_month(current_month, -settings.PPI_ACTUAL_LAG_MONTHS):
        ppi = _require(market_repo.get_ppi(last_month), f"PPI for {last_month}")
        return ppi, f"actual PPI of {last_month}"
    values = [_require(market_repo.get_ppi(m), f"PPI for {m}") for m in months]
    return sum(values) / 3, f"average of projected PPI {months[0]}..{last_month}"



def _compute_quarter_context(
    year: int,
    month: int,
    market_repo: MarketDataRepository,
    df_c: float,
    cng_q: float,
    cng_q_1: float,
    current_month: str,
) -> QuarterContext:
    """
    Build the full QuarterContext for the month being predicted.
 
      MC_Q / PPI_Q    -> the predicted month itself
      MC_Q-1          -> last month of the previous quarter when that quarter
                         is covered by actuals, else the average of its three
                         months (same rule as PPI_Q-1)
      PPI_Q-1         -> last month of the previous quarter when that quarter
                         is fully covered by PPI actuals, else the average of
                         its three months
      CNG_Q / CNG_Q-1 -> fixed values supplied in the request
 
    Raises ValueError if any required data is missing.
    """
    quarter = _quarter_of_month(month)
    prev_year, prev_quarter = _prev_quarter(year, quarter)
    # Current (predicted) month 
    target_month = f"{year}-{month:02d}"

    # Compute MC_Q = (LME + Midwest) + MC_OFFSET for the month being predicted
    mc_result = _compute_mc(target_month, market_repo)

    # PPI_Q = PPI of the month being predicted, Here PPI is fetched for the target month which is being predicted.
    ppi_q = _require(market_repo.get_ppi(target_month), f"PPI for {target_month}")

    # AMS_Q = (MC_Q x DF_C) + CNG_Q, Here AMS calculated using the MC_Q value computed above and the provided CNG_Q
    ams_q = (mc_result.mc * df_c) + cng_q

    # Previous quarter 
    last_month_q_prev = _last_month_of_quarter(prev_year, prev_quarter)

    # ── MODIFICATION ────────────────────────────────────────
    # MC_Q-1 = (LME + Midwest) + MC_OFFSET for the last month of the previous quarter when that quarter is covered by actuals, else the average of its three months (same rule as PPI_Q-1).
    # otherwise the 3-month average (same rule as PPI_Q-1).
    if last_month_q_prev <= _shift_month(current_month, -settings.PPI_ACTUAL_LAG_MONTHS):
        mc_result_prev = _compute_mc(last_month_q_prev, market_repo)
    else:
        prev_mc_list = [
            _compute_mc(m, market_repo) for m in _quarter_months(prev_year, prev_quarter)
        ]
        # Average MC, LME and Midwest separately here
        mc_result_prev = _MCResult(
            mc=sum(r.mc for r in prev_mc_list) / 3,
            lme=sum(r.lme for r in prev_mc_list) / 3,
            midwest=sum(r.midwest for r in prev_mc_list) / 3,
        )
    # ───────────────────────────────────────────────────────────


    # Previous quarter PPI
    # PPI_Q-1 = average of the three months of previous quarter if it is projected data
    ppi_q_prev, ppi_q_prev_basis = _ppi_prev_quarter(
        market_repo, prev_year, prev_quarter, current_month
    )

    # AMS_Q-1 = (MC_Q-1 x DF_C) + CNG_Q-1 Calculate AMS_Q-1 using the MC_Q-1 value computed above and the provided CNG_Q-1
    ams_q_prev = (mc_result_prev.mc * df_c) + cng_q_1
    if ppi_q_prev == 0:
        raise ValueError(
            f"PPI_Q-1 is zero for {_quarter_label(prev_year, prev_quarter)} "
            "- cannot compute PPI_Factor"
        )

    # PPI_Factor = (PPI_Q - PPI_Q-1) / PPI_Q-1, which is used in the main formula for price prediction. The AMS_delta is calculated as the difference between AMS_Q and AMS_Q-1.
    ppi_factor = (ppi_q - ppi_q_prev) / ppi_q_prev
    ams_delta = ams_q - ams_q_prev
    logger.debug(
        "QuarterContext %s: MC_Q=%.4f AMS_Q=%.4f | MC_Q-1(%s)=%.4f AMS_Q-1=%.4f | "
        "PPI_Q-1=%.4f (%s) PPI_Factor=%.6f AMS_delta=%.4f",
        target_month,
        mc_result.mc, ams_q, last_month_q_prev, mc_result_prev.mc, ams_q_prev,
        ppi_q_prev, ppi_q_prev_basis, ppi_factor, ams_delta,
    )
    # Here returning the QuarterContext object with all the computed values for the current and previous quarters, including MC, LME, Midwest, PPI, CNG, AMS, PPI_Factor, and AMS_delta.
    return QuarterContext(
       quarter_label=_quarter_label(year, quarter),
        mc_q=round(mc_result.mc, 6),
        lme_q=round(mc_result.lme, 6),
       midwest_q=round(mc_result.midwest, 6),
        ppi_q=round(ppi_q, 4),
        cng_q=round(cng_q, 6),
        ams_q=round(ams_q, 6),
       prev_quarter_label=_quarter_label(prev_year, prev_quarter),
        mc_q_1=round(mc_result_prev.mc, 6),
        lme_q_1=round(mc_result_prev.lme, 6),
       midwest_q_1=round(mc_result_prev.midwest, 6),
        ppi_q_1=round(ppi_q_prev, 4),
        cng_q_1=round(cng_q_1, 6),
        ams_q_1=round(ams_q_prev, 6),
        ppi_factor=round(ppi_factor, 8),
        ams_delta=round(ams_delta, 6),
    )


# ---------------------------------------------------------------------------
# MONTH LABEL HELPER
# ---------------------------------------------------------------------------
_MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]


def _month_label(year_month: str) -> str:
    """'2026-10' -> 'October 2026'."""
    year, month = year_month.split("-")
    return f"{_MONTH_NAMES[int(month) - 1]} {year}"


def _advance_month(year: int, month: int) -> tuple[int, int]:
    """Return (year, month) one calendar month later."""
    if month == 12:
        return year + 1, 1
    return year, month + 1


def _prev_month(year: int, month: int) -> tuple[int, int]:
    """Return (year, month) one calendar month earlier."""
    if month == 1:
        return year - 1, 12
    return year, month - 1


def _shift_month(year_month: str, delta: int) -> str:
    """Shift a 'YYYY-MM' key by *delta* calendar months (delta may be negative)."""
    year, month = map(int, year_month.split("-"))
    total = year * 12 + (month - 1) + delta
    return f"{total // 12}-{total % 12 + 1:02d}"


def _price_month_error(
    parts: PartRepository,
    part_number: str,
    tier_1: str,
    needed_month: str,
    include_current_month: bool,
) -> ValueError:
    """
    Explain why no P_current was found for *needed_month*.
    Two causes: the price column is for a different month than the one
    needed (tells the user which month to use), or the part has no price.
    """
    data_month = parts.get_price_month()
    if data_month is not None and data_month != needed_month:
        flag = "YES" if include_current_month else "NO"
        return ValueError(
            f"Part prices in the data store are for {_month_label(data_month)}, but "
           f"include_current_month={flag} needs {_month_label(needed_month)} prices "
            "as P_current. Update the price column (and the month in its header) "
            "or change include_current_month."
        )
    return ValueError(
        f"Current price not available for part_number='{part_number}' "
        f"tier_1='{tier_1}' month={needed_month}. Make sure the price column "
        "header names its month, e.g. 'Current Price ($) Aug 2026'."
    )


def _compute_price_breakdown(
    ctx: "QuarterContext",
    current_price: float,
    predicted_price: float,
    pwt: float,
    df_c: float,
) -> PriceChangeBreakdown:
    """
    Decompose the price change into four exact, non-overlapping factors:
 
        total_change = lme_effect + midwest_effect + cng_effect + ppi_effect
 
    NOTE: the caller passes the price BEFORE the deadband, so sum_check stays 0.
    """
    total_change = round(predicted_price - current_price, 6)


    # Month-over-previous-quarter movement of each input
    lme_delta = ctx.lme_q - ctx.lme_q_1
    midwest_delta = ctx.midwest_q - ctx.midwest_q_1
    cng_delta = ctx.cng_q - ctx.cng_q_1

    # ALL the effects are computed using the new formula and the MC_MULTIPLIER is no longer used in the calculations. The effects are calculated based on the deltas of LME, Midwest, CNG, and PPI, multiplied by the part weight and correction factor where applicable.
    # ── MODIFICATION ────────────────────────────────────────
    # Removed the MC_MULTIPLIER from the calculations as per the new requirements from thapase/srikanth on 09/30/2026. 
    lme_effect = round(lme_delta * df_c * pwt, 6)#####round(lme_delta * settings.MC_MULTIPLIER * df_c * pwt, 6)
    midwest_effect = round(midwest_delta * df_c * pwt, 6)  #####round(midwest_delta * settings.MC_MULTIPLIER * df_c * pwt, 6)
    cng_effect = round(cng_delta * pwt, 6)
    ppi_effect = round(ctx.ppi_factor * current_price, 6)
    sum_of_effects = round(lme_effect + midwest_effect + cng_effect + ppi_effect, 6)
    sum_check = round(sum_of_effects - total_change, 6)

    def _pct(effect: float) -> float:
        # Percentage of base price - always stable and interpretable.
        return round((effect / current_price) * 100, 4)
    return PriceChangeBreakdown(
        base_price=round(current_price, 4), #CHANGED
       predicted_price=round(predicted_price, 4),   # CHANGED
        total_change=total_change,
        lme_effect=PriceFactorBreakdown(
            dollar_effect=lme_effect,
           percentage_of_base=_pct(lme_effect),
        ),
       midwest_effect=PriceFactorBreakdown(
            dollar_effect=midwest_effect,
           percentage_of_base=_pct(midwest_effect),
        ),
        cng_effect=PriceFactorBreakdown(
            dollar_effect=cng_effect,
           percentage_of_base=_pct(cng_effect),
        ),
        ppi_effect=PriceFactorBreakdown(
            dollar_effect=ppi_effect,
           percentage_of_base=_pct(ppi_effect),
        ),
        sum_check=sum_check,
    )


# ======================================================================
# NEW - AOP (Annual Operating Plan) helpers.
# To get AOP quantity values
# This is the new feature as per thapase/srikanth requirements on 10/20/2026

_AOP_MONTH_FORMATS = ("%B %Y", "%b %Y", "%B-%Y", "%b-%Y", "%b'%y", "%b-%y", "%b %y", "%Y-%m")
def _aop_month_key(header) -> Optional[str]:
    """Convert a header (Excel date, 'Oct 2026', 'Oct-2026', '2026-10' ...) to 'YYYY-MM'."""
    if isinstance(header, (datetime, date)):
        return f"{header.year}-{header.month:02d}"

    text = str(header or "").strip()
    for fmt in _AOP_MONTH_FORMATS:
        try:
            parsed = datetime.strptime(text, fmt)
            return f"{parsed.year}-{parsed.month:02d}"
        except ValueError:
            continue
    return None

def _aop_part_key(value) -> str:
    """Excel may store part 221134 as 221134.0 - normalise to plain text."""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value).strip()


def build_aop_lookup(rows: list[list]) -> dict[tuple[str, str, str], float]:
    """
    AOP sheet as list of rows. Row 0 = header:
        Part Number | Tier 1 | Oct 2026 | Nov 2026 | ...
    Returns {(part_number, tier_1, 'YYYY-MM'): aop_qty}. Blank / non-numeric cells are skipped.
    """
    if not rows:
        return {}
    month_cols = {i: _aop_month_key(h) for i, h in enumerate(rows[0]) if i >= 2 and _aop_month_key(h)}
    lookup: dict[tuple[str, str, str], float] = {}
    for row in rows[1:]:
        if not row or row[0] in (None, ""):
            continue
        part = _aop_part_key(row[0])
        tier = _aop_part_key(row[1]) if len(row) > 1 else ""
        for col, ym in month_cols.items():
            qty = row[col] if col < len(row) else None
            if qty in (None, ""):
                continue
            try:
                lookup[(part, tier, ym)] = float(qty)
            except (TypeError, ValueError):
                continue
    return lookup

def compute_total_aop(unit_price: float, aop_qty: Optional[float]) -> Optional[float]:
    """Total AOP value = AOP qty x unit price. No AOP qty -> None (blank cell)."""
    if aop_qty is None:
        return None
    return round(unit_price * aop_qty, 2)
# ======================================================================



# ---------------------------------------------------------------------------
# MAIN ENGINE
# ---------------------------------------------------------------------------
class ForecastEngine:
    """
    Stateless forecasting engine. Depends only on the abstract repository
    interfaces - no concrete data source is referenced here.
    """
    def __init__(
        self,
        market_repo: MarketDataRepository,
        part_repo: PartRepository,
    ) -> None:
        self._market = market_repo  # LME / Midwest / PPI source
        self._parts = part_repo     # part weight and base price source
        self._df_c = settings.DF_C  # Drauss factore sourece
        self._horizon = settings.FORECAST_HORIZON_MONTHS
    # ------------------------------------------------------------------
    # PUBLIC API
    # ------------------------------------------------------------------
    def forecast(
        self,
        part_number: str,
        tier_1: str,
        cng_q: float,
        cng_q_1: float,
        include_breakdown: bool = False,
        include_current_month: bool = False,
    ) -> ForecastResponse:
        """
        Produce a 12-month price forecast for *part_number* + *tier_1*.
        include_current_month:
            False (default) -> forecast the next 12 months, starting next month.
            True            -> start at the current month and forecast 13 months.
        Raises ValueError if the part+tier_1 combination is unknown or price unavailable.
        """

        # "Today" is taken from the UTC clock; it decides which months are
        # forecast and which market data counts as actual vs projected.
        now_utc = datetime.now(timezone.utc)
        current_month = now_utc.strftime("%Y-%m")
        # P_current = price of the month just before the first forecast month:
        #   include_current_month -> LAST month's price, forecast starts this month
        #   otherwise             -> THIS month's price, forecast starts next month
        if include_current_month:
            price_year, price_month = _prev_month(now_utc.year, now_utc.month)
            base_year_month = f"{price_year}-{price_month:02d}"
        else:
            base_year_month = current_month
        logger.info(
            "Starting forecast: part=%s  tier_1=%s  current_month=%s  price_month=%s",
            part_number, tier_1, current_month, base_year_month,
        )

        # -- Part weight -----------------------------------------------------
        # Part weight is used in the formula: P_next = P_current + [(AMS_Q - AMS_Q-1) x PWt] + (PPI_Factor x P_current)
        # Here calling the get_part_weight method of the PartRepository to fetch the part weight for the given part_number and tier_1. If the part weight is not found, a ValueError is raised with a message indicating that the part was not found and suggesting to check that both values match the data store exactly.
        pwt = self._parts.get_part_weight(part_number, tier_1)
        if pwt is None:
            raise ValueError(
                f"Part not found: part_number='{part_number}' tier_1='{tier_1}'. "
                "Check that both values match the data store exactly."
            )
        # -- Base price - always from data store ------------------------------
        # Base price is used in the formula: P_next = P_current + [(AMS_Q - AMS_Q-1) x PWt] + (PPI_Factor x P_current)
        # Here calling the get_base_price method of the PartRepository to fetch the base price for the given part_number, tier_1, and base_year_month. If the base price is not found, a ValueError is raised with a message indicating that the current price is not available for the given part_number, tier_1, and month, and suggesting to make sure the price column header names its month or change include_current_month.
        base_price = self._parts.get_base_price(part_number, tier_1, base_year_month)
        if base_price is None:
            raise _price_month_error(
                self._parts, part_number, tier_1, base_year_month, include_current_month
            )
        logger.info(
            "Part=%s  Tier1=%s PWt=%.2f lbs  P_base=%.4f $/lb",
            part_number, tier_1, pwt, base_price,
        )


        # -- Iterate over 12 future months ------------------------------------
        base_year, base_month = map(int, base_year_month.split("-"))
        current_price = base_price
        forecasts: list[MonthForecast] = []


        # First forecast month is always the month after the price month.
        # include_current_month -> 13 months (this month + next 12), else 12.
        # Here calling the _advance_month function to get the year and month of the first forecast month, which is always the month after the base price month. The number of months to forecast is determined by whether include_current_month is True or False, with 13 months being produced when include_current_month is True and 12 months otherwise.
        forecast_year, forecast_month = _advance_month(base_year, base_month)
        n_months = self._horizon + 1 if include_current_month else self._horizon

        # ── MODIFICATION ────────────────────────────────────────
        # NEW: 50-day implementation lag (2 months) and +/-2% deadband.
        # Read from settings if defined there, otherwise use these defaults.
        ## price_lag = getattr(settings, "PRICE_LAG_MONTHS", 2)
        deadband = getattr(settings, "DEADBAND_PCT", 0.02)

        # ── NEW MODIFICATION ────────────────────────────────────────
        active_base = base_price
        pending_first_month_price: Optional[float] = None

        # NEW: price calculated in each step (goes live `price_lag` months later)
        ## calculated_prices: list[float] = []

        for step in range(n_months):
            year_month_key = f"{forecast_year}-{forecast_month:02d}"
            logger.debug("Step %d: forecasting %s", step + 1, year_month_key)
            # CHANGED: base price (P_current) of this month
            #   first `price_lag` months -> original base price (e.g. Sep for Oct, Nov)
            #   later months             -> price calculated `price_lag` months ago
            #                               (Oct's price is the base for Dec, ...)

            # ── NEW MODIFICATION ────────────────────────────────────────
            position = _month_position_in_quarter(forecast_month)
            is_first_month_of_quarter = (position == 1)
            is_last_month_of_quarter = (position == 3)
            #   1st or 2nd month of the quarter -> active_base (unchanged all quarter)
            #   3rd (last) month of the quarter -> this quarter's own 1st month price
            if is_last_month_of_quarter and pending_first_month_price is not None:
                current_price = pending_first_month_price
            else:
                current_price = active_base

            # -- Build context for this month (MC_Q / PPI_Q vary per month) --
            try:
                ctx = _compute_quarter_context(
                    forecast_year, forecast_month, self._market, self._df_c,
                    cng_q, cng_q_1, current_month,
                )
            except ValueError as exc:
                raise ValueError(
                    f"Cannot compute quarter context for {year_month_key}: {exc}"
                ) from exc
            
            # -- Core formula (unchanged, applied EVERY month on its base price) --
            # P_next = P_current + [(AMS_Q - AMS_Q-1) x PWt] + (PPI_Factor x P_current)
            price_adjustment = ctx.ams_delta * pwt
            ppi_adjustment = ctx.ppi_factor * current_price
            raw_price = current_price + price_adjustment + ppi_adjustment

            # ── MODIFICATION ────────────────────────────────────────
            # -- Factor breakdown (only when requested) ------------------------
            # -- NEW: deadband (final validation step) --
            # within +/-2% of the base price -> base price is retained
            # This is the new feature as per thapase/srikanth requirements on 09/30/2026
            if current_price != 0 and abs(raw_price - current_price) / abs(current_price) <= deadband:
                predicted_price = current_price
            else:
                predicted_price = raw_price
            logger.debug(
                "%s: P_current=%.4f  AMS_delta=%.4f  PWt=%.2f "
                "price_adj=%.4f  ppi_adj=%.4f P_raw=%.4f P_predicted=%.4f",
                year_month_key,
                current_price,
                ctx.ams_delta,
                pwt,
                price_adjustment,
                ppi_adjustment,
                raw_price,
                predicted_price,
            )

            # CHANGED: uses raw_price (before deadband) so sum_check stays 0
            # Deadband (final validation step)
            breakdown = (
                _compute_price_breakdown(
                    ctx=ctx,
                   current_price=current_price,
                   predicted_price=raw_price,
                    pwt=pwt,
                    df_c=self._df_c,
                )
                if include_breakdown else None
            )
            forecasts.append(
                MonthForecast(
                   year_month=year_month_key,
                   month_label=_month_label(year_month_key),
                   predicted_price=round(predicted_price, 4),   ##### NEW
                   predicted_price_without_deadband=round(raw_price, 4),   # ####NEW
                   base_price_used=round(current_price, 4),         ###NEW
                   pwt=pwt,
                   df_c=self._df_c,
                   quarter_context=ctx,
                   price_change_breakdown=breakdown,
                   is_data_projected=(step >= 1),
                )
            )

            # ── NEW MODIFICATION ────────────────────────────────────────
            # CHANGED: remember this quarter's 1st-month result for later use at its 3rd month
            if is_first_month_of_quarter:
                pending_first_month_price = predicted_price
 
            # CHANGED: when this was the quarter's last month, the base it just used
            # becomes the active base going forward (carried into the next quarter's
            # 1st and 2nd months), and we reset for the next quarter.
            if is_last_month_of_quarter:
                active_base = current_price
                pending_first_month_price = None
    
            forecast_year, forecast_month = _advance_month(forecast_year, forecast_month)
        logger.info(
            "Forecast complete for part=%s: %d months, "
            "range %s -> %s, price %.4f -> %.4f",
            part_number,
            len(forecasts),
            forecasts[0].year_month,
            forecasts[-1].year_month,
            forecasts[0].predicted_price,
            forecasts[-1].predicted_price,
        )
        return ForecastResponse(
            part_number=part_number,
            tier_1=tier_1,
            pwt_lbs=pwt,
            base_year_month=base_year_month,
            base_price=base_price,
            forecast_generated_at=now_utc.isoformat(),
            forecasts=forecasts,
        )


    # ======================================================================
    # NEW - AOP (Annual Operating Plan) helpers.
    # This is the new feature as per thapase/srikanth requirements on 10/20/2026
    # ------------------------------------------------------------------ 
    def apply_aop( 
        self, 
        response: "ForecastResponse", 
        aop_lookup: dict[tuple[str, str, str], float], 
    ) -> dict[str, tuple[Optional[float], Optional[float]]]: 
        """ 
        Total AOP for one part, month by month. 
        Returns {year_month: (aop_qty, total_aop_value)}. 
        Uses MonthForecast.predicted_price (final price after deadband = Excel col S). 
        To use the raw price (col R) use m.predicted_price_without_deadband instead. 
        """ 
        part = _aop_part_key(response.part_number) 
        tier = _aop_part_key(response.tier_1) 
        result: dict[str, tuple[Optional[float], Optional[float]]] = {} 
        for m in response.forecasts: 
            qty = aop_lookup.get((part, tier, m.year_month)) 
            result[m.year_month] = (qty, compute_total_aop(m.predicted_price, qty)) 
        return result
    # ======================================================================
