"""
Aluminium Price Forecast Engine.
FORMULA REFERENCE
─────────────────
P_next = P_current
         + [(AMS_Q - AMS_Q-1) × PWt]
         + (PPI_Factor × P_current)
Where:
  AMS_Q     = (MC_Q  × DF_c) + CNG_Q
  AMS_Q-1   = (MC_Q-1 × DF_c) + CNG_Q-1
  MC_Q      = (LME + MWP of the month being PREDICTED) × 0.91 + 1.25
  MC_Q-1    = (LME + MWP of the LAST month of the previous quarter) × 0.91 + 1.25
              (previous to the quarter containing the predicted month)
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
───────────────
We are sitting in Feb 2026 (current month).
We predict March 2026 first, then April, …, up to Jan 2027 (12 months).
Each predicted price becomes the P_current for the next step.
MC_Q and PPI_Q change with every predicted month; MC_Q-1 and PPI_Q-1 are
anchored to the previous quarter; CNG_Q and CNG_Q-1 are fixed request inputs.
Each month uses the rolling predicted price as its P_current.
"""
import logging
from calendar import monthrange
from datetime import date, datetime, timezone
from typing import Optional,NamedTuple
from app.core.config import settings
from app.data.base import MarketDataRepository, PartRepository
from app.models.response import PriceFactorBreakdown, PriceChangeBreakdown
from app.models.response import (
    ForecastResponse,
    MonthForecast,
    QuarterContext,
)
logger = logging.getLogger(__name__)
# ─────────────────────────────────────────────────────────────────────────────
# QUARTER HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _quarter_of_month(month: int) -> int:
    """Return 1-4 for a given month number (1-12)."""
    return (month - 1) // 3 + 1

def _quarter_months(year: int, quarter: int) -> list[str]:
    """
    Return the three 'YYYY-MM' keys that make up a calendar quarter.
    Quarter 1 → Jan, Feb, Mar  (months 1, 2, 3)
    Quarter 2 → Apr, May, Jun  (months 4, 5, 6)
    Quarter 3 → Jul, Aug, Sep  (months 7, 8, 9)
    Quarter 4 → Oct, Nov, Dec  (months 10, 11, 12)
    """
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
    return f"Q{quarter}-{year}"

# ─────────────────────────────────────────────────────────────────────────────
# DATA FETCH HELPERS  (raise ValueError if data is unavailable)
# ─────────────────────────────────────────────────────────────────────────────
def _require(value: Optional[float], description: str) -> float:
    if value is None:
        raise ValueError(f"Required data unavailable: {description}")
    return value


class _MCResult(NamedTuple):
    mc:      float   # (LME + Midwest) × 0.91 + 1.25  — used in formula
    lme:     float   # LME for the month              — used in breakdown
    midwest: float   # Midwest for the month          — used in breakdown

def _compute_mc(
    year_month: str,
    market_repo: MarketDataRepository,
) -> _MCResult:
    """
    MC = (LME + Midwest premium) × MC_MULTIPLIER + MC_OFFSET for a single
    'YYYY-MM' month (0.91 and 1.25 by default, see config).
    Returns a named tuple with the combined value AND the two raw components
    separately so the caller can use them for factor breakdown without
    re-fetching data.
    """
    lme = _require(market_repo.get_lme(year_month), f"LME for {year_month}")
    mwp = _require(
        market_repo.get_midwest_premium(year_month), f"Midwest premium for {year_month}"
    )
    mc = (lme + mwp) * settings.MC_MULTIPLIER + settings.MC_OFFSET
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
    obtained (for logging).

    Published PPI actuals lag the current month by PPI_ACTUAL_LAG_MONTHS
    (2 → sitting in September, actuals run through July); later rows in the
    sheet are projections. When the whole previous quarter is covered by
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
      MC_Q / PPI_Q     → the predicted month itself
      MC_Q-1           → last month of the previous quarter
      PPI_Q-1          → last month of the previous quarter when that quarter
                         is fully covered by PPI actuals, else the average of
                         its three months
      CNG_Q / CNG_Q-1  → fixed values supplied in the request
    Raises ValueError if any required data is missing.
    """
    quarter = _quarter_of_month(month)
    prev_year, prev_quarter = _prev_quarter(year, quarter)
    # ── Current (predicted) month ─────────────────────────────────────────
    target_month = f"{year}-{month:02d}"
    mc_result = _compute_mc(target_month, market_repo)
    ppi_q = _require(market_repo.get_ppi(target_month), f"PPI for {target_month}")
    ams_q = (mc_result.mc * df_c) + cng_q

    # ── Previous quarter ──────────────────────────────────────────────────
    last_month_q_prev = _last_month_of_quarter(prev_year, prev_quarter)
    mc_result_prev = _compute_mc(last_month_q_prev, market_repo)
    ppi_q_prev, ppi_q_prev_basis = _ppi_prev_quarter(
        market_repo, prev_year, prev_quarter, current_month
    )
    ams_q_prev = (mc_result_prev.mc * df_c) + cng_q_1
    # ── Derived factors ──────────────────────────────────────────────────
    if ppi_q_prev == 0:
        raise ValueError(
            f"PPI_Q-1 is zero for {_quarter_label(prev_year, prev_quarter)} "
            "— cannot compute PPI_Factor"
        )
    ppi_factor = (ppi_q - ppi_q_prev) / ppi_q_prev
    ams_delta = ams_q - ams_q_prev
    logger.debug(
        "QuarterContext %s: MC_Q=%.4f AMS_Q=%.4f | MC_Q-1(%s)=%.4f AMS_Q-1=%.4f | "
        "PPI_Q-1=%.4f (%s) PPI_Factor=%.6f AMS_delta=%.4f",
        target_month,
        mc_result.mc, ams_q, last_month_q_prev, mc_result_prev.mc, ams_q_prev,
        ppi_q_prev, ppi_q_prev_basis, ppi_factor, ams_delta,
    )
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
# ─────────────────────────────────────────────────────────────────────────────
# MONTH LABEL HELPER
# ─────────────────────────────────────────────────────────────────────────────
_MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]
def _month_label(year_month: str) -> str:
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
    """Explain why no P_current was found for *needed_month*."""
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
    Decompose the price change into four exact, non-overlapping factors.
    Formula proof:
      total_change = predicted - current
                   = [(AMS_Q - AMS_Q-1) × PWt] + [PPI_Factor × P_current]
                   = [(MC_Q - MC_Q-1) × DF_c + (CNG_Q - CNG_Q-1)] × PWt
                     + PPI_Factor × P_current
                   = (lme_delta + midwest_delta) × 0.91 × DF_c × PWt
                     + cng_delta × PWt
                     + ppi_factor × P_current
    All four effects are mathematically exact and always sum to total_change.
    """
    
    total_change = round(predicted_price - current_price, 6)
    lme_delta     = ctx.lme_q     - ctx.lme_q_1
    midwest_delta = ctx.midwest_q - ctx.midwest_q_1
    cng_delta     = ctx.cng_q     - ctx.cng_q_1
    # MC_Q - MC_Q-1 = MC_MULTIPLIER × (lme_delta + midwest_delta); offset cancels
    lme_effect     = round(lme_delta     * settings.MC_MULTIPLIER * df_c * pwt, 6)
    midwest_effect = round(midwest_delta * settings.MC_MULTIPLIER * df_c * pwt, 6)
    cng_effect     = round(cng_delta     * pwt,        6)
    ppi_effect     = round(ctx.ppi_factor * current_price, 6)
    sum_of_effects = round(lme_effect + midwest_effect + cng_effect + ppi_effect, 6)
    sum_check      = round(sum_of_effects - total_change, 6)

    def _pct(effect: float) -> float:
        # Percentage of base price — always stable and interpretable.
        # e.g. +0.56% means this factor added 0.56% upward pressure on the price,
        # regardless of how small the net total change is.
        return round((effect / current_price) * 100, 4)
    
    return PriceChangeBreakdown(
        base_price=round(current_price, 4),
        predicted_price=round(predicted_price, 4),
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
# ─────────────────────────────────────────────────────────────────────────────
# MAIN ENGINE
# ─────────────────────────────────────────────────────────────────────────────
class ForecastEngine:
    """
    Stateless forecasting engine.  Depends only on the abstract repository
    interfaces — no concrete data source is referenced here.
    """
    def __init__(
        self,
        market_repo: MarketDataRepository,
        part_repo: PartRepository,
    ) -> None:
        self._market = market_repo
        self._parts = part_repo
        self._df_c = settings.DF_C
        self._horizon = settings.FORECAST_HORIZON_MONTHS
    # ─────────────────────────────────────────────────────────────────────
    # PUBLIC API
    # ─────────────────────────────────────────────────────────────────────
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
        Parameters
        ---------
        part_number:
            The part to forecast.
        tier_1:
            Supplier name (e.g. "Kadon Aerospace", "NA").
            Together with part_number forms the unique lookup key.
        cng_q, cng_q_1:
            CNG_Q and CNG_Q-1 ($/lb) from the request; used unchanged for
            every forecast month.
        include_current_month:
            False (default) → forecast the next 12 months, starting next month.
            True            → start at the current month and forecast 13 months
                              (current month + next 12).
            P_current for the first step is the price of the month just before
            the first forecast month: this month's price for False, last
            month's price for True. The data store must hold prices for that
            month, otherwise a ValueError explains the mismatch.
        Returns
        ------
        ForecastResponse
        Raises
        -----
        ValueError
            If the part+tier_1 combination is unknown or price unavailable.
        """
        now_utc = datetime.now(timezone.utc)
        current_month = now_utc.strftime("%Y-%m")
        # P_current = price of the month just before the first forecast month:
        #   include_current_month → LAST month's price, forecast starts this month
        #   otherwise             → THIS month's price, forecast starts next month
        if include_current_month:
            price_year, price_month = _prev_month(now_utc.year, now_utc.month)
            base_year_month = f"{price_year}-{price_month:02d}"
        else:
            base_year_month = current_month
        logger.info(
            "Starting forecast: part=%s  tier_1=%s  current_month=%s  price_month=%s",
            part_number, tier_1, current_month, base_year_month,
        )
        # ── Part weight ───────────────────────────────────────────────────
        pwt = self._parts.get_part_weight(part_number, tier_1)
        if pwt is None:
            raise ValueError(
                f"Part not found: part_number='{part_number}' tier_1='{tier_1}'. "
                "Check that both values match the data store exactly."
            )
        # ── Base price — always from data store ───────────────────────────
        base_price = self._parts.get_base_price(part_number, tier_1, base_year_month)
        if base_price is None:
            raise _price_month_error(
                self._parts, part_number, tier_1, base_year_month, include_current_month
            )
        logger.info(
            "Part=%s  Tier1=%s  PWt=%.2f lbs  P_base=%.4f $/lb",
            part_number, tier_1, pwt, base_price,
        )
        # ── Iterate over 12 future months ─────────────────────────────────
        base_year, base_month = map(int, base_year_month.split("-"))
        current_price = base_price
        forecasts: list[MonthForecast] = []
        # First forecast month is always the month after the price month.
        # include_current_month → 13 months (this month + next 12), else 12.
        forecast_year, forecast_month = _advance_month(base_year, base_month)
        n_months = self._horizon + 1 if include_current_month else self._horizon
        for step in range(n_months):
            year_month_key = f"{forecast_year}-{forecast_month:02d}"
            logger.debug("Step %d: forecasting %s", step + 1, year_month_key)
            # ── Build context for this month (MC_Q / PPI_Q vary per month) ─
            try:
                ctx = _compute_quarter_context(
                    forecast_year, forecast_month, self._market, self._df_c,
                    cng_q, cng_q_1, current_month,
                )
            except ValueError as exc:
                raise ValueError(
                    f"Cannot compute quarter context for {year_month_key}: {exc}"
                ) from exc
            # ── Core formula ──────────────────────────────────────────────
            # P_next = P_current + [(AMS_Q - AMS_Q-1) × PWt] + (PPI_Factor × P_current)
            # ── THIS LINE IS NOT TOUCHED BY THE BREAKDOWN FEATURE ─────────
            price_adjustment = ctx.ams_delta * pwt
            ppi_adjustment   = ctx.ppi_factor * current_price
            predicted_price  = current_price + price_adjustment + ppi_adjustment
            logger.debug(
                "%s: P_current=%.4f  AMS_delta=%.4f  PWt=%.2f  "
                "price_adj=%.4f  ppi_adj=%.4f  P_predicted=%.4f",
                year_month_key,
                current_price,
                ctx.ams_delta,
                pwt,
                price_adjustment,
                ppi_adjustment,
                predicted_price,
            )
                
                
            # ── Factor breakdown (only when requested) ────────────────────
            breakdown = (
                _compute_price_breakdown(
                    ctx=ctx,
                    current_price=current_price,
                    predicted_price=predicted_price,
                    pwt=pwt,
                    df_c=self._df_c,
                )
                if include_breakdown else None
            )
            forecasts.append(
                MonthForecast(
                    year_month=year_month_key,
                    month_label=_month_label(year_month_key),
                    predicted_price=round(predicted_price, 4),
                    base_price_used=round(current_price, 4),
                    pwt=pwt,
                    df_c=self._df_c,
                    quarter_context=ctx,
                    price_change_breakdown=breakdown,
                    is_data_projected=(step >= 1),
                )
            )
            # Roll price forward for the next iteration
            current_price = predicted_price
            forecast_year, forecast_month = _advance_month(forecast_year, forecast_month)
        logger.info(
            "Forecast complete for part=%s: %d months, "
            "range %s → %s, price %.4f → %.4f",
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
