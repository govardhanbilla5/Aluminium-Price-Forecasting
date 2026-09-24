"""
Pydantic response models.
Every field that feeds the formula is surfaced so the caller can
validate and debug individual intermediate values.
"""
from typing import Optional
from pydantic import BaseModel, ConfigDict, Field

class QuarterContext(BaseModel):
    """Quarterly aggregates used in the price formula for one forecast step.

    Previous-quarter fields are named *_q_1 in Python and serialised as
    *_q-1 in JSON (e.g. ppi_q_1 → "ppi_q-1").
    """
    # Allow construction by Python field name (mc_q_1=...) as well as alias.
    model_config = ConfigDict(populate_by_name=True)
    # Current quarter
    quarter_label: str = Field(
        ..., description="Human-readable label, e.g. 'Q1-2026'"
    )
    mc_q: float = Field(
        ..., description="MC_Q: (LME + Midwest of the predicted month) × 0.91 + 1.25 ($/lb)"
    )
    lme_q: float = Field(
        ..., description="LME of the predicted month ($/lb)"
    )
    midwest_q: float = Field(
        ..., description="Midwest premium of the predicted month ($/lb)"
    )
    ppi_q: float = Field(
        ..., description="PPI_Q: PPI of the predicted month"
    )
    cng_q: float = Field(
        ..., description="CNG_Q: from the request, same for every month ($/lb)"
    )
    ams_q: float = Field(
        ..., description="AMS_Q = (MC_Q × DF_c) + CNG_Q ($/lb)"
    )
    # Previous quarter
    prev_quarter_label: str = Field(
        ..., description="Human-readable label, e.g. 'Q4-2025'"
    )
    mc_q_1: float = Field(
        ..., alias="mc_q-1", description="MC_Q-1: (LME + Midwest of last month of previous quarter) × 0.91 + 1.25 ($/lb)"
    )
    lme_q_1: float = Field(
        ..., alias="lme_q-1", description="LME of last month of previous quarter ($/lb)"
    )
    midwest_q_1: float = Field(
        ..., alias="midwest_q-1", description="Midwest premium of last month of previous quarter ($/lb)"
    )
    ppi_q_1: float = Field(
        ..., alias="ppi_q-1", description=(
            "PPI_Q-1: actual PPI of the last month of the previous quarter when "
            "that quarter is fully covered by published actuals (which end 2 "
            "months before the current month); otherwise the average of its "
            "three (projected) months"
        )
    )
    cng_q_1: float = Field(
        ..., alias="cng_q-1", description="CNG_Q-1: from the request, same for every month ($/lb)"
    )
    ams_q_1: float = Field(
        ..., alias="ams_q-1", description="AMS_Q-1 = (MC_Q-1 × DF_c) + CNG_Q-1 ($/lb)"
    )
    # Derived factors
    ppi_factor: float = Field(
        ..., description="PPI_Factor = (PPI_Q - PPI_Q-1) / PPI_Q-1"
    )
    ams_delta: float = Field(
        ..., description="AMS_Q - AMS_Q-1 ($/lb)"
    )

class PriceFactorBreakdown(BaseModel):
    """
    Contribution of a single factor to the price for one forecast month.
    dollar_effect        : how many dollars this factor added (or subtracted).
                           Positive = pushed price up, Negative = pushed price down.
    percentage_of_base   : dollar_effect expressed as a % of the base (current) price.
                           e.g. -0.56% means this factor exerted a 0.56% downward
                           pressure on the price. Always stable and interpretable —
                           unaffected by how large or small the net price change is.
    """
    dollar_effect:      float = Field(..., description="Dollar contribution of this factor ($)")
    percentage_of_base: float = Field(
        ...,
        description="Contribution as % of base price (dollar_effect / base_price × 100)"
    )

class PriceChangeBreakdown(BaseModel):
    """
    Full decomposition of what drove the price change for one forecast month.
    The four factors are mathematically exact and always sum to total_change:
      total_change = lme_effect + midwest_effect + cng_effect + ppi_effect
    Derivation:
      P_next - P_current
        = [(AMS_Q - AMS_Q-1) × PWt]  +  [PPI_Factor × P_current]
        = [(MC_Q - MC_Q-1) × DF_c + (CNG_Q - CNG_Q-1)] × PWt  +  PPI_effect
      MC = (LME + Midwest) × 0.91 + 1.25, so the 1.25 cancels:
      MC_Q - MC_Q-1 = 0.91 × [(lme_q - lme_q-1) + (midwest_q - midwest_q-1)]
      So:
        lme_effect     = (lme_q - lme_q-1) × 0.91 × DF_c × PWt
        midwest_effect = (midwest_q - midwest_q-1) × 0.91 × DF_c × PWt
        cng_effect     = (CNG_Q - CNG_Q-1) × PWt
        ppi_effect     = PPI_Factor × P_current
    """
    base_price:      float = Field(..., description="P_current used as base for this month ($)")
    predicted_price: float = Field(..., description="Predicted price for this month ($)")
    total_change:    float = Field(..., description="predicted_price - base_price ($)")
    lme_effect:     PriceFactorBreakdown = Field(..., description="Contribution from LME price movement")
    midwest_effect: PriceFactorBreakdown = Field(..., description="Contribution from Midwest premium movement")
    cng_effect:     PriceFactorBreakdown = Field(..., description="Contribution from CNG cost movement")
    ppi_effect:     PriceFactorBreakdown = Field(..., description="Contribution from PPI inflation")
    sum_check:      float = Field(
        ...,
        description=(
            "lme + midwest + cng + ppi — should equal total_change exactly. "
            "Any non-zero value here indicates a rounding artefact."
        ),
    )
    
class MonthForecast(BaseModel):
    """Complete forecast result for a single month."""
    year_month: str = Field(
        ..., description="Month being forecast, format YYYY-MM"
    )
    month_label: str = Field(
        ..., description="Human-readable month label, e.g. 'March 2026'"
    )
    predicted_price: float = Field(
        ..., description="Forecasted aluminium price for this part ($/lb)"
    )
    base_price_used: float = Field(
        ...,
        description=(
            "P_current used as input to forecast this month. "
            "For the first forecast month this is the known actual price."
        ),
    )
    pwt: float = Field(..., description="Part weight in lbs (PWt)")
    df_c: float = Field(..., description="Density/conversion factor DF_c (constant)")
    quarter_context: QuarterContext = Field(
        ..., description="All intermediate quarterly variables used in the formula"
    )
    price_change_breakdown: Optional[PriceChangeBreakdown] = Field(
        default=None,
        description=(
            "Factor-level decomposition of what drove the price change. "
            "Present on the /forecast-excel single-part endpoint. "
            "Null on batch and ML endpoints."
        ),
    )
    is_data_projected: bool = Field(
        ...,
        description=(
            "True when one or more input values for this month came from "
            "projected (not actual) data."
        ),
    )
class ForecastResponse(BaseModel):
    """Top-level API response for a 12-month price forecast."""
    part_number: str
    tier_1: str = Field(..., description="Tier 1 supplier name")
    pwt_lbs: float = Field(..., description="Part weight in lbs")
    base_year_month: str = Field(
        ...,
        description=(
            "Month of the known price used as the first P_current: the current "
            "month by default, or last month when include_current_month='YES'"
        ),
    )
    base_price: float = Field(
        ..., description="Known price at base_year_month ($/lb)"
    )
    forecast_generated_at: str = Field(
        ..., description="UTC timestamp when the forecast was computed"
    )
    forecasts: list[MonthForecast] = Field(
        ...,
        description=(
            "Month-by-month forecast: next 12 months, or current month + "
            "next 12 (13 total) when include_current_month='YES'"
        ),
    )
class ErrorResponse(BaseModel):
    """Standard error envelope."""
    error: str
    detail: Optional[str] = None
