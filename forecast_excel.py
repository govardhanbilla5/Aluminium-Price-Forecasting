"""
Excel-backed forecast endpoint  –  /api/v1/forecast-excel
Routes
──────
POST /api/v1/forecast-excel
    Single-part forecast. Body: { part_number, tier_1 }
GET  /api/v1/forecast-excel/parts
    All parts from aluminium_data.xlsx (part_number, tier_1, weight, price).
POST /api/v1/forecast-excel/reload
    Clear in-memory workbook cache so edits to the Excel are picked up
    without restarting the server.
"""
import logging
from fastapi import APIRouter, Depends, HTTPException, status
from app.data.base import MarketDataRepository, PartRepository
from app.data.excel_store import (
    ExcelMarketDataRepository,
    ExcelPartRepository,
    _invalidate_cache,
    get_all_parts,
)
from app.models.request import ForecastRequest
from app.models.response import ErrorResponse, ForecastResponse
from app.services.forecast_engine import ForecastEngine
logger = logging.getLogger(__name__)
router = APIRouter(prefix="/forecast-excel", tags=["forecast-excel"])
# ─────────────────────────────────────────────────────────────────────────────
# DEPENDENCY INJECTION
# ─────────────────────────────────────────────────────────────────────────────
def get_market_repo() -> MarketDataRepository:
    return ExcelMarketDataRepository()
def get_part_repo() -> PartRepository:
    return ExcelPartRepository()
def get_forecast_engine(
    market_repo: MarketDataRepository = Depends(get_market_repo),
    part_repo: PartRepository = Depends(get_part_repo),
) -> ForecastEngine:
    return ForecastEngine(market_repo=market_repo, part_repo=part_repo)
# ─────────────────────────────────────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────────────────────────────────────
@router.get(
    "/parts",
    summary="List all parts from Excel",
    response_description="All part + Tier 1 combinations with weight and current price",
)
def list_parts_excel() -> dict:
    """
    Return every row from the Parts sheet of aluminium_data.xlsx,
    including Part Number, Tier 1, Weight (lbs), and Current Price ($).
    """
    try:
        parts = get_all_parts()
        return {"parts": parts, "total": len(parts), "source": "Excel"}
    except Exception as exc:
        logger.exception("Failed to read parts from Excel")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"Could not read Excel file: {exc}",
        )
@router.post(
    "/reload",
    summary="Reload Excel data from disk",
)
def reload_excel() -> dict:
    """Clear the in-memory cache. Next request will re-read the workbook."""
    _invalidate_cache()
    return {"status": "cache cleared", "message": "Excel will be reloaded on the next request."}
@router.post(
    "",
    response_model=ForecastResponse,
    summary="Forecast aluminium price for one part (data from Excel)",
    responses={
        200: {"description": "12-month forecast with all intermediate variables"},
        404: {"model": ErrorResponse, "description": "Part + Tier 1 not found in Excel"},
        422: {"model": ErrorResponse, "description": "Price unavailable or invalid input"},
        500: {"model": ErrorResponse, "description": "Internal error"},
    },
)
def get_forecast_excel(
    payload: ForecastRequest,
    engine: ForecastEngine = Depends(get_forecast_engine),
) -> ForecastResponse:
    """
    Forecast aluminium price for **part_number + tier_1** for the next 12 months.
    Both `part_number` and `tier_1` must match a row in the **Parts sheet**
    of `aluminium_data.xlsx` exactly. The current price is always read from
    that sheet — it cannot be overridden via the API.

    **cng_q** and **cng_q-1** (required, $/lb): CNG values used unchanged
    for every forecast month.

    **include_current_month** (optional, `"YES"` / `"NO"`, default `"NO"`):
    - `"NO"`  → 12 months, starting next month.
    - `"YES"` → 13 months, starting with the current month.
    """
    logger.info(
        "Forecast request: part_number=%s  tier_1=%s",
        payload.part_number, payload.tier_1,
    )
    try:
        result = engine.forecast(
            part_number=payload.part_number,
            tier_1=payload.tier_1,
            cng_q=payload.cng_q,
            cng_q_1=payload.cng_q_1,
            include_breakdown=True,   # factor breakdown only on single-part endpoint
            include_current_month=(payload.include_current_month == "YES"),
        )
    except FileNotFoundError as exc:
        raise HTTPException(status_code=status.HTTP_500_INTERNAL_SERVER_ERROR, detail=str(exc))
    except ValueError as exc:
        error_msg = str(exc)
        logger.warning("Forecast rejected: %s", error_msg)
        if "Part not found" in error_msg:
            raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail=error_msg)
        raise HTTPException(status_code=status.HTTP_422_UNPROCESSABLE_ENTITY, detail=error_msg)
    except Exception:
        logger.exception("Unexpected error for part=%s tier_1=%s", payload.part_number, payload.tier_1)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An unexpected error occurred. See server logs.",
        )
    logger.info(
        "Forecast served: part=%s tier_1=%s  price_range=%.4f–%.4f",
        payload.part_number, payload.tier_1,
        result.forecasts[0].predicted_price,
        result.forecasts[-1].predicted_price,
    )
    return result
