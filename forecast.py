"""
Forecast API endpoints  –  /api/v1/forecast

Routes
──────
POST /api/v1/forecast
    Forecast aluminium price for the next 12 months.

    Body
    ────
    part_number   : str            – Part number to forecast (required)
    current_price : float | null   – Current month price in $/lb (optional;
                                     falls back to data store, then errors)

GET /api/v1/forecast/parts
    List all known part numbers and their weights.
"""

import logging
from typing import Annotated

from fastapi import APIRouter, Depends, HTTPException, Path, status

from app.data.base import MarketDataRepository, PartRepository
from app.models.request import ForecastRequest
from app.models.response import ErrorResponse, ForecastResponse
from app.services.forecast_engine import ForecastEngine

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/forecast", tags=["forecast"])


# ─────────────────────────────────────────────────────────────────────────────
# DEPENDENCY INJECTION
# ─────────────────────────────────────────────────────────────────────────────

from app.data.hardcoded_store import (
    HardcodedMarketDataRepository,
    HardcodedPartRepository,
    _PART_WEIGHT_MAP,
)


def get_market_repo() -> MarketDataRepository:
    return HardcodedMarketDataRepository()


def get_part_repo() -> PartRepository:
    return HardcodedPartRepository()


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
    summary="List available part numbers",
    response_description="All part numbers present in the part repository",
)
def list_parts() -> dict:
    """Return all part numbers and weights known to the system."""
    parts = [
        {"part_number": pn, "weight_lbs": w}
        for pn, w in _PART_WEIGHT_MAP.items()
    ]
    return {"parts": parts, "total": len(parts)}


@router.post(
    "",
    response_model=ForecastResponse,
    summary="Forecast aluminium price for the next 12 months",
    responses={
        200: {"description": "Successful forecast with per-month breakdown"},
        404: {"model": ErrorResponse, "description": "Part number not found"},
        422: {"model": ErrorResponse, "description": "Current price missing or invalid input"},
        500: {"model": ErrorResponse, "description": "Internal forecasting error"},
    },
)
def get_forecast(
    payload: ForecastRequest,
    engine: ForecastEngine = Depends(get_forecast_engine),
) -> ForecastResponse:
    """
    Forecast aluminium price for **part_number** for the next 12 months.

    The current month is determined automatically from the server clock.

    **current_price** ($/lb) is optional:
    - If provided → used directly as the base price for this month.
    - If omitted → the system looks it up from the internal data store.
    - If neither is available → a 422 is returned asking you to supply it.

    Every response month includes all intermediate formula variables
    (MC_Q, AMS_Q, AMS_Q-1, PPI_Factor, CNG_Q, …) for full auditability.
    """
    logger.info(
        "Forecast request: part_number=%s  current_price=%s",
        payload.part_number,
        payload.current_price if payload.current_price is not None else "<not supplied>",
    )

    try:
        result = engine.forecast(
            part_number=payload.part_number,
            current_price=payload.current_price,
        )

    except ValueError as exc:
        error_msg = str(exc)
        logger.warning("Forecast rejected for %s: %s", payload.part_number, error_msg)

        if "Unknown part number" in error_msg:
            raise HTTPException(
                status_code=status.HTTP_404_NOT_FOUND,
                detail=error_msg,
            )
        # Missing price or bad data → 422 with a clear message
        raise HTTPException(
            status_code=status.HTTP_422_UNPROCESSABLE_ENTITY,
            detail=error_msg,
        )

    except Exception:
        logger.exception("Unexpected error during forecast for part=%s", payload.part_number)
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An unexpected error occurred. See server logs for details.",
        )

    logger.info(
        "Forecast served: part=%s  months=%d  price_range=%.4f–%.4f",
        payload.part_number,
        len(result.forecasts),
        result.forecasts[0].predicted_price,
        result.forecasts[-1].predicted_price,
    )
    return result
