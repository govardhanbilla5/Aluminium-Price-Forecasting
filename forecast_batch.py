"""
Batch forecast endpoint  –  /api/v1/forecast-batch

POST /api/v1/forecast-batch
    Upload Excel with 'Part Number' and 'Tier 1' columns.
    Returns a downloadable Excel workbook with Summary + 12 monthly sheets.
"""

import logging
from datetime import datetime, timezone

from fastapi import APIRouter, Depends, HTTPException, UploadFile, File, Form, status
from fastapi.responses import Response

from app.data.base import MarketDataRepository, PartRepository
from app.data.excel_store import ExcelMarketDataRepository, ExcelPartRepository
from app.models.response import ErrorResponse
from app.services.batch_forecast import build_forecast_workbook, read_parts_from_upload
from app.services.forecast_engine import ForecastEngine

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/forecast-batch", tags=["forecast-batch"])


def get_market_repo() -> MarketDataRepository:
    return ExcelMarketDataRepository()

def get_part_repo() -> PartRepository:
    return ExcelPartRepository()

def get_engine(
    market_repo: MarketDataRepository = Depends(get_market_repo),
    part_repo: PartRepository         = Depends(get_part_repo),
) -> ForecastEngine:
    return ForecastEngine(market_repo=market_repo, part_repo=part_repo)


@router.post(
    "",
    summary="Batch forecast — upload part list, download result workbook",
    response_class=Response,
    responses={
        200: {
            "description": (
                "Excel workbook with Summary sheet + 12 monthly sheets. "
                "Each sheet contains all (Part Number, Tier 1) pairs with "
                "all intermediate variables and predicted prices."
            ),
            "content": {
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {}
            },
        },
        400: {"description": "Invalid input file"},
        500: {"description": "Internal forecasting error"},
    },
)
async def forecast_batch(
    file: UploadFile = File(
        ...,
        description=(
            "Excel file (.xlsx) with two columns: 'Part Number' and 'Tier 1'. "
            "Duplicate (Part Number, Tier 1) pairs are silently deduplicated."
        ),
    ),
    include_current_month: str = Form(
        "NO",
        description=(
            "'YES' → start at the current month (13 monthly sheets). "
            "'NO' (default) → start next month (12 monthly sheets)."
        ),
    ),
    cng_q: float = Form(
        ..., description="CNG_Q ($/lb), used unchanged for every forecast month."
    ),
    cng_q_1: float = Form(
        ...,
        alias="cng_q-1",
        description="CNG_Q-1 ($/lb), used unchanged for every forecast month.",
    ),
    engine: ForecastEngine = Depends(get_engine),
) -> Response:
    """
    Upload an Excel file with **Part Number** and **Tier 1** columns and
    receive a fully formatted forecast workbook in return.

    **Input format:**

    | Part Number | Tier 1            |
    |-------------|-------------------|
    | 09-0052-003 | Kadon Aerospace   |
    | 2190-1015   | NA                |
    | 09-0052-003 | Point Precision   |

    **Output workbook:**
    - **Summary** — all pairs × 12 months side by side
    - **12 monthly sheets** — one per forecast month, 18 columns including
      all intermediate variables (MC_Q, AMS_Q, PPI Factor, CNG, etc.)
      and the chained Predicted Price

    Parts not found in the data store show an **ERROR row** instead of
    stopping the entire batch.

    **cng_q** and **cng_q-1** (required form fields, $/lb): CNG values used
    unchanged for every forecast month.

    **include_current_month** (optional form field, `YES` / `NO`, default `NO`):
    `YES` adds the current month as the first forecast month (13 sheets).
    """
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Only .xlsx files are accepted.",
        )

    include_current_month = include_current_month.strip().upper()
    if include_current_month not in ("YES", "NO"):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="include_current_month must be 'YES' or 'NO'.",
        )

    logger.info("Batch forecast request: file=%s", file.filename)

    try:
        file_bytes = await file.read()
    except Exception as exc:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Could not read uploaded file: {exc}",
        )

    try:
        part_tier_pairs = read_parts_from_upload(file_bytes)
    except ValueError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc))

    if not part_tier_pairs:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="No valid (Part Number, Tier 1) pairs found in the uploaded file.",
        )

    logger.info("Processing batch of %d part+tier pairs", len(part_tier_pairs))

    try:
        workbook_bytes = build_forecast_workbook(
            part_tier_pairs=part_tier_pairs,
            engine=engine,
            cng_q=cng_q,
            cng_q_1=cng_q_1,
            include_current_month=(include_current_month == "YES"),
        )
    except ValueError as exc:
        raise HTTPException(status_code=status.HTTP_400_BAD_REQUEST, detail=str(exc))
    except Exception:
        logger.exception("Unexpected error during batch forecast")
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail="An unexpected error occurred. See server logs.",
        )

    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")
    output_filename = f"forecast_{today}.xlsx"

    logger.info(
        "Batch complete: %d pairs → %d bytes", len(part_tier_pairs), len(workbook_bytes)
    )

    return Response(
        content=workbook_bytes,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{output_filename}"',
            "Content-Length": str(len(workbook_bytes)),
            "X-Parts-Count": str(len(part_tier_pairs)),
        },
    )
