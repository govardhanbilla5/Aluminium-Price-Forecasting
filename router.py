"""
Aggregate all v1 endpoint routers.

To remove the ML endpoint in future:
  1. Delete app/api/v1/endpoints/forecast_ml.py
  2. Delete app/services/ml_forecast_engine.py
  3. Remove the two ML lines below — nothing else changes.
"""

from fastapi import APIRouter

from app.api.v1.endpoints.forecast_excel import router as forecast_excel_router
from app.api.v1.endpoints.forecast_batch import router as forecast_batch_router

api_v1_router = APIRouter(prefix="/api/v1")
api_v1_router.include_router(forecast_excel_router)   # formula-based
api_v1_router.include_router(forecast_batch_router)   # batch file upload/download
