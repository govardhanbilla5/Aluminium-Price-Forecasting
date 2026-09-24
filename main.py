"""
Aluminium Price Forecast API — entry point.

Run with:
    uvicorn main:app --reload --port 8000

Interactive docs:
    http://localhost:8000/docs
    http://localhost:8000/redoc
"""

import logging
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from app.api.v1.router import api_v1_router
from app.core.config import settings
from app.core.logging_config import configure_logging

# ── Logging must be configured before the routers import anything ──────────
configure_logging(debug=settings.DEBUG)
logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# LIFESPAN  (startup / shutdown hooks)
# ─────────────────────────────────────────────────────────────────────────────

@asynccontextmanager
async def lifespan(app: FastAPI):
    logger.info("══════════════════════════════════════════")
    logger.info("  %s  v%s  starting up", settings.APP_NAME, settings.APP_VERSION)
    logger.info("  DF_c=%.2f  Horizon=%d months", settings.DF_C, settings.FORECAST_HORIZON_MONTHS)
    logger.info("══════════════════════════════════════════")
    yield
    logger.info("%s shutting down", settings.APP_NAME)


# ─────────────────────────────────────────────────────────────────────────────
# APP FACTORY
# ─────────────────────────────────────────────────────────────────────────────

app = FastAPI(
    title=settings.APP_NAME,
    version=settings.APP_VERSION,
    description=(
        "Forecasts aluminium part prices for the next 12 months using a "
        "quarterly formula driven by LME prices, Midwest premium, PPI, and CNG."
    ),
    lifespan=lifespan,
    docs_url="/docs",
    redoc_url="/redoc",
)

# ── CORS  (restrict origins in production via environment variable) ─────────
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],   # tighten in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL EXCEPTION HANDLER  (catch-all safety net)
# ─────────────────────────────────────────────────────────────────────────────

@app.exception_handler(Exception)
async def global_exception_handler(request: Request, exc: Exception):
    logger.exception("Unhandled exception on %s %s", request.method, request.url)
    return JSONResponse(
        status_code=500,
        content={"error": "Internal server error", "detail": str(exc)},
    )


# ─────────────────────────────────────────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────────────────────────────────────────

app.include_router(api_v1_router)


@app.get("/health", tags=["health"])
def health_check():
    """Liveness probe — returns 200 if the service is running."""
    return {"status": "ok", "version": settings.APP_VERSION}
