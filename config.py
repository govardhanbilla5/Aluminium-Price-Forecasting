"""
Application configuration and physical constants.

All tunable constants live here so they can be changed in one place
without touching business logic. In production these can be driven by
environment variables via pydantic-settings.
"""

from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    """App-level settings. Override via environment variables."""

    APP_NAME: str = "Aluminium Price Forecast API"
    APP_VERSION: str = "1.0.0"
    DEBUG: bool = False

    # ── Physical / contractual constants ──────────────────────────────────
    # Density / conversion factor (fixed per contract)
    DF_C: float = 1.44

    # MC = (LME + Midwest) × MC_MULTIPLIER + MC_OFFSET
    MC_MULTIPLIER: float = 0.91
    MC_OFFSET: float = 1.25

    # Actual (published) PPI values reach up to this many months before the
    # current month; anything later in the sheet is a projection.
    # Sitting in September with lag 2 → actuals through July.
    PPI_ACTUAL_LAG_MONTHS: int = 2

    # How many future months to forecast
    FORECAST_HORIZON_MONTHS: int = 12

    class Config:
        env_file = ".env"
        case_sensitive = True


settings = Settings()
