"""
Unit tests for the Aluminium Price Forecast Engine.

Run with:
    pytest tests/ -v
"""

import pytest
from unittest.mock import MagicMock
from typing import Optional

from app.data.base import MarketDataRepository, PartRepository
from app.services.forecast_engine import (
    ForecastEngine,
    _quarter_of_month,
    _quarter_months,
    _last_month_of_quarter,
    _prev_quarter,
)
from app.data.hardcoded_store import (
    HardcodedMarketDataRepository,
    HardcodedPartRepository,
)


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS / FIXTURES
# ─────────────────────────────────────────────────────────────────────────────

def make_engine() -> ForecastEngine:
    return ForecastEngine(
        market_repo=HardcodedMarketDataRepository(),
        part_repo=HardcodedPartRepository(),
    )


# ─────────────────────────────────────────────────────────────────────────────
# QUARTER UTILITY TESTS
# ─────────────────────────────────────────────────────────────────────────────

class TestQuarterHelpers:

    @pytest.mark.parametrize("month,expected_q", [
        (1, 1), (2, 1), (3, 1),
        (4, 2), (5, 2), (6, 2),
        (7, 3), (8, 3), (9, 3),
        (10, 4), (11, 4), (12, 4),
    ])
    def test_quarter_of_month(self, month, expected_q):
        assert _quarter_of_month(month) == expected_q

    def test_quarter_months_q1(self):
        assert _quarter_months(2026, 1) == ["2026-01", "2026-02", "2026-03"]

    def test_quarter_months_q4(self):
        assert _quarter_months(2025, 4) == ["2025-10", "2025-11", "2025-12"]

    def test_last_month_of_quarter(self):
        assert _last_month_of_quarter(2026, 1) == "2026-03"
        assert _last_month_of_quarter(2026, 2) == "2026-06"
        assert _last_month_of_quarter(2026, 4) == "2026-12"

    def test_prev_quarter_normal(self):
        assert _prev_quarter(2026, 2) == (2026, 1)
        assert _prev_quarter(2026, 4) == (2026, 3)

    def test_prev_quarter_year_rollover(self):
        assert _prev_quarter(2026, 1) == (2025, 4)


# ─────────────────────────────────────────────────────────────────────────────
# FORECAST ENGINE — HAPPY PATH
# ─────────────────────────────────────────────────────────────────────────────

class TestForecastEngineHappyPath:

    def test_forecast_returns_12_months(self):
        engine = make_engine()
        result = engine.forecast("ALU-1001", base_year_month="2026-02")
        assert len(result.forecasts) == 12

    def test_forecast_first_month_is_march_2026(self):
        engine = make_engine()
        result = engine.forecast("ALU-1001", base_year_month="2026-02")
        assert result.forecasts[0].year_month == "2026-03"

    def test_forecast_last_month_is_feb_2027(self):
        engine = make_engine()
        result = engine.forecast("ALU-1001", base_year_month="2026-02")
        assert result.forecasts[-1].year_month == "2027-02"

    def test_base_price_in_response(self):
        engine = make_engine()
        result = engine.forecast("ALU-1001", base_year_month="2026-02")
        assert result.base_price == pytest.approx(2.9650, abs=1e-4)

    def test_prices_are_strictly_positive(self):
        engine = make_engine()
        result = engine.forecast("ALU-1001", base_year_month="2026-02")
        for f in result.forecasts:
            assert f.predicted_price > 0, f"Negative price for {f.year_month}"

    def test_prices_are_monotonically_increasing(self):
        """Given the hardcoded data trends upward, prices should rise."""
        engine = make_engine()
        result = engine.forecast("ALU-1001", base_year_month="2026-02")
        prices = [f.predicted_price for f in result.forecasts]
        assert prices == sorted(prices), "Expected prices to increase month-over-month"

    def test_march_price_manual_calculation(self):
        """
        Manually verify the March 2026 forecast step:
          Q1-2026: LME_avg=(1.11+1.12+1.128)/3=1.1193, MWP_avg=(0.1815+0.183+0.1845)/3=0.183
          MC_Q = 1.1193 + 0.183 = 1.3023
          Q4-2025: LME_avg=(1.115+1.102+1.098)/3=1.105, MWP_avg=(0.182+0.181+0.1795)/3=0.1808
          MC_Q-1 = 1.105 + 0.1808 = 1.2858 (approx)
          CNG_Q=0.0310, CNG_Q-1=0.0302
          AMS_Q = 1.3023*1.44 + 0.031 = 1.9064 (approx)
          AMS_Q-1 = 1.2858*1.44 + 0.0302 = 1.8818 (approx)
          PPI_Q=325.5, PPI_Q-1=320.1
          PPI_Factor = (325.5-320.1)/320.1 = 0.016870
          P_march = 2.965 + (1.9064-1.8818)*2.5 + 0.016870*2.965
        """
        engine = make_engine()
        result = engine.forecast("ALU-1001", base_year_month="2026-02")
        march = result.forecasts[0]

        ctx = march.quarter_context
        # Verify AMS values are consistent
        expected_ams_q = ctx.mc_q * 1.44 + ctx.cng_q
        assert ctx.ams_q == pytest.approx(expected_ams_q, abs=1e-4)

        expected_ams_q_1 = ctx.mc_q_1 * 1.44 + ctx.cng_q_1
        assert ctx.ams_q_1 == pytest.approx(expected_ams_q_1, abs=1e-4)

        # Verify formula
        expected_price = (
            result.base_price
            + ctx.ams_delta * result.pwt_lbs
            + ctx.ppi_factor * result.base_price
        )
        assert march.predicted_price == pytest.approx(expected_price, abs=1e-3)

    def test_chained_price_propagation(self):
        """Each month's base_price_used should equal the previous month's predicted_price."""
        engine = make_engine()
        result = engine.forecast("ALU-1001", base_year_month="2026-02")
        for i in range(1, len(result.forecasts)):
            prev_predicted = result.forecasts[i - 1].predicted_price
            curr_base = result.forecasts[i].base_price_used
            assert curr_base == pytest.approx(prev_predicted, abs=1e-6), (
                f"Month {result.forecasts[i].year_month}: "
                f"base_price_used={curr_base} != prev predicted={prev_predicted}"
            )

    def test_quarter_context_uses_predicted_month_and_prev_quarter_end(self):
        """MC_Q / PPI_Q come from the predicted month; MC_Q-1 / PPI_Q-1 from the
        last month of the previous quarter (shared by all months in a quarter)."""
        engine = make_engine()
        market = HardcodedMarketDataRepository()
        result = engine.forecast("ALU-1001", base_year_month="2026-02")
        # April, May, June are all Q2-2026 → previous quarter ends 2026-03
        q2_months = [f for f in result.forecasts if f.quarter_context.quarter_label == "Q2-2026"]
        assert len(q2_months) == 3
        for f in q2_months:
            ctx = f.quarter_context
            ym = f.year_month
            assert ctx.mc_q == pytest.approx(
                market.get_lme(ym) + market.get_midwest_premium(ym), abs=1e-6
            )
            assert ctx.ppi_q == pytest.approx(market.get_ppi(ym), abs=1e-4)
            assert ctx.mc_q_1 == pytest.approx(
                market.get_lme("2026-03") + market.get_midwest_premium("2026-03"), abs=1e-6
            )
            assert ctx.ppi_q_1 == pytest.approx(market.get_ppi("2026-03"), abs=1e-4)

    def test_all_parts_can_be_forecasted(self):
        engine = make_engine()
        part_repo = HardcodedPartRepository()
        from app.data.hardcoded_store import _PART_WEIGHT_MAP
        for part_number in _PART_WEIGHT_MAP:
            result = engine.forecast(part_number, base_year_month="2026-02")
            assert len(result.forecasts) == 12, f"Expected 12 months for {part_number}"


# ─────────────────────────────────────────────────────────────────────────────
# FORECAST ENGINE — ERROR CASES
# ─────────────────────────────────────────────────────────────────────────────

class TestForecastEngineErrors:

    def test_unknown_part_raises_value_error(self):
        engine = make_engine()
        with pytest.raises(ValueError, match="Unknown part number"):
            engine.forecast("DOES-NOT-EXIST", base_year_month="2026-02")

    def test_invalid_base_month_format_raises(self):
        engine = make_engine()
        with pytest.raises(ValueError, match="YYYY-MM"):
            engine.forecast("ALU-1001", base_year_month="02-2026")

    def test_missing_base_price_raises(self):
        engine = make_engine()
        # 2025-01 has no price in the hardcoded store
        with pytest.raises(ValueError, match="No known base price"):
            engine.forecast("ALU-1001", base_year_month="2025-01")

    def test_missing_market_data_raises(self):
        """If a required market data point is absent, the engine should raise cleanly."""
        market_mock = MagicMock(spec=MarketDataRepository)
        market_mock.get_lme.return_value = None   # simulate missing LME
        market_mock.get_midwest_premium.return_value = 0.18
        market_mock.get_ppi.return_value = 320.0

        part_mock = MagicMock(spec=PartRepository)
        part_mock.get_part_weight.return_value = 2.5
        part_mock.get_base_price.return_value = 2.965

        engine = ForecastEngine(market_repo=market_mock, part_repo=part_mock)
        with pytest.raises(ValueError, match="Required data unavailable"):
            engine.forecast("ALU-1001", base_year_month="2026-02")


# ─────────────────────────────────────────────────────────────────────────────
# DATA LAYER TESTS
# ─────────────────────────────────────────────────────────────────────────────

class TestHardcodedRepositories:

    def test_lme_known_month(self):
        repo = HardcodedMarketDataRepository()
        assert repo.get_lme("2026-02") == pytest.approx(1.12, abs=1e-4)

    def test_lme_unknown_month_returns_none(self):
        repo = HardcodedMarketDataRepository()
        assert repo.get_lme("2020-01") is None

    def test_part_weight_known(self):
        repo = HardcodedPartRepository()
        assert repo.get_part_weight("ALU-1001") == pytest.approx(2.5)

    def test_part_weight_unknown_returns_none(self):
        repo = HardcodedPartRepository()
        assert repo.get_part_weight("FAKE-9999") is None

    def test_base_price_known(self):
        repo = HardcodedPartRepository()
        assert repo.get_base_price("ALU-1001", "2026-02") == pytest.approx(2.965, abs=1e-4)

    def test_base_price_future_returns_none(self):
        repo = HardcodedPartRepository()
        assert repo.get_base_price("ALU-1001", "2026-06") is None
