"""
Abstract interfaces for every data source used by the forecast engine.

DESIGN INTENT
─────────────
Business logic (forecast_engine.py) depends ONLY on these abstract classes.
Concrete implementations live in separate modules and are injected at startup.

Tier 1 (supplier name) + Part Number together form the unique key for
part master lookups. The same part number can exist under multiple Tier 1
suppliers with different weights and prices.
"""

from abc import ABC, abstractmethod
from typing import Optional


class MarketDataRepository(ABC):
    """Provides historical AND projected market prices (all in $/lb)."""

    @abstractmethod
    def get_lme(self, year_month: str) -> Optional[float]:
        """LME aluminium spot price for the given month ($/lb)."""

    @abstractmethod
    def get_midwest_premium(self, year_month: str) -> Optional[float]:
        """Midwest premium for the given month ($/lb)."""

    @abstractmethod
    def get_ppi(self, year_month: str) -> Optional[float]:
        """PPI index value for the given month (dimensionless)."""

class PartRepository(ABC):
    """
    Provides part master data.

    All lookups require BOTH part_number AND tier_1 because the same
    part can be supplied by multiple Tier 1 suppliers with different
    weights and prices.
    """

    @abstractmethod
    def get_part_weight(self, part_number: str, tier_1: str) -> Optional[float]:
        """
        Return the weight (PWt) in lbs for the given part + supplier.
        Returns None if the combination is not found.
        """

    @abstractmethod
    def get_base_price(self, part_number: str, tier_1: str, year_month: str) -> Optional[float]:
        """
        Return the known price for a part + supplier in a given month ($/lb).
        Returns None if not available.
        """

    def get_price_month(self) -> Optional[str]:
        """
        Return the 'YYYY-MM' month the stored part prices belong to, or None
        if the source cannot tell. Not abstract so existing repositories keep
        working; used to explain price-month mismatches.
        """
        return None
