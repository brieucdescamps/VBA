"""
Market conventions for swaps and swaptions across currencies and indices.
"""

import os
import sys
from typing import Any, Dict, Tuple, Union

sys.path.append(os.path.abspath("."))
from enumerations import (
    DayCountConvention,
    DiscountRateIndex,
    FloatingRateIndex,
    Frequency,
    RollMethod,
    ScheduleProjection,
    SwaptionPremiumType,
    SwaptionSettlement,
)


class InterestRatesConvention:
    """
    Provides standardized market conventions for interest rate swaps and swaptions
    across major currencies and rate indices.

    These conventions cover payment frequencies, day count conventions,
    business day adjustments, calendars, and related settlement rules.

    This class is intended to serve as a centralized reference to ensure consistency
    across pricing engines, risk models, and strategy definitions.
    """

    # Market conventions for standard fixed/floating interest rate swaps
    # Indexed by (currency, floating rate index key)
    swap_market_convention: Dict[Tuple[str, str], dict] = {
        # USD 3M LIBOR
        ("USD", FloatingRateIndex.USD_3M_LIBOR.to_string_key()): {
            "spot_lag": 2,
            "fixed_leg_frequency": Frequency.SEMIANNUALLY,
            "fixed_leg_day_count_convention": DayCountConvention.THIRTY_360,
            "fixed_leg_calendar": "LNB,NYB",
            "fixed_leg_roll_method_convention": RollMethod.MODIFIED_FOLLOWING,
            "fixed_leg_schedule_projection_convention": ScheduleProjection.BACKWARD,
            "fixed_leg_pay_lag": 2,
            "floating_leg_frequency": Frequency.QUARTERLY,
            "floating_leg_day_count_convention": DayCountConvention.ACT_360,
            "floating_leg_calendar": "LNB,NYB",
            "floating_leg_roll_method_convention": RollMethod.MODIFIED_FOLLOWING,
            "floating_leg_schedule_projection_convention": ScheduleProjection.BACKWARD,
            "floating_leg_pay_lag": 2,
            "floating_leg_reset_calendar": "LNB",
            "floating_leg_reset_lag": 2,
            "floating_leg_set_in_arrears": False,
        },
        # USD SOFR
        ("USD", FloatingRateIndex.USD_SOFR.to_string_key()): {
            "spot_lag": 2,
            "fixed_leg_frequency": Frequency.ANNUALLY,
            "fixed_leg_day_count_convention": DayCountConvention.ACT_360,
            "fixed_leg_calendar": "NYB",
            "fixed_leg_roll_method_convention": RollMethod.MODIFIED_FOLLOWING,
            "fixed_leg_schedule_projection_convention": ScheduleProjection.BACKWARD,
            "fixed_leg_pay_lag": 2,
            "floating_leg_frequency": Frequency.ANNUALLY,
            "floating_leg_day_count_convention": DayCountConvention.ACT_360,
            "floating_leg_calendar": "NYB",
            "floating_leg_roll_method_convention": RollMethod.MODIFIED_FOLLOWING,
            "floating_leg_schedule_projection_convention": ScheduleProjection.BACKWARD,
            "floating_leg_pay_lag": 2,
            "floating_leg_reset_calendar": "NYB",
            "floating_leg_reset_lag": 2,
            "floating_leg_set_in_arrears": False,
        },
        # EUR 6M EURIBOR
        ("EUR", FloatingRateIndex.EUR_6M_EURIBOR.to_string_key()): {
            "spot_lag": 2,
            "fixed_leg_frequency": Frequency.ANNUALLY,
            "fixed_leg_day_count_convention": DayCountConvention.THIRTY_E_360_ISDA,
            "fixed_leg_calendar": "TGT",
            "fixed_leg_roll_method_convention": RollMethod.MODIFIED_FOLLOWING,
            "fixed_leg_schedule_projection_convention": ScheduleProjection.BACKWARD,
            "fixed_leg_pay_lag": 0,
            "floating_leg_frequency": Frequency.SEMIANNUALLY,
            "floating_leg_day_count_convention": DayCountConvention.ACT_360,
            "floating_leg_calendar": "TGT",
            "floating_leg_roll_method_convention": RollMethod.MODIFIED_FOLLOWING,
            "floating_leg_schedule_projection_convention": ScheduleProjection.BACKWARD,
            "floating_leg_pay_lag": 0,
            "floating_leg_reset_calendar": "TGT",
            "floating_leg_reset_lag": 2,
            "floating_leg_set_in_arrears": False,
        },
        ("JPY", FloatingRateIndex.JPY_6M_LIBOR.to_string_key()): {
            "spot_lag": 2,
            "fixed_leg_frequency": Frequency.SEMIANNUALLY,
            "fixed_leg_day_count_convention": DayCountConvention.ACT_365,
            "fixed_leg_calendar": "LNB, TKB",
            "fixed_leg_roll_method_convention": RollMethod.MODIFIED_FOLLOWING,
            "fixed_leg_schedule_projection_convention": ScheduleProjection.BACKWARD,
            "fixed_leg_pay_lag": 0,
            "floating_leg_frequency": Frequency.SEMIANNUALLY,
            "floating_leg_day_count_convention": DayCountConvention.ACT_360,
            "floating_leg_calendar": "LNB, TKB",
            "floating_leg_roll_method_convention": RollMethod.MODIFIED_FOLLOWING,
            "floating_leg_schedule_projection_convention": ScheduleProjection.BACKWARD,
            "floating_leg_pay_lag": 0,
            "floating_leg_reset_calendar": "LNB",
            "floating_leg_reset_lag": 2,
            "floating_leg_set_in_arrears": False,
        },
        ("JPY", FloatingRateIndex.JPY_TONAR.to_string_key()): {
            "spot_lag": 2,
            "fixed_leg_frequency": Frequency.ANNUALLY,
            "fixed_leg_day_count_convention": DayCountConvention.ACT_365,
            "fixed_leg_calendar": "TKB",
            "fixed_leg_roll_method_convention": RollMethod.MODIFIED_FOLLOWING,
            "fixed_leg_schedule_projection_convention": ScheduleProjection.BACKWARD,
            "fixed_leg_pay_lag": 0,
            "floating_leg_frequency": Frequency.ANNUALLY,
            "floating_leg_day_count_convention": DayCountConvention.ACT_365,
            "floating_leg_calendar": "TKB",
            "floating_leg_roll_method_convention": RollMethod.MODIFIED_FOLLOWING,
            "floating_leg_schedule_projection_convention": ScheduleProjection.BACKWARD,
            "floating_leg_pay_lag": 0,
            "floating_leg_reset_calendar": "TKB",
            "floating_leg_reset_lag": 0,
            "floating_leg_set_in_arrears": True,
        },
    }

    # Market conventions for physically settled European swaptions
    # Indexed by (currency, floating rate index key)
    swaption_market_convention: Dict[Tuple[str, str], dict] = {
        ("USD", FloatingRateIndex.USD_3M_LIBOR.to_string_key()): {
            "settlement": SwaptionSettlement.PHYSICAL,
            "premium_type": SwaptionPremiumType.FORWARD,
            "expiry_day_count_convention": DayCountConvention.ACT_365,
            "calendar": "LNB,NYB",
            "roll_method": RollMethod.MODIFIED_FOLLOWING,
            "expiry_notification_lag": 2,
            "expiry_pay_lag": 2,
            "underlying_swap_calendar": "LNB,NYB",
        },
        ("USD", FloatingRateIndex.USD_SOFR.to_string_key()): {
            "settlement": SwaptionSettlement.PHYSICAL,
            "premium_type": SwaptionPremiumType.FORWARD,
            "expiry_day_count_convention": DayCountConvention.ACT_365,
            "calendar": "NYB",
            "roll_method": RollMethod.MODIFIED_FOLLOWING,
            "expiry_notification_lag": 2,
            "expiry_pay_lag": 2,
            "underlying_swap_calendar": "NYB",
        },
        ("EUR", FloatingRateIndex.EUR_6M_EURIBOR.to_string_key()): {
            "settlement": SwaptionSettlement.PHYSICAL,
            "premium_type": SwaptionPremiumType.FORWARD,
            "expiry_day_count_convention": DayCountConvention.ACT_365,
            "calendar": "TGT",
            "roll_method": RollMethod.MODIFIED_FOLLOWING,
            "expiry_notification_lag": 2,
            "expiry_pay_lag": 2,
            "underlying_swap_calendar": "TGT",
        },
        ("GBP", FloatingRateIndex.GBP_SONIA.to_string_key()): {
            "settlement": SwaptionSettlement.PHYSICAL,
            "premium_type": SwaptionPremiumType.FORWARD,
            "expiry_day_count_convention": DayCountConvention.ACT_365,
            "calendar": "LNB",
            "roll_method": RollMethod.MODIFIED_FOLLOWING,
            "expiry_notification_lag": 2,
            "expiry_pay_lag": 2,
            "underlying_swap_calendar": "LNB",
        },
        ("GBP", FloatingRateIndex.GBP_6M_LIBOR.to_string_key()): {
            "settlement": SwaptionSettlement.PHYSICAL,
            "premium_type": SwaptionPremiumType.FORWARD,
            "expiry_day_count_convention": DayCountConvention.ACT_365,
            "calendar": "LNB",
            "roll_method": RollMethod.MODIFIED_FOLLOWING,
            "expiry_notification_lag": 2,
            "expiry_pay_lag": 2,
            "underlying_swap_calendar": "LNB",
        },
        ("JPY", FloatingRateIndex.JPY_6M_LIBOR.to_string_key()): {
            "settlement": SwaptionSettlement.PHYSICAL,
            "premium_type": SwaptionPremiumType.FORWARD,
            "expiry_day_count_convention": DayCountConvention.ACT_365,
            "calendar": "LNB,TKB",
            "roll_method": RollMethod.MODIFIED_FOLLOWING,
            "expiry_notification_lag": 2,
            "expiry_pay_lag": 2,
            "underlying_swap_calendar": "LNB,TKB",
        },
        ("JPY", FloatingRateIndex.JPY_TONAR.to_string_key()): {
            "settlement": SwaptionSettlement.PHYSICAL,
            "premium_type": SwaptionPremiumType.FORWARD,
            "expiry_day_count_convention": DayCountConvention.ACT_365,
            "calendar": "TKB",
            "roll_method": RollMethod.MODIFIED_FOLLOWING,
            "expiry_notification_lag": 2,
            "expiry_pay_lag": 2,
            "underlying_swap_calendar": "TKB",
        },
    }

    @staticmethod
    def get_swap_market_convention(
        currency: str,
        floating_rate_index: Union[FloatingRateIndex, DiscountRateIndex],
    ) -> Dict[str, Any]:
        """
        Retrieve the swap market convention for a given currency and floating rate index.

        Args:
            currency (str): The currency code (e.g., 'USD', 'EUR').
            floating_rate_index (FloatingRateIndex or DiscountRateIndex): The index for the floating leg.

        Returns:
            dict: A dictionary describing the fixed and floating leg conventions for swaps.

        Raises:
            KeyError: If the currency/index pair is not defined.
        """
        return InterestRatesConvention.swap_market_convention[
            (currency, floating_rate_index.to_string_key())
        ]

    @staticmethod
    def get_swaption_market_convention(
        currency: str,
        floating_rate_index: Union[FloatingRateIndex, DiscountRateIndex],
    ) -> Dict[str, Any]:
        """
        Retrieve the swaption market convention for a given currency and floating rate index.

        Args:
            currency (str): The currency code (e.g., 'USD', 'EUR').
            floating_rate_index (FloatingRateIndex or DiscountRateIndex): The index of the swap underlying the swaption.

        Returns:
            dict: A dictionary of expiry and settlement-related conventions for swaptions.

        Raises:
            KeyError: If the currency/index pair is not defined.
        """
        return InterestRatesConvention.swaption_market_convention[
            (currency, floating_rate_index.to_string_key())
        ]
