"""
This module defines enumerations used in financial calculations and schedule generation.

Enumerations included:
    - DayCountConvention: Represents various day count conventions (ACT/360, ACT/365, 30/360, and 30E/360 ISDA).
    - Frequency: Represents scheduling frequencies (daily, weekly, monthly, quarterly, semiannually, annually).
    - RollMethod: Represents date rolling methods for adjusting dates in calendars (following, modified following, preceding).
    - ScheduleProjection: Represents the projection direction for schedule generation (backward, forward).

These standardized enumerations are used throughout the application to ensure consistent values in calculations and schedule logic.
"""

from enum import Enum, auto


class DayCountConvention(Enum):
    """
    Enumeration representing day count conventions used in financial calculations.

    Members:
        ACT_360: Actual/360 convention.
        ACT_365: Actual/365 convention.
        THIRTY_360: 30/360 convention.
        THIRTY_E_360_ISDA: 30E/360 convention as per ISDA.
    """

    ACT_360 = auto()  # Actual/360 convention.
    ACT_365 = auto()  # Actual/365 convention.
    THIRTY_360 = auto()  # 30/360 convention.
    THIRTY_E_360_ISDA = auto()  # 30E/360 ISDA convention.


class Frequency(Enum):
    """
    Enumeration representing scheduling frequencies.

    Members:
        DAILY: Daily frequency.
        BUSINESS: Frequency expressed in a specified number of business days.
        WEEKLY: Weekly frequency.
        MONTHLY: Monthly frequency.
        QUARTERLY: Quarterly frequency.
        SEMIANNUALLY: Semiannual frequency.
        ANNUALLY: Annual frequency.
    """

    DAILY = auto()  # Daily frequency.
    BUSINESS = auto()  # Frequency expressed in a specified number of business days.
    WEEKLY = auto()  # Weekly frequency.
    MONTHLY = auto()  # Monthly frequency.
    QUARTERLY = auto()  # Quarterly frequency.
    SEMIANNUALLY = auto()  # Semiannual frequency.
    ANNUALLY = auto()  # Annual frequency.


class RollMethod(Enum):
    """
    Enumeration representing date rolling methods used to adjust dates in calendars.

    Members:
        FOLLOWING: Roll forward to the next business day.
        MODIFIED_FOLLOWING: Roll forward to the next business day, unless that adjustment changes the month, then roll backward.
        PRECEDING: Roll backward to the previous business day.
    """

    FOLLOWING = auto()  # Roll forward to the next business day.
    MODIFIED_FOLLOWING = auto()  # Roll forward unless it changes the month
    PRECEDING = auto()  # Roll backward to the previous business day.


class ScheduleProjection(Enum):
    """
    Enumeration representing the projection direction for a schedule.

    Members:
        BACKWARD: Project the schedule backwards.
        FORWARD: Project the schedule forwards.
    """

    BACKWARD = auto()  # Project schedule backwards.
    FORWARD = auto()  # Project schedule forwards.
