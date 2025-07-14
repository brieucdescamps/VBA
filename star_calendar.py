"""
Calendar class

This module defines a Calendar class for a given location (e.g., NYB, LON) with holiday tracking.
It implements:
  - Singleton pattern (one calendar per location)
  - Dynamic holiday loading (with add/remove logic and as-of-date support)
  - Optimized business day computations using NumPy's busday functions
  - Convenience methods for date offsetting (calendar days, weeks, months, years, and business days)
"""

import calendar
import enum
from typing import ClassVar, Dict, List, Optional, Tuple

import numpy as np

from enumerations.calendar_enumerations import Frequency, RollMethod

# Weekday offset to align with ISO: Monday=0, Sunday=6
_WEEKDAY_EPOCH_SHIFT = 3  # 1970-01-01 was a Thursday


class HolidayAction(enum.Enum):
    """Represents an action applied to a holiday date."""

    ADD = 1
    REMOVE = 2


class Calendar:
    """
    Business day calendar for a given location (e.g., NYB, LON), with holiday tracking.

    Implements:
      - Singleton pattern (one calendar per location)
      - Dynamic holiday loading (with add/remove logic and as-of-date support)
      - Optimized business day computations using NumPy's busday functions
      - Convenience methods for date offsetting (calendar days, weeks, months, years, and business days)
    """

    _instances: ClassVar[Dict[str, "Calendar"]] = {}
    _holiday_records_cache: ClassVar[
        Dict[str, List[Tuple[np.datetime64, np.datetime64, HolidayAction]]]
    ] = {}
    _processed_holidays_cache: ClassVar[Dict[Tuple[str, np.datetime64], np.ndarray]] = (
        {}
    )

    def __new__(cls, location_code: str):
        if location_code not in cls._instances:
            cls._instances[location_code] = super().__new__(cls)
        return cls._instances[location_code]

    def __init__(self, location_code: str):
        if not hasattr(self, "location_code"):
            self.location_code = location_code
            # Initialize a cache for busdaycalendar objects.
            self._busdaycalendar_cache: Dict[np.datetime64, np.busdaycalendar] = {}
        elif self.location_code == location_code:
            return
        if location_code not in self._holiday_records_cache:
            self._holiday_records_cache[location_code] = self._load_holiday_records(
                location_code
            )

    def __repr__(self):
        return f"Calendar(location_code='{self.location_code}')"

    @staticmethod
    def _load_holiday_records(
        location_code: str,
    ) -> List[Tuple[np.datetime64, np.datetime64, HolidayAction]]:
        # Simulate static or DB-loaded holidays with versioning support.
        if location_code == "NYB":
            records = [
                ("2025-01-01", "2024-01-15", HolidayAction.ADD),
                ("2025-01-20", "2024-02-01", HolidayAction.ADD),
                ("2025-02-17", "2024-02-01", HolidayAction.ADD),
                ("2025-05-26", "2024-03-10", HolidayAction.ADD),
                ("2025-07-04", "2024-03-10", HolidayAction.ADD),
                ("2025-09-01", "2024-03-10", HolidayAction.ADD),
                ("2025-11-27", "2024-08-15", HolidayAction.ADD),
                ("2025-12-25", "2024-08-15", HolidayAction.ADD),
                ("2025-03-17", "2024-01-10", HolidayAction.ADD),
                ("2025-03-17", "2024-06-20", HolidayAction.REMOVE),
                ("2025-10-12", "2024-01-15", HolidayAction.ADD),
                ("2025-10-12", "2024-07-01", HolidayAction.REMOVE),
                ("2025-10-14", "2024-07-01", HolidayAction.ADD),
            ]
        elif location_code == "LON":
            records = [
                ("2025-01-01", "2024-01-10", HolidayAction.ADD),
                ("2025-04-18", "2024-02-05", HolidayAction.ADD),
                ("2025-04-21", "2024-02-05", HolidayAction.ADD),
                ("2025-05-05", "2024-04-01", HolidayAction.ADD),
                ("2025-05-26", "2024-04-01", HolidayAction.ADD),
                ("2025-08-25", "2024-04-01", HolidayAction.ADD),
                ("2025-12-25", "2024-07-15", HolidayAction.ADD),
                ("2025-12-26", "2024-07-15", HolidayAction.ADD),
                ("2025-06-03", "2024-02-01", HolidayAction.ADD),
                ("2025-06-03", "2024-05-15", HolidayAction.REMOVE),
            ]
        else:
            records = []
        return sorted(
            [(np.datetime64(d), np.datetime64(r), a) for d, r, a in records],
            key=lambda x: x[1],
        )

    def _get_effective_holidays(
        self,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
    ) -> np.ndarray:
        """
        Return the effective holidays array.

        If a holidays array is provided, it is returned; otherwise, fetch holidays using get_holidays(as_of_date).
        """
        if holidays is not None:
            return holidays
        return self.get_holidays(as_of_date)

    def _get_busdaycalendar_cached(
        self, as_of_date: Optional[np.datetime64] = None
    ) -> np.busdaycalendar:
        """
        Return a cached NumPy busdaycalendar object for the given as_of_date.

        If not cached, compute it from the effective holidays and cache the result.
        """
        if as_of_date is None:
            as_of_date = np.datetime64("today", "D")
        if as_of_date not in self._busdaycalendar_cache:
            holidays = self.get_holidays(as_of_date)
            self._busdaycalendar_cache[as_of_date] = np.busdaycalendar(
                holidays=holidays.tolist()
            )
        return self._busdaycalendar_cache[as_of_date]

    def get_holidays(self, as_of_date: Optional[np.datetime64] = None) -> np.ndarray:
        """Return an array of holidays known as of `as_of_date`."""
        if as_of_date is None:
            as_of_date = np.datetime64("today", "D")
        # Ensure holiday records exist for this location.
        if self.location_code not in self._holiday_records_cache:
            self._holiday_records_cache[self.location_code] = (
                self._load_holiday_records(self.location_code)
            )
        key = (self.location_code, as_of_date)
        if key not in self._processed_holidays_cache:
            records = self._holiday_records_cache[self.location_code]
            self._processed_holidays_cache[key] = self._process_holiday_records(
                records, as_of_date
            )
        return self._processed_holidays_cache[key]

    def _process_holiday_records(
        self,
        records: List[Tuple[np.datetime64, np.datetime64, HolidayAction]],
        as_of_date: np.datetime64,
    ) -> np.ndarray:
        """
        Apply holiday add/remove actions up to `as_of_date`.

        This vectorized version converts records into a structured array for filtering.
        """
        dtype = np.dtype(
            [
                ("holiday", "datetime64[D]"),
                ("record_date", "datetime64[D]"),
                ("action", "U10"),
            ]
        )
        rec_arr = np.array([(d, r, a.name) for d, r, a in records], dtype=dtype)
        mask = rec_arr["record_date"] <= as_of_date
        rec_filtered = rec_arr[mask]
        to_add = rec_filtered["holiday"][rec_filtered["action"] == "ADD"]
        to_remove = rec_filtered["holiday"][rec_filtered["action"] == "REMOVE"]
        active = np.setdiff1d(np.unique(to_add), np.unique(to_remove))
        return np.sort(active)

    def _get_busdaycalendar(
        self, as_of_date: Optional[np.datetime64] = None
    ) -> np.busdaycalendar:
        """Return a NumPy busdaycalendar object with holidays as of a specific date."""
        return np.busdaycalendar(holidays=self.get_holidays(as_of_date))

    # --- Cache Management Methods ---

    @classmethod
    def clear_all_caches(cls):
        """Clear all calendar instances and caches."""
        cls._instances.clear()
        cls._holiday_records_cache.clear()
        cls._processed_holidays_cache.clear()

    @classmethod
    def clear_location_cache(cls, location_code: str):
        """
        Clear processed holiday caches related to a specific location.

        Note: This does not clear the raw holiday records.
        """
        cls._processed_holidays_cache = {
            k: v
            for k, v in cls._processed_holidays_cache.items()
            if k[0] != location_code
        }

    @classmethod
    def clear_processed_cache(cls):
        """Clear only the processed holidays (keep raw record cache)."""
        cls._processed_holidays_cache.clear()

    @classmethod
    def reload_holiday_records(cls, location_code: Optional[str] = None):
        """
        Force reload of holiday records for a specific location.

        If location_code is provided, only that location's raw records are reloaded.
        Otherwise, all locations are reloaded.
        """
        if location_code is None:
            cls._holiday_records_cache.clear()
            cls._processed_holidays_cache.clear()
        else:
            cls._holiday_records_cache[location_code] = cls._load_holiday_records(
                location_code
            )
            cls._processed_holidays_cache = {
                k: v
                for k, v in cls._processed_holidays_cache.items()
                if k[0] != location_code
            }

    # --- Core Logic Using NumPy Business Day Functions ---

    def is_holiday(
        self,
        dt: np.datetime64,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
    ) -> bool:
        """Return True if `dt` is in the holiday list."""
        if holidays is None:
            holidays = self.get_holidays(as_of_date)
        return np.datetime64(dt, "D") in holidays

    def is_weekend(self, dt: np.datetime64) -> bool:
        """Check if a date falls on a weekend (Saturday or Sunday)."""
        weekday = (np.datetime64(dt, "D").view("int64") + _WEEKDAY_EPOCH_SHIFT) % 7
        return weekday >= 5

    def is_business_day(
        self,
        dt: np.datetime64,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
    ) -> bool:
        """Check if a date is a business day using np.is_busday."""
        dt = np.datetime64(dt, "D")
        if holidays is None:
            holidays = self.get_holidays(as_of_date)
        return np.is_busday(dt, holidays=holidays)

    def add_days(
        self,
        start: np.datetime64,
        num_days: int,
        use_business: bool = False,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
        roll_method: Optional[RollMethod] = None,
    ) -> np.datetime64:
        """
        Add a specified number of calendar days to a given start date.

        If use_business is True, the resulting date is adjusted to a business day (using roll_method)
        if it is not already a business day.

        Args:
            start (np.datetime64): The starting date.
            num_days (int): The number of days to add (can be negative).
            use_business (bool): If True, adjust the result to a business day.
            as_of_date (Optional[np.datetime64]): Optional reference date for holiday context.
            holidays (Optional[np.ndarray]): Optional array of holidays.
            roll_method (Optional[RollMethod]): Rolling convention to apply if adjustment is needed.

        Returns:
            np.datetime64: The resulting date.
        """
        start = np.datetime64(start, "D")
        candidate = start + np.timedelta64(num_days, "D")
        if use_business:
            if roll_method is None:
                roll_method = RollMethod.MODIFIED_FOLLOWING
            if not self.is_business_day(candidate, as_of_date, holidays):
                candidate = self.add_business_days(
                    candidate, 0, as_of_date=as_of_date, holidays=holidays
                )
        return candidate

    def add_week(
        self,
        start: np.datetime64,
        num_weeks: int,
        use_business: bool = False,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
        roll_method: Optional[RollMethod] = None,
    ) -> np.datetime64:
        """
        Add a specified number of calendar weeks to a given start date.

        If use_business is True, the resulting date is adjusted to a business day.

        Args:
            start (np.datetime64): The starting date.
            num_weeks (int): The number of weeks to add (can be negative).
            use_business (bool): If True, adjust the result to a business day.
            as_of_date (Optional[np.datetime64]): Optional reference date for holiday context.
            holidays (Optional[np.ndarray]): Optional array of holidays.
            roll_method (Optional[RollMethod]): Rolling convention for adjustment.

        Returns:
            np.datetime64: The resulting date.
        """
        return self.add_days(
            start, num_weeks * 7, use_business, as_of_date, holidays, roll_method
        )

    def add_month(
        self,
        start: np.datetime64,
        num_months: int,
        use_business: bool = False,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
        roll_method: Optional[RollMethod] = None,
    ) -> np.datetime64:
        """
        Add a specified number of months to a given start date.

        Attempts to preserve the original day-of-month. If that day does not exist in the target month,
        falls back to the last day of that month. Optionally, if use_business is True, the result is adjusted
        to a business day.

        Args:
            start (np.datetime64): The starting date.
            num_months (int): The number of months to add (can be negative).
            use_business (bool): If True, adjust the result to a business day.
            as_of_date (Optional[np.datetime64]): Optional reference date for holiday context.
            holidays (Optional[np.ndarray]): Optional array of holidays.
            roll_method (Optional[RollMethod]): Rolling convention for adjustment.

        Returns:
            np.datetime64: The resulting date.
        """
        start = np.datetime64(start, "D")
        original_day = int(str(start).split("-")[-1])
        new_month = start.astype("datetime64[M]") + np.timedelta64(num_months, "M")
        new_month_str = np.datetime_as_string(new_month, unit="M")
        year, month = map(int, str(new_month_str).split("-"))
        last_day_num = calendar.monthrange(year, month)[1]
        if original_day > last_day_num:
            candidate_date = np.datetime64(
                f"{year}-{month:02d}-{last_day_num:02d}", "D"
            )
        else:
            candidate_date = np.datetime64(
                f"{year}-{month:02d}-{original_day:02d}", "D"
            )
        if use_business:
            if roll_method is None:
                roll_method = RollMethod.MODIFIED_FOLLOWING
            if not self.is_business_day(candidate_date, as_of_date, holidays):
                candidate_date = self.add_business_days(
                    candidate_date, 0, as_of_date=as_of_date, holidays=holidays
                )
        return candidate_date

    def add_year(
        self,
        start: np.datetime64,
        num_years: int,
        use_business: bool = False,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
        roll_method: Optional[RollMethod] = None,
    ) -> np.datetime64:
        """
        Add a specified number of years to a given start date.

        Implemented by adding 12 * num_years months.
        Optionally adjusts to a business day if use_business is True.

        Args:
            start (np.datetime64): The starting date.
            num_years (int): The number of years to add (can be negative).
            use_business (bool): If True, adjust the result to a business day.
            as_of_date (Optional[np.datetime64]): Optional reference date for holiday context.
            holidays (Optional[np.ndarray]): Optional array of holidays.
            roll_method (Optional[RollMethod]): Rolling convention for adjustment.

        Returns:
            np.datetime64: The resulting date.
        """
        return self.add_month(
            start, num_years * 12, use_business, as_of_date, holidays, roll_method
        )

    def add_business_days(
        self,
        start: np.datetime64,
        num_days: int,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
    ) -> np.datetime64:
        """
        Add or subtract business days from a start date.

        Args:
            start (np.datetime64): The starting date.
            num_days (int): The number of business days to add (can be negative).
            as_of_date (Optional[np.datetime64]): Optional reference date for holiday context.
            holidays (Optional[np.ndarray]): Optional array of holidays.

        Returns:
            np.datetime64: The resulting date.
        """
        if holidays is None:
            holidays = self.get_holidays(as_of_date)
        return np.busday_offset(start, num_days, roll="forward", holidays=holidays)

    def subtract_business_days(
        self,
        start: np.datetime64,
        num_days: int,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
    ) -> np.datetime64:
        """Subtract business days (alias for add_business_days with negative input)."""
        return self.add_business_days(start, -num_days, as_of_date, holidays)

    def get_next_business_day(
        self,
        start: np.datetime64,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
    ) -> np.datetime64:
        """Return the next business day after `start`."""
        return self.add_business_days(start, 1, as_of_date, holidays)

    def get_previous_business_day(
        self,
        start: np.datetime64,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
    ) -> np.datetime64:
        """Return the previous business day before `start`."""
        return self.add_business_days(start, -1, as_of_date, holidays)

    def business_days_between(
        self,
        start: np.datetime64,
        end: np.datetime64,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
    ) -> int:
        """
        Count the number of business days between two dates (inclusive).

        Args:
            start (np.datetime64): The start date.
            end (np.datetime64): The end date.
            as_of_date (Optional[np.datetime64]): Optional reference date for holiday context.
            holidays (Optional[np.ndarray]): Optional array of holidays.

        Returns:
            int: The count of business days.
        """
        if holidays is None:
            holidays = self.get_holidays(as_of_date)
        return int(
            np.busday_count(
                start, end + np.timedelta64(1, "D"), holidays=holidays
            ).item()
        )

    def get_business_days_in_range(
        self,
        start: np.datetime64,
        end: np.datetime64,
        as_of_date: Optional[np.datetime64] = None,
        holidays: Optional[np.ndarray] = None,
    ) -> np.ndarray:
        """
        Return an array of business days between two dates (inclusive), using np.is_busday.

        Args:
            start (np.datetime64): The start date.
            end (np.datetime64): The end date.
            as_of_date (Optional[np.datetime64]): Optional reference date for holiday context.
            holidays (Optional[np.ndarray]): Optional array of holidays.

        Returns:
            np.ndarray: Array of business days.
        """
        if end < start:
            return np.array([], dtype="datetime64[D]")
        if holidays is None:
            holidays = self.get_holidays(as_of_date)
        dates = np.arange(start, end + np.timedelta64(1, "D"))
        return dates[np.is_busday(dates, holidays=holidays)]

    def get_holidays_in_range(
        self,
        start: np.datetime64,
        end: np.datetime64,
        as_of_date: Optional[np.datetime64] = None,
    ) -> np.ndarray:
        """
        Return an array of holidays in a given range as of a specified date.

        Args:
            start (np.datetime64): The start date.
            end (np.datetime64): The end date.
            as_of_date (Optional[np.datetime64]): Optional reference date for holiday context.

        Returns:
            np.ndarray: Array of holidays.
        """
        holidays = self.get_holidays(as_of_date)
        return holidays[(holidays >= start) & (holidays <= end)]

    # --- Dynamic Modification Methods ---

    def add_holiday(
        self, holiday_date: np.datetime64, record_date: Optional[np.datetime64] = None
    ) -> None:
        """
        Add a holiday and invalidate processed holiday caches for this location.

        Args:
            holiday_date (np.datetime64): The holiday date to add.
            record_date (Optional[np.datetime64]): The record date for adding this holiday; defaults to "today".
        """
        holiday_date = np.datetime64(holiday_date, "D")
        record_date = np.datetime64(record_date or "today", "D")
        self._holiday_records_cache.setdefault(self.location_code, []).append(
            (holiday_date, record_date, HolidayAction.ADD)
        )
        # Invalidate processed holiday cache for this location.
        keys_to_remove = [
            key
            for key in self._processed_holidays_cache
            if key[0] == self.location_code
        ]
        for key in keys_to_remove:
            del self._processed_holidays_cache[key]

    def remove_holiday(
        self, holiday_date: np.datetime64, record_date: Optional[np.datetime64] = None
    ) -> None:
        """
        Remove a holiday and invalidate processed holiday caches for this location.

        Args:
            holiday_date (np.datetime64): The holiday date to remove.
            record_date (Optional[np.datetime64]): The record date for removing this holiday; defaults to "today".
        """
        holiday_date = np.datetime64(holiday_date, "D")
        record_date = np.datetime64(record_date or "today", "D")
        self._holiday_records_cache.setdefault(self.location_code, []).append(
            (holiday_date, record_date, HolidayAction.REMOVE)
        )
        keys_to_remove = [
            key
            for key in self._processed_holidays_cache
            if key[0] == self.location_code
        ]
        for key in keys_to_remove:
            del self._processed_holidays_cache[key]

    def generate_schedule(
        self,
        start_date: np.datetime64,
        end_date: np.datetime64,
        frequency: Frequency,
        interval: int = 1,
        nth_day: Optional[int] = None,
        nth_business_day: Optional[int] = None,
        day_of_week: Optional[int] = None,
        roll_method: Optional[RollMethod] = None,
        use_business: bool = False,
    ) -> np.ndarray:
        """
        Generate a schedule between start_date and end_date.

        If use_business is True, business day calendars (weekdays and holidays) are used.

        Frequencies:
        - Frequency.DAILY: Every 'interval' days.
        - Frequency.BUSINESS: Every 'interval' business days.
        - Frequency.WEEKLY: One day per week (or every 'interval' weeks), targeting a weekday.
        - Frequency.MONTHLY: One day per month (or every 'interval' months) using nth_day or nth_business_day.
        - Frequency.ANNUALLY: One day per year (or every 'interval' years) using the month/day from start_date.

        If a generated candidate is not a business day and use_business is True,
        it is adjusted according to roll_method (defaulting to MODIFIED_FOLLOWING).

        Returns:
            np.ndarray: Array of np.datetime64[D] representing the schedule.
        """
        start_date = np.datetime64(start_date, "D")
        end_date = np.datetime64(end_date, "D")
        if roll_method is None:
            roll_method = RollMethod.MODIFIED_FOLLOWING

        roll_option_mapping = {
            RollMethod.FOLLOWING: "forward",
            RollMethod.PRECEDING: "backward",
            RollMethod.MODIFIED_FOLLOWING: "forward",
        }
        roll_option = roll_option_mapping.get(roll_method, "forward")

        if use_business:
            bus_cal = self._get_busdaycalendar_cached(as_of_date=None)

        if frequency == Frequency.DAILY:
            if use_business:
                n_bd = np.busday_count(
                    start_date, end_date + np.timedelta64(1, "D"), busdaycal=bus_cal
                )
                offsets = np.arange(0, n_bd, interval)
                schedule = np.busday_offset(
                    start_date, offsets, roll=roll_option, busdaycal=bus_cal
                )
                return schedule
            else:
                return np.arange(
                    start_date,
                    end_date + np.timedelta64(1, "D"),
                    np.timedelta64(interval, "D"),
                )

        elif frequency == Frequency.BUSINESS:
            if use_business:
                n_bd = np.busday_count(
                    start_date, end_date + np.timedelta64(1, "D"), busdaycal=bus_cal
                )
                offsets = np.arange(0, n_bd, interval)
                schedule = np.busday_offset(
                    start_date, offsets, roll=roll_option, busdaycal=bus_cal
                )
                return schedule
            else:
                holidays = self.get_holidays()
                all_days = np.arange(start_date, end_date + np.timedelta64(1, "D"), 1)
                is_weekday = ((all_days.astype(int) + 3) % 7) < 5
                is_not_holiday = ~np.isin(all_days, holidays)
                bd = all_days[is_weekday & is_not_holiday]
                return bd[::interval]

        elif frequency == Frequency.WEEKLY:
            if use_business:
                total_bd = np.busday_count(
                    start_date, end_date + np.timedelta64(1, "D"), busdaycal=bus_cal
                )
                offsets = np.arange(total_bd)
                all_bd = np.busday_offset(
                    start_date, offsets, roll=roll_option, busdaycal=bus_cal
                )
                if day_of_week is None:
                    day_of_week = int((start_date.astype(int) + 3) % 7)
                weekdays = (all_bd.astype(int) + 3) % 7
                candidates = all_bd[weekdays == day_of_week]
                return candidates[::interval]
            else:
                if day_of_week is None:
                    day_of_week = int((start_date.astype(int) + 3) % 7)
                start_wd = int((start_date.astype(int) + 3) % 7)
                offset = (day_of_week - start_wd) % 7
                first_valid = start_date + np.timedelta64(offset, "D")
                schedule = np.arange(
                    first_valid,
                    end_date + np.timedelta64(1, "D"),
                    np.timedelta64(7 * interval, "D"),
                )
                return schedule

        elif frequency == Frequency.MONTHLY:
            start_month = start_date.astype("datetime64[M]")
            end_month = end_date.astype("datetime64[M]")
            month_range = np.arange(
                start_month, end_month + np.timedelta64(1, "M"), dtype="datetime64[M]"
            )
            if nth_business_day is not None:
                first_days = month_range.astype("datetime64[D]")
                candidate = np.busday_offset(
                    first_days,
                    nth_business_day - 1,
                    roll=roll_option,
                    busdaycal=bus_cal,
                )
                mask = candidate.astype("datetime64[M]") != first_days.astype(
                    "datetime64[M]"
                )
                if np.any(mask):
                    last_days = (month_range + np.timedelta64(1, "M")).astype(
                        "datetime64[D]"
                    ) - np.timedelta64(1, "D")
                    last_bd = np.busday_offset(
                        last_days, 0, roll="backward", busdaycal=bus_cal
                    )
                    candidate[mask] = last_bd[mask]
            else:
                target_day = (
                    nth_day if nth_day is not None else int(str(start_date)[-2:])
                )
                # Vectorized candidate generation:
                month_str = np.datetime_as_string(month_range, unit="M")
                last_days = (month_range + np.timedelta64(1, "M")).astype(
                    "datetime64[D]"
                ) - np.timedelta64(1, "D")
                last_day_str = np.array(
                    [s[-2:] for s in np.datetime_as_string(last_days, unit="D")]
                )
                last_day_int = last_day_str.astype(int)
                candidate_day_int = np.where(
                    target_day <= last_day_int, target_day, last_day_int
                )
                candidate_day_str = np.char.mod("%02d", candidate_day_int)
                candidate_str = np.char.add(
                    np.char.add(month_str, "-"), candidate_day_str
                )
                candidate = candidate_str.astype("datetime64[D]")
                if use_business:
                    mask = ~np.is_busday(candidate, busdaycal=bus_cal)
                    if np.any(mask):
                        candidate[mask] = np.busday_offset(
                            candidate[mask], 0, roll=roll_option, busdaycal=bus_cal
                        )
            candidate = candidate[(candidate >= start_date) & (candidate <= end_date)]
            return candidate

        elif frequency == Frequency.ANNUALLY:
            start_year = int(str(start_date)[:4])
            end_year = int(str(end_date)[:4])
            years = np.arange(start_year, end_year + 1, interval)
            month_day = "-".join(str(start_date).split("-")[1:])
            years_str = np.char.mod("%04d", years)
            candidate_str = np.char.add(np.char.add(years_str, "-"), month_day)
            candidate = candidate_str.astype("datetime64[D]")
            if use_business:
                mask = ~np.is_busday(candidate, busdaycal=bus_cal)
                if np.any(mask):
                    candidate[mask] = np.busday_offset(
                        candidate[mask], 0, roll=roll_option, busdaycal=bus_cal
                    )
            candidate = candidate[(candidate >= start_date) & (candidate <= end_date)]
            return candidate

        else:
            raise ValueError(f"Unsupported frequency: {frequency}")


# Example usage
if __name__ == "__main__":
    calendar = Calendar("NYB")
    start_dt = np.datetime64("1990-01-05")
    end_dt = np.datetime64("2100-12-31")
    mid_as_of = np.datetime64("2024-06-25")
    test_as_of = mid_as_of
    holi_days = calendar.get_holidays(test_as_of)
    test_date = np.datetime64("2025-07-03")
    next_bday = calendar.get_next_business_day(test_date, holidays=holi_days)
    print(f"Next business day after {test_date} (expected: 2025-07-07): {next_bday}")
    business_days = calendar.get_business_days_in_range(
        start_dt, end_dt, holidays=holi_days
    )
    schedule = calendar.generate_schedule(
        start_dt,
        end_dt,
        Frequency.MONTHLY,
        3,
        None,
        None,
        None,
        None,
        True,
    )
    print(f"Total business days in 2025 (NYB): {schedule}")
