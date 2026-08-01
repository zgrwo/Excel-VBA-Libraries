"""Cross-validate DateTimeUtils functions against Python reference implementations.

Usage: python tests/build_DateTimeUtils.py
"""

import os
import sys
from datetime import date, datetime, timedelta
import calendar

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tests.crossval.build_common import CrossValRunner
from tests.test_utils import SRC_DIR, VBA_CORE_DIR, VBA_CORE_IMPORT_ORDER

MODULE_PATHS = [os.path.join(VBA_CORE_DIR, name + ".cls")
                for name in VBA_CORE_IMPORT_ORDER]
MODULE_PATHS.append(os.path.join(SRC_DIR, "DateTimeUtils.bas"))


# =============================================================================
# Helpers
# =============================================================================

def _dt(y, m, d):
    """Create a datetime for COM marshaling (VBA Date literal)."""
    return datetime(y, m, d)


def _py_iso_week_num(d):
    """ISO 8601 week number from a datetime."""
    return d.isocalendar()[1]


def _py_easter(year):
    """Compute Easter Sunday using the Anonymous Gregorian algorithm."""
    a = year % 19
    b = year // 100
    c = year % 100
    d_val = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d_val - g + 15) % 30
    i = c // 4
    k = c % 4
    ll = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * ll) // 451
    month = (h + ll - 7 * m + 114) // 31
    day = ((h + ll - 7 * m + 114) % 31) + 1
    return date(year, month, day)


# =============================================================================
# Test Cases
# =============================================================================

TEST_CASES = [

    # ---- ISOWeekNum ----
    {
        "name": "ISOWeekNum_jan1_2025",
        "func": "ISOWeekNum",
        "args": lambda: (_dt(2025, 1, 1),),
        "py_ref": lambda a: _py_iso_week_num(a[0]),
        "result_type": "scalar",
    },
    {
        "name": "ISOWeekNum_dec30_2024",
        "func": "ISOWeekNum",
        "args": lambda: (_dt(2024, 12, 30),),
        "py_ref": lambda a: _py_iso_week_num(a[0]),
        "result_type": "scalar",
    },

    # ---- IsLeapYear ----
    {"name": "IsLeapYear_2024", "func": "IsLeapYear",
     "args": lambda: (2024,), "py_ref": lambda a: calendar.isleap(a[0]), "result_type": "bool"},
    {"name": "IsLeapYear_2023", "func": "IsLeapYear",
     "args": lambda: (2023,), "py_ref": lambda a: calendar.isleap(a[0]), "result_type": "bool"},
    {"name": "IsLeapYear_2000", "func": "IsLeapYear",
     "args": lambda: (2000,), "py_ref": lambda a: calendar.isleap(a[0]), "result_type": "bool"},
    {"name": "IsLeapYear_1900", "func": "IsLeapYear",
     "args": lambda: (1900,), "py_ref": lambda a: calendar.isleap(a[0]), "result_type": "bool"},

    # ---- DaysInMonth ----
    {"name": "DaysInMonth_feb_leap", "func": "DaysInMonth",
     "args": lambda: (_dt(2024, 2, 1),),
     "py_ref": lambda a: calendar.monthrange(a[0].year, a[0].month)[1], "result_type": "scalar"},
    {"name": "DaysInMonth_feb_nonleap", "func": "DaysInMonth",
     "args": lambda: (_dt(2023, 2, 1),),
     "py_ref": lambda a: 28, "result_type": "scalar"},
    {"name": "DaysInMonth_jan", "func": "DaysInMonth",
     "args": lambda: (_dt(2024, 1, 1),),
     "py_ref": lambda a: 31, "result_type": "scalar"},

    # ---- Quarter ----
    {"name": "Quarter_q1", "func": "Quarter",
     "args": lambda: (_dt(2024, 2, 15),), "py_ref": lambda a: 1, "result_type": "scalar"},
    {"name": "Quarter_q3", "func": "Quarter",
     "args": lambda: (_dt(2024, 7, 15),), "py_ref": lambda a: 3, "result_type": "scalar"},

    # ---- DayOfYear ----
    {"name": "DayOfYear_jan1", "func": "DayOfYear",
     "args": lambda: (_dt(2024, 1, 1),), "py_ref": lambda a: 1, "result_type": "scalar"},
    {"name": "DayOfYear_dec31_leap", "func": "DayOfYear",
     "args": lambda: (_dt(2024, 12, 31),), "py_ref": lambda a: 366, "result_type": "scalar"},
    {"name": "DayOfYear_dec31_nonleap", "func": "DayOfYear",
     "args": lambda: (_dt(2023, 12, 31),), "py_ref": lambda a: 365, "result_type": "scalar"},

    # ---- FirstDayOfMonth / LastDayOfMonth ----
    {"name": "FirstDayOfMonth", "func": "FirstDayOfMonth",
     "args": lambda: (_dt(2024, 2, 15),),
     "py_ref": lambda a: _dt(2024, 2, 1), "result_type": "scalar"},
    {"name": "LastDayOfMonth_feb_leap", "func": "LastDayOfMonth",
     "args": lambda: (_dt(2024, 2, 15),),
     "py_ref": lambda a: _dt(2024, 2, 29), "result_type": "scalar"},
    {"name": "LastDayOfMonth_feb_nonleap", "func": "LastDayOfMonth",
     "args": lambda: (_dt(2023, 2, 15),),
     "py_ref": lambda a: _dt(2023, 2, 28), "result_type": "scalar"},

    # ---- AgeYears ----
    {"name": "AgeYears_35", "func": "AgeYears",
     "args": lambda: (_dt(1990, 6, 15), _dt(2025, 6, 15)),
     "py_ref": lambda a: 35, "result_type": "scalar"},
    {"name": "AgeYears_not_yet", "func": "AgeYears",
     "args": lambda: (_dt(1990, 12, 31), _dt(2025, 6, 15)),
     "py_ref": lambda a: 34, "result_type": "scalar"},

    # ---- StartOfWeek ----
    {"name": "StartOfWeek_monday", "func": "StartOfWeek",
     "args": lambda: (_dt(2025, 6, 11), 2),  # vbMonday=2
     "py_ref": lambda a: _dt(2025, 6, 9), "result_type": "scalar"},

    # ---- Easter ----
    {"name": "Easter_2025", "func": "Easter",
     "args": lambda: (2025,), "py_ref": lambda a: _py_easter(a[0]),
     "result_type": "scalar"},
    {"name": "Easter_2024", "func": "Easter",
     "args": lambda: (2024,), "py_ref": lambda a: _py_easter(a[0]),
     "result_type": "scalar"},
    {"name": "Easter_2000", "func": "Easter",
     "args": lambda: (2000,), "py_ref": lambda a: _py_easter(a[0]),
     "result_type": "scalar"},

    # ---- IsWeekend ----
    {"name": "IsWeekend_saturday", "func": "IsWeekend",
     "args": lambda: (_dt(2025, 6, 14),), "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "IsWeekend_wednesday", "func": "IsWeekend",
     "args": lambda: (_dt(2025, 6, 11),), "py_ref": lambda a: False, "result_type": "bool"},

    # ---- UnixToDate / DateToUnix ----
    {"name": "DateToUnix_known", "func": "DateToUnix",
     "args": lambda: (_dt(2025, 6, 10),),
     "py_ref": lambda a: int((datetime(2025, 6, 10) - datetime(1970, 1, 1)).total_seconds()),
     "result_type": "scalar", "tol": 1.0},

    # ---- DaysInYear ----
    {"name": "DaysInYear_leap", "func": "DaysInYear",
     "args": lambda: (2024,), "py_ref": lambda a: 366, "result_type": "scalar"},
    {"name": "DaysInYear_nonleap", "func": "DaysInYear",
     "args": lambda: (2023,), "py_ref": lambda a: 365, "result_type": "scalar"},

    # ---- FirstDayOfQuarter / LastDayOfQuarter ----
    {"name": "FirstDayOfQuarter_q2", "func": "FirstDayOfQuarter",
     "args": lambda: (_dt(2024, 5, 15),),
     "py_ref": lambda a: _dt(2024, 4, 1), "result_type": "scalar"},
    {"name": "LastDayOfQuarter_q1", "func": "LastDayOfQuarter",
     "args": lambda: (_dt(2024, 2, 15),),
     "py_ref": lambda a: _dt(2024, 3, 31), "result_type": "scalar"},

    # =========================================================================
    # Missing function coverage (P2-3 additions)
    # =========================================================================

    # ---- FiscalYear ----
    {"name": "FiscalYear_default", "func": "FiscalYear",
     "args": lambda: (_dt(2024, 6, 15),),
     "py_ref": lambda a: 2024, "result_type": "scalar"},
    {"name": "FiscalYear_jul_start", "func": "FiscalYear",
     "args": lambda: (_dt(2024, 6, 15), 7),  # June < July → FY=2023
     "py_ref": lambda a: 2023, "result_type": "scalar"},
    {"name": "FiscalYear_jul_prev", "func": "FiscalYear",
     "args": lambda: (_dt(2024, 3, 15), 7),
     "py_ref": lambda a: 2023, "result_type": "scalar"},

    # ---- NthWeekday — n-th weekday of month ----
    {"name": "NthWeekday_2nd_mon", "func": "NthWeekday",
     "args": lambda: (2024, 6, 2, 2),  # 2nd Monday of June 2024
     "py_ref": lambda a: _py_nth_weekday(a[0], a[1], a[2], a[3]),
     "result_type": "scalar"},
    {"name": "NthWeekday_4th_fri", "func": "NthWeekday",
     "args": lambda: (2024, 11, 4, 6),  # 4th Friday of Nov 2024
     "py_ref": lambda a: _py_nth_weekday(a[0], a[1], a[2], a[3]),
     "result_type": "scalar"},

    # ---- WeekOfMonth ----
    {"name": "WeekOfMonth_mid", "func": "WeekOfMonth",
     "args": lambda: (_dt(2024, 6, 15),),
     "py_ref": lambda a: 3, "result_type": "scalar"},
    {"name": "WeekOfMonth_first", "func": "WeekOfMonth",
     "args": lambda: (_dt(2024, 6, 1),),
     "py_ref": lambda a: 1, "result_type": "scalar"},

    # ---- AddMonthsSafe ----
    {"name": "AddMonthsSafe_normal", "func": "AddMonthsSafe",
     "args": lambda: (_dt(2024, 1, 15), 3),
     "py_ref": lambda a: _dt(2024, 4, 15), "result_type": "scalar"},
    {"name": "AddMonthsSafe_eom", "func": "AddMonthsSafe",
     "args": lambda: (_dt(2024, 1, 31), 1),  # Jan 31 + 1month = Feb 29(leap)
     "py_ref": lambda a: _dt(2024, 2, 29), "result_type": "scalar"},

    # ---- DateRange (date arrays not comparable through COM float conversion) ----
    {"name": "DateRange_days", "func": "DateRange",
     "args": lambda: (_dt(2024, 6, 1), _dt(2024, 6, 3), "d"),
     "py_ref": lambda a: [_dt(2024, 6, 1), _dt(2024, 6, 2), _dt(2024, 6, 3)],
     "result_type": "array",
     "skip_if": True,
     "skip_reason": "Date arrays marshal as pywintypes.datetime through COM; float conversion fails"},

    # ---- WorkdaysBetween ----
    {"name": "WorkdaysBetween_5days", "func": "WorkdaysBetween",
     "args": lambda: (_dt(2024, 6, 3), _dt(2024, 6, 7)),  # Mon-Fri = 5 workdays
     "py_ref": lambda a: 5, "result_type": "scalar"},
    {"name": "WorkdaysBetween_cross_weekend", "func": "WorkdaysBetween",
     "args": lambda: (_dt(2024, 6, 7), _dt(2024, 6, 10)),  # Fri-Mon = 2 workdays
     "py_ref": lambda a: 2, "result_type": "scalar"},

    # ---- EndOfWeek ----
    {"name": "EndOfWeek_wed_to_sun", "func": "EndOfWeek",
     "args": lambda: (_dt(2024, 6, 12), 2),
     "py_ref": lambda a: _py_end_of_week(a[0], a[1]),
     "result_type": "scalar"},

    # ---- NextWorkday ----
    {"name": "NextWorkday_fwd1", "func": "NextWorkday",
     "args": lambda: (_dt(2024, 6, 3), 1),  # Mon→Tue
     "py_ref": lambda a: _py_next_workday(a[0], a[1], None),
     "result_type": "scalar"},
    {"name": "NextWorkday_skip_wknd", "func": "NextWorkday",
     "args": lambda: (_dt(2024, 6, 7), 1),  # Fri→Mon
     "py_ref": lambda a: _py_next_workday(a[0], a[1], None),
     "result_type": "scalar"},

    # ---- IsHoliday (skip: date arrays marshal inconsistently through COM) ----
    {"name": "IsHoliday_in_list", "func": "IsHoliday",
     "args": lambda: (_dt(2024, 12, 25), [_dt(2024, 12, 25), _dt(2024, 1, 1)]),
     "py_ref": lambda a: True, "result_type": "bool",
     "skip_if": True,
     "skip_reason": "Date arrays marshal as pywintypes.datetime through COM; verified manually in Excel"},
    {"name": "IsHoliday_not_found", "func": "IsHoliday",
     "args": lambda: (_dt(2024, 12, 24), [_dt(2024, 12, 25)]),
     "py_ref": lambda a: False, "result_type": "bool",
     "skip_if": True,
     "skip_reason": "Date arrays marshal as pywintypes.datetime through COM; verified manually in Excel"},
    # =========================================================================
    # Boundary / edge cases
    # =========================================================================

    # ---- ISOWeekNum boundaries ----
    {"name": "ISOWeekNum_2028_12_31", "func": "ISOWeekNum",
     "args": lambda: (_dt(2028, 12, 31),),  # Sunday, ISO week 52 of 2028 (week 1 starts 2029-01-01 Mon)
     "py_ref": lambda a: 52, "result_type": "scalar"},
    {"name": "ISOWeekNum_2026_01_01", "func": "ISOWeekNum",
     "args": lambda: (_dt(2026, 1, 1),),  # Thursday, ISO week 1 of 2026
     "py_ref": lambda a: 1, "result_type": "scalar"},
    {"name": "ISOWeekNum_jan1_2017", "func": "ISOWeekNum",
     "args": lambda: (_dt(2017, 1, 1),),  # Sunday, ISO week 52 of 2016
     "py_ref": lambda a: 52, "result_type": "scalar"},

    # ---- IsLeapYear century boundaries ----
    {"name": "IsLeapYear_2100", "func": "IsLeapYear",
     "args": lambda: (2100,), "py_ref": lambda a: False, "result_type": "bool"},
    {"name": "IsLeapYear_2400", "func": "IsLeapYear",
     "args": lambda: (2400,), "py_ref": lambda a: True, "result_type": "bool"},

    # ---- DaysInMonth boundaries ----
    {"name": "DaysInMonth_dec", "func": "DaysInMonth",
     "args": lambda: (_dt(2024, 12, 1),), "py_ref": lambda a: 31, "result_type": "scalar"},
    {"name": "DaysInMonth_apr", "func": "DaysInMonth",
     "args": lambda: (_dt(2024, 4, 1),), "py_ref": lambda a: 30, "result_type": "scalar"},

    # ---- Quarter boundaries ----
    {"name": "Quarter_mar31", "func": "Quarter",
     "args": lambda: (_dt(2024, 3, 31),), "py_ref": lambda a: 1, "result_type": "scalar"},
    {"name": "Quarter_apr1", "func": "Quarter",
     "args": lambda: (_dt(2024, 4, 1),), "py_ref": lambda a: 2, "result_type": "scalar"},
    {"name": "Quarter_dec31", "func": "Quarter",
     "args": lambda: (_dt(2024, 12, 31),), "py_ref": lambda a: 4, "result_type": "scalar"},

    # ---- DayOfYear leap year boundaries ----
    {"name": "DayOfYear_feb29_leap", "func": "DayOfYear",
     "args": lambda: (_dt(2024, 2, 29),), "py_ref": lambda a: 60, "result_type": "scalar"},
    {"name": "DayOfYear_mar1_leap", "func": "DayOfYear",
     "args": lambda: (_dt(2024, 3, 1),), "py_ref": lambda a: 61, "result_type": "scalar"},

    # ---- FirstDayOfMonth / LastDayOfMonth boundaries ----
    {"name": "FirstDayOfMonth_jan", "func": "FirstDayOfMonth",
     "args": lambda: (_dt(2024, 1, 15),), "py_ref": lambda a: _dt(2024, 1, 1), "result_type": "scalar"},
    {"name": "LastDayOfMonth_jan", "func": "LastDayOfMonth",
     "args": lambda: (_dt(2024, 1, 15),), "py_ref": lambda a: _dt(2024, 1, 31), "result_type": "scalar"},
    {"name": "LastDayOfMonth_dec", "func": "LastDayOfMonth",
     "args": lambda: (_dt(2024, 12, 1),), "py_ref": lambda a: _dt(2024, 12, 31), "result_type": "scalar"},

    # ---- AgeYears boundary ----
    {"name": "AgeYears_same_day", "func": "AgeYears",
     "args": lambda: (_dt(1990, 6, 15), _dt(1990, 6, 15)),
     "py_ref": lambda a: 0, "result_type": "scalar"},
    {"name": "AgeYears_one_day_short", "func": "AgeYears",
     "args": lambda: (_dt(1990, 6, 15), _dt(1991, 6, 14)),
     "py_ref": lambda a: 0, "result_type": "scalar"},

    # ---- StartOfWeek / EndOfWeek boundaries ----
    {"name": "StartOfWeek_already_mon", "func": "StartOfWeek",
     "args": lambda: (_dt(2025, 6, 9), 2),
     "py_ref": lambda a: _dt(2025, 6, 9), "result_type": "scalar"},
    {"name": "EndOfWeek_already_sun", "func": "EndOfWeek",
     "args": lambda: (_dt(2025, 6, 15), 2),
     "py_ref": lambda a: _dt(2025, 6, 15), "result_type": "scalar"},

    # ---- IsWeekend Sunday boundary ----
    {"name": "IsWeekend_sunday", "func": "IsWeekend",
     "args": lambda: (_dt(2025, 6, 15),), "py_ref": lambda a: True, "result_type": "bool"},
    {"name": "IsWeekend_monday", "func": "IsWeekend",
     "args": lambda: (_dt(2025, 6, 16),), "py_ref": lambda a: False, "result_type": "bool"},

    # ---- DateToUnix epoch boundaries ----
    {"name": "DateToUnix_epoch", "func": "DateToUnix",
     "args": lambda: (_dt(1970, 1, 1),),
     "py_ref": lambda a: 0, "result_type": "scalar", "tol": 1.0},
    {"name": "DateToUnix_pre_epoch", "func": "DateToUnix",
     "args": lambda: (_dt(1969, 12, 31),),
     "py_ref": lambda a: -86400, "result_type": "scalar", "tol": 1.0},

    # ---- Easter boundaries ----
    {"name": "Easter_2035", "func": "Easter",
     "args": lambda: (2035,), "py_ref": lambda a: _py_easter(a[0]),
     "result_type": "scalar"},

    # ---- FiscalYear boundaries ----
    {"name": "FiscalYear_at_start_month", "func": "FiscalYear",
     "args": lambda: (_dt(2024, 7, 1), 7),  # July 1 with July start -> FY 2024
     "py_ref": lambda a: 2024, "result_type": "scalar"},
    {"name": "FiscalYear_dec_jan_start", "func": "FiscalYear",
     "args": lambda: (_dt(2024, 2, 15), 1),  # Feb with Jan start -> FY 2024
     "py_ref": lambda a: 2024, "result_type": "scalar"},

    # ---- NthWeekday boundaries ----
    {"name": "NthWeekday_1st_wed", "func": "NthWeekday",
     "args": lambda: (2024, 6, 1, 4),  # 1st Wednesday of June 2024
     "py_ref": lambda a: _py_nth_weekday(a[0], a[1], a[2], a[3]),
     "result_type": "scalar"},

    # ---- WeekOfMonth boundary ----
    {"name": "WeekOfMonth_day31", "func": "WeekOfMonth",
     "args": lambda: (_dt(2024, 3, 31),),  # 5th Sunday of March 2024
     "py_ref": lambda a: 5, "result_type": "scalar"},

    # ---- AddMonthsSafe boundaries ----
    {"name": "AddMonthsSafe_cross_year", "func": "AddMonthsSafe",
     "args": lambda: (_dt(2024, 11, 15), 3),  # Nov + 3 = Feb 2025
     "py_ref": lambda a: _dt(2025, 2, 15), "result_type": "scalar"},
    {"name": "AddMonthsSafe_neg_months", "func": "AddMonthsSafe",
     "args": lambda: (_dt(2024, 6, 15), -3),
     "py_ref": lambda a: _dt(2024, 3, 15), "result_type": "scalar"},

    # ---- WorkdaysBetween boundaries ----
    {"name": "WorkdaysBetween_same_day", "func": "WorkdaysBetween",
     "args": lambda: (_dt(2024, 6, 5), _dt(2024, 6, 5)),  # inclusive of both ends
     "py_ref": lambda a: 1, "result_type": "scalar"},

    # ---- NextWorkday boundaries ----
    {"name": "NextWorkday_n5", "func": "NextWorkday",
     "args": lambda: (_dt(2024, 6, 3), 5),  # Mon + 5 = next Mon
     "py_ref": lambda a: _py_next_workday(a[0], a[1], None),
     "result_type": "scalar"},

    # ---- UDF wrapper ----
    {"name": "UDF_DT_ISOWEEKNUM", "func": "UDF_DT_ISOWEEKNUM",
     "args": lambda: (_dt(2025, 1, 1),),
     "py_ref": lambda a: 1,
     "result_type": "scalar"},

    # =====================================================================
    # Migrated from VBA Test_DateTimeUtils — coverage gaps (2026-06-16)
    # =====================================================================

    # ---- UnixToDate — round-trip and known timestamp ----
    {"name": "UnixToDate_roundtrip", "func": "UnixToDate",
     "args": lambda: (int((datetime(2024, 6, 15) - datetime(1970, 1, 1)).total_seconds()),),
     "py_ref": lambda a: datetime(2024, 6, 15),
     "result_type": "scalar", "tol": 1.0},
    {"name": "UnixToDate_known", "func": "UnixToDate",
     "args": lambda: (1735689600,),  # 2025-01-01 00:00:00 UTC
     "py_ref": lambda a: datetime(2025, 1, 1),
     "result_type": "scalar", "tol": 1.0},

    # ---- Age / DateDiffParts — Dictionary-returning, uses dict_keys compare mode ----
    {"name": "Age_keys", "func": "Age",
     "args": lambda: (datetime(1990, 6, 15),),
     "py_ref": lambda a: ["years", "months", "days", "totalYears", "totalMonths", "totalDays"],
     "compare_mode": "dict_keys"},
    {"name": "Age_values", "func": "Age",
     "args": lambda: (datetime(1990, 6, 15),),
     "py_ref": lambda a: [("years", 30, 50), ("months", 0, 200),
                          ("totalMonths", 300, 600), ("totalDays", 9000, 20000)],
     "compare_mode": "dict_keys_range"},
    {"name": "DateDiffParts_keys", "func": "DateDiffParts",
     "args": lambda: (datetime(2020, 1, 1), datetime(2021, 1, 1)),
     "py_ref": lambda a: ["years", "months", "days", "totalDays", "totalMonths"],
     "compare_mode": "dict_keys"},
    {"name": "DateDiffParts_values", "func": "DateDiffParts",
     "args": lambda: (datetime(2020, 1, 1), datetime(2021, 1, 1)),
     "py_ref": lambda a: [("years", 1.0, 1.0), ("months", 0.0, 0.0),
                          ("days", 0.0, 0.0), ("totalDays", 365, 366)],
     "compare_mode": "dict_keys_range"},

    # ---- Coverage: NthWeekday, WorkdaysBetween, Easter, DateRange ----
    {"name": "DaysInMonth_jan", "func": "DaysInMonth",
     "args": lambda: (2024, 1), "py_ref": lambda a: 31, "result_type": "scalar"},
    {"name": "DaysInMonth_apr", "func": "DaysInMonth",
     "args": lambda: (2024, 4), "py_ref": lambda a: 30, "result_type": "scalar"},
    {"name": "Easter_2025", "func": "Easter",
     "args": lambda: (2025,), "py_ref": lambda a: date(2025, 4, 20),
     "result_type": "scalar"},
    {"name": "Easter_2030", "func": "Easter",
     "args": lambda: (2030,), "py_ref": lambda a: date(2030, 4, 21),
     "result_type": "scalar"},
    {"name": "NthWeekday_1stMon", "func": "NthWeekday",
     "args": lambda: (2024, 2, 1, 2), "py_ref": lambda a: date(2024, 2, 5),
     "result_type": "scalar"},
    {"name": "DateRange_day_step", "func": "DateRange",
     "args": lambda: (date(2024, 1, 1), date(2024, 1, 3), "d"),
     "py_ref": lambda a: [date(2024, 1, 1), date(2024, 1, 2), date(2024, 1, 3)],
     "result_type": "array", "tol": 1.0,
     "skip_if": True,
     "skip_reason": "COM无法传递Python datetime.date对象，VBA函数行为正确"},

]


# Missing function Python references
def _py_nth_weekday(y, m, n, dow):
    """n-th weekday (dow: VBA 1=Sun..7=Sat) in month m of year y."""
    import calendar
    py_dow = (dow - 2) % 7  # VBA weekday -> Python calendar (0=Mon..6=Sun)
    cal = calendar.monthcalendar(y, m)
    cnt = 0
    for week in cal:
        if week[py_dow] != 0:
            cnt += 1
            if cnt == n:
                return datetime(y, m, week[py_dow])
    return None


def _py_end_of_week(d, start_day):
    """End of week: last day of the week containing d (start_day: 1=Sun..7=Sat)."""
    py_dow = (start_day - 2) % 7  # VBA → Python (0=Mon..6=Sun)
    days_until_end = (6 - py_dow) - ((d.weekday() - py_dow) % 7)
    if days_until_end < 0:
        days_until_end += 7
    return d + timedelta(days=days_until_end)


def _py_next_workday(start, n, holidays):
    """Next workday: start + n workdays, skipping weekends and optional holidays.

    NOTE: VBA NextWorkday uses Int(startDate) to strip time components.
    This Python ref does not truncate — test cases use date-only datetimes
    so this is equivalent, but could diverge for datetime-with-time inputs.
    """
    holiday_set = set()
    if holidays:
        for h in holidays:
            if hasattr(h, 'date'):
                holiday_set.add(h.date())
            else:
                holiday_set.add(h)
    current = start
    direction = 1 if n >= 0 else -1
    remaining = abs(n)
    while remaining > 0:
        current += timedelta(days=direction)
        if current.weekday() < 5:  # Mon-Fri
            if hasattr(current, 'date'):
                cd = current.date()
            else:
                cd = current
            if cd not in holiday_set:
                remaining -= 1
    return current


def main() -> int:
    runner = CrossValRunner("DateTimeUtils", MODULE_PATHS)
    runner.run_all(TEST_CASES)
    passed, failed = runner.print_summary()
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main())
