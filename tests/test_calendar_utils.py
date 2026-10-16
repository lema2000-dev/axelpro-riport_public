from calendar import monthrange
from datetime import date, timedelta
import pytest
from axelpro_riport.calendar_utils import get_month_weeks


@pytest.mark.parametrize("year", [2024, 2026, 2027])
@pytest.mark.parametrize("month", range(1, 13))
def test_month_contains_every_day_once_in_calendar_order(year, month):
    weeks = get_month_weeks(year, month)
    days = [day for week in weeks for day in week]
    assert days == [
        date(year, month, n) for n in range(1, monthrange(year, month)[1] + 1)
    ]
    for week in weeks:
        assert len({day.isocalendar()[:2] for day in week}) == 1
        assert all(b - a == timedelta(days=1) for a, b in zip(week, week[1:]))


def test_shared_week_keeps_iso_number_at_year_boundary():
    assert get_month_weeks(2026, 12)[-1][0].isocalendar()[:2] == (2026, 53)
    assert get_month_weeks(2027, 1)[0][0].isocalendar()[:2] == (2026, 53)
