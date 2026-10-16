from datetime import date, datetime
from decimal import Decimal
import pytest
from axelpro_riport.models import DailyExchangeRate


@pytest.mark.parametrize("rate", ["0", "-1", "NaN", "Infinity", "1E999"])
def test_existing_rate_validation_rejects_unusable_values(rate):
    with pytest.raises(ValueError):
        DailyExchangeRate(date(2026, 1, 5), date(2026, 1, 2), Decimal(rate))


def test_future_source_is_allowed_but_datetime_is_rejected():
    record = DailyExchangeRate(date(2026, 1, 5), date(2026, 1, 6), Decimal(400))
    assert record.source_date == date(2026, 1, 6)
    with pytest.raises(ValueError):
        DailyExchangeRate(datetime(2026, 1, 5), date(2026, 1, 2), Decimal(400))
