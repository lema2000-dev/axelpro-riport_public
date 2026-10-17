from datetime import date
from decimal import Decimal
from io import BytesIO
from xml.sax.saxutils import escape
from xml.etree import ElementTree as ET
from urllib.error import URLError
import pytest
from sales_report import mnb_rates


def test_fetch_extracts_eur_not_first_currency_and_keeps_source_date(monkeypatch):
    payload = '<MNBCurrentExchangeRates><Day date="2026-10-09"><Rate curr="USD" unit="1">350,00</Rate><Rate curr="EUR" unit="100">40125,00</Rate></Day></MNBCurrentExchangeRates>'
    response = (
        '<s:Envelope xmlns:s="http://schemas.xmlsoap.org/soap/envelope/">'
        '<s:Body><GetCurrentExchangeRatesResponse xmlns="http://www.mnb.hu/webservices/">'
        "<GetCurrentExchangeRatesResult>"
        + escape(payload)
        + "</GetCurrentExchangeRatesResult></GetCurrentExchangeRatesResponse></s:Body></s:Envelope>"
    ).encode()

    def fake_open(request, timeout):
        assert request.full_url == "http://www.mnb.hu/arfolyamok.asmx"
        assert timeout == 15 and request.get_method() == "POST"
        assert "MNBArfolyamServiceSoap/GetCurrentExchangeRates" in request.get_header(
            "Soapaction"
        )
        assert (
            ET.fromstring(request.data).find(
                ".//{http://www.mnb.hu/webservices/}GetCurrentExchangeRates"
            )
            is not None
        )
        return BytesIO(response)

    monkeypatch.setattr(mnb_rates, "urlopen", fake_open)
    result = mnb_rates.fetch_current_rate(today=date(2026, 10, 11))
    assert result.rate == Decimal("401.25")
    assert result.source_date == date(2026, 10, 9)
    assert result.application_date == date(2026, 10, 11)


def test_network_failure_is_passed_to_caller(monkeypatch):
    def fail(*args, **kwargs):
        raise URLError("offline")

    monkeypatch.setattr(mnb_rates, "urlopen", fail)
    with pytest.raises(URLError):
        mnb_rates.fetch_current_rate()
