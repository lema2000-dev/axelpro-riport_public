"""Az aktuális EUR-árfolyam lekérése az MNB-től."""

from datetime import date
from decimal import Decimal
from urllib.request import Request, urlopen
from xml.etree import ElementTree as ET

from .models import DailyExchangeRate


_ENDPOINT = "http://www.mnb.hu/arfolyamok.asmx"
_NAMESPACE = "http://www.mnb.hu/webservices/"


def _parse_current_rate(
    response: bytes,
    application_date: date,
) -> DailyExchangeRate:
    """Az EUR-árfolyam és a forrásdátum kiolvasása a válaszból."""
    envelope = ET.fromstring(response)

    result = envelope.find(
        f".//{{{_NAMESPACE}}}GetCurrentExchangeRatesResult"
    )
    document = ET.fromstring(result.text)

    day = document.find("Day")
    euro_rate = day.find("Rate[@curr='EUR']")

    return DailyExchangeRate(
        application_date=application_date,
        source_date=date.fromisoformat(day.attrib["date"]),
        rate=(
            Decimal(euro_rate.text.replace(",", "."))
            / Decimal(euro_rate.attrib["unit"])
        ),
    )


def fetch_current_rate(
    *,
    today: date | None = None,
) -> DailyExchangeRate:
    """Az aktuális EUR-árfolyam lekérése a megadott alkalmazási naphoz."""
    if today is None:
        today = date.today()

    body = (
        '<?xml version="1.0" encoding="utf-8"?>'
        '<soap:Envelope '
        'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">'
        "<soap:Body>"
        f'<GetCurrentExchangeRates xmlns="{_NAMESPACE}"/>'
        "</soap:Body>"
        "</soap:Envelope>"
    ).encode("utf-8")

    request = Request(
        url=_ENDPOINT,
        data=body,
        headers={
            "Content-Type": "text/xml; charset=utf-8",
            "SOAPAction": (
                f'"{_NAMESPACE}MNBArfolyamServiceSoap/'
                'GetCurrentExchangeRates"'
            ),
        },
    )

    with urlopen(request, timeout=15) as response:
        return _parse_current_rate(
            response=response.read(),
            application_date=today,
        )