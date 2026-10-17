"""Worksheet names, template ranges and report target cells."""

MONTH_NAMES = (
    "January",
    "February",
    "March",
    "April",
    "May",
    "June",
    "July",
    "August",
    "September",
    "October",
    "November",
    "December",
)

WEEK_START_ROW = 1

WEEK_END_ROW = 40

WEEK_START_COLUMN = 1

WEEK_END_COLUMN = 15

MONTHLY_CLOSE_START_ROW = 43

MONTHLY_CLOSE_END_ROW = 85

MONTHLY_CLOSE_START_COLUMN = 1

MONTHLY_CLOSE_END_COLUMN = 3

PROCESSED_DOCUMENTS_SHEET = "_imported_documents"

DAILY_EXCHANGE_RATES_SHEET = "_daily_exchange_rates"

DAILY_TARGET_ROWS = {
    "used_parts": 8,
    "new_parts": 9,
    "labor": 10,
    "workshop_used_parts": 23,
    "workshop_new_parts": 24,
    "purchases": 27,
    "shipping": 32,
    "trailer": 40,
}

MONTHLY_TARGET_ROWS = {
    "consumables": 82,
    "diagnostics": 83,
    "air_conditioning": 84,
}
