from calendar import Calendar
from datetime import date

def get_month_weeks(year: int, month: int) -> list[list[date]]:
    calendar = Calendar(firstweekday=0)

    return [
        [
            day for day in week if day.year == year and day.month == month
        ]
        for week in calendar.monthdatescalendar(year, month)
    ]

