from datetime import datetime, timedelta
import calendar
from typing import List

class CourseScheduler:
    def __init__(self, year: int, holidays: List[str]):
        self.year = year
        # Convert holiday strings to datetime objects with proper format
        self.holidays = []
        for holiday in holidays:
            try:
                # Expecting DD.MM.YYYY format
                holiday_date = datetime.strptime(holiday, '%d.%m.%Y')
                self.holidays.append(holiday_date)
            except ValueError:
                print(f"Invalid holiday format: {holiday}")
        print(f"Initialized with {len(self.holidays)} holidays: {self.holidays}")  # Debug print

    def is_business_day(self, date: datetime) -> bool:
        """Check if date is a business day (Mon-Fri and not a holiday)."""
        # First check if it's a weekend
        if date.weekday() >= 5:  # 5 = Saturday, 6 = Sunday
            return False

        # Then check if it's a holiday
        for holiday in self.holidays:
            if (date.day == holiday.day and
                date.month == holiday.month and
                date.year == holiday.year):
                return False
        return True

    def can_schedule_course(self, start_date: datetime, duration: int) -> bool:
        """Check if a course can be scheduled starting from given date."""
        if not self.is_business_day(start_date):
            return False

        current_date = start_date
        business_days = 0
        allow_cross_period = duration > 5
        week_start = start_date - timedelta(days=start_date.weekday())

        while business_days < duration:
            if not allow_cross_period:
                # Courses lasting 5 days or fewer must stay in the same work week.
                if current_date - week_start >= timedelta(days=5):
                    return False

                # Courses lasting 5 days or fewer must stay in the same month.
                if current_date.month != start_date.month:
                    return False

            if self.is_business_day(current_date):
                business_days += 1
            current_date += timedelta(days=1)

        return True

    def get_available_start_days(self, month: int, duration: int) -> List[datetime]:
        """Get all possible start dates for a course in given month."""
        available_dates = []

        _, last_day = calendar.monthrange(self.year, month)
        current_date = datetime(self.year, month, 1)
        end_date = datetime(self.year, month, last_day)

        while current_date <= end_date:
            if self.can_schedule_course(current_date, duration):
                available_dates.append(current_date)
            current_date += timedelta(days=1)

        return available_dates

    def format_date_range(self, start_date: datetime, duration: int) -> str:
        """Format date range according to requirements."""
        if duration == 1:
            return start_date.strftime('%d.%m.%Y')

        business_days = 0
        current_date = start_date
        while business_days < duration:
            if self.is_business_day(current_date):
                business_days += 1
            if business_days < duration:
                current_date += timedelta(days=1)

        return f"{start_date.strftime('%d')}-{current_date.strftime('%d.%m.%Y')}"