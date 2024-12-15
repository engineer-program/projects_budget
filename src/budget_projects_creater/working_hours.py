from calendar import monthrange
from datetime import date


AVAILABLE_YEARS = [2023,
                   2024,
                   2025,
                   2026]
MONTHS = ["Январь", 
            "Февраль", 
            "Март", 
            "Апрель", 
            "Май", 
            "Июнь", 
            "Июль", 
            "Август", 
            "Сентябрь", 
            "Октябрь", 
            "Ноябрь", 
            "Декабрь"]

# сокращенные рабочие дни в 2024 году
SHORT_DAYS = {"Январь": [], 
                "Февраль": [22], 
                "Март": [7], 
                "Апрель": [], 
                "Май": [8], 
                "Июнь": [11], 
                "Июль": [], 
                "Август": [], 
                "Сентябрь": [], 
                "Октябрь": [], 
                "Ноябрь": [2], 
                "Декабрь": []}
# праздничные и выходные дни в 2024 году
HOLIDAYS = {"Январь": [1, 2, 3, 4, 5, 6, 7, 8], 
                "Февраль": [23], 
                "Март": [8], 
                "Апрель": [29, 30], 
                "Май": [1, 9, 10], 
                "Июнь": [12], 
                "Июль": [], 
                "Август": [], 
                "Сентябрь": [], 
                "Октябрь": [], 
                "Ноябрь": [4], 
                "Декабрь": [30, 31]}
# внеплановые рабочие дни в 2024 году
OTHER_WORK_DAYS = {"Январь": [], 
                    "Февраль": [], 
                    "Март": [], 
                    "Апрель": [27], 
                    "Май": [], 
                    "Июнь": [], 
                    "Июль": [], 
                    "Август": [], 
                    "Сентябрь": [], 
                    "Октябрь": [], 
                    "Ноябрь": [], 
                    "Декабрь": [28]}


def is_working_day(date):
    day, month, dayweek = date.day, date.month, date.weekday()
    if day not in HOLIDAYS[MONTHS[month-1]] and (day in SHORT_DAYS[MONTHS[month-1]] or day in OTHER_WORK_DAYS[MONTHS[month-1]] or dayweek < 5):
        return True
    else:
        return False
    

def is_short_day(date):
    day, month = date.day, date.month
    if day in SHORT_DAYS[MONTHS[month-1]]:
        return True
    else:
        return False
    
    
def duration_work_day(date):
    if is_working_day(date):
        if is_short_day(date):
            return 7
        else:
            return 8
    else:
        return 0


def get_month_work_hours(month: str, year: int) -> int:
    res = 0
    num_month = MONTHS.index(month) + 1
    month_days = monthrange(year, num_month)[1]
    for day in range(1, month_days+1):
        res += duration_work_day(date(year, num_month, day))
    return res


class DateTimeSpent:
    def __init__(self):
        self._date_time_spent = {}


    def add_project_work_hours(self, name_project: str, hours: float) -> None:
        self._date_time_spent[name_project] = self._date_time_spent.get(name_project, 0) + hours


    def get_date_time_spent(self):
        return self._date_time_spent
    

    def get_project_work_hours(self, name_project: str) -> float:
        return self._date_time_spent.get(name_project, 0)
    

class WorkingHours:
    def __init__(self):
        self._work_hours = {}
    

    def get_work_hours(self):
        return self._work_hours


    def add_work_hours(self, date, name_project: str, hours: float) -> None:
        if date in self.get_work_hours():
            self._work_hours[date].add_project_work_hours(name_project, hours)
        else:
            self._work_hours[date] = DateTimeSpent()
            self._work_hours[date].add_project_work_hours(name_project, hours)

    
    def get_date_work_hours(self, date) -> dict:
        return self._work_hours.get(date, DateTimeSpent()).get_date_time_spent()
    

    def get_date_project_work_hours(self, date, name_project: str) -> float:
        return self._work_hours.get(date, DateTimeSpent()).get_project_work_hours(name_project)

          
class AttentanceEmployee:
    def __init__(self):
        self._attendance = {}
    

    def add_attendance(self, date, attendance) -> None:
        self._attendance[date] = self._attendance.get(date, [])
        self._attendance[date].append(attendance)


    def get_date_attendance(self, date):
        return self._attendance.get(date, [])
    

    def get_attendance(self):
        return self._attendance
