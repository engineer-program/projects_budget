from datetime import datetime
from src.budget_projects_creater.working_hours import WorkingHours, AttentanceEmployee


class Employee(WorkingHours, AttentanceEmployee):
    def __init__(self, last_name: str, first_name: str, patronymic=''):
        super().__init__()
        self._first_name = first_name.strip().lower()
        self._last_name = last_name.strip().lower()
        if patronymic is None:
            self._patronymic = ''
        else:
            self._patronymic = patronymic.strip().lower()
        self._position = None
        self._date_employment = None
        self._rate = None


    def get_first_name(self):
        return self._first_name
    
    
    def get_last_name(self):
        return self._last_name


    def get_patronymic(self):
        return self._patronymic


    def get_fi(self):
        return f'{self.get_last_name().capitalize()} {self.get_first_name().capitalize()}'
    

    def get_full_name(self):
        res = f'{self.get_last_name().capitalize()} {self.get_first_name().capitalize()}'
        if self._patronymic != '':
            res += f' {self.get_patronymic().capitalize()}'
        return res
    

    def get_fio(self):
        res = f'{self.get_last_name().capitalize()} {self.get_first_name()[0].upper()}.'
        if self._patronymic != '':
            res += f'{self.get_patronymic()[0].upper()}.'
        return res
    

    def set_position(self, position):
        self._position = position


    def get_position(self):
        return self._position
    

    def set_rate(self, rate):
        if rate is not None:
            self._rate = float(str(rate).replace(',', '.'))
                

    def get_rate(self):
        return self._rate
    
    
    def set_date_employment(self, date_employment):
        if date_employment is not None:
            self._date_employment = datetime.strptime(date_employment, '%Y-%m-%d')


    def get_date_employment(self):
        return self._date_employment
