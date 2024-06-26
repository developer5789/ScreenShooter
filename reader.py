from collections import defaultdict
import openpyxl
import datetime
import os



class  Msk_Reader:
    """Класс читает эксель с рейсами и заполняет словарь self.routes_dict"""

    def __init__(self, app, file_path=None):
        """Инициализация аттрибутов

        Аргументы:
            file_path(str, None): путь до файла с рейсами,
            app(App): главное окно приложения.
        """
        self.app = app
        self.file_path = file_path
        self.final_report_path = None
        self.routes_dict = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        self.wb = None
        self.total = 0
        self.paths_json = {}

    def read(self):
        """Запускает чтение файла эксель с рейсами."""
        self.wb = openpyxl.load_workbook(self.file_path, keep_links=False, data_only=True)
        sheet = self.wb.active
        counter = 0
        step = sheet.max_row // 10
        step_counter = step
        for row in sheet.rows:
            counter += 1
            step_counter -= 1
            if step_counter == 0:
                self.app.event_generate('<<Updated>>', when='tail')
                step_counter = 0
            if type(row[2].value) == datetime.datetime:
                screen = self.get_screen_value(row[18].value)
                if screen in ('', '1', '0'):
                    self.total += 1
                    date = self.convert_date_to_str(row[2].value)
                    route = self.get_int(row[1].value)
                    bus_numb = self.get_bus_numb(row[11].value)
                    self.routes_dict[date][route][bus_numb].append(
                        {
                            'direction': self.get_int(row[3].value),
                            'start_time': self.convert_time_to_str(row[7].value),
                            'finish_time': self.convert_time_to_str(row[8].value),
                            'screen': screen,
                            'row_numb': counter,
                            'bus_numb': bus_numb,
                            'route': route,
                            'route_numb': row[0].value
                        }
                    )

        self.app.load_window.progress_text_var.set(f'Файл {self.file_path} прочитан')
        if self.report_type == 'НС':
            self.read_json()
        self.app.load_window.progress_var.set(100)
        self.app.load_window.ok_btn['state'] = 'normal'


    @staticmethod
    def get_int(value):
        try:
            return int(value)
        except Exception:
            return value if value is not None else ''


    @staticmethod
    def convert_time_to_str(time_obj: datetime.time):
        """Возвращает строку со временем в формате часы:минуты

            Аргументы:
                time_obj(datetime.time): значение, считанное из колонки со временем.
        """
        if time_obj is None or time_obj == '':
            return ''
        if time_obj and type(time_obj) == str:
            return time_obj
        if type(time_obj) == datetime.time or type(time_obj) == datetime.datetime:
            return time_obj.strftime('%H:%M')

    @staticmethod
    def convert_date_to_str(date_obj: datetime.datetime):
        """Возвращает строку с датой в формате день.месяц.год

        Аргументы :
            st(str, datetime.datetime): значение, считанное из колонки с датой.
        """
        if type(date_obj) == datetime.datetime:
            return date_obj.strftime('%d.%m.%Y')
        elif date_obj and '.' in date_obj:
            return date_obj.strip()
        else:
            return ''

    def get_route(self):
        """Функция-генератор данных о рейсах из словаря self.routes_dict."""
        for date in self.routes_dict:
            for route in self.routes_dict[date]:
                for bus_numb in self.routes_dict[date][route]:
                    for position, dict in enumerate(self.routes_dict[date][route][bus_numb]):
                        yield position, dict, date

    def get_screen_value(self, value):
        if type(value) == int:
            return str(value)
        if not value:
            return ''
        return str(value)


class Reader:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.final_report_path = None
        self.dict_problems = defaultdict(lambda: defaultdict(list))
        self.wb = None
        self.date = None
        self.total = 0

    def read(self):
        self.wb = openpyxl.load_workbook(self.file_path)
        sheet = self.wb.active
        counter = 0
        for row in sheet.rows:
            counter += 1
            if row[0].value != 'Дата':
                if self.date is None and row[0].value:
                    self.convert_str_to_date(row[0].value)
                route = row[1].value
                if row[5].value:
                    self.total += 1
                    bus_numb = row[5].value
                    self.dict_problems[route][bus_numb].append(
                        {
                            'date': f'{row[0].value.day:>02}.{row[0].value.month:>02}.{row[0].value.year}'
                            if type(row[0].value) == datetime.datetime else row[0].value,
                            'direction': row[2].value,
                            'plan': self.convert_str_to_time(row[3].value),
                            'fact': self.convert_str_to_time(row[4].value),
                            'problem': str(row[7].value).strip() if row[7].value is not None else '',
                            'screen': row[8].value.strip() if row[8].value is not None else "",
                            'row_numb': counter,
                            'queue': "",
                            'bus_numb': row[5].value,
                            'route': row[1].value,
                        }
                    )

        for route in self.dict_problems:
            for lst_flights in self.dict_problems[route].values():
                lst_flights.sort(key=lambda item: self.get_seconds(item['plan']))

    @staticmethod
    def get_seconds(st):
        if st:
            hours, minutes = map(int, st.split(':'))
            res = hours * 60 + minutes
            return res
        return 0

    def convert_str_to_time(self, st: datetime.time):
        if st is None or st == '':
            return ''
        if type(st) == datetime.time or type(st) == datetime.datetime:
            return st.strftime('%H:%M')
        try:
            return datetime.time(*[int(val.strip()) for val in st.split(':')]).strftime('%H:%M')
        except Exception:
            return ''

    def convert_str_to_date(self, st):
        if type(st) == datetime.datetime:
            self.date = st.strftime('%d.%m.%Y')
        elif st.strip() and '.' in st:
            self.date = st.strip()

    def read_final_reports(self): # надо переделать под разные даты
        wb = openpyxl.load_workbook(self.final_report_path, read_only=True)
        sheet = wb.active
        cur_bus = None
        search_res = False
        for row in sheet.values:
            if type(row[0]) == datetime.datetime and row[0].strftime('%d.%m.%Y') == self.date:
                if not search_res:
                    search_res = True
                route = self.convert_route_into_numb(row[2].replace('-КРС', ''))
                bus_numb = row[1]
                queue_ = str(row[10])
                if route in self.dict_problems and bus_numb in self.dict_problems[route] and bus_numb != cur_bus:
                    for flight in self.dict_problems[route][bus_numb]:
                        flight['queue'] = queue_
                    cur_bus = bus_numb
            if search_res and row[0].strftime('%d.%m.%Y') != self.date:
                break

    def convert_route_into_numb(self, st):
        try:
            return int(st)
        except Exception:
            return st

    def get_route(self):
        for route in self.dict_problems:
            for bus_numb in self.dict_problems[route]:
                for position, dict in enumerate(self.dict_problems[route][bus_numb]):
                    yield position, dict