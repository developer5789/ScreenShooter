from collections import defaultdict
import openpyxl
import datetime

class Reader:
    """Класс читает эксель с рейсами и заполняет словарь self.routes_dict"""
    def __init__(self, app,  file_path=None):
        """Инициализация аттрибутов

        Аргументы:
            file_path(str, None): путь до файла с рейсами,
            app(App): главное окно приложения.
        """
        self.app = app
        self.file_path = file_path
        self.routes_dict = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        self.wb = None
        self.total = 0
        self.report_type = None

    def read(self):
        """Запускает чтение файла эксель с рейсами."""
        self.wb = openpyxl.load_workbook(self.file_path)
        self.report_type = self.app.load_window.combobox_value.get()
        self.app.load_window.progress_text_var.set(f'Загрузка файла...')
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
            if type(row[2].value) == datetime.datetime and row[18].value is None:
                self.total += 1
                date = self.convert_date_to_str(row[2].value)
                route = self.get_int(row[1].value)
                bus_numb = self.get_bus_numb(row[11].value)
                self.routes_dict[date][route][bus_numb].append(
                    {
                        'direction': self.get_int(row[3].value),
                        'start_time': self.convert_time_to_str(row[7].value),
                        'finish_time': self.convert_time_to_str(row[8].value),
                        'screen': None,
                        'row_numb': counter,
                        'bus_numb': bus_numb,
                        'route': route,
                        'route_numb': row[0].value
                    }
                )

        self.app.load_window.progress_text_var.set(f'Файл {self.file_path} прочитан')
        self.app.load_window.progress_var.set(100)
        self.app.load_window.ok_btn['state'] = 'normal'

    @staticmethod
    def get_int(value):
        try:
            return int(value)
        except Exception:
            return value

    @staticmethod
    def get_bus_numb(bus_numb: str):
        """Возращает значение гос.номера ТС с пробелами."""
        if not bus_numb:
            return ''
        bus_numb = bus_numb.replace(' ', '')
        bus_numb = f'{bus_numb[:2]} {bus_numb[2:-2]} {bus_numb[-2:]}'
        return bus_numb

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
