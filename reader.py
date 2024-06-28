from collections import defaultdict
import openpyxl
import datetime


class Reader:
    def __init__(self, app, file_path=None):
        self.app = app
        self.file_path = file_path
        self.final_report_path = None
        self.dict_problems = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        self.wb = None
        self.dates = []
        self.total = 0

    def read(self):
        self.wb = openpyxl.load_workbook(self.file_path)
        sheet = self.wb.active
        counter = 0
        for row in sheet.rows:
            counter += 1
            if row[0].value != 'Дата':
                route_date = str(row[0].value).strip() if row[0].value is not None else None

                if route_date not in self.dates:
                    self.dates.append(route_date)

                route = row[1].value
                self.total += 1
                bus_numb = row[5].value
                self.dict_problems[route_date][route][bus_numb].append(
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
                        'bus_numb': row[5].value if row[5].value is not None else "",
                        'route': row[1].value,
                    }
                    )

        for route_date in self.dict_problems:
            for route in self.dict_problems[route_date]:
                for lst_flights in self.dict_problems[route_date][route].values():
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

    def convert_date_to_str(self, st):
        if type(st) in (datetime.datetime, datetime.date):
            return st.strftime('%d.%m.%Y')
        elif type(st) == int:
            return
        elif st.strip() and '.' in st:
            return st.strip()

    def read_final_reports(self):
        wb = openpyxl.load_workbook(self.final_report_path, read_only=True)
        sheet = wb.active
        cur_bus = None
        for row in sheet.values:
            if type(row[0]) == datetime.datetime and row[0].strftime('%d.%m.%Y') in self.dates:
                route_date = row[0].strftime('%d.%m.%Y')
                route = self.convert_route_into_numb(row[2].replace('-КРС', ''))
                bus_numb = row[1]
                queue_ = str(row[10])
                if (route in self.dict_problems[route_date] and bus_numb in self.dict_problems[route_date][route]
                        and bus_numb != cur_bus):
                    for flight in self.dict_problems[route_date][route][bus_numb]:
                        flight['queue'] = queue_
                    cur_bus = bus_numb

    def convert_route_into_numb(self, st):
        try:
            return int(st)
        except Exception:
            return st

    def get_route(self):
        """Функция-генератор данных о рейсах из словаря self.dict_broblems."""
        for date in self.dict_problems:
            for route in self.dict_problems[date]:
                for bus_numb in self.dict_problems[date][route]:
                    for position, dict in enumerate(self.dict_problems[date][route][bus_numb]):
                        yield position, dict
