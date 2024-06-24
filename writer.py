import datetime
import json
import os

class Writer:
    """Класс записывает данные о проделанной работе в файл эксель."""
    def __init__(self, reader):
        self.reader = reader
        self.screen_paths = {}

    def write(self):
        """Запускает запись данных в обьект self.reader.wb."""
        for i in range(self.reader.app.table.table_size):
            item_id = str(i)
            values = self.reader.app.table.item(item_id)['values']
            row_numb = values[10]

            self.reader.wb.active[f'B{row_numb}'] = values[8]
            self.reader.wb.active[f'C{row_numb}'] = values[1]
            self.reader.wb.active[f'D{row_numb}'] = values[2]
            self.reader.wb.active[f'E{row_numb}'] = values[3]
            self.reader.wb.active[f'F{row_numb}'] = values[4]
            self.reader.wb.active[f'H{row_numb}'] = values[5]
            self.reader.wb.active[f'I{row_numb}'] = values[6]
        self.reader.wb.save(self.reader.file_path)

        self.save()

    @staticmethod
    def convert_into_datetime(st_datetime: str, format=None):
        """Возвращает обьект datetime

        Аргументы:
            st_datetime(str): строка с датой,
            format(str): формат даты
        """
        try:
            return datetime.datetime.strptime(st_datetime.strip(), format)
        except ValueError:
            return st_datetime

    def save(self):
        if self.rd.report_type == 'НС':
            try:
               self.write_to_json(self.screen_paths)
            except Exception:
                pass

        self.rd.wb.save(self.rd.file_path)


