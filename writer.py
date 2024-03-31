import datetime

class Writer:
    """Класс записывает данные о проделанной работе в файл эксель."""
    def __init__(self, reader):
        self.rd = reader

    def write(self):
        """Запускает запись данных в обьект self.reader.wb."""
        for i in range(self.rd.app.table.table_size):
            item_id = str(i)
            if item_id in self.rd.app.table.edited_items:
                values = self.rd.app.table.item(item_id)['values']
                row_numb = values[10]
                start_time = self.convert_into_datetime(values[3], '%H.%M.%S')
                self.rd.wb.active[f'B{row_numb}'] = values[1]
                self.rd.wb.active[f'C{row_numb}'] = self.convert_into_datetime(values[0], '%d.%m.%Y')
                self.rd.wb.active[f'D{row_numb}'] = values[2]
                self.rd.wb.active[f'E{row_numb}'] = start_time
                self.rd.wb.active[f'F{row_numb}'] = start_time
                self.rd.wb.active[f'H{row_numb}'] = start_time
                self.rd.wb.active[f'I{row_numb}'] = self.convert_into_datetime(values[4], '%H.%M.%S')
                self.rd.wb.active[f'L{row_numb}'] = values[5]
                self.rd.wb.active[f'S{row_numb}'] = values[6]

        self.rd.wb.save(self.rd.file_path)

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
