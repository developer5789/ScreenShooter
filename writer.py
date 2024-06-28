import datetime


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
            problem = self.convert_into_datetime(values[6])

            self.reader.wb.active[f'B{row_numb}'] = values[13]
            self.reader.wb.active[f'C{row_numb}'] = values[2]
            self.reader.wb.active[f'D{row_numb}'] = values[3]
            self.reader.wb.active[f'E{row_numb}'] = values[4]
            self.reader.wb.active[f'F{row_numb}'] = values[5]
            self.reader.wb.active[f'H{row_numb}'] = problem
            self.reader.wb.active[f'I{row_numb}'] = values[7]
        self.reader.wb.save(self.reader.file_path)


    @staticmethod
    def convert_into_datetime(st):
        """Возвращает обьект datetime

        Аргументы:
            st_datetime(str): строка с датой,
            format(str): формат даты
        """
        if type(st) == str and ":" in st:
            try:
                return datetime.datetime.strptime(st.strip(), "%Y-%m-%d %H:%M:%S")
            except ValueError:
                return st
        else:
            return st



