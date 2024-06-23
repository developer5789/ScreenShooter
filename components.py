import time
from threading import Thread
import pyautogui as pg
from collections import defaultdict
import datetime
from datetime import timedelta
import os
import tkinter as tk
from tkinter import PhotoImage as Img
from tkinter import ttk
from tkinter import filedialog as fd
from myclasses import MyCheckboxTreeview, TableEntry
import json
from selenium.common.exceptions import NoSuchWindowException
from loger import Loger
from messages import show_inf, show_error, askyesnocancel, last_item_message
from image_editor import ImageEditor


def activate_buttons(*btns):
    """Переводит кнопки в активное состояние."""
    for btn in btns:
        btn['state'] = 'normal'


def block_buttons(*btns):
    """Делает кнопки неактивными."""
    for btn in btns:
        btn['state'] = 'disabled'


config = {} # нужен ли?
screen_paths = defaultdict(lambda: defaultdict(list))


class ConfigWindow(tk.Toplevel):
    """Класс описывает окно пользовательских настроек"""

    def __init__(self, *args, **kwargs):
        """Инициализация элементов внутри окна и аттрибутов"""
        super().__init__(*args, **kwargs)
        self.geometry("400x200")
        self.title('Настройки')
        self.resizable(False, False)
        self.grab_set()
        self.path_frame = ttk.Frame(self)
        self.path_label = ttk.Label(self.path_frame, text="Путь до папки профиля Chrome:")
        self.path_entry = tk.Entry(self.path_frame, justify='left', width=60)
        self.cords_frame = ttk.Frame(self, padding=3)
        self.cords_label = tk.Label(self.cords_frame, text="Область скриншота:")
        self.cords_entry = tk.Entry(self.cords_frame, width=35)
        self.timeout_frame = ttk.Frame(self, padding=3)
        self.timeout_label = tk.Label(self.timeout_frame, text="Timeout (сек): ")
        self.timeout_var = tk.IntVar()
        self.timeout_spinbox = tk.Spinbox(self.timeout_frame, state='readonly',
                                          to=10.0, textvariable=self.timeout_var)
        self.btn_frame = ttk.Frame(self, padding=2)
        self.btn_set = ttk.Button(self.btn_frame, text='Применить', command=self.apply)
        self.pack_items()
        self.protocol('WM_DELETE_WINDOW', self.destroy)
        self.values = []
        self.fill_out_fields()

    def pack_items(self):
        """Размещает элементы внутри окна."""
        self.path_frame.grid(row=0, column=0, padx=20)
        self.cords_frame.grid(row=1, column=0, padx=20, sticky='nswe')
        self.btn_frame.grid(row=3, column=0, pady=5, sticky='nswe', )
        self.timeout_frame.grid(row=2, column=0, padx=20, sticky='nswe')

        self.path_label.pack(anchor='w', pady=3)
        self.path_entry.pack()
        self.cords_label.pack(anchor='w', pady=3)
        self.cords_entry.pack(anchor='w')
        self.timeout_label.pack(anchor='w')
        self.timeout_spinbox.pack(anchor='w')
        self.btn_set.pack(side=tk.RIGHT, padx=15)

    def fill_out_fields(self):
        """Заполняет поля значениями из словаря 'config'"""
        profile_path = config['profile_path']
        screen_cords = str(config['screen_cords']).strip('[]')
        timeout_val = config['timeout']
        self.values.extend((profile_path, screen_cords, timeout_val))
        self.path_entry.insert('end', profile_path)
        self.cords_entry.insert('end', screen_cords)
        self.timeout_var.set(timeout_val)

    def apply(self):
        """Сравнивает новые настройки со старыми и устанавливает новые,если отличаются."""
        new_values = [self.path_entry.get(),
                      [int(val.strip()) for val in self.cords_entry.get().split(',')],
                      self.timeout_var.get()
                      ]
        if self.values != new_values:
            self.set_settings(new_values)
        self.destroy()

    def set_settings(self, new_values: list):
        """Записывает новые значения пользовательских настроек """
        global x, y, width, height, timeout, profile_path
        for i, key in enumerate(config):
            config[key] = new_values[i]
        timeout = config["timeout"]
        x, y, width, height = config["screen_cords"]
        profile_path = config["profile_path"]
        self.save_settings()

    @staticmethod
    def save_settings():
        """Сохраняет настройки, записывая в файл 'config.json'"""
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)


class FilterWindow(tk.Toplevel):
    """Класс описывает окно фильтрации по столбцам таблицы."""

    def __init__(self, app, colname):
        """Инициализация элементов окна,аттрибутов, настройка параметров

        """
        super().__init__()
        self.app = app
        self.colname = colname
        self.geometry("295x320")
        self.title(f'Фильтрация')
        self.resizable(False, False)
        self.grab_set()
        self.geometry("+{}+{}".format(self.app.winfo_rootx() + 50, self.app.winfo_rooty() + 50))
        self.btn_frame = ttk.Frame(self)
        self.table_frame = ttk.Frame(self)
        self.entry_value = tk.StringVar()
        self.entry_value.trace_add("write", self.update_table)
        self.search_entry = ttk.Entry(self, justify='left', width=29, font=('Arial', 9), textvariable=self.entry_value)
        self.table = MyCheckboxTreeview(self, self.table_frame, show='tree', height=8)
        self.scroll = ttk.Scrollbar(self.table_frame, command=self.table.yview)
        self.table.config(yscrollcommand=self.scroll.set)
        self.ok_btn = ttk.Button(self.btn_frame, text='Ок', width=5, command=self.apply)
        self.cancel_btn = ttk.Button(self.btn_frame, text='Отмена', command=self.destroy)
        self.img = Img(file=r'icons/filter_remove.png')
        self.remove_filter_btn = ttk.Button(self, image=self.img,
                                            compound='image', takefocus=False, command=self.delete_filter)
        self.pack_items()

    def delete_filter(self):
        filter_vals = self.app.table.filters[self.colname]
        if filter_vals:
            self.app.table.del_filter(self.colname)
            self.table.overwrite_table()

    def apply(self):
        filtered_values = self.table.get_checked_val()
        previous_vals = self.table.values
        if filtered_values != previous_vals:
            self.set_filter_icon()
            self.app.table.filters[self.colname] = filtered_values
            self.app.table.filter_items()

        self.destroy()

    def update_table(self, *args):
        value = self.entry_value.get()
        if value:
            self.table.filter(value)
        else:
            self.table.return_initial_state()
        self.table.set_btn_state()

    def pack_items(self):
        """Размещает элементы внутри окна."""
        self.search_entry.grid(row=0, column=0, sticky='nswe', pady=10, padx=10)
        self.table.grid(row=0, column=0)
        self.scroll.grid(row=0, column=1, sticky='ns')
        self.ok_btn.grid(row=0, column=0, padx=15)
        self.cancel_btn.grid(row=0, column=1, padx=10)
        self.btn_frame.grid(row=2, column=0, pady=10, columnspan=2, sticky='nswe')
        self.remove_filter_btn.grid(row=0, column=1)
        self.table_frame.grid(row=1, column=0, columnspan=2, pady=10)

    def set_filter_icon(self):
        if self.app.table.filters[self.colname] is None:
            filter_icon = self.app.table.icons['filter']
            self.app.table.heading(self.colname, image=filter_icon) # н № #


class Table(ttk.Treeview):
    """Класс описывает таблицу с информацией о рейсах"""

    columns = ("date", 'route', "direction", "start", 'finish', 'bus_numb',
               'screen', 'screen_path', 'position', 'edited', 'row_id', 'colour', 'route_id', 'route_dir')

    def __init__(self, app, **kwargs):
        """Инициализация аттрибутов, колонок таблицы, создание стилей."""
        super().__init__(**kwargs)
        self.app = app
        self.icons = {
            'direction': Img(file=r'icons/icons8-направление-22.png'),
            'date': Img(file=r'icons/icons8-расписание-22.png'),
            'start': Img(file=r'icons/icons8-начало-22.png'),
            'finish': Img(file=r'icons/icons8-end-function-button-on-computer-keybord-layout-22.png'),
            'bus_numb': Img(file=r'icons/icons8-автобус-22.png'),
            'screen': Img(file=r'icons/icons8-скриншот-22.png'),
            'route': Img(file=r'icons/icons8-автобусный-маршрут-20.png'),
            'filter': Img(file=r'icons/icons8-фильтр-22.png')

        }
        self.heading('date', text='Дата', anchor='w', image=self.icons['date'],
                     command=lambda: FilterWindow(self.app, 'date'))
        self.heading('route', text='Маршрут', anchor='w', image=self.icons['route'],
                     command=lambda: FilterWindow(self.app, 'route'))
        self.heading('direction', text='Направление', anchor='w', image=self.icons['direction'],
                     command=lambda: FilterWindow(self.app, 'direction'))
        self.heading("start", text="Начало", anchor='w', image=self.icons['start'],
                     command=lambda: FilterWindow(self.app, 'start'))
        self.heading("finish", text="Конец", anchor='w', image=self.icons['finish'],
                     command=lambda: FilterWindow(self.app, 'finish'))
        self.heading("bus_numb", text="Гос.номер", anchor='w', image=self.icons['bus_numb'],
                     command=lambda: FilterWindow(self.app, 'bus_numb'))
        self.heading('screen', text="Скрин", anchor='w', image=self.icons['screen'],
                     command=lambda: FilterWindow(self.app, 'screen'))
        self.column("date", width=140, stretch=True)
        self.column("route", width=130, stretch=True)
        self.column("direction", stretch=True, width=155)
        self.column("start", stretch=True, width=80)
        self.column("finish", stretch=True, width=90)
        self.column("bus_numb", stretch=True, width=115)
        self.column("screen", stretch=True, width=90)
        self.tag_configure('bold', font=('Arial', 13, 'bold'))
        self.tag_configure('gray_colored', background='#D3D3D3')
        self.tag_configure('green_colored', background='#98FB98')
        self.tag_configure('white_colored', background='white')
        self.tag_configure('orange_colored', background='#F0E68C')
        self.tag_configure('red_colored', background='#FFA07A')
        self.style = ttk.Style()
        self.style.configure('Treeview', font=('Arial', 13), rowheight=60, separator=100)
        self.style.map('Treeview', background=[('selected', '#D3D3D3')], foreground=[('selected', 'black')])
        self.heading_style = ttk.Style()
        self.heading_style.configure('Treeview.Heading', font=('Arial', 12))
        self.bind('<<TreeviewSelect>>', self.select_item)
        self.bind('<Double-Button-1>', self.click_cell)
        self.bind()
        self.table_size = 0
        self.current_item = '0'
        self.editing_cell = None
        self.filters = {col: None for col in self['displaycolumns']}
        self.focused_route = None
        self.empty_val = 0


    def click_cell(self, event):
        col, selected_item = self.identify_column(event.x), self.focus()
        text = self.item(selected_item)['values'][int(col[1:]) - 1]
        box = list(self.bbox(selected_item, col))
        box[0], box[1] = box[0] + event.widget.winfo_x(), box[1] + event.widget.winfo_y()
        self.editing_cell = TableEntry(selected_item, col, box, self.app,
                                       font=('Arial', 13), background='#98FB98')
        self.editing_cell.insert(0, text)

    def del_filter(self, colname):
        self.filters[colname] = None
        self.filter_items()
        self.heading(colname, image=self.app.table.icons[colname])

    def match(self, filtered_cols, row_vals):
        for col in filtered_cols:
            col_indx = self['displaycolumns'].index(col)
            col_mark_vals = self.filters[col]
            value = str(row_vals[col_indx])
            if value not in col_mark_vals:
                return False
        return True

    def filter_items(self):
        filtered_cols = [col for col in self.filters if self.filters[col]]
        for i in range(self.table_size):
            item_id = str(i)
            values = self.item(item_id)['values']
            if self.match(filtered_cols, values):
                self.move(item_id, '', i)
            else:
                self.detach(item_id)

    def select_item(self, event):
        self.current_item = self.selection()[0]
        self.check()

    def next_item(self):
        next_item = self.next(self.current_item)
        if next_item:
            self.selection_set((next_item,))
            self.yview_scroll(1, 'units')
        return next_item

    def del_command(self):
        values = self.item(self.current_item)['values']
        screen_path, root_dir = values[7], values[13]
        if screen_path:
            ans = self.askdel(screen_path, root_dir, self.app, self.current_item)
            self.next_item() if ans is not None else None

    def askdel(self, screen_path, root_dir, parent, item):
        ans = askyesnocancel(parent)
        if ans:
            self.del_screen(screen_path, root_dir, item)
            self.set(item, 6, 0)
            self.color('red_colored', item)
        elif ans is not None:
            self.del_screen(screen_path, root_dir, item)
            self.set(self.current_item, 6, '')
            self.color('white_colored', item)
        return ans

    def del_screen(self, screen_path: str, root_dir: str, item: str):
        """Удаление скриншота

        Аргументы:
            screen_path(str): путь к скрину
            root_dir(str): путь к директории скрина
        """
        if self.app.rd.report_type == 'НС':
            screen_name = screen_path.split('\\')[-1]
            numb, indx = screen_name[:-4].split(' - ')
            numb_dict = screen_paths[root_dir][numb]

            if int(indx) < numb_dict['max_value']:
                numb_dict['empty_positions'].append(indx)
            else:
                numb_dict['max_value'] -= 1

        os.remove(screen_path)
        self.set(item, 7, '')

    def execute_command(self, values=None, action=None, switch=True): #надо посмотреть
        """Выполняет переданную команду

        Аргументы:
            values(list, None): список значений ячеек строки,
            action(str, None): номер исполняемой команды (1-скрин)
            """
        if values is None:
            values = self.item(self.current_item)['values']
        screen = str(values[6])
        if action:
            screen_path = self.make_screenshot(values)
            self.set(self.current_item, 6, action)
            self.set(self.current_item, 7, screen_path)
            self.color('green_colored', self.current_item)

        else:
            self.del_screen(values[7], values[13], self.current_item) if values[7] else None
            self.set(self.current_item, 6, action)
            self.color('red_colored', self.current_item)

        if switch:
            self.next_item()
        if not screen:
            self.app.res_panel.add_route()

    def check(self):
        """Проверка на наличие скрина."""
        screen_path = self.item(self.current_item)['values'][7]
        if screen_path:
            self.app.btn_panel.show_btn['state'] = 'normal'
        else:
            self.app.btn_panel.show_btn['state'] = 'disabled'

    def show_screen(self):
        """Открывает скрин текущей строки таблицы."""
        screen_path = self.item(self.current_item)['values'][7]
        os.startfile(screen_path)

    def color(self, colour, item):
        """Заливка после выполнения команды"""

        self.item(item, tags=(colour,))


    def fill_out_table(self, rd):
        """Заполняет таблицу значениями из прочитанного документа excel

         Аргументы:
            rd(Reader): обьект класса Reader.
        """
        counter = -1
        for position, flight, date in rd.get_route():
            counter += 1
            item = str(counter)
            route, screen, route_numb = flight['route'], flight['screen'], flight['route_numb']
            root_dir = self.get_root_dir(date, route)
            values = (date, route, flight['direction'], flight['start_time'],
                      flight['finish_time'], flight['bus_numb'], screen, '', position, 0,
                      flight['row_numb'], 'white_colored', route_numb, root_dir)
            self.insert('', 'end', values=values, iid=item)

            if rd.report_type == 'Комиссия':
                self.find_screen(item, screen, root_dir, route_numb)
            else:
                self.find_screen_in_json(rd, values, item)

        self.app.res_panel.set_progress(rd.total, rd.total - self.empty_val, self.empty_val)

        self.table_size = counter + 1
        if self.table_size:
            self.selection_set(('0',))
        else:
            block_buttons(*self.app.btn_panel.buttons)
        rd.routes_dict.clear()

    def get_values(self):
        """Возвращает список значений ячеек текущей строки"""
        return self.item(self.current_item)['values']

    def find_screen(self, item, screen, root_dir, route_numb):

        if screen == '1':
            screen_path = rf'{root_dir}\{route_numb}.jpg'
            if os.path.exists(screen_path):
                self.set(item, 7, screen_path)
                self.item(item, tags=('green_colored',))
            else:
                self.item(item, tags=('orange_colored',))

        elif screen == '0':
            self.item(item, tags=('red_colored',))

        else:
            self.empty_val += 1

    def find_screen_in_json(self, rd, values, item):
        route_id = ';'.join(list(map(str, values[:6])))
        screen = values[6]
        if route_id in rd.paths_json:
            screen_path = rd.paths_json[route_id]
            exist_path = os.path.exists(screen_path)

            if screen == '1' and exist_path:
                self.set(item, 7, screen_path)
                self.item(item, tags=('green_colored',))
            elif screen == '1' and not exist_path:
                self.item(item, tags=('orange_colored',))
            elif screen == '0':
                self.item(item, tags=('red_colored',))
            else:
                self.empty_val += 1

        else:
            if screen == '1':
                self.item(item, tags=('orange_colored',))
            elif screen == '0':
                self.item(item, tags=('red_colored',))
            else:
                self.empty_val += 1

    def make_screenshot(self, values: list):
        """Делает скрин трека

        Аргументы:
            values(list): список значений ячеек строки,
            overwrite(bool): параметр, показывающий перезаписывается ли скрин.
        """
        screen_path = values[7]

        if not screen_path:
            screen_path = self.get_screen_path(values)

        screen = pg.screenshot(region=(x, y, width, height))
        screen.save(screen_path)

        return screen_path
    def get_screen_path(self, values: list):
        """Возвращает путь до сделанного скриншота

        Аргументы:
            values(list): список значений ячеек строки
        """
        root_dir = values[13]
        if not os.path.exists(root_dir):
            os.makedirs(root_dir)

        if self.app.rd.report_type == 'НС':
            bus_numb = values[5]
            indx = self.get_screen_indx(bus_numb, root_dir)
            return root_dir + '\\' + f'{bus_numb} - {indx}.jpg'
        else:
            route_numb = values[12]
            screen_path = root_dir + '\\' + f'{route_numb}.jpg'

        return screen_path

    @staticmethod
    def get_screen_indx(bus_numb: str, root_dir: str):
        """Возвращает порядковый номер скриншота. Используется при разборе
         нештатных ситуациий т.к в именовании должен присутствовать порядковый номер скрина.

         Аргументы:
            bus_numb(str): гос.номер автобуса,
            root_dir(str): путь до директории скрина.
        """
        if screen_paths[root_dir][bus_numb]:
            empty_spaces: list = screen_paths[root_dir][bus_numb]['empty_positions']
            if empty_spaces:
                indx = empty_spaces.pop()
            else:
                screen_paths[root_dir][bus_numb]['max_value'] += 1
                indx = screen_paths[root_dir][bus_numb]['max_value']
        else:
            screen_paths[root_dir][bus_numb] = {
                'empty_positions': [],
                'max_value': 1
            }
            indx = screen_paths[root_dir][bus_numb]['max_value']
        return indx

    def get_root_dir(self, date: str, route: (str, int)):
        """Возвращает путь до директории скрина

        Аргументы:
            date(str): дата выпуска ТС,
            route(str, int): код маршрута,
        """
        month_numb = int(date.split('.')[1])
        path = fr'скрины\{self.app.rd.report_type}\{months[month_numb]}_{route}'
        return path

    @staticmethod
    def get_datetime_str(date_st, time_st, start=True):
        """Делает смещение времени начала и конца рейса на 2 минуты.
          Возвращает строку со временем в формате 'Год-Месяц-День Часы:Минуты

           Аргументы:
                date_st(str): строка с датой,
                time_st(str): строка со временем,
                start(bool): начало или конец рейса.
        """
        datetime_st = f'{date_st} {time_st}'
        datetime_obj = datetime.datetime.strptime(datetime_st, '%d.%m.%Y %H:%M')
        if datetime_obj.hour < 2:
            datetime_obj += datetime.timedelta(hours=24)

        if start:
            datetime_obj -= datetime.timedelta(minutes=2)
        else:
            datetime_obj += datetime.timedelta(minutes=2)

        return datetime_obj.strftime('%Y-%m-%d %H:%M')

    def get_paths(self):
        root_path = rf'скрины\{self.app.rd.report_type}'
        for dir in os.listdir(root_path):
            path = root_path + '\\' + dir
            numb_dict = defaultdict(list)
            for screen_name in os.listdir(path):
                numb, indx = screen_name[:-4].split(' - ')
                numb_dict[numb].append(int(indx))
            for numb in numb_dict:
                lst_indx = numb_dict[numb]
                max_indx = max(lst_indx)
                empty_pos = []
                for i in range(1, max_indx + 1):
                    if i not in lst_indx:
                        empty_pos.append(i)
                numb_dict[numb] = {
                    'empty_positions': empty_pos,
                    'max_value': max_indx
                }
            path = root_path + '\\' + dir
            screen_paths[path] = numb_dict

class LoadWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.geometry("850x250")
        self.title('ScreenShooter')
        self.resizable(False, False)
        self.grab_set()
        self.report_name = None
        self.final_report_name = None
        self.btn_upload = ttk.Button(self, text='Загрузить', state='disabled', command=self.close)
        self.frame_1 = None
        self.frame_2 = None
        # self.frame_3 = ttk.Frame(self, padding=3)

    def select_file(self, report_type=None):
        if report_type == 'unaccounted_flights':
            self.report_name = fd.askopenfilename()
            self.text_1.insert(1.0, self.report_name)
            self.check()
        else:
            self.final_report_name = fd.askopenfilename()
            self.text_2.insert(1.0, self.final_report_name)
            self.check()

    def check(self):
        if self.final_report_name and self.report_name:
            self.btn_upload['state'] = 'normal'

    def close(self):
        self.app.rd.file_path = self.report_name
        self.app.rd.final_report_path = self.final_report_name
        Thread(target=self.app.run_reader).start()

    def set(self):
        self.frame_1 = ttk.Frame(self, padding=3)
        self.frame_2 = ttk.Frame(self, padding=3)
        btn_1 = ttk.Button(self.frame_1, text='Найти', command=lambda: self.select_file('unaccounted_flights'))
        btn_2 = ttk.Button(self.frame_2, text='Найти', command=lambda: self.select_file())
        label_1 = ttk.Label(self.frame_1, text='Файл с неучтёнными рейсами:', justify='left')
        label_2 = ttk.Label(self.frame_2, text='Файл со всеми рейсами:', justify='left')
        self.text_1 = tk.Text(self.frame_1, height=1, width=65, background='white', font=('Arial', 11))
        self.text_2 = tk.Text(self.frame_2, height=1, width=65, background='white', font=('Arial', 11))
        label_1.grid(row=0, column=0, columnspan=2, sticky='we')
        self.text_1.grid(row=1, column=0)
        btn_1.grid(row=1, column=1)
        label_2.grid(row=0, column=0, columnspan=2, sticky='we')
        self.text_2.grid(row=1, column=0)
        btn_2.grid(row=1, column=4)
        self.frame_1.pack()
        self.frame_2.pack()
        self.btn_upload.pack(side='right', padx=10)


    def end(self):
        self.destroy()
        self.app.table.fill_out_table(self.app.rd)

class ResultPanel:
    """Класс описывает панель результатов, которая включает в себя
    счётчики скорости, прогнозируемого времени на работу, отработанного времени,
    общего кол-ва рейсов, разобранных рейсов, оставшихся рейсов.
    """

    def __init__(self, root: tk.Tk):
        """Инициализация элементов панели и аттрибутов."""
        self.root = root
        self.main_frame = ttk.Frame(root, relief='groove')
        self.progress_var = tk.IntVar()
        self.progressbar = ttk.Progressbar(orient="horizontal", length=700, variable=self.progress_var)
        self.speed_counter = tk.StringVar(value='0')
        self.routes_counter = tk.IntVar()
        self.completed_counter = tk.IntVar()
        self.remaining_counter = tk.IntVar()
        self.worked_hours_counter = tk.StringVar(value='00:00:00')
        self.predict_counter = tk.StringVar(value='00:00:00')
        self.time_counter = 0
        self.state = 0
        self.last_action = 0

    def prepare_panel(self):
        """Размещает элементы на панели."""
        speed_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)
        predict_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)
        routes_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)
        completed_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)
        remaining_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)
        worked_hours_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)

        speed_label = ttk.Label(speed_frame, text="Cкорость(cкр/мин):", font=("Arial", 12))
        predict_label = ttk.Label(predict_frame, text="Прогноз:", font=("Arial", 12))
        routes_label = ttk.Label(routes_frame, text="Всего:", font=("Arial", 12))
        completed_label = ttk.Label(completed_frame, text="Сделано:", font=("Arial", 12))
        remaining_label = ttk.Label(remaining_frame, text="Осталось:", font=("Arial", 12))
        worked_hours_label = ttk.Label(worked_hours_frame, text="Отработано:", font=("Arial", 12))

        speed_label.pack(side=tk.LEFT)
        predict_label.pack(side=tk.LEFT)
        routes_label.pack(side=tk.LEFT)
        completed_label.pack(side=tk.LEFT)
        remaining_label.pack(side=tk.LEFT)
        remaining_label.pack(side=tk.LEFT)
        worked_hours_label.pack(side=tk.LEFT)

        speed_display = ttk.Label(speed_frame, textvariable=self.speed_counter, font=("Arial", 18))
        flights_display = ttk.Label(routes_frame, textvariable=self.routes_counter, font=("Arial", 18))
        completed_display = ttk.Label(completed_frame, textvariable=self.completed_counter, font=("Arial", 18))
        remaining_display = ttk.Label(remaining_frame, textvariable=self.remaining_counter, font=("Arial", 18))
        worked_hours_display = ttk.Label(worked_hours_frame, textvariable=self.worked_hours_counter, font=("Arial", 18))
        predict_display = ttk.Label(predict_frame, textvariable=self.predict_counter, font=("Arial", 18))

        speed_display.pack(side=tk.RIGHT)
        flights_display.pack(side=tk.RIGHT)
        completed_display.pack(side=tk.RIGHT)
        remaining_display.pack(side=tk.RIGHT)
        predict_display.pack(side=tk.RIGHT)
        worked_hours_display.pack(side=tk.RIGHT)

        speed_frame.grid(row=0, column=0, padx=3, sticky='NSEW')
        worked_hours_frame.grid(row=0, column=1, padx=3, sticky='NSEW')
        predict_frame.grid(row=0, column=2, padx=3, sticky='NSEW')
        routes_frame.grid(row=0, column=3, padx=3, sticky='NSEW')
        completed_frame.grid(row=0, column=4, padx=3, sticky='NSEW')
        remaining_frame.grid(row=0, column=5, padx=3, sticky='NSEW')

        self.main_frame.columnconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.columnconfigure(2, weight=1)
        self.main_frame.columnconfigure(3, weight=1)
        self.main_frame.columnconfigure(4, weight=1)
        self.main_frame.columnconfigure(5, weight=1)

        self.main_frame.grid(row=0, column=0, sticky='nsew')
        self.main_frame.rowconfigure(1, weight=1)

    def set_progress(self, total, completed_items, remaining_items):
        self.routes_counter.set(total)
        self.remaining_counter.set(remaining_items)
        self.completed_counter.set(completed_items)
        progress_value = 100 * self.completed_counter.get() / self.routes_counter.get()
        self.progress_var.set(int(progress_value))

    def add_route(self):
        """Увеличивает кол-во разобранных рейсов на единицу."""
        self.completed_counter.set(self.completed_counter.get() + 1)
        self.remaining_counter.set(self.routes_counter.get() - self.completed_counter.get())
        speed = self.calc_speed()
        predict_value = 60 * self.remaining_counter.get() / speed
        self.predict_counter.set(str(timedelta(seconds=int(predict_value))))
        progress_value = 100 * self.completed_counter.get() / self.routes_counter.get()
        self.progress_var.set(int(progress_value))
        if not self.remaining_counter.get():
            show_inf()

    def subtract_route(self):
        """Уменьшает кол-во разобранных рейсов на единицу."""
        self.completed_counter.set(self.completed_counter.get() - 1)
        self.remaining_counter.set(self.routes_counter.get() - self.completed_counter.get())
        speed = self.speed_counter.get()
        predict_value = 60 * self.remaining_counter.get() / float(speed)
        self.predict_counter.set(str(timedelta(seconds=int(predict_value))))
        progress_value = 100 * self.completed_counter.get() / self.routes_counter.get()
        self.progress_var.set(int(progress_value))

    def run_time(self):
        """Запускает отсчет отработанного времени."""
        self.time_counter += 1
        self.worked_hours_counter.set(str(timedelta(seconds=self.time_counter)))
        if self.state:
            self.root.after(1000, self.run_time)

    def start(self):
        """Запускает работу счетчиков панели результатов."""
        self.last_action = time.time()
        block_buttons(self.root.btn_panel.start_btn)
        activate_buttons(*self.root.btn_panel.buttons[1:])
        self.state = 1
        self.run_time()

    def pause(self):
        """Останавливает работу счетчиков."""
        block_buttons(*self.root.btn_panel.buttons[1:])
        activate_buttons(self.root.btn_panel.start_btn)
        self.state = 0

    def calc_speed(self):
        """Вычисляет скорость работы."""
        try:
            action_time = time.time()
            speed = round(60 / (action_time - self.last_action), 1)
            if not int(speed):
                speed = 0.1
            self.speed_counter.set(str(speed))
            self.last_action = action_time
            return speed
        except ZeroDivisionError:
            return 15


class ButtonPanel:
    """Класс описывает панель с кнопками на графическом интерфейсе."""

    def __init__(self, root):
        """Инициализация аттрибутов

            Аргументы:
                root(App): главное окно приложения.
        """
        self.root = root
        self.btn_frame = ttk.Frame(root)
        self.autoclicker_frame = ttk.Frame(self.btn_frame)
        self.button_style = ttk.Style()
        self.button_style.configure("mystyle.TButton", font='Arial 13', padding=10)
        self.icons = {
            'screen_icon': Img(file=r'icons/icons8-камера-50.png'),
            'show_icon': Img(file=r'icons/icons8-глаз-50.png'),
            'cancel_icon': Img(file=r'icons/icons8-отмена-50.png'),
            'video_icon': Img(file=r'icons/icons8-видеозвонок-64.png')

        }
        self.start_btn = None
        self.break_btn = None
        self.screen_btn = None
        self.del_btn = None
        self.show_btn = None
        self.video_btn = None
        self.buttons = []
        self.init_buttons()

    def init_buttons(self):
        """Инициализация всех кнопок на панели."""
        self.start_btn = ttk.Button(self.btn_frame, text="Работать", state='disabled',
                                    command=self.root.res_panel.start, takefocus=False, style='mystyle.TButton'
                                    )
        self.break_btn = ttk.Button(self.btn_frame, text="Перерыв", state='disabled',
                                    command=self.root.res_panel.pause, takefocus=False, style='mystyle.TButton'
                                    )
        self.screen_btn = ttk.Button(self.btn_frame, text="Скрин", image=self.icons['screen_icon'], compound='right',
                                     command=lambda: self.root.table.execute_command(action='1'), state='disabled',
                                     takefocus=False, style='mystyle.TButton'
                                     )
        self.video_btn = ttk.Button(self.btn_frame, text="Видео", image=self.icons['video_icon'], compound='right',
                                   state='disabled', takefocus=False, style='mystyle.TButton'
                                   )
        self.cancel_btn = ttk.Button(self.btn_frame, image=self.icons['cancel_icon'], state='disabled',
                                     compound='image', takefocus=False, command=self.root.table.del_command,
                                     )
        self.show_btn = ttk.Button(self.btn_frame, image=self.icons['show_icon'], compound='image',
                                   takefocus=False, command=lambda: ImageEditor(self.root), state='disabled',
                                   )

        self.buttons.extend((self.start_btn, self.break_btn, self.screen_btn, self.cancel_btn))

    def pack_buttons(self):
        """Размещает все кнопки на панели."""
        self.autoclicker_frame.grid(row=4, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
        self.btn_frame.grid(row=2, column=1, sticky='nsew', pady=130)
        self.start_btn.grid(row=1, column=0, sticky='nsew', pady=10, padx=10)
        self.break_btn.grid(row=1, column=1, sticky='nsew', pady=10, padx=10)
        self.video_btn.grid(row=2, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
        self.screen_btn.grid(row=3, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
        self.cancel_btn.grid(row=3, column=4, pady=10, padx=20)
        self.show_btn.grid(row=2, column=4, pady=10, padx=20)

