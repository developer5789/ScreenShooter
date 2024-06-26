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

    columns = ("date", "queue", "direction", "start_plan", 'start_fact', 'bus_numb', 'problem',
               'screen', 'screen_path', 'position', 'row_id', 'colour', 'route_dir', 'route')

    def __init__(self, app, **kwargs):
        """Инициализация аттрибутов, колонок таблицы, создание стилей."""
        super().__init__(**kwargs)
        self.app = app
        self.icons = {
            'direction': Img(file=r'icons/icons8-направление-22.png'),
            'date': Img(file=r'icons/icons8-дата-22.png'),
            'start_plan': Img(file=r'icons/icons8-расписание-22.png'),
            'start_fact': Img(file=r'icons/free-icon-shrug-5894030.png'),
            'bus_numb': Img(file=r'icons/icons8-автобус-22.png'),
            'screen': Img(file=r'icons/icons8-скриншот-22.png'),
            'queue': Img(file=r'icons/icons8-автобусный-маршрут-20.png'),
            'filter': Img(file=r'icons/icons8-фильтр-22.png'),
            'problem': Img(file=r'icons/icons8-внимание-22.png')

        }
        self.heading('date', text='Дата', anchor='w', image=self.icons['date'],
                     command=lambda: FilterWindow(self.app, 'date'))
        self.heading('queue', text='Наряд', anchor='w', image=self.icons['queue'],
                     command=lambda: FilterWindow(self.app, 'queue'))
        self.heading('direction', text='Направление', anchor='w', image=self.icons['direction'],
                     command=lambda: FilterWindow(self.app, 'direction'))
        self.heading("start_plan", text="План", anchor='w', image=self.icons['start_plan'],
                     command=lambda: FilterWindow(self.app, 'start_plan'))
        self.heading("start_fact", text="Факт", anchor='w', image=self.icons['start_fact'],
                     command=lambda: FilterWindow(self.app, 'start_fact'))
        self.heading("bus_numb", text="ТС", anchor='w', image=self.icons['bus_numb'],
                     command=lambda: FilterWindow(self.app, 'bus_numb'))
        self.heading('problem', text="Проблема", anchor='w', image=self.icons['problem'],
                     command=lambda: FilterWindow(self.app, 'problem'))
        self.heading('screen', text="Скрин", anchor='w', image=self.icons['screen'],
                     command=lambda: FilterWindow(self.app, 'screen'))

        self.column("date", width=140, stretch=True)
        self.column("queue", width=90, stretch=True)
        self.column("direction", stretch=True, width=140)
        self.column("start_plan", stretch=True, width=60)
        self.column("start_fact", stretch=True, width=60)
        self.column("bus_numb", stretch=True, width=90)
        self.column("problem", stretch=True, width=250)
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
        self.loading_label = None
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
        checked_routes, filtered_routes = 0, 0
        for i in range(self.table_size):
            item_id = str(i)
            values = self.item(item_id)['values']
            if self.match(filtered_cols, values):
                self.move(item_id, '', i)
                if values[7].strip():
                    checked_routes += 1
                filtered_routes += 1
            else:
                self.detach(item_id)

        self.app.res_panel.set_progress(filtered_routes, checked_routes, filtered_routes - checked_routes)

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

    def cancel(self):
        screen, screen_path = self.get_values()[7:9]
        if screen or screen_path:
            self.del_screen(screen_path)
            self.set(self.current_item, 7, '')
            self.next_item()
            self.color('white_colored', self.current_item)
            if screen:
                self.app.res_panel.subtract_route()


    def del_screen(self, screen_path):
        """Удаление скриншота

        Аргументы:
            screen_path(str): путь к скрину
            root_dir(str): путь к директории скрина
        """

        if screen_path:
            os.remove(screen_path)
            self.set(self.current_item, 8, '')


    def execute_command(self, action=None): #надо посмотреть
        """Выполняет переданную команду

        Аргументы:
            values(list, None): список значений ячеек строки,
            action(str, None): номер исполняемой команды (1-скрин)
            """

        values = self.item(self.current_item)['values']
        screen, screen_path = str(values[7]), values[8]
        if action == "Скрин":
            screen_path = self.make_screenshot(values)
            self.set(self.current_item, 7, "Есть")
            self.set(self.current_item, 8, screen_path)
            self.color('green_colored', self.current_item)

        if action == "Видео":
            self.del_screen(screen_path) if values[8] else None
            self.set(self.current_item, 7, action)
            self.color('green_colored', self.current_item)


        self.next_item()
        if not screen:
            self.app.res_panel.add_route()

    def check(self):
        """Проверка на наличие скрина."""
        screen_path = self.item(self.current_item)['values'][8]
        if screen_path:
            self.app.btn_panel.show_btn['state'] = 'normal'
        else:
            self.app.btn_panel.show_btn['state'] = 'disabled'

    def show_screen(self):
        """Открывает скрин текущей строки таблицы."""
        screen_path = self.item(self.current_item)['values'][8]
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
        for position, flight in rd.get_route():
            counter += 1
            item = str(counter)
            route_queue = f'{flight["route"]}{"__" + flight["queue"] if flight["queue"] else ""}'
            values = (flight['date'], route_queue, flight['direction'], flight['plan'],
                      flight['fact'], flight['bus_numb'], flight['problem'], flight['screen'], '', position,
                      flight['row_numb'], 'white_colored', "root_dir", flight['route'])
            self.insert('', 'end', values=values, iid=item)

            self.find_screen(item, flight)

        self.app.res_panel.set_progress(rd.total, rd.total - self.empty_val, self.empty_val)

        self.table_size = counter + 1
        if self.table_size:
            self.selection_set(('0',))
        else:
            block_buttons(*self.app.btn_panel.buttons)
        rd.dict_problems.clear()

    def get_values(self):
        """Возвращает список значений ячеек текущей строки"""
        return self.item(self.current_item)['values']

    def find_screen(self, item, flight):
        screen = flight['screen']
        if screen == "Есть":
            screen_name = f"{flight['plan'].replace(':', '_')} {flight['bus_numb']} {flight['problem']}.jpg"
            screen_path = f"скрины\\{flight['date']}\\{flight['route']}\\{screen_name}"
            if os.path.exists(screen_path):
                self.set(item, 8, screen_path)
                self.item(item, tags=('green_colored',))
            else:
                self.item(item, tags=('orange_colored',))

        elif screen == 'Видео':
            self.item(item, tags=('green_colored',))

        elif not screen:
            self.empty_val += 1


    # def make_screenshot(self, values: list):
    #     """Делает скрин трека
    #
    #     Аргументы:
    #         values(list): список значений ячеек строки,
    #         overwrite(bool): параметр, показывающий перезаписывается ли скрин.
    #     """
    #     screen_path = values[7]
    #
    #     if not screen_path:
    #         screen_path = self.get_screen_path(values)
    #
    #     screen = pg.screenshot(region=(x, y, width, height))
    #     screen.save(screen_path)
    #
    #     return screen_path
    #

    def make_screenshot(self, values):
        date, route = values[0], values[13]

        if date not in os.listdir('скрины'):
            os.mkdir(rf'скрины\{date}')
        if str(route) not in os.listdir(fr'скрины\{date}'):
            os.mkdir(rf'скрины\{date}\{route}')
        screnshot_name = f'{values[3].replace(":", "_")} {values[5]} {values[6]}'
        pg.screenshot(rf'скрины\{date}\{route}\{screnshot_name}.jpg')

        return rf'скрины\{date}\{route}\{screnshot_name}.jpg'


    def show_loading_label(self):
        self.loading_label = ttk.Label(self, text="Загрузка данных...", background='white', font=("Arial", 14))
        self.loading_label.place(relx=.5, rely=.5, anchor="c")





class LoadWindow(tk.Toplevel):
    def __init__(self, app):
        super().__init__()
        self.app = app
        self.geometry("700x200")
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
        self.destroy()
        Thread(target=self.app.run_reader).start()


    def set(self):
        self.frame_1 = ttk.Frame(self, padding=3)
        self.frame_2 = ttk.Frame(self, padding=3)
        btn_1 = ttk.Button(self.frame_1, text='Найти', command=lambda: self.select_file('unaccounted_flights'))
        btn_2 = ttk.Button(self.frame_2, text='Найти', command=lambda: self.select_file())
        label_1 = ttk.Label(self.frame_1, text='Файл с неучтёнными рейсами:', justify='left')
        label_2 = ttk.Label(self.frame_2, text='Файл со всеми рейсами:', justify='left')
        self.text_1 = tk.Text(self.frame_1, height=1, width=60, background='white', font=('Arial', 10))
        self.text_2 = tk.Text(self.frame_2, height=1, width=60, background='white', font=('Arial', 10))
        label_1.grid(row=0, column=0, columnspan=2, sticky='we')
        self.text_1.grid(row=1, column=0)
        btn_1.grid(row=1, column=1)
        label_2.grid(row=0, column=0, columnspan=2, sticky='we')
        self.text_2.grid(row=1, column=0)
        btn_2.grid(row=1, column=4)
        self.frame_1.pack()
        self.frame_2.pack()
        self.btn_upload.pack(side='right', padx=10)



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

    def prepare_panel(self):
        """Размещает элементы на панели."""
        label_style = ttk.Style()
        label_style.configure("My.TLabel")

        speed_frame = ttk.Frame(self.main_frame, borderwidth=2, padding=2, style='My.TLabel', relief='groove')
        predict_frame = ttk.Frame(self.main_frame, borderwidth=2, padding=1, style='My.TLabel', relief='groove')
        routes_frame = ttk.Frame(self.main_frame,  borderwidth=2, padding=1, style='My.TLabel', relief='groove')
        completed_frame = ttk.Frame(self.main_frame, borderwidth=2, padding=1, style='My.TLabel', relief='groove')
        remaining_frame = ttk.Frame(self.main_frame,  borderwidth=2, padding=1, style='My.TLabel', relief='groove')
        worked_hours_frame = ttk.Frame(self.main_frame, borderwidth=2, padding=1, style='My.TLabel', relief='groove')

        speed_label = ttk.Label(speed_frame, text="Cкорость(cкр/мин)", font=("Arial", 12, 'bold'), style='My.TLabel')
        predict_label = ttk.Label(predict_frame, text="Прогноз", font=("Arial", 12, 'bold'), style='My.TLabel')
        routes_label = ttk.Label(routes_frame, text="Всего", font=("Arial", 12, 'bold'), style='My.TLabel')
        completed_label = ttk.Label(completed_frame, text="Сделано", font=("Arial", 12, 'bold'), style='My.TLabel')
        remaining_label = ttk.Label(remaining_frame, text="Осталось", font=("Arial", 12, 'bold'), style='My.TLabel')
        worked_hours_label = ttk.Label(worked_hours_frame, text="Отработано", font=("Arial", 12, 'bold'), style='My.TLabel')

        speed_label.pack()
        worked_hours_label.pack()
        predict_label.pack()
        routes_label.pack()
        completed_label.pack()
        remaining_label.pack()

        speed_display = ttk.Label(speed_frame, textvariable=self.speed_counter, foreground='#000080',
                                                                                font=("Arial", 18), relief='sunken')
        flights_display = ttk.Label(routes_frame, textvariable=self.routes_counter, foreground='#000080', font=("Arial",  18), relief='sunken')
        completed_display = ttk.Label(completed_frame, textvariable=self.completed_counter, foreground='#000080', font=("Arial", 18), relief='sunken')
        remaining_display = ttk.Label(remaining_frame, textvariable=self.remaining_counter, foreground='#000080', font=("Arial", 18), relief='sunken')
        worked_hours_display = ttk.Label(worked_hours_frame, textvariable=self.worked_hours_counter,foreground='#000080', font=("Arial", 18), relief='sunken')
        predict_display = ttk.Label(predict_frame, textvariable=self.predict_counter, foreground='#000080', font=("Arial", 18), relief='sunken')

        speed_display.pack(pady=3)
        flights_display.pack(pady=3)
        completed_display.pack(pady=3)
        remaining_display.pack(pady=3)
        predict_display.pack(pady=3)
        worked_hours_display.pack(pady=3)

        speed_frame.grid(row=0, column=0, sticky='NSEW')
        worked_hours_frame.grid(row=0, column=1, sticky='NSEW')
        predict_frame.grid(row=0, column=2,  sticky='NSEW')
        routes_frame.grid(row=0, column=3, sticky='NSEW')
        completed_frame.grid(row=0, column=4,  sticky='NSEW')
        remaining_frame.grid(row=0, column=5,  sticky='NSEW')

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

        speed = self.speed_counter.get()
        if speed != "0":
            predict_value = 60 * self.remaining_counter.get() / float(speed)
            self.predict_counter.set(str(timedelta(seconds=int(predict_value))))


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
        speed = 2 if self.speed_counter.get() == '0' else self.speed_counter.get()
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
                                     command=lambda: self.root.table.execute_command(action='Скрин'), state='disabled',
                                     takefocus=False, style='mystyle.TButton'
                                     )
        self.video_btn = ttk.Button(self.btn_frame, text="Видео", image=self.icons['video_icon'], compound='right',
                                    command=lambda: self.root.table.execute_command(action='Видео'),
                                    state='disabled', takefocus=False, style='mystyle.TButton'
                                   )
        self.cancel_btn = ttk.Button(self.btn_frame, image=self.icons['cancel_icon'], state='disabled',
                                     compound='image', takefocus=False, command=self.root.table.cancel,
                                     )
        self.show_btn = ttk.Button(self.btn_frame, image=self.icons['show_icon'], compound='image',
                                   takefocus=False, command=lambda: ImageEditor(self.root), state='disabled',
                                   )

        self.buttons.extend((self.start_btn, self.break_btn, self.screen_btn, self.cancel_btn, self.video_btn))

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

