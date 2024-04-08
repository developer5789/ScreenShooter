import time
from threading import Thread
import pyautogui as pg
from collections import defaultdict
import datetime
from datetime import timedelta
import os
import tkinter as tk
from tkinter import  PhotoImage as Img
from tkinter import ttk
from tkinter import filedialog as fd
import json
from Аutoclicker import AutoClicker
from selenium.common.exceptions import NoSuchWindowException
from loger import Loger
from messages import show_inf, show_error

def get_config():
    """Получаем найстроки пользователя из файла config.json"""
    global config, x, y, width, height, profile_path, timeout
    with open('config.json', encoding='utf-8') as f:
        config = json.load(f)
    x, y, width, height = config['screen_cords']
    profile_path = config['profile_path']
    timeout = config['timeout']

def activate_buttons(*btns):
    """Переводит кнопки в активное состояние."""
    for btn in btns:
        btn['state'] = 'normal'


def block_buttons(*btns):
    """Делает кнопки неактивными."""
    for btn in btns:
        btn['state'] = 'disabled'

months = {
    1: 'Январь',
    2: 'Февраль',
    3: 'Март',
    4: 'Апрель',
    5: 'Май',
    6: 'Июнь',
    7: 'Июль',
    8: 'Август',
    9: 'Сентябрь',
    10: 'Октябрь',
    11: 'Ноябрь',
    12: 'Декабрь'
}
config = {}
x, y, width, height = None, None, None, None
timeout = None
profile_path = ''
get_config()
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
                                          to=10.0,  textvariable=self.timeout_var)
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


class EditWindow(tk.Toplevel):
    """Класс описывает окно редактирования строк таблицы."""

    def __init__(self, app, values: list):
        """Инициализация элементов окна,аттрибутов, настройка параметров

        Аргументы:
            values(list): список из значений ячеек строки таблицы.
            app(App): главное окно приложения.

        """
        super().__init__()
        self.app = app
        self.values = values
        self.geometry("330x450")
        self.title('Редактировать')
        self.resizable(False, False)
        self.grab_set()
        self.geometry("+{}+{}".format(self.app.winfo_rootx() + 50, self.app.winfo_rooty() + 50))
        self.date_frame = ttk.Frame(self)
        self.date_label = ttk.Label(self.date_frame, text="Дата:")
        self.date_entry = tk.Entry(self.date_frame, justify='left', width=35)
        self.route_frame = ttk.Frame(self, padding=3)
        self.route_label = tk.Label(self.route_frame, text="Маршрут:")
        self.route_entry = tk.Entry(self.route_frame, width=35)
        self.direction_frame = ttk.Frame(self, padding=3)
        self.direction_label = tk.Label(self.direction_frame, text="Направление:")
        self.direction_entry = tk.Entry(self.direction_frame, width=35)
        self.start_frame = ttk.Frame(self, padding=3)
        self.start_label = tk.Label(self.start_frame, text="Время начала рейса:")
        self.start_entry = tk.Entry(self.start_frame, width=35)
        self.end_frame = ttk.Frame(self, padding=3)
        self.end_label = tk.Label(self.end_frame, text="Время окончания рейса:")
        self.end_entry = tk.Entry(self.end_frame, width=35)
        self.bus_numb_frame = ttk.Frame(self, padding=3)
        self.bus_numb_label = tk.Label(self.bus_numb_frame, text="Гос.номер:")
        self.bus_numb_entry = tk.Entry(self.bus_numb_frame, width=35)
        self.screen_frame = ttk.Frame(self, padding=3)
        self.screen_label = tk.Label(self.screen_frame, text="Скрин:")
        self.screen_entry = tk.Entry(self.screen_frame, width=35)
        self.btn_frame = ttk.Frame(self, padding=3)
        self.btn_edit = ttk.Button(self.btn_frame, text='Редактировать', command=self.edit)
        self.fields = [self.date_entry, self.route_entry, self.direction_entry, self.start_entry,
                       self.end_entry, self.bus_numb_entry, self.screen_entry]
        self.pack_items()
        self.fill_out_fields()
        self.protocol('WM_DELETE_WINDOW', self.destroy)


    def pack_items(self):
        """Размещает элементы внутри окна."""
        self.date_frame.grid(row=0, column=0, padx=20)
        self.route_frame.grid(row=1, column=0, padx=20)
        self.direction_frame.grid(row=2, column=0, padx=20)
        self.start_frame.grid(row=3, column=0, padx=20)
        self.end_frame.grid(row=4, column=0, padx=20)
        self.bus_numb_frame.grid(row=5, column=0, padx=20)
        self.screen_frame.grid(row=6, column=0, padx=20)
        self.btn_frame.grid(row=7, column=0, pady=20, sticky='nswe', )

        self.date_label.pack(anchor='w')
        self.date_entry.pack()
        self.route_label.pack(anchor='w')
        self.route_entry.pack()
        self.direction_label.pack(anchor='w')
        self.direction_entry.pack()
        self.start_label.pack(anchor='w')
        self.start_entry.pack()
        self.end_label.pack(anchor='w')
        self.end_entry.pack()
        self.bus_numb_label.pack(anchor='w')
        self.bus_numb_entry.pack()
        self.screen_label.pack(anchor='w')
        self.screen_entry.pack()
        self.btn_edit.pack(side=tk.RIGHT, padx=15)

        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)
        self.rowconfigure(3, weight=1)
        self.rowconfigure(4, weight=1)
        self.rowconfigure(5, weight=1)
        self.rowconfigure(6, weight=1)
        self.rowconfigure(7, weight=1)
        self.columnconfigure(0, weight=1)

    def fill_out_fields(self):
        """Заполняет поля окна значениями из self.values."""
        for i, field in enumerate(self.fields):
            field.insert(0, str(self.values[i]))

    def get_data(self):
        """Возвращает список из значений полей окна."""
        res = []
        for field in self.fields:
            value = field.get()
            try:
                value = int(value)
            except ValueError:
                pass
            res.append(value)
        return res

    def edit(self):
        """Заносит новые значения в таблицу."""
        item_id = str(self.app.table.current_item)
        old_values = self.app.table.item(str(self.app.table.current_item))['values']
        edited_values = self.get_data()
        if old_values[:7] != edited_values:
            for i, value in enumerate(edited_values):
                self.app.table.set(str(self.app.table.current_item), i, value)
            self.change_counter(old_values, edited_values)
            if item_id not in self.app.table.edited_items:
                self.app.table.edited_items[item_id] = {
                    'old_values': old_values,
                }
        self.destroy()


    def change_counter(self, old_values, edited_values):
        """Меняет счетчик разобранных рейсов

        Аргументы:
            old_values(list): старые значения ячеек строки
            edited_values(list): значения ячеек строки после редактирования
        """
        if not old_values[:7][-1] and str(edited_values[-1]).strip():
            self.app.res_panel.add_route()
        if old_values[:7][-1] and not str(edited_values[-1]).strip():
            self.app.res_panel.subtract_route()


class Table(ttk.Treeview):
    """Класс описывает таблицу с информацией о рейсах"""

    columns = ("date", 'route', "direction", "start", 'finish', 'bus_numb',
               'screen', 'screen_path', 'position', 'edited', 'row_id', 'colour', 'route_id', 'route_dir')

    def __init__(self, app, **kwargs):
        """Инициализация аттрибутов, колонок таблицы, создание стилей."""
        super().__init__(**kwargs)
        self.app = app
        self.icons = {
            'icon_direction': Img(file=r'icons/icons8-направление-22.png'),
            'icon_time': Img(file=r'icons/icons8-расписание-22.png'),
            'icon_start': Img(file=r'icons/icons8-начало-22.png'),
            'icon_finish': Img(file=r'icons/icons8-end-function-button-on-computer-keybord-layout-22.png'),
            'icon_bus': Img(file=r'icons/icons8-автобус-22.png'),
            'icon_screen': Img(file=r'icons/icons8-скриншот-22.png'),
            'route_icon': Img(file=r'icons/icons8-автобусный-маршрут-20.png')

        }
        self.heading('date', text='Дата', anchor='w', image=self.icons['icon_time'])
        self.heading('route', text='Маршрут', anchor='w', image=self.icons['route_icon'])
        self.heading('direction', text='Направление', anchor='w', image=self.icons['icon_direction'], )
        self.heading("start", text="Начало", anchor='w', image=self.icons['icon_start'])
        self.heading("finish", text="Конец", anchor='w', image=self.icons['icon_finish'])
        self.heading("bus_numb", text="Гос.номер", anchor='w', image=self.icons['icon_bus'])
        self.heading('screen', text="Скрин", anchor='w', image=self.icons['icon_screen'])
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
        self.tag_configure('blue_colored', background='#00BFFF')
        self.tag_configure('red_colored', background='#FFA07A')
        self.style = ttk.Style()
        self.style.configure('Treeview', font=('Arial', 13), rowheight=60, separator=100)
        self.style.map('Treeview', background=[('selected', '#D3D3D3')], foreground=[('selected', 'black')])
        self.heading_style = ttk.Style()
        self.heading_style.configure('Treeview.Heading', font=('Arial', 12))
        self.bind('<<TreeviewSelect>>', self.select_item)
        self.bind('<Double-Button-1>', self.click_item)
        self.table_size = 0
        self.current_item = 0
        self.edited_items = {}
        self.autoclicker = AutoClicker(profile_path)
        self.current_bus_numb = None

    def run_autoclicker(self):
        """Запускается автоматический режим работы."""
        self.autoclicker.state = 1
        self.autoclicker.skip = False
        block_buttons(self.app.btn_panel.play_btn)
        activate_buttons(self.app.btn_panel.pause_btn, self.app.btn_panel.skip_btn)

        while self.autoclicker.state:
            try:
                values = self.item(str(self.current_item))['values']
                bus_numb = values[5]
                datetime_from = self.get_datetime_str(values[0], values[3])
                datetime_to = self.get_datetime_str(values[0], values[4], start=False)
                reset = True if bus_numb == self.current_bus_numb else False
                self.autoclicker(bus_numb, datetime_from, datetime_to, reset)
                time.sleep(2)
                track = self.autoclicker.check_track()

                if self.autoclicker.skip: # если скипнуть на последнем элементе?
                    self.autoclicker.skip = False
                    self.next_item()
                elif self.autoclicker.state and track:
                    self.autoclicker.focus_on_track(self)
                    self.execute_command(values, 1)
                    if not reset:
                        self.current_bus_numb = bus_numb
                elif self.autoclicker.state and track is not None:
                    self.execute_command(values, 0)
                    if not reset:
                        self.current_bus_numb = bus_numb
                elif self.autoclicker.state and track is None:
                    continue
                else:
                    break

                if self.current_item == self.table_size - 1:
                    break

            except NoSuchWindowException as err:
                self.autoclicker.pause()
                activate_buttons(self.app.btn_panel.play_btn)
                show_error('Потеряна связь с браузером! Сделайте перезагрузку!')
                Loger.enter_in_log(err)

            except Exception as err:
                self.autoclicker.pause()
                activate_buttons(self.app.btn_panel.play_btn)
                show_error('Возникла ошибка при построении трека!Попробуйте перезагрузить страницу!')
                Loger.enter_in_log(err)

        block_buttons(self.app.btn_panel.skip_btn)

    def select_item(self, event):
        self.current_item = int(self.selection()[0])
        self.check()

    def next_item(self):
        if int(self.current_item) + 1 < self.table_size:
            next_id = str(int(self.current_item) + 1)
            self.selection_set((next_id, ))

    def cancel(self):
        """Отмена всех операций, сделанных со строкой таблицы:
        удаление скриншота, значения ячеек возвращаются к исходным, удаление заливки
        """
        item_id = str(self.current_item)

        if item_id in self.edited_items:
            values = self.item(item_id)['values']
            screen_path, screen, root_dir = values[7], values[6], values[13]
            if screen_path:
                self.del_screen(screen_path, root_dir)
            if item_id + 1 < self.table_size:
                self.current_item += 1
                self.color(-1)
                self.check()
                self.yview_scroll(1, 'units')



    def del_screen(self, screen_path: str, root_dir: str):
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

    def click_item(self, event):
        try:
            values = self.item(str(self.current_item))['values']
            bus_numb = values[5]
            datetime_from = self.get_datetime_str(values[0], values[3])
            datetime_to = self.get_datetime_str(values[0], values[4], start=False)
            reset = True if bus_numb == self.current_bus_numb else False
            self.autoclicker(bus_numb, datetime_from, datetime_to, reset)

            if not reset:
                self.current_bus_numb = bus_numb

        except NoSuchWindowException as err:
            show_error('Потеряна связь с браузером! Сделайте перезагрузку!')
            Loger.enter_in_log(err)

        except Exception as err:
            show_error('Упс! Возникла ошибка при построении трека!')
            Loger.enter_in_log(err)


    def execute_command(self, values=None, action=None):
        """Выполняет переданную команду

        Аргументы:
            values(list, None): список значений ячеек строки,
            action(str, None): номер исполняемой команды (1-скрин)
            """
        if values is None:
            values = self.item(str(self.current_item))['values']
        screen = values[7]
        if action and screen:
            screen_path = self.make_screenshot(values, True)
            self.set(str(self.current_item), 6, action)
            self.set(str(self.current_item), 7, screen_path)
            self.down()
        elif action and not screen:
            screen_path = self.make_screenshot(values)
            self.append_to_edited(values)
            self.set(str(self.current_item), 7, screen_path)
            self.set(str(self.current_item), 6, action)
            self.down()
        elif not action and not screen:
            self.set(str(self.current_item), 6, action)
            self.append_to_edited(values)
            self.down()
        else:
            pass


    def append_to_edited(self, values: list):
        """При редактировании добавляет элемент таблицы в self.edited_items.
        Сохраняются исходные значения ячеек строки таблицы.
        """
        if str(self.current_item) not in self.edited_items:
            self.edited_items[str(self.current_item)] = {
                'old_values': values,
            }


    def check(self):
        """Проверка на наличие скрина."""
        screen_path = self.item(str(self.current_item))['values'][7]
        if screen_path:
            self.app.btn_panel.show_btn['state'] = 'normal'
        else:
            self.app.btn_panel.show_btn['state'] = 'disabled'

    def show_screen(self):
        """Открывает скрин текущей строки таблицы."""
        screen_path = self.item(str(self.current_item))['values'][7]
        os.startfile(screen_path)

    def color(self):
        """Заливка после выполнения команды"""

        if self.item(str(self.current_item))['values'][6]:
            self.item(str(self.current_item ), tags=('green_colored',))
        else:
            self.item(str(self.current_item), tags=('red_colored',))


    def down(self):
        """Переключение на следующую строку таблицы после исполнения команды."""
        if self.current_item + 1 < self.table_size:
            self.color()
            self.next_item()
            self.yview_scroll(1, 'units')
        else:
            self.color()
            self.check()

    def fill_out_table(self, rd):
        """Заполняет таблицу значениями из прочитанного документа excel

         Аргументы:
            rd(Reader): обьект класса Reader.
        """
        counter = -1
        for position, flight, date in rd.get_route():
            counter += 1
            route = flight['route']
            root_dir = self.get_root_dir(date, route)
            values = (date, route, flight['direction'], flight['start_time'],
                      flight['finish_time'], flight['bus_numb'], '', '', position, 0,
                      flight['row_numb'], 'white_colored', flight['route_numb'], root_dir)
            self.insert('', 'end', values=values, iid=str(counter))
        self.table_size = counter + 1
        if self.table_size:
            self.selection_set(('0', ))
        else:
            block_buttons(*self.app.btn_panel.buttons)
        rd.routes_dict.clear()

    def get_values(self):
        """Возвращает список значений ячеек текущей строки"""
        return self.item(str(self.current_item))['values']

    def make_screenshot(self, values: list, overwrite=False):
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
        if not overwrite:
            self.app.res_panel.add_route()

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
    """Класс описывает окно  для выбора файла с рейсами и
    отображения """
    def __init__(self, app):
        """Инициализация элементов, аттрибутов, настройка параметров окна."""
        super().__init__()
        self.app = app
        self.geometry("850x250")
        self.title('ScreenShooter')
        self.resizable(False, False)
        self.grab_set()
        self.report_name = None
        self.final_report_name = None
        self.upload_btn = ttk.Button(self, text='Загрузить', state='disabled', command=self.next_step)
        self.frame_1 = ttk.Frame(self, padding=3)
        self.frame_2 = ttk.Frame(self, padding=3)
        self.frame_3 = ttk.Frame(self, padding=3)
        self.text_1 = tk.Text(self.frame_2, height=1, width=65, background='white', font=('Arial', 11))
        self.btn_1 = ttk.Button(self.frame_2, text='Найти', command=lambda: self.select_file())
        self.label_1 = ttk.Label(self.frame_1, text='Тип:', justify='left')
        self.label_2 = ttk.Label(self.frame_2, text='Файл с неучтёнными рейсами:')
        self.combobox_value = tk.StringVar()
        self.combobox = ttk.Combobox(self.frame_1, textvariable=self.combobox_value,
                                     values=('НС', 'Комиссия'))
        self.progress_var = tk.IntVar(value=0)
        self.progress_text_var = tk.StringVar()
        self.progressbar = ttk.Progressbar(self.frame_3, orient="horizontal", length=650, variable=self.progress_var,
                                           style='TProgressbar')
        self.res_label = ttk.Label(self.frame_3, textvariable=self.progress_text_var, justify='left')
        self.ok_btn = ttk.Button(self, text='Ок', state='disabled', command=self.end)

    def select_file(self):
        """Выбирает путь до файла с рейсами."""
        self.report_name = fd.askopenfilename()
        self.text_1.insert(1.0, self.report_name)
        self.upload_btn['state'] = 'normal'

    def next_step(self):
        """Размещает новые виджеты и переходит к этапу чтения файла с рейсами."""
        self.del_first_step()
        self.pack_items(2)
        self.app.rd.file_path = self.report_name
        self.app.rd.final_report_path = self.final_report_name
        Thread(target=self.app.run_reader).start()

    def pack_items(self, step: int):
        """Размещает элементы внутри окна

        Аргументы:
            step(int): параметр, обозначающий этап загрузки файла.
        """
        if step == 1:
            self.label_1.grid(row=0, column=0, columnspan=2, sticky='we')
            self.combobox.grid(row=1, column=0, sticky='we')
            self.label_2.grid(row=2, column=0, columnspan=2, sticky='we')
            self.text_1.grid(row=3, column=0)
            self.btn_1.grid(row=3, column=1)
            self.frame_1.pack(ipadx=225, pady=10 )
            self.frame_2.pack()
            self.upload_btn.pack(side='right', padx=10)
        else:
            self.geometry("800x250")
            self.res_label.grid(row=0, column=0)
            self.progressbar.grid(row=1, column=0, columnspan=4, padx=70)
            self.frame_3.grid(row=0, columnspan=3, pady=60)
            self.ok_btn.grid(row=1, column=2)

    def del_first_step(self):
        """Удаляет виджеты этапа выбора файла."""
        self.frame_1.destroy()
        self.frame_2.destroy()
        self.upload_btn.destroy()

    def end(self):
        """Закрывает окно загрузки и переходит к заполнению таблицы."""
        self.destroy()
        self.app.table.fill_out_table(self.app.rd)

    def update_progressbar(self, event):
        """Увеличивает значение полосы прогресса на 10"""
        cur_value = self.progress_var.get()
        self.progress_var.set(cur_value + 10)


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
            past_minutes = self.time_counter / 60
            speed = round(self.completed_counter.get() / past_minutes, 1)
            self.speed_counter.set(str(speed))
            return speed
        except ZeroDivisionError:
            return 100.0

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
            'pass_step_icon': Img(file=r'icons/icons8-пропустить-шаг-64.png'),
            'show_icon': Img(file=r'icons/icons8-глаз-50.png'),
            'cancel_icon': Img(file=r'icons/icons8-отмена-50.png'),
            'play_icon': Img(file=r'icons/icons8-треугольник-50.png'),
            'pause_icon': Img(file=r'icons/icons8-пауза-50.png'),
            'stop_icon': Img(file=r'icons/icons8-стоп-50.png'),
            'edit_icon': Img(file=r'icons/icons8-редактировать-50.png'),
            'chrome_icon': Img(file=r'icons/icons8-google-chrome-50.png'),
            'download_icon': Img(file=r'icons/icons8-скачать-50.png'),
            'settings_icon': Img(file=r'icons/icons8-настройка-50.png'),
            'skip_icon': Img(file=r'icons/icons8-пропустить-шаг-64.png')

        }
        self.start_btn = None
        self.break_btn = None
        self.screen_btn = None
        self.play_btn = None
        self.pause_btn = None
        self.stop_btn = None
        self.cancel_btn = None
        self.edit_btn = None
        self.show_btn = None
        self.chrome_btn = None
        self.settings_btn = None
        self.skip_btn = None
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
        self.skip_btn = ttk.Button(self.btn_frame, text="Пропустить", image=self.icons['skip_icon'], compound='right',
                                   command=self.root.table.autoclicker.skip_route, state='disabled',
                                   takefocus=False, style='mystyle.TButton'
                                   )
        self.play_btn = ttk.Button(self.autoclicker_frame, image=self.icons['play_icon'],
                                   compound='image', takefocus=False, state='disabled',
                                   command=lambda: Thread(target=self.root.table.run_autoclicker).start()
                                   )
        self.pause_btn = ttk.Button(self.autoclicker_frame, image=self.icons['pause_icon'], state='disabled',
                                    compound='image', takefocus=False, command=self.root.pause_autoclicker,
                                    )
        self.stop_btn = ttk.Button(self.autoclicker_frame, image=self.icons['stop_icon'], state='disabled',
                                   compound='image', takefocus=False, command=self.root.close_autoclicker,
                                   )
        self.cancel_btn = ttk.Button(self.btn_frame, image=self.icons['cancel_icon'], state='disabled',
                                     compound='image', takefocus=False, command=self.root.table.hide_item,
                                     )
        self.edit_btn = ttk.Button(self.btn_frame, image=self.icons['edit_icon'],
                                   compound='image', takefocus=False, state='disabled',
                                   command=lambda: EditWindow(self.root, self.root.table.get_values())
                                   )
        self.show_btn = ttk.Button(self.btn_frame, image=self.icons['show_icon'], compound='image',
                                   takefocus=False, command=self.root.table.show_screen, state='disabled',
                                   )
        self.chrome_btn = ttk.Button(self.btn_frame, image=self.icons['chrome_icon'],
                                     compound='image', takefocus=False, state='disabled',
                                     command=lambda: Thread(target=self.root.run_webdriver).start()
                                     )
        self.settings_btn = ttk.Button(self.btn_frame, image=self.icons['settings_icon'], compound='image',
                                       command=lambda: ConfigWindow(), takefocus=False
                                       )
        self.buttons.extend((self.start_btn, self.break_btn, self.screen_btn, self.cancel_btn,
                             self.edit_btn, self.play_btn,
                             self.stop_btn, self.chrome_btn))

    def pack_buttons(self):
        """Размещает все кнопки на панели."""
        self.autoclicker_frame.grid(row=4, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
        self.btn_frame.grid(row=2, column=1, sticky='nsew', pady=130)
        self.start_btn.grid(row=1, column=0, sticky='nsew', pady=10, padx=10)
        self.break_btn.grid(row=1, column=1, sticky='nsew', pady=10, padx=10)
        self.skip_btn.grid(row=2, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
        self.screen_btn.grid(row=3, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
        self.play_btn.grid(row=0, column=0, padx=30, pady=10)
        self.cancel_btn.grid(row=3, column=4, pady=10, padx=20)
        self.edit_btn.grid(row=5, column=4, pady=10, padx=30, sticky='nsew')
        self.show_btn.grid(row=2, column=4, pady=10, padx=20)
        self.pause_btn.grid(row=0, column=1, padx=30, pady=10)
        self.stop_btn.grid(row=0, column=2, padx=30, pady=10)
        self.chrome_btn.grid(row=4, column=4, pady=10, padx=20)
        self.settings_btn.grid(row=6, column=4, pady=10, padx=30, sticky='nsew')
