import time
from threading import Thread
import pyautogui as pg
from collections import defaultdict
import openpyxl
import datetime
from datetime import timedelta
import os
import tkinter as tk
from tkinter import messagebox, PhotoImage as Img
from tkinter import ttk
from tkinter import filedialog as fd
import keyboard
import json
from Аutoclicker import AutoClicker
from selenium.common.exceptions import (ElementNotInteractableException,
                                        ElementClickInterceptedException, NoSuchElementException)

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
editing_state = False
screen_paths = defaultdict(lambda: defaultdict(list))


class ConfigWindow(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        global editing_state
        super().__init__(*args, **kwargs)
        self.geometry("400x200")
        self.title('Настройки')
        self.resizable(False, False)
        self.grab_set()
        self.geometry("+{}+{}".format(app.winfo_rootx() + 50, app.winfo_rooty() + 50))
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
        # editing_state = True
        # app.res_panel.state = 0
        self.protocol('WM_DELETE_WINDOW', self.finish_editing)
        self.values = []
        self.fill_out_fields()

    def finish_editing(self):
        global editing_state
        app.res_panel.state = 1
        editing_state = False
        self.destroy()

    def pack_items(self):
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

        # self.rowconfigure(0, weight=1)
        # self.rowconfigure(1, weight=1)
        # self.rowconfigure(2, weight=1)
        # self.rowconfigure(3, weight=1)
        # self.rowconfigure(4, weight=1)
        # self.rowconfigure(5, weight=1)
        # self.rowconfigure(6, weight=1)
        # self.rowconfigure(7, weight=1)
        # self.columnconfigure(0, weight=1)

    def fill_out_fields(self):
        profile_path = config['profile_path']
        screen_cords = str(config['screen_cords']).strip('[]')
        timeout_val = config['timeout']
        self.values.extend((profile_path, screen_cords, timeout_val))
        self.path_entry.insert('end', profile_path)
        self.cords_entry.insert('end', screen_cords)
        self.timeout_var.set(timeout_val)

    def apply(self):
        new_values = [self.path_entry.get(),
                    [int(val.strip()) for val in self.cords_entry.get().split(',')],
                    self.timeout_var.get()
        ]
        if self.values != new_values:
            self.set_settings(new_values)
        self.destroy()

    def set_settings(self, new_values):
        global x, y, width, height, timeout, profile_path
        for i, key in enumerate(config):
            config[key] = new_values[i]
        timeout = config["timeout"]
        x, y, width, height = config["screen_cords"]
        profile_path = config["profile_path"]
        self.save_settings()

    @staticmethod
    def save_settings():
        with open('config.json', 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=4)


class EditWindow(tk.Toplevel):
    def __init__(self, values=None):
        global editing_state
        super().__init__()
        self.values = values
        self.geometry("330x450")
        self.title('Редактировать')
        self.resizable(False, False)
        self.grab_set()
        self.geometry("+{}+{}".format(app.winfo_rootx() + 50, app.winfo_rooty() + 50))
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
        editing_state = True
        app.res_panel.state = 0
        self.protocol('WM_DELETE_WINDOW', self.finish_editing)

    def finish_editing(self):
        global editing_state
        app.res_panel.state = 1
        editing_state = False
        self.destroy()

    def pack_items(self):
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
        for i, field in enumerate(self.fields):
            field.insert(0, str(self.values[i]))

    def get_data(self):
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
        item_id = str(app.table.current_item)
        old_values = app.table.item(str(app.table.current_item))['values']
        edited_values = self.get_data()
        if old_values[:7] != edited_values:
            for i, value in enumerate(edited_values):
                app.table.set(str(app.table.current_item), i, value)
            self.change_counter(old_values, edited_values)
            if item_id not in app.table.edited_items:
                app.table.edited_items[item_id] = {
                    'old_values': old_values,
                }
        self.finish_editing()

    @staticmethod
    def change_counter(old_values, edited_values):
        if not old_values[:7][-1] and str(edited_values[-1]).strip():
            app.res_panel.add_flight()
        if old_values[:7][-1] and not str(edited_values[-1]).strip():
            app.res_panel.subtract_flight()


class Table(ttk.Treeview):
    columns = ("date", 'route', "direction", "start", 'finish', 'bus_numb',
               'screen', 'screen_path', 'position', 'edited', 'row_id', 'colour', 'route_id', 'route_dir')

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
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
        self.style = ttk.Style()
        self.style.configure('Treeview', font=('Arial', 13), rowheight=60, separator=100)
        self.heading_style = ttk.Style()
        self.heading_style.configure('Treeview.Heading', font=('Arial', 12))
        self.bind('<<TreeviewSelect>>', self.item_selected)
        self.bind('<Double-Button-1>', self.click_item)
        self.table_size = 0
        self.current_item = 0
        self.edited_items = {}
        self.autoclicker = AutoClicker(profile_path)
        self.current_bus_numb = None

    def run_autoclicker(self):
        self.autoclicker.state = 1
        block_buttons(app.play_btn)
        activate_buttons(app.pause_btn, app.skip_btn)
        self.autoclicker.update_widgets()
        while self.current_item < self.table_size and self.autoclicker.state:
            try:
                values = self.item(str(self.current_item))['values']
                route, bus_numb = str(values[1]), values[5]
                datetime_obj = get_datetime_obj(values[0], values[3])
                reset = True if bus_numb == self.current_bus_numb else False
                self.autoclicker(route, bus_numb, datetime_obj, reset)
                time.sleep(timeout)

                if self.autoclicker.skip:
                    self.autoclicker.skip = False
                    self.current_item += 1
                    self.color(-1)
                elif self.autoclicker.state:
                    self.take_action(values, '1')
                    if not reset:
                        self.current_bus_numb = bus_numb
                else:
                    break

            except (ElementNotInteractableException, ElementClickInterceptedException):
                self.autoclicker.pause()
                activate_buttons(app.play_btn)
                show_error('Возникла ошибка при обращении к элементу!')
            except (NoSuchElementException, AttributeError):
                self.autoclicker.pause()
                activate_buttons(app.play_btn)
                show_error('Элемент не найден!')
            except Exception:
                self.autoclicker.pause()
                activate_buttons(app.play_btn)
                show_error('Возникла ошибка!')

        block_buttons(app.skip_btn)

    def cancel(self):  # надо посмотреть
        item_id = str(self.current_item)
        values = self.item(str(self.current_item))['values']
        screen_path, screen, root_dir = values[7], values[6], values[13]
        if item_id in self.edited_items:
            if screen_path:
                self.del_screen(screen_path, root_dir)
            for i, value in enumerate(self.edited_items[item_id]['old_values']):
                self.set(item_id, i, value)
            if self.current_item + 1 < self.table_size:
                self.current_item += 1
                self.color(-1)
                self.check()
                self.yview_scroll(1, 'units')
            if screen:
                app.res_panel.subtract_flight()
            del self.edited_items[item_id]

    @staticmethod
    def del_screen(screen_path, root_dir):
        if rd.report_type == 'НС':
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
            if not self.autoclicker.widgets:
                self.autoclicker.update_widgets()

            values = self.item(str(self.current_item))['values']
            route, bus_numb = str(values[1]), values[5]
            datetime_obj = get_datetime_obj(values[0], values[3])
            reset = True if bus_numb == self.current_bus_numb else False
            self.autoclicker(route, bus_numb, datetime_obj, reset)

            if not reset:
                self.current_bus_numb = bus_numb
        except (ElementNotInteractableException, ElementClickInterceptedException):
            show_error('Возникла ошибка при обращении к элементу!')

        except (NoSuchElementException, AttributeError):
            show_error('Элемент не найден!')

        except Exception:
            show_error('Возникла ошибка!')

    def item_selected(self, event):
        if len(self.selection()):
            offset = self.current_item - int(self.selection()[0])
            if offset:
                selected_item = self.current_item - offset
                self.selection_remove(str(selected_item))
                self.current_item = selected_item
                self.color(offset)
                self.check()
            else:
                self.selection_remove(str(self.current_item))

    def up(self):
        if self.current_item - 1 >= 0:
            exc = self.item(str(self.current_item))['values'][10]
            self.current_item -= 1
            self.color(1, exc)
            self.check()
            self.yview_scroll(-1, 'units')

    def take_action(self, values=None, action=None):
        if self.current_item > self.table_size - 1:
            return
        if values is None:
            values = self.item(str(self.current_item))['values']
        if action == '1' and values[6]:
            screen_path = self.make_screenshot(values, True)
            self.set(str(self.current_item), 6, action)
            self.set(str(self.current_item), 7, screen_path)
            self.set(str(self.current_item), 9, 1)
            self.down(True)
        elif action == '1' and not values[6]:
            self.append_to_edited(values)
            screen_path = self.make_screenshot(values)
            self.set(str(self.current_item), 7, screen_path)
            self.set(str(self.current_item), 6, action)
            self.down()
        else:
            self.down()

    def append_to_edited(self, values):
        if str(self.current_item) not in self.edited_items:
            self.edited_items[str(self.current_item)] = {
                'old_values': values,
            }

    def check(self):
        screen_path = self.item(str(self.current_item))['values'][7]
        if screen_path:
            app.show_btn['state'] = 'normal'
        else:
            app.show_btn['state'] = 'disabled'

    def show_screen(self):
        screen_path = self.item(str(self.current_item))['values'][7]
        os.startfile(screen_path)

    def color_after_action(self, offset, exc):
        self.item(str(self.current_item), tags=('gray_colored',))
        if self.item(str(self.current_item + offset))['values'][6] and not exc:
            self.item(str(self.current_item + offset), tags=('green_colored',))
            self.set(str(self.current_item + offset), 11, 'green_colored')
        elif self.item(str(self.current_item + offset))['values'][6] and exc:
            self.item(str(self.current_item + offset), tags=('blue_colored',))
            self.set(str(self.current_item + offset), 11, 'blue_colored')
        # else:
        #     self.item(str(self.current_item + offset), tags=('white_colored',))
        #     self.set(str(self.current_item + offset), 11, 'white_colored')

    def color(self, offset):
        colour = self.item(str(self.current_item + offset))['values'][11]
        self.item(str(self.current_item), tags=('gray_colored',))
        self.item(str(self.current_item + offset), tags=(colour,))

    def down(self, edited=False):
        if self.current_item + 1 < self.table_size:
            exc = self.item(str(self.current_item))['values'][9]
            self.current_item += 1
            self.color_after_action(-1, exc)
            self.check()
            self.yview_scroll(1, 'units')
        else:
            if edited:
                self.set(str(self.current_item), 11, 'blue_colored')
            else:
                self.set(str(self.current_item), 11, 'green_colored')

    def fill_out_table(self, rd):
        counter = -1
        for position, flight, date in rd.get_flight():
            counter += 1
            route = flight['route']
            root_dir = self.get_root_dir(date, route)
            values = (date, route, flight['direction'], flight['start_time'],
                      flight['finish_time'], flight['bus_numb'], '', '', position, 0,
                      flight['row_numb'], 'white_colored', flight['route_numb'], root_dir)
            self.insert('', 'end', values=values, iid=str(counter))
        self.table_size = counter + 1
        if self.table_size:
            self.item(str(self.current_item), tags=('gray_colored',))
        else:
            block_buttons(*app.buttons)
        rd.dict_problems.clear()

    def make_screenshot(self, values, exc=False, edited=False):
        if not edited:
            screen_path = self.get_screen_path(values)
        else:
            screen_path = self.item(str(self.current_item))['values'][7]

        screen = pg.screenshot(region=(x, y, width, height))
        screen.save(screen_path)
        if not exc:
            app.res_panel.add_flight()

        return screen_path

    def get_screen_path(self, values):
        root_dir = values[13]
        if not os.path.exists(root_dir):
            os.makedirs(root_dir)

        if rd.report_type == 'НС':
            bus_numb = values[5]
            indx = self.get_screen_indx(bus_numb, root_dir)
            return root_dir + '\\' + f'{bus_numb} - {indx}.jpg'
        else:
            route_numb = values[12]
            screen_path = root_dir + '\\' + f'{route_numb}.jpg'

        return screen_path

    @staticmethod
    def get_screen_indx(bus_numb, root_dir):
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

    @staticmethod
    def get_root_dir(date, route):
        month_numb = int(date.split('.')[1])
        path = fr'скрины\{rd.report_type}\{months[month_numb]}_{route}'
        return path


class LoadWindow(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.geometry("850x250")
        self.title('ScreenShooter')
        self.resizable(False, False)
        self.grab_set()
        self.report_name = None
        self.final_report_name = None
        self.btn_upload = ttk.Button(self, text='Загрузить', state='disabled', command=self.close)
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
        self.report_name = fd.askopenfilename()
        self.text_1.insert(1.0, self.report_name)
        self.btn_upload['state'] = 'normal'

    def close(self):
        self.del_first_step()
        self.set_upload_step()
        rd.file_path = self.report_name
        rd.final_report_path = self.final_report_name
        Thread(target=app.run_reader).start()

    def set(self):
        self.label_1.grid(row=0, column=0, columnspan=2, sticky='we')
        self.combobox.grid(row=1, column=0, sticky='we')
        self.label_2.grid(row=2, column=0, columnspan=2, sticky='we')
        self.text_1.grid(row=3, column=0)
        self.btn_1.grid(row=3, column=1)
        self.frame_1.pack(ipadx=225, pady=10 )
        self.frame_2.pack()
        self.btn_upload.pack(side='right', padx=10)

    def del_first_step(self):
        self.frame_1.destroy()
        self.frame_2.destroy()
        self.btn_upload.destroy()

    def set_upload_step(self):
        self.geometry("800x250")
        self.res_label.grid(row=0, column=0)
        self.progressbar.grid(row=1, column=0, columnspan=4, padx=70)
        self.frame_3.grid(row=0, columnspan=3, pady=60)
        self.ok_btn.grid(row=1, column=2)

    def end(self):
        self.destroy()
        app.table.fill_out_table(rd)

    def update_progressbar(self, event):
        cur_value = self.progress_var.get()
        self.progress_var.set(cur_value + 10)


class DownloadWindow(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.root = app
        self.geometry("700x250")
        self.title('Загрузка рейсов')
        self.resizable(False, False)
        self.grab_set()
        self.frame_1 = ttk.Frame(self, padding=3)
        self.progress_var = tk.IntVar(value=0)
        self.progress_text_var = tk.StringVar(value='Загрузка рейсов с РНИС...')
        self.progressbar = ttk.Progressbar(self.frame_1, orient="horizontal", length=500, variable=self.progress_var,
                                           style='TProgressbar')
        self.res_label = ttk.Label(self.frame_1, textvariable=self.progress_text_var, justify='left')
        self.bind('<<Updated>>', self.update_progressbar)
        self.pack_widgets()

    def pack_widgets(self):
        self.res_label.grid(row=0, column=0)
        self.progressbar.grid(row=1, column=0, columnspan=4, padx=70)
        self.frame_1.grid(row=0, columnspan=3, pady=60)

    def update_progressbar(self, event):
        cur_value = self.progress_var.get()
        self.progress_var.set(cur_value + 10)

    def end(self):
        self.destroy()


class ResultPanel:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.main_frame = ttk.Frame(root, relief='groove')
        self.progress_var = tk.IntVar()
        self.progressbar = ttk.Progressbar(orient="horizontal", length=700, variable=self.progress_var)
        self.speed_counter = tk.StringVar(value='0')
        self.flights_counter = tk.IntVar()
        self.completed_counter = tk.IntVar()
        self.remaining_counter = tk.IntVar()
        self.worked_hours_counter = tk.StringVar(value='00:00:00')
        self.predict_counter = tk.StringVar(value='00:00:00')
        self.time_counter = 0
        self.state = 0

    def prepare_panel(self):
        speed_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)
        predict_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)
        flights_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)
        completed_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)
        remaining_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)
        worked_hours_frame = ttk.Frame(self.main_frame, relief='groove', borderwidth=1, padding=3)

        speed_label = ttk.Label(speed_frame, text="Cкорость(cкр/мин):", font=("Arial", 12))
        predict_label = ttk.Label(predict_frame, text="Прогноз:", font=("Arial", 12))
        flights_label = ttk.Label(flights_frame, text="Всего:", font=("Arial", 12))
        completed_label = ttk.Label(completed_frame, text="Сделано:", font=("Arial", 12))
        remaining_label = ttk.Label(remaining_frame, text="Осталось:", font=("Arial", 12))
        worked_hours_label = ttk.Label(worked_hours_frame, text="Отработано:", font=("Arial", 12))

        speed_label.pack(side=tk.LEFT)
        predict_label.pack(side=tk.LEFT)
        flights_label.pack(side=tk.LEFT)
        completed_label.pack(side=tk.LEFT)
        remaining_label.pack(side=tk.LEFT)
        remaining_label.pack(side=tk.LEFT)
        worked_hours_label.pack(side=tk.LEFT)

        speed_display = ttk.Label(speed_frame, textvariable=self.speed_counter, font=("Arial", 18))
        flights_display = ttk.Label(flights_frame, textvariable=self.flights_counter, font=("Arial", 18))
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
        flights_frame.grid(row=0, column=3, padx=3, sticky='NSEW')
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

    def add_flight(self):
        self.completed_counter.set(self.completed_counter.get() + 1)
        self.remaining_counter.set(self.flights_counter.get() - self.completed_counter.get())
        speed = self.calc_speed()
        predict_value = 60 * self.remaining_counter.get() / speed
        self.predict_counter.set(str(timedelta(seconds=int(predict_value))))
        progress_value = 100 * self.completed_counter.get() / self.flights_counter.get()
        self.progress_var.set(int(progress_value))
        if not self.remaining_counter.get():
            show_inf()

    def subtract_flight(self):
        self.completed_counter.set(self.completed_counter.get() - 1)
        self.remaining_counter.set(self.flights_counter.get() - self.completed_counter.get())
        speed = self.speed_counter.get()
        predict_value = 60 * self.remaining_counter.get() / float(speed)
        self.predict_counter.set(str(timedelta(seconds=int(predict_value))))
        progress_value = 100 * self.completed_counter.get() / self.flights_counter.get()
        self.progress_var.set(int(progress_value))

    def run_time(self):
        self.time_counter += 1
        self.worked_hours_counter.set(str(timedelta(seconds=self.time_counter)))
        if self.state:
            self.root.after(1000, self.run_time)
        if editing_state:
            self.root.after(1000, self.run_time)

    def start(self):
        block_buttons(app.start_btn)
        activate_buttons(*app.buttons[1:])
        self.state = 1
        self.run_time()

    def pause(self):
        block_buttons(*app.buttons[1:])
        activate_buttons(app.start_btn)
        self.state = 0

    def calc_speed(self):
        try:
            past_minutes = self.time_counter / 60
            speed = round(self.completed_counter.get() / past_minutes, 1)
            self.speed_counter.set(str(speed))
            return speed
        except ZeroDivisionError:
            return 100.0


def show_error(message):
    messagebox.showerror('Ошибка', message)


def show_inf():
    messagebox.showinfo('Уведомление', 'Ура!Разбор рейсов закончен!')


def open_warning():
    messagebox.showwarning(title="Предупреждение", message="У вас открыт файл эксель с неучтенными рейсами."
                                                           "Для корректного завершения программы закройте его и нажмите 'Ок'")


def get_value(val):
    if type(val) == datetime.time:
        return f"{val.hour:>02}:{val.minute:>02}"
    return val


def find_report():
    for file in os.listdir():
        if file.endswith('.xlsx'):
            return file
    raise FileNotFoundError


class Reader:
    def __init__(self, file_path=None):
        self.file_path = file_path
        self.dict_problems = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        self.wb = None
        self.total = 0
        self.report_type = None

    def read(self):
        self.wb = openpyxl.load_workbook(self.file_path)
        self.report_type = app.load_window.combobox_value.get()
        app.load_window.progress_text_var.set(f'Загрузка файла...')
        sheet = self.wb.active
        counter = 0
        step = sheet.max_row // 10
        step_counter = step
        for row in sheet.rows:
            counter += 1
            step_counter -= 1
            if step_counter == 0:
                app.event_generate('<<Updated>>', when='tail')
                time.sleep(0.5)
                step_counter = 0
            if type(row[2].value) == datetime.datetime and row[18].value is None:
                self.total += 1
                date = self.convert_str_to_date(row[2].value)
                route = row[1].value
                bus_numb = self.get_bus_numb(row[11].value)
                self.dict_problems[date][route][bus_numb].append(
                    {
                        'direction': row[3].value,
                        'start_time': self.convert_str_to_time(row[7].value),
                        'finish_time': self.convert_str_to_time(row[8].value),
                        'screen': None,
                        'row_numb': counter,
                        'bus_numb': bus_numb,
                        'route': route,
                        'route_numb': row[0].value
                    }
                )

        app.load_window.progress_text_var.set(f'Файл {self.file_path} прочитан')
        app.load_window.progress_var.set(100)
        app.load_window.ok_btn['state'] = 'normal'
        activate_buttons()

    @staticmethod
    def get_bus_numb(bus_numb: str):
        if not bus_numb:
            return ''
        bus_numb = bus_numb.replace(' ', '')
        bus_numb = f'{bus_numb[:2]} {bus_numb[2:-2]} {bus_numb[-2:]}'
        return bus_numb


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
            return st.strftime('%d.%m.%Y')
        elif st.strip() and '.' in st:
            return st.strip()

    def convert_route_into_numb(self, st):
        try:
            return int(st)
        except Exception:
            return st

    def get_flight(self):
        for date in self.dict_problems:
            for route in self.dict_problems[date]:
                for bus_numb in self.dict_problems[date][route]:
                    for position, dict in enumerate(self.dict_problems[date][route][bus_numb]):
                        yield position, dict, date


class Writer:
    def __init__(self, reader: Reader):
        self.reader = reader

    def write(self):
        for i in range(app.table.table_size):
            item_id = str(i)
            if item_id in app.table.edited_items:
                values = app.table.item(item_id)['values']
                row_numb = values[10]
                start_time = self.convert_into_datetime(values[3], '%H.%M.%S')
                self.reader.wb.active[f'B{row_numb}'] = values[1]
                self.reader.wb.active[f'C{row_numb}'] = self.convert_into_datetime(values[0], '%d.%m.%Y')
                self.reader.wb.active[f'D{row_numb}'] = values[2]
                self.reader.wb.active[f'E{row_numb}'] = start_time
                self.reader.wb.active[f'F{row_numb}'] = start_time
                self.reader.wb.active[f'H{row_numb}'] = start_time
                self.reader.wb.active[f'I{row_numb}'] = self.convert_into_datetime(values[4], '%H.%M.%S')
                self.reader.wb.active[f'L{row_numb}'] = values[5]
                self.reader.wb.active[f'S{row_numb}'] = values[6]

        self.reader.wb.save(self.reader.file_path)

    @staticmethod
    def convert_into_datetime(st_datetime, format=None):
        try:
            return datetime.datetime.strptime(st_datetime.strip(), format)
        except ValueError:
            return st_datetime


def activate_buttons(*btns):
    for btn in btns:
        btn['state'] = 'normal'


def block_buttons(*btns):
    for btn in btns:
        btn['state'] = 'disabled'


def run_webdriver():
    if not app.table.autoclicker:
        app.table.autoclicker.run_webdriver()


def get_datetime_obj(date_st, time_st):
    datetime_st = f'{date_st} {time_st}'
    datetime_obj = datetime.datetime.strptime(datetime_st, '%d.%m.%Y %H:%M') - datetime.timedelta(minutes=2)
    if datetime_obj.hour < 2:
        datetime_obj += datetime.timedelta(hours=24)
    return datetime_obj


def make_dirs():
    if 'скрины' not in os.listdir():
        os.mkdir('скрины')
    if rd.report_type not in os.listdir('скрины'):
        os.mkdir(fr'скрины\{rd.report_type}')


def get_paths():
    root_path = rf'скрины\{rd.report_type}'
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


class App(tk.Tk):

    def __init__(self, *args, **kwargs):
        self.get_config()
        super().__init__(*args, **kwargs)
        self.title('ScreenShooter')
        self.load_window = None
        self.res_panel = ResultPanel(self)
        self.table = Table(column=Table.columns, show='headings', padding=10,
                           displaycolumns=('date', "route", "direction", "start", 'finish', 'bus_numb', 'screen'))
        self.scroll = tk.Scrollbar(command=self.table.yview)
        self.table.config(yscrollcommand=self.scroll.set)
        self.btn_frame = ttk.Frame(self)
        self.autoclicker_frame = ttk.Frame(self.btn_frame)
        self.bind('<<Updated>>', lambda event: app.load_window.update_progressbar(event))
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
        self.download_btn = None
        self.settings_btn = None
        self.buttons = []
        self.button_style = ttk.Style()
        self.button_style.configure("mystyle.TButton", font='Arial 13', padding=10)
        self.init_buttons()
        self.pack()
        self.after(1000, self.show_load_window)
        self.protocol('WM_DELETE_WINDOW', self.finish)

    def init_buttons(self):
        self.start_btn = ttk.Button(self.btn_frame, text="Работать", state='disabled',
                                    command=self.res_panel.start, takefocus=False, style='mystyle.TButton'
                                    )  # button1
        self.break_btn = ttk.Button(self.btn_frame, text="Перерыв", state='disabled', command=self.res_panel.pause,
                                    takefocus=False, style='mystyle.TButton'
                                    )  # button2
        self.screen_btn = ttk.Button(self.btn_frame, text="Скрин", image=self.icons['screen_icon'], compound='right',
                                     command=lambda: self.table.take_action(action='1'), state='disabled',
                                     takefocus=False, style='mystyle.TButton'
                                     )  # button3
        self.skip_btn = ttk.Button(self.btn_frame, text="Пропустить", image=self.icons['skip_icon'], compound='right',
                                     command=self.table.autoclicker.skip_route, state='disabled',
                                     takefocus=False, style='mystyle.TButton'
                                     )
        self.play_btn = ttk.Button(self.autoclicker_frame, image=self.icons['play_icon'],
                                   compound='image', takefocus=False, state='disabled',
                                   command=lambda: Thread(target=self.table.run_autoclicker).start()
                                   )  # button4
        self.pause_btn = ttk.Button(self.autoclicker_frame, image=self.icons['pause_icon'], state='disabled',
                                    compound='image', takefocus=False, command=self.pause_autoclicker,
                                    )  # button8
        self.stop_btn = ttk.Button(self.autoclicker_frame, image=self.icons['stop_icon'], state='disabled',
                                   compound='image', takefocus=False, command=self.table.autoclicker.stop,
                                   )  # button9
        self.cancel_btn = ttk.Button(self.btn_frame, image=self.icons['cancel_icon'], state='disabled',
                                     compound='image', takefocus=False, command=self.table.cancel,
                                     )  # button5
        self.edit_btn = ttk.Button(self.btn_frame, image=self.icons['edit_icon'],
                                   compound='image', takefocus=False, state='disabled',
                                   command=lambda: EditWindow(self.table.item(str(self.table.current_item))['values'])
                                   )  # button6 поменять вызов command
        self.show_btn = ttk.Button(self.btn_frame, image=self.icons['show_icon'], compound='image',
                                   takefocus=False, command=self.table.show_screen, state='disabled',
                                   )  # button7
        self.chrome_btn = ttk.Button(self.btn_frame, image=self.icons['chrome_icon'],
                                     compound='image', takefocus=False, state='disabled',
                                     command=lambda: Thread(target=run_webdriver).start()
                                     )  # button10
        self.download_btn = ttk.Button(self.btn_frame, image=self.icons['download_icon'], compound='image', state='disabled',
                                       takefocus=False, command=lambda: Thread(target=self.download_routes).start()
                                       )  # button11
        self.settings_btn = ttk.Button(self.btn_frame, image=self.icons['settings_icon'], compound='image',
                                       command=lambda: ConfigWindow(), takefocus=False

                                       )
        self.buttons.extend((self.start_btn, self.break_btn, self.screen_btn, self.cancel_btn,
                             self.edit_btn, self.download_btn, self.play_btn, self.pause_btn,
                             self.stop_btn, self.chrome_btn))

    def pack(self):
        self.table.grid(row=2, column=2, columnspan=4, sticky='NSEW')
        self.scroll.grid(row=2, column=6, sticky='ns')
        self.res_panel.prepare_panel()
        self.res_panel.main_frame.grid(row=0, column=0, columnspan=6, sticky='nsew')
        self.res_panel.progressbar.grid(row=1, column=0, columnspan=6, sticky='ew', pady=10)
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
        self.download_btn.grid(row=6, column=4, pady=10, padx=30, sticky='nsew')
        self.settings_btn.grid(row=7, column=4, pady=10, padx=30, sticky='nsew')

        self.table.rowconfigure(0, pad=15)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(2, weight=1)

    def pause_autoclicker(self):
        self.table.autoclicker.pause()
        block_buttons(self.pause_btn, self.skip_btn)
        activate_buttons(self.play_btn)
        self.table.autoclicker.widgets.clear()

    def finish(self):
        try:
            wr = Writer(rd)
            wr.write()
        except PermissionError:
            if rd.wb:
                while True:
                    try:
                        open_warning()
                        wr = Writer(rd)
                        wr.write()
                        break
                    except PermissionError:
                        continue
        except AttributeError:
            pass
        finally:
            if self.table.autoclicker:
                self.table.autoclicker.stop()
            self.destroy()

    def run_reader(self):
        try:
            rd.read()
            activate_buttons(self.start_btn)
            self.res_panel.flights_counter.set(rd.total)
            self.res_panel.remaining_counter.set(rd.total)
            make_dirs()
            if rd.report_type == 'НС':
                get_paths()
        except PermissionError:
            show_error("Открыт файл эксель с неучтенными рейсами. Закройте файл и перезапустите приложение!")
        except Exception:
            show_error("Возникла ошибка! Перезапустите приложение!")

    def show_load_window(self):
        self.load_window = LoadWindow(self)
        self.load_window.set()

    def download_routes(self):
        dwn_window = DownloadWindow()
        self.table.autoclicker.download_routes(dwn_window)

    @staticmethod
    def get_config():
        global config, x, y, width, height, profile_path, timeout
        with open('config.json', encoding='utf-8') as f:
            config = json.load(f)
        x, y, width, height = config['screen_cords']
        profile_path = config['profile_path']
        timeout = config['timeout']


try:
    app = App()
    rd = Reader()
    app.mainloop()
except PermissionError:
    pass

