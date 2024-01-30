import time
from threading import Thread
import pyautogui as pg
from collections import defaultdict
from functools import partial
import openpyxl
import datetime
from datetime import timedelta
import os
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog as fd
import keyboard
import json
from Аutoclicker import AutoClicker
from selenium.common.exceptions import ElementNotInteractableException


END = False
X, Y, WIDTH, HEIGHT = None, None, None, None
with open('cords.json', 'r') as f:
    cords_dict = json.load(f)
    X, Y = cords_dict['left_top']
    WIDTH = cords_dict['width']
    HEIGHT = cords_dict['height']
editing_state = False


class EditWindow(tk.Toplevel):
    def __init__(self, values=None):
        global editing_state
        super().__init__()
        self.values = values
        self.geometry("330x450")
        self.title('Редактировать')
        self.resizable(False, False)
        self.grab_set()
        self.geometry("+{}+{}".format(root.winfo_rootx() + 50, root.winfo_rooty() + 50))
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
        res_panel.state = 0
        self.protocol('WM_DELETE_WINDOW', self.finish_editing)

    def finish_editing(self):
        global editing_state
        res_panel.state = 1
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
        self.btn_frame.grid(row=7, column=0, pady=20, sticky='nswe',)

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
        item_id = str(table.current_item)
        old_values = table.item(str(table.current_item))['values']
        edited_values = self.get_data()
        if old_values[:7] != edited_values:
            for i, value in enumerate(edited_values):
                table.set(str(table.current_item), i, value)
            self.change_counter(old_values, edited_values)
            if item_id not in table.edited_items:
                table.edited_items[item_id] = {
                    'old_values': old_values,
                }
        self.finish_editing()

    @staticmethod
    def change_counter(old_values, edited_values):
        if not old_values[:7][-1] and str(edited_values[-1]).strip():
            res_panel.add_flight()
        if old_values[:7][-1] and not str(edited_values[-1]).strip():
            res_panel.subtract_flight()


class Table(ttk.Treeview):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.heading('date', text='Дата', anchor='w', image=icon_plan)
        self.heading('route', text='Маршрут', anchor='w', image=icon_route)
        self.heading('direction', text='Направление', anchor='w', image=icon_direction, )
        self.heading("start", text="Начало", anchor='w', image=icon_start)
        self.heading("finish", text="Конец", anchor='w', image=icon_finish)
        self.heading("bus_numb", text="Гос.номер", anchor='w', image=icon_bus)
        self.heading('screen', text="Скрин", anchor='w', image=icon_screen)
        self.tag_configure('bold', font=('Arial', 13, 'bold'))
        self.column("date", width=140, stretch=True)
        self.column("route", width=130, stretch=True)
        self.column("direction", stretch=True, width=155)
        self.column("start", stretch=True, width=80)
        self.column("finish", stretch=True, width=90)
        self.column("bus_numb", stretch=True, width=115)
        self.column("screen", stretch=True, width=90)
        self.tag_configure('gray_colored', background='#D3D3D3')
        self.tag_configure('green_colored', background='#98FB98')
        self.tag_configure('white_colored', background='white')
        self.tag_configure('blue_colored', background='#00BFFF')
        self.table_size = 0
        self.current_item = 0
        self.edited_items = {}
        self.autoclicker = AutoClicker()
        self.current_bus_numb = None

    def run_autoclicker(self):
        self.autoclicker.state = 1
        while self.current_item < self.table_size and self.autoclicker.state:
            try:
                values = self.item(str(self.current_item))['values']
                route, bus_numb = str(values[1]), values[5]
                datetime_obj = get_datetime_obj(values[0], values[3])
                reset = True if bus_numb == self.current_bus_numb else False
                self.autoclicker(route, bus_numb, datetime_obj, reset)
                time.sleep(2)
                self.take_action(values, 'Есть')
                if not reset:
                    self.current_bus_numb = bus_numb
            except ElementNotInteractableException:
                self.autoclicker.pause()
                print('Ошибка, элемент не кликается!')

    def cancel(self): # надо посмотреть
        item_id = str(self.current_item)
        values = self.item(str(self.current_item))['values']
        screen_path, screen = values[7], values[6]
        if item_id in self.edited_items:
            if screen_path:
                os.remove(screen_path)
            for i, value in enumerate(self.edited_items[item_id]['old_values']):
                self.set(item_id, i, value)
            if self.current_item + 1 < self.table_size:
                self.current_item += 1
                self.color(-1)
                self.check()
                self.yview_scroll(1, 'units')
            if screen:
                res_panel.subtract_flight()
            del self.edited_items[item_id]

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
        if action == 'Есть' and values[6]:
            screen_path = self.make_screenshot(values, True)
            self.set(str(self.current_item), 6, action)
            self.set(str(self.current_item), 7, screen_path)
            self.set(str(self.current_item), 9, 1)
            self.down(True)
        elif action == 'Есть' and not values[6]:
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
            button7['state'] = 'normal'
        else:
            button7['state'] = 'disabled'

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
            values = (date, route, flight['direction'], flight['start_time'],
                      flight['finish_time'], flight['bus_numb'], '', '', position, 0,
                      flight['row_numb'], 'white_colored', flight['route_numb'])
            self.insert('', 'end', values=values, iid=str(counter))
        self.table_size = counter + 1
        if self.table_size:
            self.item(str(self.current_item), tags=('gray_colored',))
        else:
            block_buttons(*buttons)

    def make_screenshot(self, values, exc=False, edited=False):
        bus_numb = values[5]
        date = values[0]
        if date not in os.listdir('скрины'):
            os.mkdir(rf'скрины\{date}')
        if not edited:
            screen_name = f'{bus_numb[:2]} {bus_numb[2:-2]} {bus_numb[-2:]}'
            screen_name += ' - ' + self.get_order(screen_name, date)
            screen_path = rf'скрины\{date}\{screen_name}.jpg'
        else:
            screen_path = self.item(str(self.current_item))['values'][7]
        screen = pg.screenshot(region=(X, Y, WIDTH, HEIGHT))
        screen.save(screen_path)
        if not exc:
            res_panel.add_flight()
        return screen_path

    def get_order(self, formated_numb, date):
        path = rf'скрины\{date}'
        order = 1
        for scr_name in os.listdir(path):
            if formated_numb in scr_name:
                order += 1
        return str(order)


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
        self.label_1 = ttk.Label(self.frame_1, text='Тип файла:', justify='left')
        self.label_2 = ttk.Label(self.frame_2, text='Файл с неучтёнными рейсами:', justify='left')
        self.combobox_value = tk.StringVar()
        self.combobox = ttk.Combobox(self.frame_1, textvariable=self.combobox_value,
                                     values=('НС', 'Комиссия'), justify='left')
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
        Thread(target=run).start()

    def set(self):
        self.label_1.grid(row=0, column=0, columnspan=2, sticky='we')
        self.combobox.grid(row=1, column=0)
        self.label_2.grid(row=2, column=0, columnspan=2, sticky='we')
        self.text_1.grid(row=3, column=0)
        self.btn_1.grid(row=3, column=1)
        self.frame_1.pack(ipadx=280, pady=10)
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
        table.fill_out_table(rd)

    def update_progressbar(self, event):
        cur_value = self.progress_var.get()
        self.progress_var.set(cur_value + 10)


class DownloadWindow(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.root = root
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
        block_buttons(button1)
        activate_buttons(*buttons[1:])
        self.state = 1
        self.run_time()

    def pause(self):
        block_buttons(*buttons[1:])
        activate_buttons(button1)
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
        self.final_report_path = None
        self.dict_problems = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        self.wb = None
        self.date = None
        self.total = 0
        self.report_type = None

    def read(self):
        self.wb = openpyxl.load_workbook(self.file_path)
        self.report_type = load_window.combobox_value.get()
        load_window.progress_text_var.set(f'Загрузка файла...')
        sheet = self.wb.active
        counter = 0
        step = sheet.max_row//10
        step_counter = step
        for row in sheet.rows:
            counter += 1
            step_counter -= 1
            if step_counter == 0:
                root.event_generate('<<Updated>>', when='tail')
                time.sleep(0.5)
                step_counter = 0
            if type(row[2].value) == datetime.datetime and row[18].value is None:
                self.total += 1
                date = self.convert_str_to_date(row[2].value)
                route = row[1].value
                bus_numb = row[11].value
                self.dict_problems[date][route][bus_numb].append(
                    {
                        'direction': row[3].value,
                        'start_time': self.convert_str_to_time(row[4].value),
                        'finish_time': self.convert_str_to_time(row[8].value),
                        'screen': None,
                        'row_numb': counter,
                        'bus_numb': bus_numb,
                        'route': route,
                        'route_numb': row[0].value
                    }
                )

        load_window.progress_text_var.set(f'Файл {self.file_path} прочитан')
        load_window.progress_var.set(100)
        load_window.ok_btn['state'] = 'normal'
        activate_buttons()

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
        for i in range(table.table_size):
            item_id = str(i)
            if item_id in table.edited_items:
                values = table.item(item_id)['values']
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
                self.reader.wb.active[f'T{row_numb}'] = values[6]

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
    table.autoclicker.run_webdriver()


def get_datetime_obj(date_st, time_st):
    datetime_st = f'{date_st} {time_st}'
    datetime_obj = datetime.datetime.strptime(datetime_st, '%d.%m.%Y %H:%M') - datetime.timedelta(minutes=2)
    if datetime_obj.hour < 2:
        datetime_obj += datetime.timedelta(hours=24)
    return datetime_obj


def new_window():
    global load_window
    load_window = LoadWindow(root)
    load_window.set()


def download_routes():
    dwn_window = DownloadWindow()
    table.autoclicker.download_routes(dwn_window)


load_window = None
flight = None
root = tk.Tk()
res_panel = ResultPanel(root)
res_panel.prepare_panel()
res_panel.main_frame.grid(row=0, column=0, columnspan=6, sticky='nsew')
res_panel.progressbar.grid(row=1, column=0, columnspan=6, sticky='ew', pady=10)
button_frame = ttk.Frame(root)
icon_route = tk.PhotoImage(file=r'icons/icons8-автобусный-маршрут-20.png')
icon_direction = tk.PhotoImage(file=r'icons/icons8-направление-22.png')
icon_plan = tk.PhotoImage(file=r'icons/icons8-расписание-22.png')
icon_start = tk.PhotoImage(file=r'icons/icons8-начало-22.png')
icon_finish = tk.PhotoImage(file=r'icons/icons8-end-function-button-on-computer-keybord-layout-22.png')
icon_bus = tk.PhotoImage(file=r'icons/icons8-автобус-22.png')
icon_problem = tk.PhotoImage(file=r'icons/icons8-внимание-22.png')
icon_screen = tk.PhotoImage(file=r'icons/icons8-скриншот-22.png')
video_icon = tk.PhotoImage(file=r'icons/icons8-видеозвонок-64.png')
screen_icon = tk.PhotoImage(file=r'icons/icons8-камера-50.png')
pass_step_icon = tk.PhotoImage(file=r'icons/icons8-пропустить-шаг-64.png')
up_icon = tk.PhotoImage(file=r'icons/icons8-стрелка-вверх-50-2.png')
down_icon = tk.PhotoImage(file=r'icons/icons8-стрелка-вниз-50-2.png')
show_icon = tk.PhotoImage(file=r'icons/icons8-глаз-50.png')
cancel_icon = tk.PhotoImage(file=r'icons/icons8-отмена-50.png')
play_icon = tk.PhotoImage(file=r'icons/icons8-треугольник-50.png')
pause_icon = tk.PhotoImage(file=r'icons/icons8-пауза-50.png')
stop_icon = tk.PhotoImage(file=r'icons/icons8-стоп-50.png')
edit_icon = tk.PhotoImage(file=r'icons/icons8-редактировать-50.png')
chrome_icon = tk.PhotoImage(file=r'icons/icons8-google-chrome-50.png')
download_icon = tk.PhotoImage(file=r'icons/icons8-скачать-50.png')
table_style = ttk.Style()
table_style.configure('Treeview', font=('Arial', 13), rowheight=60, separator=100)
heading_style = ttk.Style()
heading_style.configure('Treeview.Heading', font=('Arial', 12))

button_style = ttk.Style()
button_style.configure("mystyle.TButton",
                       font='Arial 13',
                       padding=10,

                       )

columns = ("date", 'route', "direction", "start", 'finish',
           'bus_numb', 'screen', 'screen_path', 'position', 'edited',
           'row_id', 'colour', 'route_id')
table = Table(column=columns, show='headings', padding=10,
              displaycolumns=('date', "route", "direction", "start", 'finish', 'bus_numb', 'screen'))
table.bind('<<TreeviewSelect>>', table.item_selected)
autoclicker_frame = ttk.Frame(button_frame)
button1 = ttk.Button(button_frame, text="Работать", state='disabled', command=res_panel.start, takefocus=False,
                     style='mystyle.TButton')
button2 = ttk.Button(button_frame, text="Перерыв", state='disabled', command=res_panel.pause, takefocus=False,
                     style='mystyle.TButton')
button3 = ttk.Button(button_frame, text="Скрин", command=lambda: table.take_action('Есть'),
                     image=screen_icon, compound='right', state='disabled', takefocus=False, style='mystyle.TButton')
button4 = ttk.Button(autoclicker_frame, image=play_icon,
                     compound='image', takefocus=False, command=lambda: Thread(target=table.run_autoclicker).start(),
                                                                        )
button5 = ttk.Button(button_frame, image=cancel_icon,
                     compound='image', takefocus=False, command=table.cancel,
                                                                        )
button6 = ttk.Button(button_frame, image=edit_icon,
                     compound='image', takefocus=False, command=lambda: EditWindow(table.item(str(table.current_item))['values']))
button7 = ttk.Button(button_frame, image=show_icon,
                     compound='image', takefocus=False, command=table.show_screen, state='disabled')
button8 = ttk.Button(autoclicker_frame, image=pause_icon,
                     compound='image', takefocus=False, command=table.autoclicker.pause,
                                                                        )
button9 = ttk.Button(autoclicker_frame, image=stop_icon,
                     compound='image', takefocus=False, command=table.autoclicker.stop,
                                                                        )
button10 = ttk.Button(button_frame, image=chrome_icon,
                     compound='image', takefocus=False, command=lambda: thread_1.start(),
                                                                        )
button11 = ttk.Button(button_frame, image=download_icon,
                     compound='image', takefocus=False, command=lambda: Thread(target=download_routes).start()
                      )

buttons = [button1, button2, button3]
button1.grid(row=1, column=0, sticky='nsew', pady=10, padx=10)
button2.grid(row=1, column=1, sticky='nsew', pady=10, padx=10)
button3.grid(row=2, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
button4.grid(row=0, column=0, padx=30, pady=10)
button8.grid(row=0, column=1, padx=30, pady=10)
button9.grid(row=0, column=2, padx=30, pady=10)
button6.grid(row=5, column=4, pady=10, padx=30, sticky='nsew')
autoclicker_frame.grid(row=3, column=0, columnspan=2, sticky='nsew', pady=5, padx=5)
button10.grid(row=3, column=4, pady=10, padx=20)
button5.grid(row=2, column=4, pady=10, padx=20)
button7.grid(row=1, column=4, pady=10, padx=20)
button11.grid(row=6, column=4, pady=10, padx=30, sticky='nsew')
button_frame.grid(row=2, column=1, sticky='nsew', pady=130)
table.grid(row=2, column=2, columnspan=4, sticky='NSEW')
table.rowconfigure(0, pad=15)
scroll = tk.Scrollbar(command=table.yview)
scroll.grid(row=2, column=6, sticky='ns')
table.config(yscrollcommand=scroll.set)
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=1)
root.grid_rowconfigure(2, weight=1)
root.grid_columnconfigure(2, weight=1)
root.title('ScreenShooter')
root.bind('<<Updated>>', lambda event: load_window.update_progressbar(event))


def run():
    try:
        rd.read()
        activate_buttons(button1)
        res_panel.flights_counter.set(rd.total)
        res_panel.remaining_counter.set(rd.total)
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
        else:
            show_error("Открыт файл эксель с неучтенными рейсами.Закройте файл и перезапустите приложение!")


def finish():
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
        else:
            show_error("Открыт файл эксель с неучтенными рейсами.Закройте файл и перезапустите приложение!")
    finally:
        root.destroy()

try:
    os.mkdir('скрины') if 'скрины' not in os.listdir() else None
    rd = Reader()
    thread_1 = Thread(target=run_webdriver)
    root.after(1000, new_window)
    root.protocol('WM_DELETE_WINDOW', finish)
    root.mainloop()
except FileNotFoundError:
    show_error("Не найден файл эксель с неучтенными рейсами.Загрузите файл и перезапустите приложение!")
except PermissionError:
    pass
except AttributeError:
    pass