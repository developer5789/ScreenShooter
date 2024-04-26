from tkinter import ttk
import tkinter as tk
from PIL import Image, ImageTk
from messages import last_screen_message


class ImageEditor(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.app = args[0]
        self.table = self.app.table
        self.current_item = str(self.table.current_item)
        self.label_style = ttk.Style()
        self.label_style.configure("My.TLabel",  # имя стиля
                                font="helvetica 10",  # шрифт
)
        self.icons = {
                    'del': tk.PhotoImage(file='icons/icons8-отходы-50.png'),
                    'back': tk.PhotoImage(file='icons/icons8-налево-50.png'),
                    'next': tk.PhotoImage(file='icons/icons8-направо-50.png'),
                    'no_screen': tk.PhotoImage(file='icons/no_screen.png')
                      }
        self.main_frame = tk.Frame(self)
        self.btn_frame = tk.Frame(self)
        self.canvas = tk.Canvas(self, width=1560, height=850, bg="#F5F5F5")
        self.del_btn = ttk.Button(self.btn_frame, image=self.icons['del'],
                                   compound='image', takefocus=False, command=self.del_screen)
        self.back_btn = ttk.Button(self.btn_frame, image=self.icons['back'], command=self.previous_image,
                                  compound='image', takefocus=False)
        self.next_btn = ttk.Button(self.btn_frame, image=self.icons['next'], command=self.next_image,
                                  compound='image', takefocus=False)

        self.image_path = self.table.item(self.current_item)['values'][7]
        self.image = None
        self.image_tk = None
        self.start_value = tk.StringVar()
        self.finish_value = tk.StringVar()
        self.bus_numb_value = tk.StringVar()
        self.route_value = tk.StringVar()
        self.pack_items()
        self.open_image(self.image_path)

    def pack_items(self):
        label_style = ttk.Style()
        label_style.configure("My.TLabel")
        route_frame = ttk.Frame(self.main_frame,  borderwidth=1, padding=3, style='My.TLabel')
        start_frame = ttk.Frame(self.main_frame,  borderwidth=1, padding=3, style='My.TLabel')
        finish_frame = ttk.Frame(self.main_frame,  borderwidth=1, padding=3, style='My.TLabel')
        bus_numb_frame = ttk.Frame(self.main_frame,  borderwidth=1, padding=3, style='My.TLabel')

        route_label = ttk.Label(route_frame, text="Маршрут", font=("Arial", 12, 'bold'), style='My.TLabel')
        start_label = ttk.Label(start_frame, text="Начало", font=("Arial", 12, 'bold'), style='My.TLabel')
        finish_label = ttk.Label(finish_frame, text="Конец", font=("Arial", 12, 'bold'), style='My.TLabel')
        bus_numb_label = ttk.Label(bus_numb_frame, text="Гос.номер", font=("Arial", 12, 'bold'), style='My.TLabel')

        route_label.pack()
        start_label.pack()
        finish_label.pack()
        bus_numb_label.pack()

        start_display = tk.Label(start_frame, textvariable=self.start_value,
                                 font=("Arial", 14), foreground='#000080', relief='sunken')
        finish_display = tk.Label(finish_frame, textvariable=self.finish_value,
                                  font=("Arial", 14), foreground='#000080', relief='sunken')
        bus_numb_display = tk.Label(bus_numb_frame, textvariable=self.bus_numb_value,
                                    font=("Arial", 14), foreground='#000080', relief='sunken')
        route_display = tk.Label(route_frame, textvariable=self.route_value,
                                 font=("Arial", 14),  foreground='#000080', relief='sunken')

        start_display.pack()
        finish_display.pack()
        bus_numb_display.pack()
        route_display.pack()

        route_frame.grid(row=0, column=0, padx=5, sticky='NSEW')
        start_frame.grid(row=0, column=1, padx=5, sticky='NSEW')
        finish_frame.grid(row=0, column=2, padx=5, sticky='NSEW')
        bus_numb_frame.grid(row=0, column=3, padx=5, sticky='NSEW')
        self.del_btn.pack(side=tk.LEFT, padx=10, pady=5)
        self.back_btn.pack(side=tk.LEFT, padx=10, pady=5)
        self.next_btn.pack(side=tk.LEFT, padx=10, pady=5)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)

        self.main_frame.grid(row=0, column=0)
        self.canvas.grid(row=1, column=0, pady=10)
        self.btn_frame.grid(row=2, column=0)

    @staticmethod
    def find_image(func):
        def wrapper(*args):
            self = args[0]
            switch_func = func(*args)
            item = self.current_item
            while item:
                item = switch_func(item)
                if item:
                    filename = self.table.item(item)['values'][7]
                    if filename:
                        self.current_item = item
                        self.open_image(filename)
                        return
            last_screen_message(self)
        return wrapper

    @find_image
    def next_image(self):
        return self.table.next

    @find_image
    def previous_image(self):
        return self.table.prev

    def open_image(self, filepath):
        if filepath:
            self.image_path = filepath
            image = Image.open(filepath)
            self.image_tk = ImageTk.PhotoImage(image)
            img_name = filepath.split("\\")[-1]
            self.title(f"Режим просмотра скринов: {img_name}")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.image_tk)
            self.set_values(self.current_item)

    def set_values(self, item_id):
        values = self.table.item(item_id)['values']
        date, route, start, finish, bus_numb = values[0:2] + values[3:6]
        datetime_start = self.table.get_datetime_str(date, start)
        datetime_finish = self.table.get_datetime_str(date, finish, start=False)
        self.route_value.set(route)
        self.start_value.set(datetime_start)
        self.finish_value.set(datetime_finish)
        self.bus_numb_value.set(bus_numb)

    def del_screen(self):
        values = self.table.item(self.current_item)['values']
        screen_path, root_dir = values[7], values[13]
        if screen_path:
            res = self.table.askdel(screen_path, root_dir, self, self.current_item)
            if res is not None:
                self.image_tk = self.icons['no_screen']
                self.title(f"Режим просмотра скринов: скриншот не найден")
                self.canvas.create_image(780, 425, image=self.image_tk)





