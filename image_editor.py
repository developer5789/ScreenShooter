from tkinter import ttk
from tkinter import *
from PIL import Image, ImageTk
import os


class ImageEditor(Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.app = args[0]
        self.table = self.app.table
        self.current_item = str(self.table.current_item)
        self.title("Режим просмотра скринов")
        self.label_style = ttk.Style()
        self.label_style.configure("My.TLabel",  # имя стиля
                                font="helvetica 10",  # шрифт
)
        self.icons = {
                    'del': PhotoImage(file='icons/icons8-отходы-50.png'),
                    'back': PhotoImage(file='icons/icons8-налево-50.png'),
                    'next': PhotoImage(file='icons/icons8-направо-50.png'),
                      }
        self.main_frame = Frame(self)
        self.btn_frame = Frame(self)
        self.canvas = Canvas(self, width=1560, height=800, bg="#F5F5F5")
        self.del_btn = ttk.Button(self.btn_frame, image=self.icons['del'],
                                   compound='image', takefocus=False)
        self.back_btn = ttk.Button(self.btn_frame, image=self.icons['back'], command=self.previous_image,
                                  compound='image', takefocus=False)
        self.next_btn = ttk.Button(self.btn_frame, image=self.icons['next'], command=self.next_image,
                                  compound='image', takefocus=False)

        self.image_file = None
        self.image = None
        self.image_tk = None
        self.images = os.listdir('скрины/Комиссия/Декабрь_80')
        self.start_value = StringVar()
        self.finish_value = StringVar()
        self.bus_numb_value = StringVar()
        self.route_value = StringVar()
        self.pack_items()
        self.open_image()

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

        start_display = Label(start_frame, textvariable=self.start_value, font=("Arial", 14), foreground='#000080', relief='sunken')
        finish_display = Label(finish_frame, textvariable=self.finish_value, font=("Arial", 14), foreground='#000080', relief='sunken')
        bus_numb_display = Label(bus_numb_frame, textvariable=self.bus_numb_value, font=("Arial", 14), foreground='#000080', relief='sunken')
        route_display = Label(route_frame, textvariable=self.route_value, font=("Arial", 14),  foreground='#000080', relief='sunken')

        start_display.pack()
        finish_display.pack()
        bus_numb_display.pack()
        route_display.pack()

        route_frame.grid(row=0, column=0, padx=5, sticky='NSEW')
        start_frame.grid(row=0, column=1, padx=5, sticky='NSEW')
        finish_frame.grid(row=0, column=2, padx=5, sticky='NSEW')
        bus_numb_frame.grid(row=0, column=3, padx=5, sticky='NSEW')
        self.del_btn.pack(side=LEFT, padx=10, pady=5)
        self.back_btn.pack(side=LEFT, padx=10, pady=5)
        self.next_btn.pack(side=LEFT, padx=10, pady=5)

        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)
        self.rowconfigure(2, weight=1)

        self.main_frame.grid(row=0, column=0)
        self.canvas.grid(row=1, column=0, pady=10)
        self.btn_frame.grid(row=2, column=0)


    def next_image(self):
        next_item = self.table.next(self.current_item)
        if next_item:
            self.current_item = next_item
            self.open_image()

    def previous_image(self):
        previous_item = self.table.prev(self.current_item)
        if previous_item:
            self.current_item = previous_item
            self.open_image()

    def open_image(self):
        # открытие изображения для редактирования
        filename = self.table.item(self.current_item)['values'][7]
        if filename:
            self.image_file = filename
            image = Image.open(filename)
            self.image_tk = ImageTk.PhotoImage(image)
            # self.canvas.configure(width=self.image.width, height=self.image.width)
            self.canvas.create_image(0, 0, anchor=NW, image=self.image_tk)
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


