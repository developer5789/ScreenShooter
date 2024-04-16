from components import *
from reader import Reader
from writer import Writer
from messages import open_warning


class App(tk.Tk):
    """Класс описывает главное окно приложения."""
    def __init__(self, *args, **kwargs):
        """Инициализация компонентов приложения."""
        super().__init__(*args, **kwargs)
        self.title('ScreenShooter')
        self.rd = Reader(self)
        self.load_window = None
        self.res_panel = ResultPanel(self)
        self.table = Table(self, column=Table.columns, show='headings', padding=10,
                           displaycolumns=('date', "route", "direction", "start", 'finish', 'bus_numb', 'screen'))
        self.scroll = tk.Scrollbar(command=self.table.yview)
        self.table.config(yscrollcommand=self.scroll.set)
        self.btn_panel = ButtonPanel(self)
        self.bind('<<Updated>>', lambda event: app.load_window.update_progressbar(event))
        self.after(1000, self.show_load_window)
        self.protocol('WM_DELETE_WINDOW', self.finish)
        self.pack()



    def pack(self):
        """Размещение компонентов внутри главного окна."""
        self.table.grid(row=2, column=2, columnspan=4, sticky='NSEW')
        self.scroll.grid(row=2, column=6, sticky='ns')
        self.res_panel.prepare_panel()
        self.res_panel.main_frame.grid(row=0, column=0, columnspan=6, sticky='nsew')
        self.res_panel.progressbar.grid(row=1, column=0, columnspan=6, sticky='ew', pady=10)
        self.btn_panel.pack_buttons()

        self.table.rowconfigure(0)
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(2, weight=1)

    def pause_autoclicker(self):
        """Останавливает автокликер."""
        self.table.autoclicker.pause()
        block_buttons(self.btn_panel.pause_btn, self.btn_panel.skip_btn)
        activate_buttons(self.btn_panel.play_btn)

    def close_autoclicker(self):
        """Останавливает автокликер, закрывает браузер."""
        self.table.autoclicker.stop()
        self.btn_panel.play_btn['state'] = 'normal'
        self.btn_panel.pause_btn['state'] = 'disabled'

    def finish(self):
        """Операции после закрытия главного окна приложения."""
        try:
            wr = Writer(self.rd)
            wr.write()
        except PermissionError:
            if self.rd.wb:
                while True:
                    try:
                        open_warning()
                        wr = Writer(self.rd)
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

    def run_webdriver(self):
        if not self.table.autoclicker:
            self.table.autoclicker.run_webdriver()

    def run_reader(self):
        """Запуск чтения файла с рейсами."""
        try:
            self.rd.read()
            activate_buttons(self.btn_panel.start_btn)
            self.res_panel.routes_counter.set(self.rd.total)
            self.res_panel.remaining_counter.set(self.rd.total)
            self.make_dirs()
            if self.rd.report_type == 'НС':
                self.table.get_paths()

        except PermissionError as err:
            show_error("Открыт файл эксель с неучтенными рейсами. Закройте файл и перезапустите приложение!")
            Loger.enter_in_log(err)

        except Exception as err:
            show_error("Возникла ошибка! Перезапустите приложение!")
            Loger.enter_in_log(err)

    def show_load_window(self):
        """Открывает окно загрузки файла с рейсами."""
        self.load_window = LoadWindow(self)
        self.load_window.pack_items(1)


    def make_dirs(self):
        if 'скрины' not in os.listdir():
            os.mkdir('скрины')
        if self.rd.report_type not in os.listdir('скрины'):
            os.mkdir(fr'скрины\{self.rd.report_type}')

if __name__ == "__main__":
    app = App()
    app.mainloop()

