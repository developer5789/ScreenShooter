from tkinter import messagebox


def show_error(message):
    messagebox.showerror('Ошибка', message)

def last_screen_message(parent):
    messagebox.showinfo('Уведомление', 'Вы дошли до крайнего скриншота в таблице!', parent=parent)


def show_inf():
    messagebox.showinfo('Уведомление', 'Ура!Разбор рейсов закончен!')


def last_item_message():
    messagebox.showinfo('Уведомление', 'Вы дошли до крайнего элемента в таблице!')


def open_warning():
    messagebox.showwarning(title="Предупреждение", message="У вас открыт файл эксель с неучтенными рейсами."
                                                           "Для корректного завершения программы закройте его и нажмите 'Ок'")
def askyesnocancel(parent):
    return messagebox.askyesnocancel(title='Удаление скриншота', message="Скриншот будет удалён.\n(Да-0 в столбце 'скрин',"
                                                                  "нет-пустое значение.)", parent=parent)
