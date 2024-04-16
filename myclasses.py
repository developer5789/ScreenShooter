from ttkwidgets import CheckboxTreeview
from tkinter.ttk import Style, Entry

class MyCheckboxTreeview(CheckboxTreeview): # фильтрацию надо доделать
    def __init__(self, filter_window, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.style = Style()
        self.style.configure('Checkbox.Treeview', font=('Arial', 10), rowheight=25)
        self.bind("<Button-1>", self._box_click)
        self.colname = filter_window.colname
        self.main_table = filter_window.app.table
        self.values = []
        self.size = 0
        self.empty_val = False
        self.fill_out_table()

    def get_checked_val(self):
        res = []
        for item in self.get_checked():
            if item != '0':
                val = self.item(item)['text']
                res.append(val) if val != 'Пустые' else res.append('')
        return res


    def return_initial_state(self):
        if self.size != 0:
            for i in range(self.size + 1):
                self.move(str(i), '', i)


    def filter(self, value):
        matching_items = ['0']
        for i in range(1, self.size + 1):
            item_value = self.item(str(i))['text']
            if value in item_value and item_value != 'Пустые':
                matching_items.append(str(i))
            else:
                self.detach(str(i))

        if len(matching_items) > 1:
            for i, item in enumerate(matching_items):
                self.move(item, '', i)
        else:
            self.detach('0')

    def all_checked(self):
        return all([self.tag_has("checked", str(i)) for i in range(1, self.size + 1)])

    def read_main_table(self):
        col_index = self.main_table['columns'].index(self.colname)
        empty_val = False
        for item in self.main_table.get_children():
            value = self.main_table.item(item)['values'][col_index]
            value = str(value)
            if value not in self.values and value:
                self.values.append(value)
            elif not value:
                empty_val = True

        self.values.sort()
        if empty_val:
            self.values.append('')


    def fill_out_table(self):
        self.read_main_table()
        self.insert("", "end", '0', text='(Выделить все)')

        for i, val in enumerate(self.values):
            if val:
                self.insert("", "end", i + 1, text=val)
            else:
                self.insert("", 'end', i + 1, text='Пустые')


        self.size = len(self.get_children()) - 1
        if not self.size:
            self.delete('0')
            return

        mark_values = self.main_table.filters[self.colname]
        if mark_values is not None:
            self.check_items(mark_values)
        else:
            self.check_all()

    def check_items(self, values):
        for item in self.get_children():
            if item != '0':
                val = self.item(item)['text'] if self.item(item)['text'] else ''
                if val in values:
                    self.change_state(item, 'checked')

        if self.all_checked():
            self.change_state('0', 'checked')
        else:
            self.change_state('0', 'tristate')

    def _box_click(self, event):
        x, y, widget = event.x, event.y, event.widget
        elem = widget.identify("element", x, y)
        if "image" in elem:
            # a box was clicked
            item = self.identify_row(y)
            if item == '0':
                self.check_all() if self.tag_has("unchecked", item) else self.uncheck_all()
            else:
                checked_items = [item for item in self.get_checked() if item != '0']
                numb_checked = len(checked_items)
                if self.tag_has("unchecked", item) and numb_checked + 1  == self.size:
                    self.check_all()
                elif self.tag_has("unchecked", item) and numb_checked + 1 < self.size:
                    self.change_state('0', 'tristate')
                    self.change_state(item, "checked")
                elif self.tag_has("checked", item) and numb_checked - 1 == 0:
                    self.change_state('0', 'unchecked')
                    self.change_state(item, "unchecked")
                elif self.tag_has("checked", item) and numb_checked - 1 < self.size:
                    self.change_state('0', 'tristate')
                    self.change_state(item, "unchecked")


class TableEntry(Entry):
    def __init__(self, item, col, box: list, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.item = item
        self.col = int(col[1:]) - 1
        self.table = args[0].table
        self.app = args[0]
        self.bind('<Return>', self.enter)
        self.bind('<FocusOut>', self.enter)
        self.table.bind('<MouseWheel>', self.del_entry)
        self.app.table.bind('<Configure>', self.del_entry)
        self.app.scroll.bind('<MouseWheel>', self.del_entry)
        self.app.scroll.bind('<ButtonPress>', self.del_entry)
        self.place(x=box[0],
                   y=box[1],
                   width=box[2],
                   height=box[3])
        self.focus()

    def del_entry(self, event=None):
        self.table.unbind('<MouseWheel>')
        self.app.table.unbind('<Configure>')
        self.app.scroll.unbind('<MouseWheel>')
        self.app.scroll.unbind('<ButtonPress>')
        self.destroy()

    def enter(self, event):
        value = self.get()
        self.table.set(self.item, self.col, value)
        self.destroy()


