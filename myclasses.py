from ttkwidgets import CheckboxTreeview
from tkinter.ttk import Style, Entry, Label


class MyCheckboxTreeview(CheckboxTreeview): # фильтрацию надо доделать
    def __init__(self, filter_window, parent, **kwargs):
        super().__init__(parent, **kwargs)
        self.window = filter_window
        self.style = Style()
        self.style.configure('Checkbox.Treeview', font=('Arial', 10), rowheight=25)
        self.bind("<Button-1>", self._box_click)
        self.colname = filter_window.colname
        self.main_table = filter_window.app.table
        self.values = []
        self.size = 0
        self.empty_val = False
        self.fill_out_table()

    def overwrite_table(self):
        self.values = []
        self.delete(*self.get_children())
        self.fill_out_table()

    def set_btn_state(self):
        self.window.ok_btn['state'] = 'disabled' if not self.any_checked() else 'normal'

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
        return all([self.tag_has("checked", i) for i in self.get_children()])

    def any_checked(self):
        return any([self.tag_has("checked", i) for i in self.get_children()])

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

        self.check_all()

    def filters_has(self):
        return any(list(self.main_table.filters.values()))


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
            self.set_btn_state()


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

    def enter(self, event=None):
        current_value = self.table.item(self.item)["values"][self.col]
        value = self.get().strip()
        self.table.set(self.item, self.col, value)
        self.destroy()

        if current_value and not value:
            self.app.res_panel.subtract_route()
        elif not current_value and value:
            self.app.res_panel.add_route()

        self.table.editing_cell = None



