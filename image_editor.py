
import tkinter as tk
from PIL import Image, ImageTk


class ImageEditor(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.app = args[0]
        self.table = self.app.table
        self.current_item = str(self.table.current_item)
        self.canvas = tk.Canvas(self, width=1950, height=1030, bg="#F5F5F5")
        self.image_path = self.table.item(self.current_item)['values'][8]
        self.image = None
        self.image_tk = None
        self.pack_items()
        self.open_image(self.image_path)

    def pack_items(self):
        self.canvas.grid(row=0, column=0)

    def open_image(self, filepath):
        if filepath:
            self.image_path = filepath
            image = Image.open(filepath)
            self.image_tk = ImageTk.PhotoImage(image)
            img_name = filepath.split("\\")[-1]
            self.title(f"Режим просмотра скринов: {img_name}")
            self.canvas.create_image(0, 0, anchor=tk.NW, image=self.image_tk)





