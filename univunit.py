import tkinter as tk

import ttkbootstrap as ttk


class Table(tk.Frame):
    def __init__(self, parent=None, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.scrollbar = ttk.Scrollbar(self)
        self.style = ttk.Style()

        # Устанавливаем высоту строки
        self.style.configure('Treeview', rowheight=25)

        self.tree = ttk.Treeview(self, yscrollcommand=self.scrollbar.set)
        self.create_widgets()

    def create_widgets(self):
        # Создание скроллбара

        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Создание таблицы

        self.tree.pack(expand=True, fill='both')

        # Связывание скроллбара с таблицей
        self.scrollbar.config(command=self.tree.yview)

    def configure_columns(self, columns_):

        # self.tree.column('#0', width=0, stretch=tk.NO)
        # Настройка столбцов
        anchor_ = tk.CENTER
        width_ = 100
        column = 'нет наименования'
        self.tree['columns'] = [items["name"] for items in columns_]
        self.tree.column('#0', width=0, stretch=tk.NO)
        for item in columns_:
            for key, val in item.items():
                if key == 'anchor':
                    anchor_ = val
                elif key == 'width':
                    width_ = val
                elif key == 'name':
                    column = val

            # print(f"column={column}, anchor={anchor_}, width={width_}")
            self.tree.column(column, anchor=anchor_, width=width_)
            self.tree.heading(column, text=column, anchor=tk.CENTER)

    def populate_table(self, data):
        self.style.configure('Treeview.Heading', background='#F79646', foreground='white',
                             font=('Helvetica', 11, 'bold'))
        # Очистка таблицы перед заполнением
        for i in self.tree.get_children():
            self.tree.delete(i)
        # Добавление данных
        # for item in data:
        #     self.tree.insert('', 'end', values=item)

        # Чередуем цвета строк
        for i, item in enumerate(data):
            if i % 2 == 0:
                self.tree.insert('', 'end', values=item, tags=('evenrow',))
            else:
                self.tree.insert('', 'end', values=item, tags=('oddrow',))

        # Устанавливаем цвета для четных и нечетных строк
        self.tree.tag_configure('evenrow', background='#DAEEF3')
        self.tree.tag_configure('oddrow', background='#B7DEE8')

