from tkinter.ttk import Scrollbar

from reports import *
# from tkscrolledframe import ScrolledFrame
import tkinter as tk
from tkinter import (filedialog, font)

import ttkbootstrap as ttk

# from ttkbootstrap.scrolled import ScrolledFrame

themes = ['cosmo', 'flatly', 'litera', 'minty', 'lumen', 'sandstone',
          'yeti', 'pulse', 'united', 'morph', 'journal', 'darkly',
          'superhero', 'solar', 'cyborg', 'vapor', 'simplex', 'cerculean']


class Component(ttk.Frame):
    def __init__(self, parent, label_text, text_value):
        super().__init__(master=parent)

        # grid
        self.rowconfigure(0, weight=1)
        self.columnconfigure((0, 1), weight=1, uniform='a')
        ttk.Label(self,
                  text=label_text,
                  wraplength=400,
                  font=("Helvetica", 12),
                  justify=tk.RIGHT
                  ).grid(
            padx=5, row=0, column=0, sticky=tk.W)
        ttk.Label(self,
                  text=text_value,
                  background="pink",
                  border=2,
                  relief=tk.SUNKEN,
                  anchor=tk.CENTER,
                  font=("Helvetica", 14),
                  width=10).grid(row=0, column=1)

        self.pack(expand=1, fill=tk.BOTH)


class Table(tk.Frame):
    def __init__(self, parent=None, *args, **kwargs):
        tk.Frame.__init__(self, parent, *args, **kwargs)
        self.scrollbar = ttk.Scrollbar(self)
        self.tree = ttk.Treeview(self, yscrollcommand=self.scrollbar.set)
        self.create_widgets()

    def create_widgets(self):
        # Создание скроллбара

        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Создание таблицы

        self.tree.pack(expand=True, fill='both')

        # Связывание скроллбара с таблицей
        self.scrollbar.config(command=self.tree.yview)

    def configure_columns(self, columns):
        # Настройка столбцов
        self.tree['columns'] = columns
        self.tree.column('#0', width=0, stretch=tk.NO)
        for col in columns:
            self.tree.column(col, anchor=tk.CENTER, width=100)
            self.tree.heading(col, text=col)

    def populate_table(self, data):
        # Очистка таблицы перед заполнением
        for i in self.tree.get_children():
            self.tree.delete(i)
        # Добавление данных
        for item in data:
            self.tree.insert('', 'end', values=item)


class App(tk.Tk):
    def __init__(self, title, size, _theme):
        super().__init__()
        self.name_report = tk.StringVar()
        # Menu strip
        self.menu_frame = ttk.Frame(self, padding=5)
        self.ds = ttk.DateEntry(
            master=self.menu_frame,
            width=15,
            relief="solid",
            dateformat='%d-%m-%Y'
        )
        self.dp = ttk.DateEntry(self.menu_frame,
                                width=15,
                                relief="solid",
                                dateformat='%d-%m-%Y')

        self.one_hour_fte = tk.IntVar()
        self.fte = ttk.Entry(self.menu_frame, width=5, font=("Calibri", 12),
                             textvariable=self.one_hour_fte)
        self.export_excell_var = tk.IntVar()

        self.style = ttk.Style(theme=_theme)
        self.title(title)
        self['background'] = '#EBEBEB'
        self.geometry(f"{size[0]}x{size[1]}")
        # self.minsize(size[0], size[1])
        # self.maxsize(size[0], size[1])
        self.bold_font = font.Font(family='Helvetica', size=13, weight='bold')

        self.create_widgets()

        self.table = Table(self)
        self.table.pack(expand=True, fill='both')

    def create_widgets(self):
        cmb_frame = ttk.Frame(self, padding=1)

        cmb = ttk.Combobox(cmb_frame, values=[items["name"] for items in reports],
                           state="readonly",
                           height=4,
                           font=("Calibri", 12)
                           , textvariable=self.name_report
                           # , textvariable=name_report
                           )
        cmb.set('Выбор из списка отчетов')
        # cmb.bind('<<ComboboxSelected>>', cmb_function)
        cmb.pack(side=tk.LEFT, expand=True, fill=tk.X)

        cmb_frame.pack(fill=tk.X, pady=10, padx=10)

        self.ds.pack(side=tk.LEFT, padx=10)
        self.ds.entry.delete(0, tk.END)
        self.ds.entry.insert(0, '01-01-2024')

        self.dp.pack(side=tk.LEFT, padx=10)

        ttk.Label(self.menu_frame, text="FTE:", width=5,
                  border=2, font=("Calibri", 12, 'bold'),
                  background="gray", foreground='white').pack(side=tk.LEFT)

        self.fte.delete(0, tk.END)
        self.fte.insert(0, '164')
        self.fte.pack(side=tk.LEFT, padx=10)

        ttk.Checkbutton(self.menu_frame, text='Экспорт в Excel',
                        variable=self.export_excell_var,
                        onvalue=1,
                        offvalue=0).pack(side=tk.LEFT, padx=5)

        btn_clear = ttk.Button(self.menu_frame, text='Очистить',
                               command=self.btn_clear_click,
                               width=10)
        btn_clear.pack(padx=5, side=tk.LEFT)

        btn_go = ttk.Button(self.menu_frame, text='Открыть',
                            command=self.btn_go_click,
                            width=10)
        btn_go.pack(side=tk.LEFT)

        self.menu_frame.pack(fill=tk.X)

    def btn_go_click(self):
        date_begin = self.ds.entry.get()
        date_end = self.dp.entry.get()
        fte = self.one_hour_fte.get()
        export = self.export_excell_var.get()
        name_report = self.name_report.get()

        file_name = filedialog.askopenfilename()
        report_data = get_report_test(name_report=name_report, filename=file_name)

        num_report = report_data[0]
        param = {'df': report_data[1],
                 'fte': fte,
                 'reportnumber': num_report,
                 'date_end': date_end,
                 'date_begin': date_begin,
                 'export_excell': export}
        # https: // www.youtube.com / watch?v = mop6g - c5HEY & list = PLZHIeS5WrW4IhlHiQq9fmlTu4F_YIc4ek & index = 10
        fr = get_data_test(**param)
        # data = [
        #     ('Алексей', '30', 'Москва'),
        #     ('Ирина', '25', 'Санкт-Петербург'),
        #     # Добавьте больше данных
        # ]
        columns = list()
        data = list()
        if num_report == 1:
            columns = fr['columnsname']
            for item in fr['values']:
                data.append(tuple(fr['values'][item]))
        elif (num_report == 3) | (num_report == 2):
            columns = ["Наименование", "Значение"]
            for item in fr.items():
                key, val = item
                data.append((key, val))

        self.table.configure_columns(columns)
        self.table.populate_table(data)
        self.update_window_size()

    def update_window_size(self):
        # Рассчитываем общую ширину столбцов
        total_width = sum(self.table.tree.column(col, 'width') for col in self.table.tree['columns'])
        # Обновляем размер окна
        self.geometry(f'{total_width}x400')

    def btn_clear_click(self):
        for i in self.table.tree.get_children():
            self.table.tree.delete(i)
        # print(result_frame.winfo_children())
        # for item in self.winfo_children():
        #     item.destroy()


app = App("Анализ отчётов", (800, 500), 'yeti')
app.mainloop()
