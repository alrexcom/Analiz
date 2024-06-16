import tkinter as tk
from tkinter import (filedialog, font, messagebox)

import ttkbootstrap as ttk

from univunit import Table
from reports import *

themes = ['cosmo', 'flatly', 'litera', 'minty', 'lumen', 'sandstone',
          'yeti', 'pulse', 'united', 'morph', 'journal', 'darkly',
          'superhero', 'solar', 'cyborg', 'vapor', 'simplex', 'cerculean']


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

        self.fte_frame = ttk.Frame(self.menu_frame)

        self.fte_frame.pack_forget()
        self.one_hour_fte = tk.IntVar()

        self.fte = ttk.Entry(self.fte_frame, width=5, font=("Calibri", 12),
                             textvariable=self.one_hour_fte)

        # self.export_excell_var = tk.IntVar()
        # self.export=ttk.Checkbutton(self.menu_frame, text='Экспорт в Excel',
        #                 variable=self.export_excell_var,
        #                 onvalue=1,
        #                 offvalue=0)

        self.style = ttk.Style(theme=_theme)
        self.title(title)
        self['background'] = '#EBEBEB'
        self.geometry(f"{size[0]}x{size[1]}")
        # self.minsize(size[0], size[1])
        # self.maxsize(size[0], size[1])
        self.bold_font = font.Font(family='Helvetica', size=13, weight='bold')
        # self.label_info = tk.Label(self, text="Информация", font=("Calibri", 12), bg="red")
        # self.text = tk.Text(bg="darkgreen", fg="white", wrap=WORD, height=2)
        # self.text.tag_config('title', justify=tk.LEFT)
        # # Скрываю при открытии
        # self.text.pack_forget()

        self.create_widgets()

        self.table = Table(self)
        self.table.pack(expand=True, fill='both')

    def create_widgets(self):
        cmb_frame = ttk.Frame(self, padding=1)

        cmb = ttk.Combobox(cmb_frame, values=[items["name"] for items in reports],
                           state="readonly",
                           height=4,
                           font=("Calibri", 12),
                           textvariable=self.name_report
                           )
        cmb.set('Выбор из списка отчетов')
        cmb.bind('<<ComboboxSelected>>', self.cmb_function)
        cmb.pack(side=tk.LEFT, expand=True, fill=tk.X)

        cmb_frame.pack(fill=tk.X, pady=10, padx=10)

        self.ds.pack(side=tk.LEFT, padx=10)
        self.ds.entry.delete(0, tk.END)
        self.ds.entry.insert(0, '01-01-2024')

        self.dp.pack(side=tk.LEFT, padx=10)

        ttk.Label(self.fte_frame, text="FTE:", width=5, anchor=tk.CENTER,
                  border=2, font=("Calibri", 12, 'bold'),
                  background='#B7DEE8', foreground='white').pack(side=tk.LEFT)

        self.fte.delete(0, tk.END)
        self.fte.insert(0, '164')
        self.fte.pack(side=tk.LEFT, padx=10)

        # ttk.Checkbutton(self.menu_frame, text='Экспорт в Excel',
        #                 variable=self.export_excell_var,
        #                 onvalue=1,
        #                 offvalue=0).pack(side=tk.LEFT, padx=5)

        btn_go = ttk.Button(self.menu_frame, text='Открыть',
                            command=self.btn_go_click,
                            width=10)
        btn_go.pack(side=tk.LEFT)

        self.menu_frame.pack(fill=tk.X)

        # self.text.pack(fill=tk.X)

    def cmb_function(self, event):
        num_report = event.widget.current() + 1
        if num_report == 1:
            self.fte_frame.pack(side=tk.LEFT, padx=5)
        else:
            self.fte_frame.pack_forget()

    def btn_go_click(self):
        date_begin = self.ds.entry.get()
        date_end = self.dp.entry.get()
        fte = self.one_hour_fte.get()
        # export = bool(self.export_excell_var.get())
        name_report = self.name_report.get()
        file_name = ''
        try:
            file_name = filedialog.askopenfilename()
            report_data = get_reports(name_report=name_report, filename=file_name)

            data_begin = pd.to_datetime(date_begin, dayfirst=True).strftime('%Y-%m-%d')
            data_end = pd.to_datetime(date_end, dayfirst=True).strftime('%Y-%m-%d')

            num_report = report_data[0]
            param = {'df': report_data[1],
                     'fte': fte,
                     'reportnumber': num_report,
                     'date_end': data_end,
                     'date_begin': data_begin,
                     # 'export_excell': export
                     }
            # https: // www.youtube.com / watch?v = mop6g - c5HEY & list = PLZHIeS5WrW4IhlHiQq9fmlTu4F_YIc4ek & index=10
            fr = get_data_report(**param)

            self.table.configure_columns(fr['columns'])
            self.table.populate_table(fr['data'])
            self.update_window_size()
        except Exception as e:
            # self.label_info.config(text=f"Не смог открыть файл {file_name}{e}")
            # self.text.pack(fill=tk.X)
            # self.text.insert(5.0, f"Не смог открыть файл {file_name}{e}")
            messagebox.showinfo("Ошибка!", f"Не смог открыть файл {file_name}{e}")

    def update_window_size(self):
        # Рассчитываем общую ширину столбцов
        total_width = sum(self.table.tree.column(col, 'width') for col in self.table.tree['columns'])
        row_count = len(self.table.tree.get_children())
        height = (row_count * 25)+180
        # Обновляем размер окна
        self.geometry(f'{total_width}x{height}')


app = App("Анализ отчётов", (800,500), 'yeti')
app.mainloop()
