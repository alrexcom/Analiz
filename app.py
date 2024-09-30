from datetime import datetime
import tkinter as tk
from tkinter import (filedialog, font, messagebox)

import ttkbootstrap as ttk
from ttkbootstrap import DateEntry
import lukoil_query as lk
import jobdays as jdays
import calc
from reports import (get_data_report, names_reports)
from univunit import Table, Univunit
import bd_unit

DB_MANAGER = bd_unit.DatabaseManager('test.db')

themes = ['cosmo', 'flatly', 'litera', 'minty', 'lumen', 'sandstone',
          'yeti', 'pulse', 'united', 'morph', 'journal', 'darkly',
          'superhero', 'solar', 'cyborg', 'vapor', 'simplex', 'cerculean']


def prompt_file_selection():
    try:
        return filedialog.askopenfilename()
    except Exception as e:
        messagebox.showinfo("Ошибка!", f"Не удалось выбрать файл: {e}")
        return None


class App(tk.Tk):
    def __init__(self, title, size, _theme):
        super().__init__()
        self.style = ttk.Style(theme=_theme)
        self['background'] = '#EBEBEB'
        self.bold_font = font.Font(family='Helvetica', size=13, weight='bold')
        self.title(title)

        self.geometry(f"{size[0]}x{size[1]}")
        # self.one_hour_fte = None
        self.fte_frame = None
        self.ds = None
        self.dp = None
        # self.fte = None
        self.middle_fte = None
        self.name_report = tk.StringVar()

        self.create_widgets()
        self.create_menu()
        # self.table = Table(self)
        # self.table.pack(expand=True, fill='both')

    def create_table(self):
        if not hasattr(self, 'table'):
            self.table = Table(self)
            self.table.pack(expand=True, fill='both')

    def create_widgets(self):
        # Menu strip
        current_date = datetime.now()

        menu_frame = tk.Frame(self, padx=5, pady=5)
        self.ds = DateEntry(
            master=menu_frame,
            width=15,
            relief="solid",
            dateformat='%d-%m-%Y'
        )

        self.fte_frame = tk.Frame(menu_frame)
        self.fte_frame.pack_forget()
        self.one_hour_fte = tk.IntVar()
        self.fte = ttk.Entry(self.fte_frame, width=5, font=("Calibri", 12),
                             textvariable=self.one_hour_fte)

        cmb_frame = ttk.Frame(self, padding=1)

        cmb = ttk.Combobox(cmb_frame, values=names_reports(),
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

        self.ds.entry.insert(0, Univunit.get_first_day_of_quarter(current_date=current_date))

        self.dp = DateEntry(menu_frame,
                            width=15,
                            relief="solid",
                            dateformat='%d-%m-%Y')

        self.dp.pack(side=tk.LEFT, padx=10)
        self.dp.bind('<<DateEntrySelected>>', self.on_date_change)
        self.dp.entry.delete(0, tk.END)
        self.dp.entry.insert(0, Univunit.get_last_day_of_current_month(date_format='%d-%m-%Y'))

        ttk.Label(self.fte_frame, text="FTE:", width=5, anchor=tk.CENTER,
                  border=2, font=("Calibri", 12, 'bold'),
                  background='#B7DEE8', foreground='white').pack(side=tk.LEFT)

        self.middle_fte = tk.IntVar()
        ttk.Checkbutton(self.fte_frame, text="По среднему", variable=self.middle_fte,
                        onvalue=1, offvalue=0,
                        command=self.middle_click).pack(side=tk.RIGHT)

        self.set_fte_from_db()
        self.fte.pack(side=tk.LEFT, padx=10)

        btn_go = ttk.Button(menu_frame, text='Открыть',
                            command=self.btn_go_click,
                            width=10)
        btn_go.pack(side=tk.LEFT)

        menu_frame.pack(fill=tk.X)

    def on_date_change(self):
        self.set_fte_from_db()

    def cmb_function(self, event):
        num_report = event.widget.current() + 1
        if num_report == 1:
            self.toggle_fte_frame(True)
            self.set_fte_from_db()
        else:
            self.toggle_fte_frame(False)

    def toggle_fte_frame(self, show):
        if show:
            self.fte_frame.pack(side=tk.LEFT, padx=5)
        else:
            self.fte_frame.pack_forget()

    def btn_go_click(self):
        file_name = prompt_file_selection()
        if file_name:
            self.process_file(file_name)

    def process_file(self, file_name):
        try:
            param = self.get_params(file_name)
            fr = get_data_report(**param)
            self.create_table()
            self.table.configure_columns(fr['columns'])
            self.table.populate_table(fr['data'])
            self.update_window_size()
        except Exception as e:
            messagebox.showinfo("Ошибка!", f"Не смог обработать файл {file_name}: {e}")

    def middle_click(self):
        self.set_fte_from_db()
        # print(self.middle_fte.get())

    # def on_date_change(self, event):
    #     """событие на изменение dp"""
    #     self.set_fte_from_db()

    def set_fte_from_db(self):
        self.fte.config(state=tk.NORMAL)
        self.fte.delete(0, tk.END)
        data_po = self.dp.entry.get()
        if self.middle_fte.get() == 1:
            self.fte.insert(0, '164')
            self.fte.config(state='readonly = FALSE')
        else:
            self.fte.insert(0, DB_MANAGER.read_one_rec(data_po))
            self.fte.config(state='readonly')

    def get_params(self, file_name):
        # print(fte)
        return {
            'filename': file_name,
            'name_report': self.name_report.get(),
            'fte': self.one_hour_fte.get(),
            'middle_fte': self.middle_fte.get(),
            'date_end': self.dp.entry.get(),
            'date_begin': self.ds.entry.get(),
        }

    def update_window_size(self):
        # Рассчитываем общую ширину столбцов
        total_width = sum(self.table.tree.column(col, 'width') for col in self.table.tree['columns'])
        row_count = len(self.table.tree.get_children())
        height = (row_count * 25) + 180
        # Обновляем размер окна
        self.geometry(f'{total_width}x{height}')

    def create_menu(self):
        """
        Создание панели меню
        :return:
        """
        menubar = tk.Menu(self)

        # Создание выпадающего меню "Файл"
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Ввод трудозатрат Лукойл заявок", command=self.create_new_query)
        file_menu.add_command(label="Ввести рабочие дни", command=self.create_new_window)
        file_menu.add_command(label="Калькулятор  FTE", command=self.create_fte_window)
        file_menu.add_separator()  # Добавление разделителя
        file_menu.add_command(label="Выход", command=self.quit)  # Добавление команды "Выход"

        # Добавление выпадающего меню "Файл" в панель меню
        menubar.add_cascade(label="Файл", menu=file_menu)
        self.config(menu=menubar)

    def create_new_query(self):
        lk.LukoilQueries(self)

    def create_new_window(self):
        jdays.JobDaysApp(self)

    def create_fte_window(self):
        calc.CalcApp(self)


if __name__ == '__main__':
    App("Анализ отчётов", (800, 500), 'yeti').mainloop()
