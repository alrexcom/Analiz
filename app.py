from datetime import datetime
import tkinter as tk
from tkinter import (filedialog, font, messagebox)

import ttkbootstrap as ttk
from ttkbootstrap import DateEntry

import univunit
from reports import (get_data_report, names_reports)
from univunit import Table, Univunit
import bd_unit

themes = ['cosmo', 'flatly', 'litera', 'minty', 'lumen', 'sandstone',
          'yeti', 'pulse', 'united', 'morph', 'journal', 'darkly',
          'superhero', 'solar', 'cyborg', 'vapor', 'simplex', 'cerculean']

DB_MANAGER = bd_unit.DatabaseManager('test.db')


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
        JobDaysApp(self)
        
    def create_new_window(self):
        JobDaysApp(self)

    def create_fte_window(self):
        CalcApp(self)


class CalcApp(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry("400x200")
        self.title("Посчитаем фте")
        self.hours_var = tk.IntVar()
        self.var_fte = tk.StringVar()
        self.fte_on_month = tk.IntVar()
        self.Itog = None
        self.Hours = None
        self.Fte_month = None
        self.create_widgets()

    def create_widgets(self):
        frame_item = tk.Frame(self)

        tk.Label(frame_item, text='Часы', bg="#333333", fg="white", font=("Arial", 12)).grid(row=1,
                                                                                             sticky="e")
        self.Hours = tk.Entry(frame_item, textvariable=self.hours_var)
        self.Hours.grid(row=1, column=1, pady=5, padx=10, sticky="ew")

        tk.Label(frame_item, text='FTE в месяц', bg="#333333", fg="white", font=("Arial", 12)).grid(row=2, sticky="e")

        self.Fte_month = tk.Entry(frame_item, textvariable=self.fte_on_month)
        self.Fte_month.grid(row=2, column=1, pady=5, padx=10, sticky="ew")

        self.Fte_month.config(state=tk.NORMAL)
        self.Fte_month.delete(0, tk.END)
        self.Fte_month.insert(0, '164')

        tk.Button(frame_item, text='Посчитать FTE',
                  command=self.get_fte,
                  bg="#FF3399", fg="white", font=("Arial", 12)).grid(row=3, column=0, pady=10, padx=10, sticky="ew")

        tk.Button(frame_item, text='Посчитать часы',
                  command=self.get_hours,
                  bg="#FF3399", fg="white", font=("Arial", 12)).grid(row=3, column=1, pady=10, padx=10, sticky="ew")

        tk.Label(frame_item, text='FTE', bg="#333333", fg="white", font=("Arial", 12)).grid(row=4, sticky="e")
        self.Itog = tk.Entry(frame_item, textvariable=self.var_fte)
        self.Itog.grid(row=4, column=1, pady=5, padx=10, sticky="ew")

        self.Itog.config(state=tk.NORMAL)
        self.Itog.delete(0, tk.END)
        self.Itog.insert(0, '0')

        frame_item.pack(fill=tk.X, padx=50, pady=15)

    def get_hours(self):
        hours = univunit.calc_hours(fte_on_month=self.fte_on_month.get(), fte=self.var_fte.get())
        self.Hours.config(state=tk.NORMAL)
        self.Hours.delete(0, tk.END)
        self.Hours.insert(0, hours)

    def get_fte(self):
        itogo = univunit.calc_fte(fte_on_month=self.fte_on_month.get(), hours=self.hours_var.get())
        self.Itog.config(state=tk.NORMAL)
        self.Itog.delete(0, tk.END)
        self.Itog.insert(0, itogo)


class JobDaysApp(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry("500x400")
        self.title("Добавление рабочих дней для получения FTE")
        self.days_var = tk.IntVar()

        self.result_label = None

        self.create_widgets()

        self.table_fte = Table(self)
        self.table_fte.pack(expand=True, fill='both')
        self.read_all_data()

    def create_widgets(self):
        frame_item = tk.Frame(self)

        tk.Label(frame_item, text='Число рабочих дней', bg="#333333", fg="white", font=("Arial", 16)).grid(row=1,
                                                                                                           sticky="e")
        tk.Entry(frame_item, textvariable=self.days_var).grid(row=1, column=1, pady=20, padx=10, sticky="ew")

        tk.Label(frame_item, text='Месяц из даты', bg="#333333", fg="white", font=("Arial", 16)).grid(row=2, sticky="e")

        self.month_year = DateEntry(master=frame_item, width=15, relief="solid", dateformat='%d-%m-%Y')
        self.month_year.grid(row=2, column=1, pady=10, sticky="ew", padx=10)

        # Место для отображения результата входа
        self.result_label = tk.Label(frame_item, bg="#333333", fg="blue", font=("Arial", 10))
        self.result_label.grid(row=3, columnspan=2, sticky="ew")

        frame_buttons = tk.Frame(frame_item)

        tk.Button(frame_buttons, text='Добавить запись',
                  command=self.save_days,
                  bg="#FF3399", fg="white", font=("Arial", 12)).pack(side=tk.LEFT, padx=10)
        tk.Button(frame_buttons, text='Удалить запись',
                  command=self.delete_rec,
                  bg="#FF3399", fg="white", font=("Arial", 12)).pack(side=tk.LEFT)
        frame_buttons.grid(row=4, columnspan=2, pady=10)
        frame_item.pack()

    def add_month_name_first(self, tuples_list):
        """
            Добавление в кортеж имени месяца
        """
        updated_list = []
        for item in tuples_list:
            month_number, date_str, work_days = item
            month_name = datetime.strptime(date_str, '%Y-%m-%d').strftime('%B')
            updated_list.append((month_name, month_number, date_str, work_days))
        return updated_list

    def read_all_data(self):
        self.table_fte.configure_columns([{'name': 'Месяц'}, {'name': 'Дней'}, {'name': 'Период'}, {'name': 'FTE'}])

        data = self.add_month_name_first(DB_MANAGER.read_all_table())
        # data.append(('Month':data.strftime('%B')))

        self.table_fte.populate_table(data)

    def delete_rec(self):
        try:
            date_in = Univunit.first_date_of_month(self.month_year.entry.get())
            DB_MANAGER.delete_record(date_in)
            self.result_label.config(text=f"За период {date_in} удалена запись")
            self.read_all_data()
        except Exception as e:
            messagebox.showinfo('Ошибка с бд', f"Данные не сохранены: {e}")

    def save_days(self):
        try:
            days = self.days_var.get()
            if days == 0:
                raise ValueError("Количество дней не может быть нулевым.")
            date_in = Univunit.first_date_of_month(self.month_year.entry.get())
            DB_MANAGER.insert_data([(self.days_var.get(), date_in)])
            self.result_label.config(text=f"{date_in}, число:{self.days_var.get()} добавлены")
            self.read_all_data()
        except ValueError as ve:
            messagebox.showinfo('Ошибка ввода', str(ve))
        except Exception as e:
            messagebox.showinfo('Ошибка с бд', f"Данные не сохранены: {e}")


if __name__ == '__main__':
    App("Анализ отчётов", (800, 500), 'yeti').mainloop()
