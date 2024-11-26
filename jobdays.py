import tkinter as tk
from datetime import datetime
from tkinter import messagebox

import bd_unit
from univunit import Table, Univunit
from ttkbootstrap import DateEntry

DB_MANAGER = bd_unit.DatabaseManager()


class JobDaysApp(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry("500x400")
        self.title("Добавление рабочих дней для получения FTE")
        self.hours_var = tk.IntVar()
        self.middle_days_var = tk.StringVar()

        self.result_label = None
        self.mdf = None

        self.create_widgets()

        self.table_fte = Table(self)
        self.table_fte.pack(expand=True, fill='both')
        self.read_all_data()

    def create_widgets(self):
        frame_item = tk.Frame(self)

        tk.Label(frame_item, text='Среднее число рабочих дней', bg="#333333", fg="white", font=("Arial", 12)).grid(
            row=1,
            sticky="e")
        self.mdf = tk.Entry(frame_item, textvariable=self.middle_days_var)
        self.mdf.grid(row=1, column=1, pady=20, padx=10, sticky="ew")

        tk.Button(frame_item, text='Сохранить',
                  command=self.save_middle_days,
                  bg="#FF3399", fg="white", font=("Arial", 12)).grid(row=1, column=2, pady=20, padx=10, sticky="ew")

        tk.Label(frame_item, text='Месяц из даты', bg="#333333", fg="white", font=("Arial", 16)).grid(row=2, sticky="e")

        self.month_year = DateEntry(master=frame_item, width=15, relief="solid", dateformat='%d-%m-%Y')
        self.month_year.grid(row=2, column=1, pady=10, sticky="ew", padx=10)

        tk.Label(frame_item, text='Часов', bg="#333333", fg="white", font=("Arial", 16)).grid(row=3,
                                                                                              sticky="e")
        tk.Entry(frame_item, textvariable=self.hours_var).grid(row=3, column=1, pady=20, padx=10, sticky="ew")

        # Место для отображения результата входа
        self.result_label = tk.Label(frame_item, bg="#333333", fg="blue", font=("Arial", 10))
        self.result_label.grid(row=4, columnspan=2, sticky="ew")

        frame_buttons = tk.Frame(frame_item)

        tk.Button(frame_buttons, text='Добавить запись',
                  command=self.save_days,
                  bg="#FF3399", fg="white", font=("Arial", 12)).pack(side=tk.LEFT, padx=10)
        tk.Button(frame_buttons, text='Удалить запись',
                  command=self.delete_rec,
                  bg="#FF3399", fg="white", font=("Arial", 12)).pack(side=tk.LEFT)
        frame_buttons.grid(row=5, columnspan=2, pady=10)
        frame_item.pack()

    def add_month_name_first(self, tuples_list):
        """
            Добавление в кортеж имени месяца
        """
        updated_list = []
        for item in tuples_list:
            date_str, work_days = item
            month_name = datetime.strptime(date_str, '%Y-%m-%d').strftime('%B')
            updated_list.append((month_name, date_str, work_days))
        return updated_list

    def read_all_data(self):

        self.mdf.config(state=tk.NORMAL)
        self.mdf.delete(0, tk.END)

        middle_days_var = DB_MANAGER.get_middle_fte()
        if middle_days_var != '0':
            self.mdf.insert(0, str(middle_days_var))

        # self.table_fte.configure_columns([{'name': 'Месяц'}, {'name': 'Дней'}, {'name': 'Период'}, {'name': 'Часы'}])
        self.table_fte.configure_columns([{'name': 'Месяц'}, {'name': 'Период'}, {'name': 'Часы'}])
        data = self.add_month_name_first(DB_MANAGER.read_all_table())
        # data.append(('Month':data.strftime('%B')))

        self.table_fte.populate_table(data)
        # Привязываем обновление переменной num_query при выборе строки
        self.table_fte.tree.bind('<<TreeviewSelect>>', self.update_num_query)

    def update_num_query(self, event):
        self.month_year.entry.delete(0, tk.END)
        cur_item = self.table_fte.tree.item(self.table_fte.tree.focus())  # Получаем данные выбранной строки
        if cur_item:
            values = cur_item.get('values', [])  # Получаем список значений
            if values:
                self.month_year.entry.insert(0, values[1])

    def delete_rec(self):
        try:
            DB_MANAGER.delete_record(self.month_year.entry.get())
            self.result_label.config(text=f"За период {self.month_year.entry.get()} удалена запись")
            self.read_all_data()
        except Exception as e:
            messagebox.showinfo('Ошибка с бд', f"Данные не сохранены: {e}")

    def save_days(self):
        try:
            days = self.hours_var.get()
            if days == 0:
                raise ValueError("Количество дней не может быть нулевым.")
            date_in = datetime.strptime(self.month_year.entry.get(), '%d-%m-%Y')
            date_in = Univunit.first_date_of_month(date_in)
            DB_MANAGER.insert_data([(self.hours_var.get(), date_in)])
            self.result_label.config(text=f"{date_in}, число:{self.hours_var.get()} добавлены")
            self.read_all_data()
        except ValueError as ve:
            messagebox.showinfo('Ошибка ввода', str(ve))
        except Exception as e:
            messagebox.showinfo('Ошибка с бд', f"Данные не сохранены: {e}")

    def save_middle_days(self):
        try:
            days = self.middle_days_var.get()
            if days == 0:
                days = 164
            DB_MANAGER.save_middle_fte(days)
            self.result_label.config(text=f"Среднее:{days} добавлено")
            self.read_all_data()
        except ValueError as ve:
            messagebox.showinfo('Ошибка ввода', str(ve))
        except Exception as e:
            messagebox.showinfo('Ошибка с бд', f"Данные не сохранены: {e}")
