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
            date_in = datetime.strptime(self.month_year.entry.get(), '%d-%m-%Y')
            date_in = Univunit.first_date_of_month(date_in)
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
            date_in = datetime.strptime(self.month_year.entry.get(),'%d-%m-%Y')
            date_in = Univunit.first_date_of_month(date_in)
            DB_MANAGER.insert_data([(self.days_var.get(), date_in)])
            self.result_label.config(text=f"{date_in}, число:{self.days_var.get()} добавлены")
            self.read_all_data()
        except ValueError as ve:
            messagebox.showinfo('Ошибка ввода', str(ve))
        except Exception as e:
            messagebox.showinfo('Ошибка с бд', f"Данные не сохранены: {e}")
