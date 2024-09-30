import tkinter as tk
from datetime import datetime
from tkinter import messagebox

from ttkbootstrap import DateEntry

import bd_unit
from univunit import Table, Univunit

DB_MANAGER = bd_unit.DatabaseManager('test.db')


class LukoilQueries(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)

        self.geometry("500x600")
        self.title("Добавление рабочих дней для получения FTE")
        self.hours_var = tk.IntVar()
        self.num_query = tk.StringVar()
        self.result_label = None

        self.create_widgets()

        self.table_fte = Table(self)
        self.table_fte.pack(expand=True, fill='both')
        self.read_all_data()

    def create_widgets(self):
        frame_item = tk.Frame(self)

        tk.Label(frame_item, text='Заявка', bg="#333333", fg="white", font=("Arial", 16)).grid(row=1,
                                                                                               sticky="e")
        tk.Entry(frame_item, textvariable=self.num_query).grid(row=1, column=1, pady=5, padx=10, sticky="ew")

        tk.Label(frame_item, text='Часы', bg="#333333", fg="white", font=("Arial", 16)).grid(row=2,
                                                                                             sticky="e")
        tk.Entry(frame_item, textvariable=self.hours_var).grid(row=2, column=1, pady=5, padx=10, sticky="ew")

        tk.Label(frame_item, text='Дата регистрации', bg="#333333", fg="white", font=("Arial", 16)).grid(row=3,
                                                                                                         sticky="e")

        self.date_reg = DateEntry(master=frame_item, width=15, relief="solid", dateformat='%d-%m-%Y')
        self.date_reg.grid(row=3, column=1, pady=10, sticky="ew", padx=10)

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

    def read_all_data(self):
        self.table_fte.configure_columns(
            [{'name': 'Заявка'}, {'name': 'Часы'}, {'name': 'Регистрация'}, {'name': 'Квартал'}])
        data = DB_MANAGER.read_all_lukoil()
        self.table_fte.populate_table(data)

    def delete_rec(self):
        try:
            num_query = self.num_query.get()
            DB_MANAGER.delete_num_query(num_query)
            self.result_label.config(text=f"Запрос {num_query} удален")
            self.read_all_data()
        except Exception as e:
            messagebox.showinfo('Ошибка с бд', f"Данные не сохранены: {e}")

    def save_days(self):
        try:
            days = self.hours_var.get()
            if days == 0:
                raise ValueError("Количество часов не может быть нулевым.")
            date_registration = datetime.strptime(self.date_reg.entry.get(),'%d-%m-%Y')
            quoter = Univunit.get_first_day_of_quarter(date_registration)
            num_query = self.num_query.get()
            query_hours = self.hours_var.get()

            DB_MANAGER.insert_query([(num_query, query_hours, quoter, date_registration)])
            self.result_label.config(text=f" Запрос {num_query} за {date_registration} добавлен")
            self.read_all_data()
        except ValueError as ve:
            messagebox.showinfo('Ошибка ввода', str(ve))
        except Exception as e:
            messagebox.showinfo('Ошибка с бд', f"Данные не сохранены: {e}")
