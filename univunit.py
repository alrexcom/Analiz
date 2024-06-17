import tkinter as tk
from datetime import date
import ttkbootstrap as ttk
import pandas as pd

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


def get_first_day_of_quarter(current_date, date_format="%d-%m-%Y"):
    """
    current_date = date.today()
    get_first_day_of_quarter(current_date)
    :param date_format:
    :param current_date:
    :return: Первое число текущего квартала
    """
    year = current_date.year
    month = current_date.month

    if month in [1, 2, 3]:
        quarter_start = date(year, 1, 1)
    elif month in [4, 5, 6]:
        quarter_start = date(year, 4, 1)
    elif month in [7, 8, 9]:
        quarter_start = date(year, 7, 1)
    else:
        quarter_start = date(year, 10, 1)

    return quarter_start.strftime(date_format)

# # Пример использования функции:
# current_date = date.today()
# print(f"Первое число текущего квартала: {get_first_day_of_quarter(current_date)}")


def convert_date(date_str, date_format="%Y-%m-%d"):
    """
    Преобразование даты в формат
    :param date_str:
    :param date_format:
    :return:
    """
    try:
        return pd.to_datetime(date_str, dayfirst=True).strftime(date_format)
    except ValueError:
        raise ValueError("Некорректный формат даты")

