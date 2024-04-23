# -*- coding: cp1251 -*-
#  from base64 import b16decode
# from ctypes.wintypes import INT
from typing import Callable, Any

import pandas as pd
from os import path
import tkinter as tk
from tkinter import END, LEFT, WORD, Text, filedialog

# import unit1 as un

# Создаем главное окно
root = tk.Tk()
root.title("Анализ отчётов")
# root.withdraw() # Скрыть главное окно, если вы хотите, чтобы диалог был показан без главного окна
text = Text(width=200, height=50, bg="darkgreen",
            fg='white', wrap=WORD)
text.tag_config('title', justify=LEFT,
                font=("Verdana", 24, 'bold'))


def get_data(file_name, rprt, one_hour_fte=164):
    # Данные по ресурсным планам и списанию трудозатрат сотрудников за период

    global fr
    num_row_header: Callable[[Any], int] = lambda rprt: 0 if rprt == 1 else 1  # if rprt == 2 else 1

    df = pd.read_excel(file_name, header=num_row_header(rprt),
                       parse_dates=['Дата'],  date_format='%d.%m.%Y')
    if rprt == 1:
        headers = ['Проект', 'План, FTE', 'Пользователь', 'Фактические трудозатраты (час.) (Сумма)',
                   'Кол-во штатных единиц']

        fr = df[headers].loc[df['Менеджер проекта'] == 'Тапехин Алексей Александрович']

        fr['Факт, FTE'] = round(fr['Фактические трудозатраты (час.) (Сумма)'] / one_hour_fte, 2)
        fr['Часы план'] = one_hour_fte * fr['План, FTE']
        fr['Остаток часов'] = fr['Часы план'] - fr['Фактические трудозатраты (час.) (Сумма)']
        fr = fr.groupby(['Проект', 'Пользователь', 'Кол-во штатных единиц', 'План, FTE', 'Часы план',
                         'Факт, FTE', 'Остаток часов'])['Фактические трудозатраты (час.) (Сумма)'].sum()
    # Контроль заполнения факта за период
    elif rprt in (2, 3):
        # df = pd.read_excel(file_name, header=1)
        fr = df[(df['Проект'] == 'Т0133-КИС "Производственный учет и отчетность"') |
                (df['Проект'] == 'С0134-КИС "Производственный учет и отчетность"')][
            ['Проект', 'ФИО', 'Дата', 'Трудозатрады за день']]
        # fr['Дата'] = pd.to_datetime(fr['Дата'], format='%d.%m.%Y')

        if rprt == 3:
            date_range = pd.date_range(start='2024-04-01', end='2024-04-10')
            fr = fr[fr['Дата'].isin(date_range)]
            fr = fr.sort_values(['Проект', 'ФИО', 'Дата'])
        else:
            # fr = round(fr.groupby(['Проект', 'ФИО'])['Трудозатрады за день'].sum(), 2)
            fr = fr.groupby(['Проект', 'ФИО']).agg({'Дата': 'max', 'Трудозатрады за день': 'sum'})

    if export_excel.get():
        save_file = path.dirname(file_name) + '/output.xlsx'
        label.config(text=f"Файл сохранился в {save_file}")
        fr.to_excel(save_file, index=True)
    return fr


def print_text(filename, rprt):
    return f'{filename}\n \n \n {get_data(filename, rprt=rprt)}'


# Создаем функцию, которая будет вызываться при нажатии кнопки
def on_button_click():
    filename = filedialog.askopenfilename()
    label.config(text="")
    text.delete(1.0, END)
    msg = ""
    if check_box1.get():
        msg = print_text(filename, rprt=1)
    if check_box2.get():
        if export_excel.get():
            msg = print_text(filename, rprt=3)
        else:
            msg = print_text(filename, rprt=2)

    text.insert(5.0, msg)


# Создаем виджеты (например, кнопку, текстовое поле и метку)
# entry = tk.Entry(root)
# entry.pack()
# def print_selection():
#     if (var1.get() == 1):
#         label.config(text='I love Python ')

export_excel = tk.IntVar()
check_box1 = tk.IntVar()
check_box2 = tk.IntVar()
# c1 = tk.Checkbutton(root, text='Экспорт в Excel', variable=var1, onvalue=1, offvalue=0, command=print_selection)
c1 = tk.Checkbutton(root, text='Экспорт в Excel', variable=export_excel, onvalue=1, offvalue=0)
c1.pack()

chk1 = tk.Checkbutton(root, text="Отчёт Данные по ресурсным планам и списанию трудозатрат сотрудников за период",
                      variable=check_box1, onvalue=1, offvalue=0)
chk1.pack()
chk2 = tk.Checkbutton(root, text="Контроль заполнения факта за период",
                      variable=check_box2, onvalue=1, offvalue=0)
chk2.pack()

button = tk.Button(root, text="Открыть отчёт", command=on_button_click)
button.pack()

label = tk.Label(root, text="")
label.pack()
text.pack()
# Запускаем основной цикл обработки событий
root.mainloop()
