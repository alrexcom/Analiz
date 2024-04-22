# -*- coding: cp1251 -*-
#  from base64 import b16decode
# from ctypes.wintypes import INT
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




def get_data(file_name, one_hour_fte=164, rprt=1):
    global fr
    # Данные по ресурсным планам и списанию трудозатрат сотрудников за период
    if rprt == 1:
        df = pd.read_excel(file_name, header=0)
        fr = df[df['Менеджер проекта'] == 'Тапехин Алексей Александрович'][['Проект', 'План, FTE', 'Пользователь',
                                                                            'Фактические трудозатраты (час.) (Сумма)',
                                                                            'Кол-во штатных единиц']]
        fr['Факт, FTE'] = round(fr['Фактические трудозатраты (час.) (Сумма)'] / one_hour_fte, 2)
        fr['Часы план'] = one_hour_fte * fr['План, FTE']
        fr['Остаток часов'] = fr['Часы план'] - fr['Фактические трудозатраты (час.) (Сумма)']
        fr = fr.groupby(['Проект', 'Пользователь', 'Кол-во штатных единиц', 'План, FTE', 'Часы план',
                         'Факт, FTE', 'Остаток часов'])['Фактические трудозатраты (час.) (Сумма)'].sum()
    # Контроль заполнения факта за период
    elif rprt in (2, 3):
        df = pd.read_excel(file_name, header=1)
        fr = df[(df['Проект'] == 'Т0133-КИС "Производственный учет и отчетность"') |
                (df['Проект'] == 'С0134-КИС "Производственный учет и отчетность"')][
            ['Проект', 'ФИО', 'Дата', 'Трудозатрады за день']]
        if rprt == 3:
            fr = fr.sort_values(['Проект', 'ФИО', 'Дата'])
        else:
            fr = round(fr.groupby(['Проект', 'ФИО'])['Трудозатрады за день'].sum(), 2)
    return fr


def print_text(filename):
    return f'{filename}\n \n \n {fr}'


# Создаем функцию, которая будет вызываться при нажатии кнопки
def on_button_click():

    filename = filedialog.askopenfilename()
    label.config(text="")
    text.delete(1.0, END)

    if check_box1.get():
        fr = get_data(filename, rprt=1)
    if check_box2.get():
        if export_excel.get():
            fr = get_data(filename, rprt=3)
            # fr = un.get_data(filename, rprt=2)
        else:
            fr = get_data(filename, rprt=2)
        # fr = un.get_data(filename, rprt=lambda rp: 3 if export_excel.get() else 2)
    # arg = [filename, fr]
    text.insert(5.0, print_text(filename))

    if export_excel.get():
        save_file = path.dirname(filename) + '/output.xlsx'
        label.config(text=f"Файл сохранился в {save_file}")
        fr.to_excel(save_file, index=True)


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
