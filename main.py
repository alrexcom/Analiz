# -*- coding: cp1251 -*-
#  from base64 import b16decode
# from ctypes.wintypes import INT
# from typing import Callable, Any
# from unittest import result

import pandas as pd
from os import path
import tkinter as tk
from tkinter import (END, LEFT, WORD, Text,
                     filedialog, ttk)

# import unit1 as un
rprtname = ""

# Создаем главное окно
root = tk.Tk()
root.geometry("500x400")
root.title("Анализ отчётов")

frame = tk.Frame(root)

# root.withdraw() # Скрыть главное окно, если вы хотите, чтобы диалог был показан без главного окна


reports = [
    {
        "name": "Отчёт Данные по ресурсным планам и списанию трудозатрат сотрудников за период",
        "header_row": 0,
        "reportnumber": 1,
        "data_columns": ["Дата"]
    },
    {
        "name": "Контроль заполнения факта за период",
        "header_row": 1,
        "reportnumber": 2,
        "data_columns": ["Дата"]
    },
    {
        "name": "Сводный список запросов  для SLA",
        "header_row": 2,
        "reportnumber": 3,
        "data_columns": ["Дата регистрации", "Крайний срок решения", "Дата решения", "Дата закрытия"],
        # "Дата последнего назначения в группу"],
        "status": ["В ожидании", "Выполнено", "Закрыто", "Проект изменения", "Решен", "Назначен", "Выполняется",
                   "Планирование изменения", "Выполнение изменения", "Экспертиза решения", "Согласование изменения",
                   "Автроизация изменения"]  # Отмена  убрано
    }
]


def report1(df, fte):
    """
    Отчёт Данные по ресурсным планам и списанию трудозатрат сотрудников за период

    """
    headers = ['Проект', 'План, FTE', 'Пользователь', 'Фактические трудозатраты (час.) (Сумма)',
               'Кол-во штатных единиц']
    fr = df[headers].loc[df['Менеджер проекта'] == 'Тапехин Алексей Александрович']
    fr['Факт, FTE'] = round(fr['Фактические трудозатраты (час.) (Сумма)'] / fte, 2)
    fr['Часы план'] = fte * fr['План, FTE']
    fr['Остаток часов'] = fr['Часы план'] - fr['Фактические трудозатраты (час.) (Сумма)']
    fr = fr.groupby(['Проект', 'Пользователь', 'Кол-во штатных единиц', 'План, FTE', 'Часы план',
                     'Факт, FTE', 'Остаток часов'])['Фактические трудозатраты (час.) (Сумма)'].sum()
    return fr


def report2(fr):
    """
    Контроль заполнения факта за период
    """
    fr = fr.groupby(['Проект', 'ФИО']).agg({'Дата': 'max', 'Трудозатрады за день': 'sum'})
    return fr


def report3(md):
    """
            Контроль заполнения факта за период с экспортом
    """
    # date_range = pd.date_range(start='2024-04-01', end='2024-04-30')

    data_reg = md[(md['Дата'] <= Date_end.get())
                  & (md['Дата'] >= Date_begin.get())]
    # fr = fr[fr['Дата'].isin(date_range)]

    return data_reg.sort_values(['Проект', 'ФИО', 'Дата'])


def report4(df):
    # fr = df.loc[df["Услуга"] == "КИС \"Производственный учет и отчетность\""]
    # date_range = pd.date_range(Date_begin.get(), Date_end.get(), freq='D')
    mdf = df
    # df[df['Дата регистрации'] < '2023-08-31']
    data_reg = mdf[(mdf['Дата регистрации'] <= Date_end.get())
                   & (mdf['Дата регистрации'] >= Date_begin.get())]
    sum1 = data_reg.groupby(['П2С'])["Зарегистрировано в период"].sum()
    sum2 = mdf.groupby(['П2С'])['Выполнено в период'].sum()
    # sum3 = df.loc[df['Тип запроса']=='Инцидент'].groupby(['П2С']).count()
    inzindent = mdf[['П2С', 'Дата регистрации', 'Статус',
                     'Просрочено в период', 'Дата закрытия']].loc[mdf['Тип запроса'] == 'Инцидент']

    data_inzindent = inzindent[(inzindent['Дата регистрации'] <= Date_end.get())
                               & (inzindent['Дата регистрации'] >= Date_begin.get())]

    sum3 = data_inzindent[['Дата регистрации', 'П2С']].groupby('П2С').count()

    data_inzindent = inzindent[(inzindent['Статус'] == 'Закрыто') & (inzindent['Просрочено в период'] > 0)
                               & (inzindent['Дата закрытия'] <= Date_end.get())
                               & (inzindent['Дата закрытия'] >= Date_begin.get())]

    sum4 = data_inzindent[['Просрочено в период', 'П2С']].groupby(['П2С']).count()

    # data_ = mdf[['Просрочено в период', 'П2С', 'Дата регистрации']][(mdf['Просрочено в период'] > 0) &
    #                                                                 (mdf['Дата регистрации'] >= Date_begin.get())
    #                                                                 & (mdf['Дата регистрации'] <= Date_end.get())
    #                                                                 ]

    sum6 = data_reg[['Просрочено в период', 'П2С']].groupby(['П2С']).sum()

    sum7 = mdf[['Открыто на конец периода с просрочкой', 'П2С']].groupby(['П2С']).sum()

    sum8 = mdf.groupby(['П2С'])["Открыто на начало периода"].sum()
    # Открыто на конец периода с просрочкой
    slap = round((1 - (sum6['Просрочено в период']['П'] + sum7['Открыто на конец периода с просрочкой']['П'])
                  / (sum8['П'] + sum1['П'])) * 100, 2)
    slac = round((1 - (sum6['Просрочено в период']['С'] + sum7['Открыто на конец периода с просрочкой']['С'])
                  / (sum8['С'] + sum1['С'])) * 100, 2)
    ss = (f"SLA для поддержки = {slap} "
          f"SLA для сопровождения = {slac} "
          f"\n----------------------------------\n"
          f"1 Общее количество зарегистрированных заявок : {sum1}"
          f"\n-Итого:{sum1.sum()}\n\n"
          f"2 Общее количество выполненных заявок : {sum2}"
          f"\n-Итого:{sum2.sum()}\n\n"
          f"3 Общее количество зарегистрированных заявок за "
          f" отчетный период, имеющих категорию «Инцидент»: {sum3}"
          f"\n-Итого:{sum3.sum()}\n\n"
          f"4 Количество заявок за период с превышением срока выполнения, имеющих категорию «Инцидент» : {sum4}"
          f"\n-Итого:{sum4.sum()}\n\n"
          f"5 Количество заявок за период с превышением времени реакции по поддержке: 0 \n\n"
          f"6 (TUR1) Количество закрытых заявок на поддержку с нарушением сроков заявок:{sum6}"
          f"\n-Итого:{sum6.sum()}\n\n"
          f"7 (TUR2) Количество незакрытых заявок, с нарушением срока: {sum7}"
          f"\n-Итого:{sum7.sum()}\n\n"
          f"8 (TUR3) Количество перешедших с прошлого периода заявок на поддержку: {sum8}"
          f"\n-Итого:{sum8.sum()}\n\n"
          f"9 (TUR4) Количество зарегистрированных заявок по поддержке: {sum1}"
          f"\n-Итого:{sum1.sum()}"

          )
    return ss


def get_data(reportnumber, df, fte=1):
    global fr
    if reportnumber == 1:
        fr = report1(df, fte)
    elif reportnumber == 3:
        fr = report4(df)
    elif reportnumber == 2:
        fr = df[(df['Проект'] == 'Т0133-КИС "Производственный учет и отчетность"') |
                (df['Проект'] == 'С0134-КИС "Производственный учет и отчетность"')][
            ['Проект', 'ФИО', 'Дата', 'Трудозатрады за день']]
        if export_excell_var.get() == 1:
            fr = report3(fr)
        else:
            fr = report2(fr)

    return fr


def read_report():
    # global fr
    name_of_report = cmb.get()
    filename = filedialog.askopenfilename()

    try:  # Читаем файл
        for items in reports:
            if items["reportnumber"] == 3:
                df = pd.read_excel(filename, header=items['header_row'], parse_dates=items['data_columns'],
                                   date_format='%d.%m.%Y')

                df = df.loc[df["Услуга"] == "КИС \"Производственный учет и отчетность\""]
                df = df.loc[df["Статус"].isin(items['status'])]

                df["П2С"] = "П"
                df.loc[df["Тип запроса"] == 'Нестандартное', "П2С"] = "СДОП"
                df.loc[df["Тип запроса"] == 'Стандартное без согласования', "П2С"] = "С"

                fr = get_data(items["reportnumber"], df)
                text.insert(5.0, f'{filename}\n \n \n {fr}')
            elif items['name'] == name_of_report:
                df = pd.read_excel(filename, header=items['header_row'], parse_dates=items['data_columns'],
                                   date_format='%d.%m.%Y')
                if items["reportnumber"] == 1:
                    fte = get_fte()
                    fr = get_data(items["reportnumber"], df, fte)
                else:
                    fr = get_data(items["reportnumber"], df)

                if export_excell_var.get():  # Checbox
                    save_file = path.dirname(filename) + '/output.xlsx'
                    label.config(text=f"Файл сохранился в {save_file}")
                    fr.to_excel(save_file, index=True)
                text.insert(5.0, f'{filename}\n \n \n {fr}')
                # return df
    except Exception as e:
        text.insert(5.0, f"Не смог открыть файл {filename}{e}")


def dec_fte(func):
    def wrap_fn():
        if len(one_hour_fte.get()) == 0:
            raise TypeError("Нужно указать FTE")
        elif int(one_hour_fte.get()) <= 0:
            raise TypeError("fte <=0")
        else:
            return func()

    return wrap_fn


@dec_fte
def get_fte():
    try:
        return int(one_hour_fte.get())
    except ValueError as e:
        text.insert(5.0, "fte должно быть числом")
    except TypeError as e:
        text.insert(5.0, str(e))
    except Exception as e:
        text.insert(5.0, str(e))


def init():
    label.config(text="")
    text.delete(1.0, END)
    Date_begin.delete(0, END)
    Date_end.delete(0, END)
    Date_begin.insert(0, "2024-01-01")
    Date_end.insert(0, "2024-03-30")


def cmb_function(event):
    init()


def btn_go_click():
    # init()
    read_report()


# Date_begin = imp_begin.get()

frame1 = tk.Frame()

frame1.pack(pady=10)
tk.Label(frame1, text="Отчетный период", bd=1, width=80, font="italic 14 bold").pack()
tk.Label(frame1, text="C", bd=1, width=10, font="italic 10 bold").pack(side=LEFT)
Date_begin = tk.Entry(frame1, width=15, bd=1, bg="white", relief="solid", font="italic 10 bold")
Date_begin.pack(side=LEFT, padx=15)

tk.Label(frame1, text="ПО", bd=1, width=2,
         font="italic 10 bold").pack(side=LEFT)
Date_end = tk.Entry(frame1, width=15, relief="solid")
Date_end.pack(padx=15)

export_excell_var = tk.IntVar()
export_excell_checkbox = tk.Checkbutton(frame, text='Экспорт в Excel',
                                        variable=export_excell_var, onvalue=1,
                                        offvalue=0)

label1 = tk.Label(frame, text="Отчеты", font=("Helvetica", 16))

cmb = ttk.Combobox(frame, values=[items["name"] for items in reports], state="readonly", width=60)
cmb.set('Выбор из списка отчетов')
cmb.bind('<<ComboboxSelected>>', cmb_function)
cmb["state"] = "readonly"
label = tk.Label(frame, text="")
label2 = tk.Label(frame, text="FTE:", width=5, border=2)

one_hour_fte = tk.Entry(frame, width=5)

text = Text(width=200, height=50, bg="darkgreen",
            fg='white', wrap=WORD)
text.tag_config('title', justify=LEFT,
                font=("Verdana", 24, 'bold'))

btn_go = tk.Button(frame, text='Открыть', command=btn_go_click)

label2.grid(column=0, row=0, padx=0, pady=1, sticky='w')
one_hour_fte.grid(column=0, row=0, padx=40, pady=0, sticky='w')
export_excell_checkbox.grid(column=0, row=0, padx=90, sticky='w')
label1.grid(column=1, row=0, padx=0, pady=0, sticky='w')
cmb.grid(column=0, row=1, padx=1, columnspan=2)
btn_go.grid(column=2, row=1, padx=5)
label.grid(column=0, row=2, padx=1, pady=0, columnspan=2)

frame.pack(side="top")
text.pack()
# export_excell_checkbox.pack(side="left")
# Запускаем основной цикл обработки событий
frame.mainloop()
