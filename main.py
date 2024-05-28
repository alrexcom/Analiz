# -*- coding: cp1251 -*-
from reports import *
from os import path
import tkinter as tk
from tkinter import (END, LEFT, WORD, Text,
                     filedialog, ttk)
from tkcalendar import DateEntry

# Создаем главное окно
root = tk.Tk()
root.geometry("800x600")
root.title("Анализ отчётов")


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
        text.insert(5.0, f"fte должно быть числом \n {e}")
    except TypeError as e:
        text.insert(5.0, str(e))
    except Exception as e:
        text.insert(5.0, str(e))


def init():
    lbl_path.config(text="")
    text.delete(1.0, END)
    one_hour_fte.delete(0, END)
    one_hour_fte.insert(0, '164')


def cmb_function(event):
    init()


def btn_go_click():
    dp = Date_end.get()
    ds = Date_begin.get()
    fte = get_fte()
    export_excel = export_excell_var.get()
    file_name = 'errors'
    try:
        file_name = filedialog.askopenfilename()
        num_report = cmb.current() + 1

        report_data = get_report(num_report=num_report, filename=file_name)
        param = {'df': report_data, 'fte': fte, 'reportnumber': num_report, 'date_end': dp,
                 'date_begin': ds,
                 'export_excell': export_excel}
        fr = get_data(**param)

        if export_excell_var.get():  # Checbox
            save_file = path.dirname(file_name) + '/output.xlsx'
            lbl_path.config(text=f"Файл сохранился в {save_file}")
            fr.to_excel(save_file, index=True)
        text.insert(5.0, f'{file_name}\n \n \n {fr}')
    except Exception as e:
        text.insert(5.0, f"Не смог открыть файл {file_name}{e}")


export_excell_var = tk.IntVar()
lblzag = tk.Label(root, text="Анализ отчётов", font="italic 14 bold", background="gray", fg="lightgreen")

cmb = ttk.Combobox(root, values=[items["name"] for items in reports], state="readonly", width=90)
cmb.set('Выбор из списка отчетов')
cmb.bind('<<ComboboxSelected>>', cmb_function)
cmb["state"] = "readonly"

btn_go = tk.Button(root, text='Открыть', command=btn_go_click, width=20)
label1 = tk.Label(root, text="", font=("Helvetica", 16), background="gray", fg='white')
lblC = tk.Label(root, text="Период с", font="italic 10 bold", background="gray", fg='white')

# Date_begin = tk.Entry(root, width=15, bd=1, bg="white", relief="solid", font="italic 10 bold")
# Date_begin = DateEntry(root, date_pattern='dd/mm/YYYY')
Date_begin = DateEntry(root, date_pattern='YYYY-mm-dd')

lblpo = tk.Label(root, text="ПО", bd=1, width=2, font="italic 10 bold", background="gray", fg='white')
Date_end = DateEntry(root, width=15, relief="solid", date_pattern='YYYY-mm-dd')

lbl_fte = tk.Label(root, text="FTE:", width=5, border=2, background="gray", fg='white')
one_hour_fte = tk.Entry(root, width=5)

export_excell_checkbox = tk.Checkbutton(root, text='Экспорт в Excel',
                                        variable=export_excell_var, onvalue=1,
                                        offvalue=0, background="gray")
lbl_path = tk.Label(root, text="")

text = Text(bg="darkgreen", fg="white", wrap=WORD)
text.tag_config('title', justify=LEFT)
# ,              font=("Verdana", 24, 'bold'))


# grid layout
root.columnconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)
root.rowconfigure((0, 1, 2, 3), weight=1)
root.rowconfigure(4, weight=40)

lblzag.grid(column=0, row=0, columnspan=7, sticky='nwe', pady=3, padx=5)

cmb.grid(column=0, row=1, columnspan=6, sticky='nw', padx=10)
btn_go.grid(column=6, row=1, sticky='ne', padx=10)
label1.grid(column=0, row=2, columnspan=7, sticky='news', padx=5)

lblC.grid(column=0, row=2, sticky='w', padx=10)
Date_begin.grid(column=1, row=2, sticky='w')
lblpo.grid(column=2, row=2, sticky='w')
Date_end.grid(column=3, row=2, sticky='w')
lbl_fte.grid(column=4, row=2, sticky='e')
one_hour_fte.grid(column=5, row=2, sticky='w')
export_excell_checkbox.grid(column=6, row=2, sticky='e', padx=10)

lbl_path.grid(column=0, row=3, columnspan=7, sticky='nwe', padx=10)
text.grid(column=0, row=4, columnspan=7, sticky='news', padx=5, pady=5)

root.mainloop()
