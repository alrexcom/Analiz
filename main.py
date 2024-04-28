# -*- coding: cp1251 -*-
#  from base64 import b16decode
# from ctypes.wintypes import INT
from typing import Callable, Any

import pandas as pd
from os import path
import tkinter as tk
from tkinter import END, LEFT, WORD, Text, filedialog, ttk

# import unit1 as un
rprtname = ""

# ������� ������� ����
root = tk.Tk()
root.geometry("500x400")
root.title("������ �������")

frame = tk.Frame(root)


# root.withdraw() # ������ ������� ����, ���� �� ������, ����� ������ ��� ������� ��� �������� ����


def get_data(file_name, rprt, one_hour_fte=164):
    # ������ �� ��������� ������ � �������� ����������� ����������� �� ������

    global fr
    num_row_header: Callable[[Any], int] = lambda rprt: 0 if rprt == 1 else 1  # if rprt == 2 else 1

    df = pd.read_excel(file_name, header=num_row_header(rprt),
                       parse_dates=['����'], date_format='%d.%m.%Y')
    if rprt == 1:
        headers = ['������', '����, FTE', '������������', '����������� ������������ (���.) (�����)',
                   '���-�� ������� ������']

        fr = df[headers].loc[df['�������� �������'] == '������� ������� �������������']

        fr['����, FTE'] = round(fr['����������� ������������ (���.) (�����)'] / one_hour_fte, 2)
        fr['���� ����'] = one_hour_fte * fr['����, FTE']
        fr['������� �����'] = fr['���� ����'] - fr['����������� ������������ (���.) (�����)']
        fr = fr.groupby(['������', '������������', '���-�� ������� ������', '����, FTE', '���� ����',
                         '����, FTE', '������� �����'])['����������� ������������ (���.) (�����)'].sum()
    # �������� ���������� ����� �� ������
    elif rprt in (2, 3):
        # df = pd.read_excel(file_name, header=1)
        fr = df[(df['������'] == '�0133-��� "���������������� ���� � ����������"') |
                (df['������'] == '�0134-��� "���������������� ���� � ����������"')][
            ['������', '���', '����', '������������ �� ����']]
        # fr['����'] = pd.to_datetime(fr['����'], format='%d.%m.%Y')

        if rprt == 3:
            date_range = pd.date_range(start='2024-04-01', end='2024-04-10')
            fr = fr[fr['����'].isin(date_range)]
            fr = fr.sort_values(['������', '���', '����'])
        else:
            # fr = round(fr.groupby(['������', '���'])['������������ �� ����'].sum(), 2)
            fr = fr.groupby(['������', '���']).agg({'����': 'max', '������������ �� ����': 'sum'})

    if export_excell:  # Checbox
        save_file = path.dirname(file_name) + '/output.xlsx'
        label.config(text=f"���� ���������� � {save_file}")
        fr.to_excel(save_file, index=True)
    return fr


def print_text(filename, rprt):
    return f'{filename}\n \n \n {get_data(filename, rprt=rprt)}'


def cmb_function(event):
    filename = filedialog.askopenfilename()
    label.config(text="")
    text.delete(1.0, END)
    msg = ""
    if cmb.get() == reports[0]:
        msg = print_text(filename, rprt=1)
    if cmb.get() == reports[1]:
        if export_excell:  # export checkbox
            msg = print_text(filename, rprt=3)
        else:
            msg = print_text(filename, rprt=2)

    text.insert(5.0, msg)


export_excell = tk.IntVar()
reports = [
    "����� ������ �� ��������� ������ � �������� ����������� ����������� �� ������",
    "�������� ���������� ����� �� ������"
]

export_excell_checkbox = tk.Checkbutton(frame, text='������� � Excel', onvalue=True, offvalue=False,
                                        variable=export_excell, )
label1 = tk.Label(frame, text="������", font=("Helvetica", 16))
cmb = ttk.Combobox(frame, values=reports, state="readonly", width=60)
cmb.set('����� �� ������ �������')
cmb.bind('<<ComboboxSelected>>', cmb_function)
cmb["state"] = "readonly"
label = tk.Label(frame, text="")

text = Text(width=200, height=50, bg="darkgreen",
            fg='white', wrap=WORD)
text.tag_config('title', justify=LEFT,
                font=("Verdana", 24, 'bold'))

label1.grid(column=0, row=0, padx=1, pady=0)
export_excell_checkbox.grid(column=1, row=0)
cmb.grid(column=0, row=1, padx=1, columnspan=2)
label.grid(column=0, row=2, padx=1, pady=0, columnspan=2)
# text.grid(column=0, row=3, columnspan=2)




frame.pack(side="top")
text.pack()
# ��������� �������� ���� ��������� �������
frame.mainloop()
