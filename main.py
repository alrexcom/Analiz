# -*- coding: cp1251 -*-
#  from base64 import b16decode
# from ctypes.wintypes import INT
from typing import Callable, Any

import pandas as pd
from os import path
import tkinter as tk
from tkinter import END, LEFT, WORD, Text, filedialog

# import unit1 as un

# ������� ������� ����
root = tk.Tk()
root.title("������ �������")
# root.withdraw() # ������ ������� ����, ���� �� ������, ����� ������ ��� ������� ��� �������� ����
text = Text(width=200, height=50, bg="darkgreen",
            fg='white', wrap=WORD)
text.tag_config('title', justify=LEFT,
                font=("Verdana", 24, 'bold'))


def get_data(file_name, rprt, one_hour_fte=164):
    # ������ �� ��������� ������ � �������� ����������� ����������� �� ������

    global fr
    num_row_header: Callable[[Any], int] = lambda rprt: 0 if rprt == 1 else 1  # if rprt == 2 else 1

    df = pd.read_excel(file_name, header=num_row_header(rprt),
                       parse_dates=['����'],  date_format='%d.%m.%Y')
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

    if export_excel.get():
        save_file = path.dirname(file_name) + '/output.xlsx'
        label.config(text=f"���� ���������� � {save_file}")
        fr.to_excel(save_file, index=True)
    return fr


def print_text(filename, rprt):
    return f'{filename}\n \n \n {get_data(filename, rprt=rprt)}'


# ������� �������, ������� ����� ���������� ��� ������� ������
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


# ������� ������� (��������, ������, ��������� ���� � �����)
# entry = tk.Entry(root)
# entry.pack()
# def print_selection():
#     if (var1.get() == 1):
#         label.config(text='I love Python ')

export_excel = tk.IntVar()
check_box1 = tk.IntVar()
check_box2 = tk.IntVar()
# c1 = tk.Checkbutton(root, text='������� � Excel', variable=var1, onvalue=1, offvalue=0, command=print_selection)
c1 = tk.Checkbutton(root, text='������� � Excel', variable=export_excel, onvalue=1, offvalue=0)
c1.pack()

chk1 = tk.Checkbutton(root, text="����� ������ �� ��������� ������ � �������� ����������� ����������� �� ������",
                      variable=check_box1, onvalue=1, offvalue=0)
chk1.pack()
chk2 = tk.Checkbutton(root, text="�������� ���������� ����� �� ������",
                      variable=check_box2, onvalue=1, offvalue=0)
chk2.pack()

button = tk.Button(root, text="������� �����", command=on_button_click)
button.pack()

label = tk.Label(root, text="")
label.pack()
text.pack()
# ��������� �������� ���� ��������� �������
root.mainloop()
