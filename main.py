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

# ������� ������� ����
root = tk.Tk()
root.geometry("500x400")
root.title("������ �������")

frame = tk.Frame(root)

# root.withdraw() # ������ ������� ����, ���� �� ������, ����� ������ ��� ������� ��� �������� ����


reports = [
    {
        "name": "����� ������ �� ��������� ������ � �������� ����������� ����������� �� ������",
        "header_row": 0,
        "reportnumber": 1,
        "data_columns": ["����"]
    },
    {
        "name": "�������� ���������� ����� �� ������",
        "header_row": 1,
        "reportnumber": 2,
        "data_columns": ["����"]
    },
    {
        "name": "������� ������ ��������  ��� SLA",
        "header_row": 2,
        "reportnumber": 3,
        "data_columns": ["���� �����������", "������� ���� �������", "���� �������", "���� ��������"],
        # "���� ���������� ���������� � ������"],
        "status": ["� ��������", "���������", "�������", "������ ���������", "�����", "��������", "�����������",
                   "������������ ���������", "���������� ���������", "���������� �������", "������������ ���������",
                   "����������� ���������"]  # ������  ������
    }
]


def report1(df, fte):
    """
    ����� ������ �� ��������� ������ � �������� ����������� ����������� �� ������

    """
    headers = ['������', '����, FTE', '������������', '����������� ������������ (���.) (�����)',
               '���-�� ������� ������']
    fr = df[headers].loc[df['�������� �������'] == '������� ������� �������������']
    fr['����, FTE'] = round(fr['����������� ������������ (���.) (�����)'] / fte, 2)
    fr['���� ����'] = fte * fr['����, FTE']
    fr['������� �����'] = fr['���� ����'] - fr['����������� ������������ (���.) (�����)']
    fr = fr.groupby(['������', '������������', '���-�� ������� ������', '����, FTE', '���� ����',
                     '����, FTE', '������� �����'])['����������� ������������ (���.) (�����)'].sum()
    return fr


def report2(fr):
    """
    �������� ���������� ����� �� ������
    """
    fr = fr.groupby(['������', '���']).agg({'����': 'max', '������������ �� ����': 'sum'})
    return fr


def report3(fr):
    """
            �������� ���������� ����� �� ������ � ���������
    """
    date_range = pd.date_range(start='2024-04-01', end='2024-04-30')

    fr = fr[fr['����'].isin(date_range)]
    fr = fr.sort_values(['������', '���', '����'])
    return fr


def report4(df):
    # fr = df.loc[df["������"] == "��� \"���������������� ���� � ����������\""]

    sum1 = df.groupby(['�2�'])["������� �� ������ �������"].sum()
    sum2 = df.groupby(['�2�'])['��������� � ������'].sum()
    # sum3 = df.loc[df['��� �������']=='��������'].groupby(['�2�']).count()
    sum3 = df[['�2�', '��� �������']].loc[df['��� �������'] == '��������'].groupby(
        ['�2�']).count()
    ss = (f"1 ����� ���������� ������������������ ������ : {sum1}\n\n"
          f"2 ����� ���������� ����������� ������ : {sum2}\n\n"
          f"3 ����� ���������� ������������������ ������ �� "
          f"�������� ������, ������� ��������� ���������: {sum3}")
    return ss

    # print(fr.columns)
    # ss = (f"{df.groupby(['�2�'])["������� �� ������ �������"].sum()} \n"
    #       f"����� { df.loc[df['�2�'].isin(['�','�'])]["������� �� ������ �������"].sum()} \n\n"
    #       f"{df.groupby(['�2�'])['���������������� � ������'].sum()}"
    #       f"����� { df.loc[df['�2�'].isin(['�','�'])]["���������������� � ������"].sum()} \n\n")

    # info()
    # ss = f"{df.groupby(['�2�'])[['������� �� ������ �������', '���������������� � ������',
    #                              '��������� � ������', '���������� � ������', '������� �� ����� �������']].sum()}"

    # ss = [f"{df.groupby(['�2�'])[['������� �� ������ �������', '���������������� � ������',
    #                              '��������� � ������']].sum()}",
    #  f"{df.groupby(['�2�'])[[ '���������� � ������', '������� �� ����� �������']].sum()}"]

    # return f"����� �� ���� ������� �� ������ �������: {s1} \n �����: {itogo}"
    # return f"{s1} \n �����: {itogo}"


def get_data(reportnumber, df, fte=1):
    global fr
    if reportnumber == 1:
        fr = report1(df, fte)
    elif reportnumber == 3:
        fr = report4(df)
    elif reportnumber == 2:
        fr = df[(df['������'] == '�0133-��� "���������������� ���� � ����������"') |
                (df['������'] == '�0134-��� "���������������� ���� � ����������"')][
            ['������', '���', '����', '������������ �� ����']]
        if export_excell_var.get() == 1:
            fr = report3(fr)
        else:
            fr = report2(fr)

    return fr


def read_report():
    # global fr
    name_of_report = cmb.get()
    filename = filedialog.askopenfilename()
    try:  # ������ ����
        for items in reports:
            if items["reportnumber"] == 3:
                df = pd.read_excel(filename, header=items['header_row'], parse_dates=items['data_columns'],
                                   date_format='%d.%m.%Y')
                df = df.loc[df["������"] == "��� \"���������������� ���� � ����������\""]
                df = df.loc[df["������"].isin(items['status'])]

                df["�2�"] = "�"
                df.loc[df["��� �������"] == '�������������', "�2�"] = "����"
                df.loc[df["��� �������"] == '����������� ��� ������������', "�2�"] = "�"

                fr = get_data(items["reportnumber"], df)
                # text.insert(5.0, f'{filename}\n \n \n {fr}')
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
                    label.config(text=f"���� ���������� � {save_file}")
                    fr.to_excel(save_file, index=True)
            text.insert(5.0, f'{filename}\n \n \n {fr}')
                # return df
    except Exception as e:
        text.insert(5.0, f"�� ���� ������� ���� {filename}{e}")


def dec_fte(func):
    def wrap_fn():
        if len(one_hour_fte.get()) == 0:
            raise TypeError("����� ������� FTE")
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
        text.insert(5.0, "fte ������ ���� ������")
    except TypeError as e:
        text.insert(5.0, str(e))
    except Exception as e:
        text.insert(5.0, str(e))


def init():
    label.config(text="")
    text.delete(1.0, END)


def cmb_function(event):
    init()


def btn_go_click():
    init()
    read_report()


export_excell_var = tk.IntVar()
export_excell_checkbox = tk.Checkbutton(frame, text='������� � Excel',
                                        variable=export_excell_var, onvalue=1,
                                        offvalue=0)

label1 = tk.Label(frame, text="������", font=("Helvetica", 16))
cmb = ttk.Combobox(frame, values=[items["name"] for items in reports], state="readonly", width=60)
cmb.set('����� �� ������ �������')
cmb.bind('<<ComboboxSelected>>', cmb_function)
cmb["state"] = "readonly"
label = tk.Label(frame, text="")
label2 = tk.Label(frame, text="FTE:", width=5, border=2)

one_hour_fte = tk.Entry(frame, width=5)

text = Text(width=200, height=50, bg="darkgreen",
            fg='white', wrap=WORD)
text.tag_config('title', justify=LEFT,
                font=("Verdana", 24, 'bold'))

btn_go = tk.Button(frame, text='�������', command=btn_go_click)

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
# ��������� �������� ���� ��������� �������
frame.mainloop()
