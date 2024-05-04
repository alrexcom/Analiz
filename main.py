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
import datetime
# import tkcalendar as cc
from tkcalendar import DateEntry

# import unit1 as un
rprtname = ""

# ������� ������� ����
root = tk.Tk()
root.geometry("800x600")
root.title("������ �������")

# frame = tk.Frame(root)

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
        "data_columns": ["���� �����������", "������� ���� �������", "���� �������", "���� ��������"]
        # "���� ���������� ���������� � ������"],
        # "status": ["� ��������", "���������", "�������", "������ ���������", "�����", "��������", "�����������",
        #            "������������ ���������", "���������� ���������", "���������� �������", "������������ ���������",
        #            "����������� ���������"]  # ������  ������
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


def report2(df, date_begin, date_end, export_to_excell):
    """
    �������� ���������� ����� �� ������
    """

    frm = df[(df['������'] == '�0133-��� "���������������� ���� � ����������"') |
             (df['������'] == '�0134-��� "���������������� ���� � ����������"')][
        ['������', '���', '����', '������������ �� ����']]
    if export_to_excell == False:
        result = frm.groupby(['������', '���']).agg({'����': 'max', '������������ �� ����': 'sum'})
    else:
        date_range = pd.date_range(start=date_begin, end=date_end, freq='D')
        data_reg = frm[frm['����'].isin(date_range)]
        result = data_reg.sort_values(['������', '���', '����'])
    return result


def report3(df, date_end, date_begin):
    """
    ������� ������ ��������  ��� SLA
    """
    status = ["� ��������", "���������", "�������", "������ ���������", "�����", "��������", "�����������",
              "������������ ���������", "���������� ���������", "���������� �������", "������������ ���������",
              "����������� ���������"]

    try:
        mdf = df
        mdf = mdf.loc[mdf["������"] == "��� \"���������������� ���� � ����������\""]
        mdf = mdf.loc[mdf["������"].isin(status)]
        mdf["�2�"] = "�"
        mdf.loc[mdf["��� �������"] == '�������������', "�2�"] = "����"
        mdf.loc[mdf["��� �������"] == '����������� ��� ������������', "�2�"] = "�"


        # fr = df.loc[df["������"] == "��� \"���������������� ���� � ����������\""]
        date_range = pd.date_range(start=date_begin, end=date_end, freq='D')

        data_reg = mdf[mdf['���� �����������'].isin(date_range)]
        # data_reg = mdf[(mdf['���� �����������'] <= date_end)
        #                & (mdf['���� �����������'] >= date_begin)]

        sum1 = data_reg.groupby(['�2�'])["���������������� � ������"].sum()
        sum2 = mdf.groupby(['�2�'])['��������� � ������'].sum()
        # sum3 = df.loc[df['��� �������']=='��������'].groupby(['�2�']).count()
        inzindent = mdf[['�2�', '���� �����������', '������',
                         '���������� � ������', '���� ��������']].loc[mdf['��� �������'] == '��������']

        data_inzindent = inzindent[(inzindent['���� �����������'] <= date_end)
                                   & (inzindent['���� �����������'] >= date_begin)]

        sum3 = data_inzindent[['���� �����������', '�2�']].groupby('�2�').count()

        data_inzindent = inzindent[(inzindent['������'] == '�������') & (inzindent['���������� � ������'] > 0)
                                   & (inzindent['���� ��������'] <= date_end)
                                   & (inzindent['���� ��������'] >= date_begin)]

        sum4 = data_inzindent[['���������� � ������', '�2�']].groupby(['�2�']).count()

        sum6 = data_reg[['���������� � ������', '�2�']].groupby(['�2�']).sum()

        sum7 = mdf[['������� �� ����� ������� � ����������', '�2�']].groupby(['�2�']).sum()

        sum8 = mdf.groupby(['�2�'])["������� �� ������ �������"].sum()
        # ������� �� ����� ������� � ����������
        slap = round((1 - (sum6['���������� � ������']['�'] + sum7['������� �� ����� ������� � ����������']['�'])
                      / (sum8['�'] + sum1['�'])) * 100, 2)
        slac = round((1 - (sum6['���������� � ������']['�'] + sum7['������� �� ����� ������� � ����������']['�'])
                      / (sum8['�'] + sum1['�'])) * 100, 2)
        ss = (f"SLA ��� ��������� = {slap} "
              f"SLA ��� ������������� = {slac} "
              f"\n----------------------------------\n"
              f"1 ����� ���������� ������������������ ������ : {sum1}"
              f"\n-�����:{sum1.sum()}\n\n"
              f"2 ����� ���������� ����������� ������ : {sum2}"
              f"\n-�����:{sum2.sum()}\n\n"
              f"3 ����� ���������� ������������������ ������ �� "
              f" �������� ������, ������� ��������� ���������: {sum3}"
              f"\n-�����:{sum3.sum()}\n\n"
              f"4 ���������� ������ �� ������ � ����������� ����� ����������, ������� ��������� ��������� : {sum4}"
              f"\n-�����:{sum4.sum()}\n\n"
              f"5 ���������� ������ �� ������ � ����������� ������� ������� �� ���������: 0 \n\n"
              f"6 (TUR1) ���������� �������� ������ �� ��������� � ���������� ������ ������:{sum6}"
              f"\n-�����:{sum6.sum()}\n\n"
              f"7 (TUR2) ���������� ���������� ������, � ���������� �����: {sum7}"
              f"\n-�����:{sum7.sum()}\n\n"
              f"8 (TUR3) ���������� ���������� � �������� ������� ������ �� ���������: {sum8}"
              f"\n-�����:{sum8.sum()}\n\n"
              f"9 (TUR4) ���������� ������������������ ������ �� ���������: {sum1}"
              f"\n-�����:{sum1.sum()}"

              )
        return ss
    except Exception as e:
        return f"������ ���������� {e}"


def get_data(reportnumber, date_end, date_begin, df, fte, export_excell):
    frm = ''
    if reportnumber == 1:
        frm = report1(df, fte)
    elif reportnumber == 3:
        frm = report3(df, date_end=date_end, date_begin=date_begin)
    elif reportnumber == 2:
        frm = report2(df=df, date_begin=date_begin, date_end=date_end, export_to_excell=export_excell)
    return frm


def get_report(num_report, filename):
    for items in reports:
        if num_report == items['reportnumber']:
            result = pd.read_excel(filename, header=items['header_row'], parse_dates=items['data_columns'],
                                   date_format='%d.%m.%Y')
            return result


# def read_report(num_report, date_begin, date_end, filename, export_excel, fte):


# for items in reports:
#     df = pd.read_excel(filename, header=items['header_row'], parse_dates=items['data_columns'],
#                        date_format='%d.%m.%Y')
#     if items["reportnumber"] == 3 & num_report == 3:
#
#         df = df.loc[df["������"] == "��� \"���������������� ���� � ����������\""]
#         df = df.loc[df["������"].isin(items['status'])]
#
#         df["�2�"] = "�"
#         df.loc[df["��� �������"] == '�������������', "�2�"] = "����"
#         df.loc[df["��� �������"] == '����������� ��� ������������', "�2�"] = "�"
#
#         fr = get_data(reportnumber=items["reportnumber"], df=df, datebegin=databegin, dateend=dateend, fte=1)
#         text.insert(5.0, f'{filename}\n \n \n {fr}')
#     elif items["reportnumber"] == 1 & num_report == 1:
#         fte = get_fte()
#         fr = get_data(reportnumber=items["reportnumber"], df=df, datebegin=databegin, dateend=dateend,
#                       fte=fte)
#     else:
#         fr = get_data(reportnumber=items["reportnumber"], df=df, datebegin=databegin, dateend=dateend,
#                       fte=1)


# return df


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
        text.insert(5.0, f"fte ������ ���� ������ \n {e}")
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
            lbl_path.config(text=f"���� ���������� � {save_file}")
            fr.to_excel(save_file, index=True)
        text.insert(5.0, f'{file_name}\n \n \n {fr}')
    except Exception as e:
        text.insert(5.0, f"�� ���� ������� ���� {file_name}{e}")


# Date_begin = imp_begin.get()
export_excell_var = tk.IntVar()

lblzag = tk.Label(root, text="������ �������", font="italic 14 bold", background="gray", fg="lightgreen")

cmb = ttk.Combobox(root, values=[items["name"] for items in reports], state="readonly", width=90)
cmb.set('����� �� ������ �������')
cmb.bind('<<ComboboxSelected>>', cmb_function)
cmb["state"] = "readonly"

btn_go = tk.Button(root, text='�������', command=btn_go_click, width=20)
label1 = tk.Label(root, text="", font=("Helvetica", 16), background="gray", fg='white')
lblC = tk.Label(root, text="������ �", font="italic 10 bold", background="gray", fg='white')
# Date_begin = tk.Entry(root, width=15, bd=1, bg="white", relief="solid", font="italic 10 bold")
# Date_begin = DateEntry(root, date_pattern='dd/mm/YYYY')
Date_begin = DateEntry(root, date_pattern='YYYY-mm-dd')

lblpo = tk.Label(root, text="��", bd=1, width=2, font="italic 10 bold", background="gray", fg='white')
Date_end = DateEntry(root, width=15, relief="solid", date_pattern='YYYY-mm-dd')

lbl_fte = tk.Label(root, text="FTE:", width=5, border=2, background="gray", fg='white')
one_hour_fte = tk.Entry(root, width=5)

export_excell_checkbox = tk.Checkbutton(root, text='������� � Excel',
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
