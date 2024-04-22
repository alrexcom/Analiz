# -*- coding: cp1251 -*-
#  from base64 import b16decode
# from ctypes.wintypes import INT
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




def get_data(file_name, one_hour_fte=164, rprt=1):
    global fr
    # ������ �� ��������� ������ � �������� ����������� ����������� �� ������
    if rprt == 1:
        df = pd.read_excel(file_name, header=0)
        fr = df[df['�������� �������'] == '������� ������� �������������'][['������', '����, FTE', '������������',
                                                                            '����������� ������������ (���.) (�����)',
                                                                            '���-�� ������� ������']]
        fr['����, FTE'] = round(fr['����������� ������������ (���.) (�����)'] / one_hour_fte, 2)
        fr['���� ����'] = one_hour_fte * fr['����, FTE']
        fr['������� �����'] = fr['���� ����'] - fr['����������� ������������ (���.) (�����)']
        fr = fr.groupby(['������', '������������', '���-�� ������� ������', '����, FTE', '���� ����',
                         '����, FTE', '������� �����'])['����������� ������������ (���.) (�����)'].sum()
    # �������� ���������� ����� �� ������
    elif rprt in (2, 3):
        df = pd.read_excel(file_name, header=1)
        fr = df[(df['������'] == '�0133-��� "���������������� ���� � ����������"') |
                (df['������'] == '�0134-��� "���������������� ���� � ����������"')][
            ['������', '���', '����', '������������ �� ����']]
        if rprt == 3:
            fr = fr.sort_values(['������', '���', '����'])
        else:
            fr = round(fr.groupby(['������', '���'])['������������ �� ����'].sum(), 2)
    return fr


def print_text(filename):
    return f'{filename}\n \n \n {fr}'


# ������� �������, ������� ����� ���������� ��� ������� ������
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
        label.config(text=f"���� ���������� � {save_file}")
        fr.to_excel(save_file, index=True)


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
