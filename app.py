from tkinter.ttk import Scrollbar

from reports import *
# from tkscrolledframe import ScrolledFrame
import tkinter as tk
from tkinter import (filedialog, font)

import ttkbootstrap as ttk
from ttkbootstrap.scrolled import ScrolledFrame

themes = ['cosmo', 'flatly', 'litera', 'minty', 'lumen', 'sandstone',
          'yeti', 'pulse', 'united', 'morph', 'journal', 'darkly',
          'superhero', 'solar', 'cyborg', 'vapor', 'simplex', 'cerculean']


class Component(ttk.Frame):
    def __init__(self, parent, label_text, text_value):
        super().__init__(master=parent)

        # grid
        self.rowconfigure(0, weight=1)
        self.columnconfigure((0, 1), weight=1, uniform='a')
        ttk.Label(self,
                  text=label_text,
                  wraplength=400,
                  font=("Helvetica", 12),
                  justify=tk.RIGHT
                  ).grid(
            padx=5, row=0, column=0,  sticky=tk.W)
        ttk.Label(self,
                  text=text_value,
                  background="pink",
                  border=2,
                  relief=tk.SUNKEN,
                  anchor=tk.CENTER,
                  font=("Helvetica", 14),
                  width=10).grid(row=0, column=1)

        self.pack(expand=1, fill=tk.BOTH)


class App(tk.Tk):
    def __init__(self, title, size, _theme):
        super().__init__()
        self.style = ttk.Style(theme=_theme)
        self.title(title)
        # self['background'] = '#EBEBEB'
        self.geometry(f"{size[0]}x{size[1]}")
        self.minsize(size[0], size[1])
        self.maxsize(size[0], size[1])
        self.bold_font = font.Font(family='Helvetica', size=13, weight='bold')

        HeadFrame(self)
        global result_frame, table_frame

        result_frame = ScrolledFrame(self, autohide=True)

        result_frame.pack(side=tk.LEFT, expand=True, fill=tk.BOTH, pady=10)

        table_frame =ttk.Frame(self)
        table_frame.pack(expand=1, fill=tk.BOTH, pady=10, padx=10)
        self.mainloop()


class TableFrame(tk.Frame):
    def __init__(self, parent, data):
        super().__init__(parent)
        headers = data['columnsname']
        tree_scroll = Scrollbar(table_frame, orient=tk.HORIZONTAL)
        table = ttk.Treeview(table_frame, columns=headers,
                             show='headings', xscrollcommand=tree_scroll.set)

        for column in headers:
            table.heading(column, text=column)

        for row in data['values']:
            table.insert('', 'end', values=data['values'][row])


class HeadFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self.show()

    def show(self):
        global ds, dp, msg_info, one_hour_fte, export_excell_var
        cmbframe = ttk.Frame(self, padding=1)
        # name_report=tk.StringVar()
        cmb = ttk.Combobox(cmbframe, values=[items["name"] for items in reports],
                           state="readonly",
                           height=4,
                           font=("Calibri", 12)
                           # , textvariable=name_report
                           )
        cmb.set('Выбор из списка отчетов')
        cmb.bind('<<ComboboxSelected>>', cmb_function)
        cmb.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        cmbframe.pack(expand=True, fill=tk.BOTH)

        # Menu strip
        menu_frame = ttk.Frame(self, padding=5, borderwidth=5)

        ds = ttk.DateEntry(
            master=menu_frame,
            width=15,
            relief="solid",
            dateformat='%d-%m-%Y'
        )
        ds.pack(side=tk.LEFT, padx=10)
        ds.entry.delete(0, tk.END)
        ds.entry.insert(0, '01-01-2024')

        dp = ttk.DateEntry(menu_frame,
                           width=15,
                           relief="solid",
                           dateformat='%d-%m-%Y')
        dp.pack(side=tk.LEFT, padx=10)

        ttk.Label(menu_frame, text="FTE:", width=5,
                  border=2, font=("Calibri", 12, 'bold'),
                  background="gray", foreground='white').pack(side=tk.LEFT)

        one_hour_fte = tk.IntVar()

        fte = ttk.Entry(menu_frame, width=5, font=("Calibri", 12),
                        textvariable=one_hour_fte)
        fte.delete(0, tk.END)
        fte.insert(0, '164')
        fte.pack(side=tk.LEFT, padx=10)

        export_excell_var = tk.IntVar()
        ttk.Checkbutton(menu_frame, text='Экспорт в Excel',
                        variable=export_excell_var,
                        onvalue=1,
                        offvalue=0).pack(side=tk.LEFT, padx=5)

        btn_clear = ttk.Button(menu_frame, text='Очистить',
                               command=btn_clear_click,
                               width=10)
        btn_clear.pack(padx=5, side=tk.LEFT)

        btn_go = ttk.Button(menu_frame, text='Открыть',
                            command=btn_go_click,
                            width=10)
        btn_go.pack(side=tk.LEFT)

        menu_frame.pack(expand=True, fill=tk.BOTH)

        # string info
        info_frame = ttk.Frame(self, padding=1, borderwidth=5,
                               relief='raised')
        msg_info = tk.StringVar()
        ttk.Label(master=info_frame, textvariable=msg_info,
                  background='#EBEBEB',
                  foreground="red", font=("Calibri", 12),
                  text="Информация").pack(side=tk.LEFT, expand=True, fill=tk.BOTH)
        info_frame.pack(expand=True, fill=tk.BOTH)

        self.pack()


# num_report = 0


def btn_clear_click():
    # print(result_frame.winfo_children())
    for item in result_frame.winfo_children():
        item.destroy()


def cmb_function(event):
    # print(event.widget.current() + 1)
    global num_report
    num_report = event.widget.current() + 1


def btn_go_click():
    date_begin = ds.entry.get()
    date_end = dp.entry.get()
    fte = one_hour_fte.get()
    export = export_excell_var.get()
    # print(f"c {date_begin} по {date_end}")
    # msg_info.set(f"c {date_begin} по {date_end} отчет № {num_report},"
    #              f" fte ={fte} , export = {export}")
    file_name = filedialog.askopenfilename()
    report_data = get_report(num_report=num_report, filename=file_name)

    param = {'df': report_data,
             'fte': fte,
             'reportnumber': num_report,
             'date_end': date_end,
             'date_begin': date_begin,
             'export_excell': export}
    # https: // www.youtube.com / watch?v = mop6g - c5HEY & list = PLZHIeS5WrW4IhlHiQq9fmlTu4F_YIc4ek & index = 10
    fr = get_data_test(**param)
    if num_report == 1:
        # result_frame.destroy()
        TableFrame(result_frame,fr)
    else:
        for item in fr.items():
            key, val = item
            # print(key, val)
            Component(result_frame, label_text=key, text_value=val)


# ttk.window(themename='dark')


App("Анализ отчётов", (800, 500), 'yeti')
# App("Анализ отчётов", (800, 700), 'superhero')
