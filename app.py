from reports import *
from os import path
import tkinter as tk
from tkinter import (END, LEFT, WORD, Text,
                     filedialog, ttk, font)
from tkcalendar import DateEntry


class App(tk.Tk):
    def __init__(self, size):
        super().__init__()
        self.title('Анализ отчётов')
        self['background'] = '#EBEBEB'
        self.geometry(f"{size[0]}x{size[1]}")
        self.minsize(size[0], size[1])
        self.bold_font = font.Font(family='Helvetica', size=13, weight='bold')

        self.myframe = HeadFrame(self)

        # self.put_frames()
        self.mainloop()

    # def put_frames(self):
    #     HeadFrame(self)
    #     # .grid(row=0, column=0, sticky='nsew'))
    #     # self.add_table_frame = TableFrame(self).grid(row=1, column=0, sticky='nsew')


class Component(ttk.Frame):
    def __init__(self, parent, label_text, text_value):
        super().__init__(master=parent)

        # grid
        self.rowconfigure(0, weight=1)
        self.columnconfigure((0, 1), weight=1, uniform='a')
        ttk.Label(self, text=label_text, wraplength=300).grid(
            padx=10, row=0, column=0, sticky='w')
        ttk.Label(self, text=text_value, background="pink", width=100).grid(row=0, column=1, sticky='ew', padx=10)

        self.pack(expand=1, fill=tk.BOTH)

class HeadFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        # ttk.Label(self, background='gray').pack(expand=True, fill=tk.BOTH)
        # self['background'] = self.master['background']
        # self.place(relx=0, rely=0, relheight=1, relwidth=0.5)
        # self.config(width=50)
        self.create_menu()
        self.pack()

    def create_menu(self):
        # Date_farme=tk.Frame(self, bg='yellow', relief=tk.RAISED)
        # date_begin = tk.Entry(self, width=15, bd=1, bg="white", relief="solid", font="italic 10 bold")
        lblzag = tk.Label(self, text="Анализ отчётов", font="italic 14 bold", background="gray", fg="lightgreen")

        lbl_frame = ttk.Frame(self, padding=10)

        #
        date_begin = DateEntry(lbl_frame, date_pattern='YYYY-mm-dd')
        date_end = DateEntry(lbl_frame, width=15, relief="solid", date_pattern='YYYY-mm-dd')
        lblpo = tk.Label(lbl_frame, text="ПО", bd=1, width=2, font="italic 10 bold", background="gray", fg='white')
        lblC = tk.Label(lbl_frame, text="Период с", font="italic 10 bold", background="gray", fg='white')
        export_excell_checkbox = tk.Checkbutton(lbl_frame, text='Экспорт в Excel',
                                                onvalue=1,
                                                offvalue=0)

        lbl_fte = tk.Label(lbl_frame, text="FTE:", width=5, border=2, background="gray", fg='white')
        one_hour_fte = tk.Entry(lbl_frame, width=5)
        btn_go = tk.Button(lbl_frame, text='Открыть', command=btn_go_click, width=15)
        btn_clear = tk.Button(lbl_frame, text='Очистить', command=btn_go_click, width=15)

        # cmbfrm = ttk.Frame(self)
        cmb = ttk.Combobox(self, values=[items["name"] for items in reports], state="readonly", width=100, height=2,
                           font=('Helvetica', 11))
        cmb.set('Выбор из списка отчетов')
        cmb.bind('<<ComboboxSelected>>', cmb_function)
        cmb["state"] = "readonly"

        info_farame = ttk.Frame(self, borderwidth=2, relief=tk.SOLID)
        lbl_info = tk.Label(info_farame, text="Информация")

        text = Text(bg="darkgreen", fg="white", wrap=WORD)
        text.tag_config('title', justify=LEFT)
        # from main import export_excell_var

        result_frame= ttk.Frame(self)
        label1 = tk.Label(result_frame, text="", font=("Helvetica", 16), background="gray", fg='white')
        #  Каркас
        # self.columnconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)
        self.columnconfigure(0, weight=1)
        self.rowconfigure((0, 1, 2, 3), weight=1)
        self.rowconfigure(4, weight=50)

        lblzag.grid(column=0, row=0, sticky='nwe', pady=3, padx=5)
        cmb.grid(column=0, row=1, sticky='nw', padx=10)

        # label1.grid(column=0, row=2, columnspan=5, sticky='news', padx=5)

        lbl_frame.grid(column=0, row=2, sticky='nw')
        lblC.pack(side=LEFT)
        date_begin.pack(side=LEFT, padx=5)
        lblpo.pack(side=LEFT)
        date_end.pack(side=LEFT, padx=5)
        lbl_fte.pack(side=LEFT, padx=10)
        one_hour_fte.pack(side=LEFT)
        export_excell_checkbox.pack(side=LEFT, padx=10)
        btn_clear.pack(side=LEFT, padx=5)
        btn_go.pack(side=LEFT)

        info_farame.grid(column=0, row=3, sticky='nwe', padx=10)
        lbl_info.pack(side="left", padx=5, fill="x")

        result_frame.grid(column=0, row=4, sticky='news', padx=5, pady=5)
        # label1.pack(side="left", padx=5, fill="x")
        txt = "Общее количество незакрытых заявок по сопровождению на начало периода"
        Component(result_frame, label_text=txt, text_value='3')
        Component(result_frame, label_text="Сервис менеджер", text_value="Тапехин Алексей")
        Component(result_frame, label_text="Общее количество закрытых за период заявок по сопровождению"
                  , text_value='9')

def btn_go_click():
    pass


def cmb_function(event):
    pass


App((800, 700))
