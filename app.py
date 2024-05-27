from reports import *
from os import path
import tkinter as tk
from tkinter import (END, LEFT, WORD, Text,
                     filedialog, ttk, font)
from tkcalendar import DateEntry


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Анализ отчётов')
        self['background'] = '#EBEBEB'
        self.conf = {'width': 600, 'height': 800, 'padx': (10, 30), 'pady': 10}
        self.bold_font = font.Font(family='Helvetica', size=13, weight='bold')
        self.put_frames()

    def put_frames(self):
        self.add_head_frame = HeadFrame(self).grid(row=0, column=0, sticky='nsew')
        # self.add_table_frame = TableFrame(self).grid(row=1, column=0, sticky='nsew')


class HeadFrame(tk.Frame):
    def __init__(self, parent):
        super().__init__(parent)
        self['background'] = self.master['background']

    def init(self):
        self.lbl_path.config(text="")
        self.one_hour_fte.delete(0, END)
        self.one_hour_fte.insert(0, '164')

    def cmb_function(self):
        self.init()

    def btn_go_click(self):
        dp = self.Date_end.get()
        ds = self.Date_begin.get()
        fte = self.get_fte()
        export_excel = self.export_excell_var.get()
        file_name = 'errors'
        try:
            file_name = filedialog.askopenfilename()
            num_report = self.cmb.current() + 1

            report_data = reports.get_report(num_report=num_report, filename=file_name)
            param = {'df': report_data, 'fte': fte, 'reportnumber': num_report, 'date_end': dp,
                     'date_begin': ds,
                     'export_excell': export_excel}
            fr = reports.get_data(**param)

            if self.export_excell_var.get():  # Checbox
                save_file = path.dirname(file_name) + '/output.xlsx'
                self.lbl_path.config(text=f"Файл сохранился в {save_file}")
                fr.to_excel(save_file, index=True)
            # text.insert(5.0, f'{file_name}\n \n \n {fr}')
        except Exception as e:
            print(e)
            # text.insert(5.0, f"Не смог открыть файл {file_name}{e}")

    def put_widgets(self):
        self.lblzag = tk.Label(self, text="Анализ отчётов", font="italic 14 bold", background="gray", fg="lightgreen")

        self.cmb = ttk.Combobox(self, values=[items["name"] for items in reports], state="readonly", width=90)
        self.cmb.set('Выбор из списка отчетов')
        self.cmb.bind('<<ComboboxSelected>>', self.cmb_function)
        self.cmb["state"] = "readonly"

        self.btn_go = tk.Button(self, text='Открыть', command=self.btn_go_click, width=20)
        self.label1 = tk.Label(self, text="", font=("Helvetica", 16), background="gray", fg='white')
        self.lblC = tk.Label(self, text="Период с", font="italic 10 bold", background="gray", fg='white')

        self.Date_begin = DateEntry(self, date_pattern='YYYY-mm-dd')

        self.lblpo = tk.Label(self, text="ПО", bd=1, width=2, font="italic 10 bold", background="gray", fg='white')
        self.Date_end = DateEntry(self, width=15, relief="solid", date_pattern='YYYY-mm-dd')

        self.lbl_fte = tk.Label(self, text="FTE:", width=5, border=2, background="gray", fg='white')
        self.one_hour_fte = tk.Entry(self, width=5)

        self.export_excell_checkbox = tk.Checkbutton(self, text='Экспорт в Excel',
                                                     variable=export_excell_var, onvalue=1,
                                                     offvalue=0, background="gray")
        self.lbl_path = tk.Label(self, text="")
        self.columnconfigure((0, 1, 2, 3, 4, 5, 6), weight=1)
        self.rowconfigure((0, 1, 2, 3), weight=1)

        self.lblzag.grid(column=0, row=0, columnspan=7, sticky='nwe', pady=3, padx=5)

        self.cmb.grid(column=0, row=1, columnspan=6, sticky='nw', padx=10)
        self.btn_go.grid(column=6, row=1, sticky='ne', padx=10)
        self.label1.grid(column=0, row=2, columnspan=7, sticky='news', padx=5)

        self.lblC.grid(column=0, row=2, sticky='w', padx=10)
        self.Date_begin.grid(column=1, row=2, sticky='w')
        self.lblpo.grid(column=2, row=2, sticky='w')
        self.Date_end.grid(column=3, row=2, sticky='w')
        self.lbl_fte.grid(column=4, row=2, sticky='e')
        self.one_hour_fte.grid(column=5, row=2, sticky='w')
        self.export_excell_checkbox.grid(column=6, row=2, sticky='e', padx=10)

        self.lbl_path.grid(column=0, row=3, columnspan=7, sticky='nwe', padx=10)


app = App()
app.mainloop()
