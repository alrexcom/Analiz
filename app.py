import datetime
import tkinter as tk
from tkinter import (filedialog, font, messagebox)

import ttkbootstrap as ttk
from ttkbootstrap import DateEntry

from reports import (get_data_report, names_reports)
from univunit import Table, get_first_day_of_quarter, first_date_of_month

themes = ['cosmo', 'flatly', 'litera', 'minty', 'lumen', 'sandstone',
          'yeti', 'pulse', 'united', 'morph', 'journal', 'darkly',
          'superhero', 'solar', 'cyborg', 'vapor', 'simplex', 'cerculean']


def prompt_file_selection():
    try:
        return filedialog.askopenfilename()
    except Exception as e:
        messagebox.showinfo("Ошибка!", f"Не удалось выбрать файл: {e}")
        return None


class App(tk.Tk):
    def __init__(self, title, size, _theme):
        super().__init__()
        self.style = ttk.Style(theme=_theme)
        self['background'] = '#EBEBEB'
        self.bold_font = font.Font(family='Helvetica', size=13, weight='bold')
        self.title(title)

        self.geometry(f"{size[0]}x{size[1]}")

        self.fte_frame = None
        self.ds = None
        self.dp = None
        self.fte = None

        self.name_report = tk.StringVar()
        self.one_hour_fte = tk.IntVar()

        self.create_widgets()
        self.create_menu()
        self.table = Table(self)
        self.table.pack(expand=True, fill='both')

    def create_widgets(self):
        # Menu strip
        menu_frame = tk.Frame(self, padx=5, pady=5)
        self.ds = DateEntry(
            master=menu_frame,
            width=15,
            relief="solid",
            dateformat='%d-%m-%Y'
        )
        self.dp = DateEntry(menu_frame,
                            width=15,
                            relief="solid",
                            dateformat='%d-%m-%Y')

        self.fte_frame = tk.Frame(menu_frame)
        self.fte_frame.pack_forget()

        self.fte = ttk.Entry(self.fte_frame, width=5, font=("Calibri", 12),
                             textvariable=self.one_hour_fte)

        cmb_frame = ttk.Frame(self, padding=1)

        cmb = ttk.Combobox(cmb_frame, values=names_reports(),
                           state="readonly",
                           height=4,
                           font=("Calibri", 12),
                           textvariable=self.name_report
                           )
        cmb.set('Выбор из списка отчетов')
        cmb.bind('<<ComboboxSelected>>', self.cmb_function)
        cmb.pack(side=tk.LEFT, expand=True, fill=tk.X)

        cmb_frame.pack(fill=tk.X, pady=10, padx=10)

        self.ds.pack(side=tk.LEFT, padx=10)
        self.ds.entry.delete(0, tk.END)
        self.ds.entry.insert(0, get_first_day_of_quarter(current_date=datetime.date.today()))

        self.dp.pack(side=tk.LEFT, padx=10)

        ttk.Label(self.fte_frame, text="FTE:", width=5, anchor=tk.CENTER,
                  border=2, font=("Calibri", 12, 'bold'),
                  background='#B7DEE8', foreground='white').pack(side=tk.LEFT)

        self.fte.delete(0, tk.END)
        self.fte.insert(0, '164')
        self.fte.pack(side=tk.LEFT, padx=10)

        btn_go = ttk.Button(menu_frame, text='Открыть',
                            command=self.btn_go_click,
                            width=10)
        btn_go.pack(side=tk.LEFT)

        menu_frame.pack(fill=tk.X)

    def cmb_function(self, event):
        num_report = event.widget.current() + 1
        if num_report == 1:
            self.toggle_fte_frame(True)
        else:
            self.toggle_fte_frame(False)

    def toggle_fte_frame(self, show):
        if show:
            self.fte_frame.pack(side=tk.LEFT, padx=5)
        else:
            self.fte_frame.pack_forget()

    def btn_go_click(self):
        file_name = prompt_file_selection()
        if file_name:
            self.process_file(file_name)

    def process_file(self, file_name):
        try:
            param = self.get_params(file_name)
            fr = get_data_report(**param)
            self.table.configure_columns(fr['columns'])
            self.table.populate_table(fr['data'])
            self.update_window_size()
        except Exception as e:
            messagebox.showinfo("Ошибка!", f"Не смог обработать файл {file_name}: {e}")

    def get_params(self, file_name):
        return {
            'filename': file_name,
            'name_report': self.name_report.get(),
            'fte': self.one_hour_fte.get(),
            'date_end': self.dp.entry.get(),
            'date_begin': self.ds.entry.get(),
        }

    def update_window_size(self):
        # Рассчитываем общую ширину столбцов
        total_width = sum(self.table.tree.column(col, 'width') for col in self.table.tree['columns'])
        row_count = len(self.table.tree.get_children())
        height = (row_count * 25) + 180
        # Обновляем размер окна
        self.geometry(f'{total_width}x{height}')

    def create_menu(self):
        # Создание панели меню
        menubar = tk.Menu(self)

        # Создание выпадающего меню "Файл"
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Ввести рабочие дни", command=append_job_days)  # Добавление команды "Открыть"

        file_menu.add_separator()  # Добавление разделителя
        file_menu.add_command(label="Выход", command=self.quit)  # Добавление команды "Выход"

        # Добавление выпадающего меню "Файл" в панель меню
        menubar.add_cascade(label="Файл", menu=file_menu)

        # Повторите эти шаги, чтобы добавить другие выпадающие меню, если это необходимо

        # Отображение панели меню
        self.config(menu=menubar)


def append_job_days():
    # Функция добавления рабочих дней для получения FTE
    create_new_window()


def create_new_window():
    new_window = tk.Toplevel()
    new_window.geometry("500x200")
    new_window.title("Добавление рабочих дней для получения FTE")
    days_var = tk.IntVar()
    frame_item = tk.Frame(new_window)
    tk.Label(frame_item, text='Число рабочих дней', bg="#333333", fg="white", font=("Arial", 16)).grid(row=1,
                                                                                                       sticky="e")
    tk.Entry(frame_item, textvariable=days_var).grid(row=1, column=1, pady=20, padx=10, sticky="ew")

    tk.Label(frame_item, text='Месяц из даты', bg="#333333", fg="white", font=("Arial", 16)).grid(row=2, sticky="e")

    month_year = DateEntry(master=frame_item, width=15, relief="solid", dateformat='%d-%m-%Y')
    month_year.grid(row=2, column=1, pady=10, sticky="ew",padx=10)
    month_year.entry.delete(0, tk.END)
    month_year.entry.insert(0, get_first_day_of_quarter(current_date=datetime.date.today()))

    date_in = first_date_of_month(month_year.entry.get())

    tk.Button(frame_item, text='Добавить запись', command=lambda: save_days(job_days=days_var.get(), date_in=date_in),
              bg="#FF3399", fg="white", font=("Arial", 14)).grid(row=3, columnspan=2, pady=10)
    frame_item.pack()


def save_days(**params):
    # print(f"{params['date_in']}, число:{params['job_days']}")
    messagebox.showinfo("Сообщение", f"{params['date_in']}, число:{params['job_days']}")


if __name__ == '__main__':
    App("Анализ отчётов", (800, 500), 'yeti').mainloop()
