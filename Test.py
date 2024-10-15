import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
import bd_unit

DB_MANAGER = bd_unit.DatabaseManager()

def get_data_lukoil(data_fromsql):
    if not data_fromsql:
        print("Нет данных для обработки.")
        return pd.DataFrame(columns=['Месяц', 'Неделя', 'Часы', 'fte'])  # Возвращаем пустой DataFrame

    # Определяем колонки для DataFrame
    col = ['Заявка', 'Подзадача', 'Часы', 'Регистрация', 'Квартал', 'Месяц', 'Содержание']
    df = pd.DataFrame(data_fromsql, columns=col)

    # Преобразуем столбец с датами в datetime формат
    df['Регистрация'] = pd.to_datetime(df['Регистрация'], errors='coerce')  # Обработка ошибок преобразования

    # Проверка на наличие NaT после преобразования
    if df['Регистрация'].isnull().any():
        print("Некоторые даты были некорректными и будут проигнорированы.")
        df = df.dropna(subset=['Регистрация'])

    # Добавляем столбец fte
    df['fte'] = df['Часы'] / 164

    # Находим начало месяца для каждой даты
    df['month_start'] = df['Регистрация'].dt.to_period('M').dt.to_timestamp()

    # Рассчитываем, сколько недель прошло с начала месяца
    df['Неделя'] = ((df['Регистрация'] - df['month_start']).dt.days // 7) + 1

    # Группируем по номеру месяца и неделе, суммируем Часы и fte
    weekly_summary = df.groupby(['Месяц', 'Неделя'])[['Часы', 'fte']].sum().reset_index()

    # Устанавливаем новый индекс для отображения
    weekly_summary.set_index(['Месяц', 'Неделя'], inplace=True)

    # print("Группированные данные:\n", weekly_summary)  # Отладочный вывод
    weekly_summary['fte'] = weekly_summary['fte'].round(2)
    return weekly_summary

def show_data_in_table(data):
    # Создаем главное окно
    root = tk.Tk()
    root.title("Таблица Lukoil")

    # Создаем Treeview для отображения таблицы
    tree = ttk.Treeview(root, columns=['Месяц', 'Неделя', 'Часы', 'fte'], show='headings')
    tree.pack(side='top', fill='both', expand=True)

    # Настраиваем заголовки столбцов
    for column in ['Месяц', 'Неделя', 'Часы', 'fte']:
        tree.heading(column, text=column)
        tree.column(column, anchor='center')

    # Вставляем данные в таблицу
    for index, row in data.iterrows():
        # Индекс представляется как кортеж
        month_week = index
        # Вставляем значения, формируя строку для отображения
        tree.insert("", "end", values=(month_week[0], month_week[1], row['Часы'], row['fte']))

    # Добавляем кнопку для выхода из приложения
    btn_exit = tk.Button(root, text="Выход", command=root.destroy)
    btn_exit.pack(side='bottom')

    # Запускаем главный цикл приложения
    root.protocol("WM_DELETE_WINDOW", lambda: on_closing(root))  # Обработка закрытия окна
    root.mainloop()

def on_closing(root):
    if messagebox.askokcancel("Выход", "Вы действительно хотите выйти?"):
        root.destroy()

# Получаем данные и выводим в таблицу
data_fromsql = DB_MANAGER.read_all_lukoil()
weekly_summary = get_data_lukoil(data_fromsql)

if not weekly_summary.empty:
    show_data_in_table(weekly_summary)
else:
    print("Нет данных для отображения.")
