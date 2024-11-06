# import pandas as pd
from datetime import datetime

import numpy as np

from univunit import (Univunit, pd)
# , save_to_json)
import bd_unit

DB_MANAGER = bd_unit.DatabaseManager()
# (pd, convert_date)
# import Univunit as u

reports = [
    {
        "name": "Отчёт Данные по ресурсным планам и списанию трудозатрат сотрудников за период",
        "header_row": 0,
        "reportnumber": 1,
        "data_columns": ["Дата"]
    },
    {
        "name": "Контроль заполнения факта за период",
        "header_row": 1,
        "reportnumber": 2,
        "data_columns": ["Дата"]
    },
    {
        "name": "Сводный список запросов  для SLA (Сопровождение)  С0134-КИС",
        "header_row": 2,
        "reportnumber": 3,
        "data_columns": ["Дата регистрации", "Крайний срок решения", "Дата решения", "Дата закрытия",
                         "Дата последнего назначения в группу"],
        "status": ["В ожидании", "Выполнено", "Закрыто", "Проект изменения", "Решен", "Назначен", "Выполняется",
                   "Планирование изменения", "Выполнение изменения", "Экспертиза решения", "Согласование изменения",
                   "Автроизация изменения"],  # Отмена  убрано
        "support": False
    },
    {
        "name": "Сводный список запросов  для SLA (Поддержка) Т0133-КИС",
        "header_row": 2,
        "reportnumber": 4,
        "data_columns": ["Дата регистрации", "Крайний срок решения", "Дата решения", "Дата закрытия",
                         "Дата последнего назначения в группу"],
        "status": ["В ожидании", "Выполнено", "Закрыто", "Проект изменения", "Решен", "Назначен", "Выполняется",
                   "Планирование изменения", "Выполнение изменения", "Экспертиза решения", "Согласование изменения",
                   "Автроизация изменения"],  # Отмена  убрано
        "support": True
    },
    {
        "name": "Лукойл: Отчет по запросам и задачам (долго открывается)",
        "reportnumber": 5,
        "header_row": 3,
        "data_columns": ["Дата Выполнения работ", "Время назначения задачи"],
        "headers": ['ID инцидента/ЗИ', 'Исполнитель  по задаче', 'Трудозатраты по задаче (десят. часа)',
                    'Время назначения задачи', 'Категория инцидента', 'Статус', 'Содержание задачи'],
        'order by': 'Исполнитель  по задаче'
    },

]

SLA_POINTS = [
        ("Открыто на начало периода", 0,
         "( tur3 ) Общее количество незакрытых заявок по сопровождению на начало периода ", 1),
        ("Открыто на начало периода", 1, "( tur3 ) Количество перешедших с прошлого периода заявок на поддержку ", 8),
        ("Зарегистрировано в период", 1, "Общее количество зарегистрированных заявок по поддержке", 1),
        ("Зарегистрировано в период", 0, "(tur4) Общее количество зарегистрированных заявок по сопровождению ", 2),
        ("Зарегистрировано в период", 1, "(tur4) Количество зарегистрированных заявок по поддержке", 9),
        ("Выполнено в период", 0, "3 п / п Общее количество закрытых за период заявок по сопровождению ", 3),
        ("Выполнено в период", 1, "Общее количество выполненных заявок по поддержке", 2),
        ("Выполнено с просрочкой в период", 0,
         "4 (tur1) Общее количество закрытых за период заявок по сопровождению c нарушением SLA", 4),
        ("Выполнено с просрочкой в период", 1,
         "(tur1) Количество закрытых заявок на поддержку с нарушением сроков заявок", 4),
        ("Открыто на конец периода", 0, "5 п/п Общее количество незакрытых заявок по сопровождению на конец периода ",
         5),
        ("Открыто на конец периода с просрочкой", 0,
         "6 (tur2) Количество заявок за период с превышением времени реакции по сопровождению ", 6),
        ("Открыто на конец периода с просрочкой", 1, "(tur2) Количество незакрытых заявок, с нарушением срока", 6),
        ("Просрочено в период", 0, "7 п/п Количество заявок за период с превышением срока выполнения по сопровождению ",
         7),
        (
        "Комментарий к нарушению SLA", 0, "8 п/п Количество заявок за период с превышением времени диспетчеризации ", 8)
    ]

def report1(**param):
    """
    Отчёт Данные по ресурсным планам и списанию трудозатрат сотрудников за период

    """

    df = param['df']

    df_lukoil = DB_MANAGER.read_sum_lukoil(ds=param['date_begin'], dp=param['date_end'])

    date_begin = datetime.strptime(param['date_begin'], '%Y-%m-%d').strftime('%m.%Y')
    date_end = datetime.strptime(param['date_end'], '%Y-%m-%d').strftime('%m.%Y')

    df = df[(df['Дата'] >= date_begin) & (df['Дата'] <= date_end)]
    df = pd.merge(df, df_lukoil, on=['Проект', 'Пользователь', 'Дата'], how='left')
    # df = df.replace([np.nan, -np.inf], 0)
    user = 'Тапехин Алексей Александрович'

    fte = param['fte']
    fact_sum = 'Фактические трудозатраты (час.) (Сумма)'
    headers = ['Дата', 'Проект', 'Пользователь', 'Лукойл, час.', 'Фактические трудозатраты (час.) (Сумма)', 'План, FTE',
               'Заявок лукойл']

    # fr = df[headers].loc[(df['Менеджер проекта'] == user) | (df['Пользователь'] == user)]

    fr = df.loc[(df['Менеджер проекта'] == user) | (df['Пользователь'] == user), headers]
    hours_plan = fte * fr['План, FTE']
    hours_plan_week = round(hours_plan / 4, 2)
    fr['План ч. мес.'] = round(hours_plan, 2)
    fr['План ч. нед.'] = hours_plan_week
    # f"{round(hours_plan, 2)}  / {round(hours_plan/4, 2)}"

    fact_fte = round((fr[fact_sum] / fte), 2)

    fr['Факт, FTE'] = fact_fte
    fr['Лукойл, FTE'] = round((fr['Лукойл, час.'] / fte), 2)
    hours_remain = round(hours_plan - fr[fact_sum], 2)
    fr['Остаток часов'] = hours_remain
    fr['Остаток FTE'] = round((fr['Остаток часов'][fr['Остаток часов'] > 0] / fte), 2)
    fr = fr.sort_values('Пользователь')
    fr = fr.replace([np.nan, -np.inf], '')

    # fr = fr[fr['Факт, FTE'] > 0]
    fr = fr.reindex(
        columns=['Дата', 'Проект', 'Пользователь', 'План ч. мес.', 'План ч. нед.', 'Заявок лукойл', 'Лукойл, час.',
                 'Фактические трудозатраты (час.) (Сумма)',
                 'План, FTE', 'Лукойл, FTE', 'Факт, FTE', 'Остаток часов', 'Остаток FTE'])
    # Преобразование DataFrame в записи NumPy
    data = fr.to_records(index=False)
    columns = [{"name": col} for col in fr.columns]

    return {'columns': columns, 'data': data.tolist()}


def report2(**params):
    """
    Контроль заполнения факта за период

    """

    df = params['df']

    # Оптимизированная фильтрация с использованием .loc[]
    frm = df.loc[
        (df['ФИО'] == 'Тапехин Алексей Александрович') |
        (df['Проект'] == 'Т0133-КИС "Производственный учет и отчетность"'),
        ['Проект', 'ФИО', 'Дата', 'Трудозатрады за день']
    ]

    date_range = pd.date_range(start=params['date_begin'], end=params['date_end'], freq='D')
    data_reg = frm[frm['Дата'].isin(date_range)]
    result = data_reg.sort_values(['ФИО', 'Проект', 'Дата'])

    # Создание списка столбцов
    columns = [{"name": col} for col in result.columns]
    # Преобразование в список кортежей
    data = []
    for i in result.values:
        data.append(tuple(i))

    return {'columns': columns, 'data': data}


def calc(sum1, sum6, sum7, sum8, prefix):
    s7 = sum7.get(prefix, 0)
    s8 = sum8.get(prefix, 0)
    s6 = sum6.get(prefix, 0)
    s1 = sum1.get(prefix, 0)

    if s8 + s1 == 0:
        s8 = 1

    slap = round((1 - (s6 + s7) / (s8 + s1)) * 100, 2)
    return slap


def report_sla(**params):
    """
    Сводный список запросов для SLA
    """
    try:

        # status = params['status']
        # filtered_mdf = params["df"].query("Статус in @status")

        # status переменная нужна в запросе

        # Объединённая фильтрация
        # filtered_mdf = mdf.query("Услуга == 'КИС \"Производственный учет и отчетность\"' and Статус in @status")
        # filtered_mdf = mdf.query(Статус in @status")

        filtered_mdf = params["df"]
        ds = params["date_begin"]
        dp = params["date_end"]
        filtered_mdf["dateisnull"] = "0"
        filtered_mdf.loc[filtered_mdf['Дата закрытия'].isna(), "dateisnull"] = -1

        date_filter = (
                (filtered_mdf['Статус'] != "Отменено") & (filtered_mdf['Дата регистрации'] <= dp) &
                (
                        (filtered_mdf["dateisnull"] == "-1")
                        | (filtered_mdf['Дата регистрации'] >= ds)
                        | (filtered_mdf['Дата закрытия'] >= ds)
                )
        )

        filtered_mdf = filtered_mdf[date_filter]
        # date_filter = (
        #                   ((filtered_mdf['Дата регистрации'] >= params["date_begin"]) &
        #         (filtered_mdf['Дата регистрации'] <= params["date_end"]))
        #                | ((filtered_mdf['Открыто на начало периода'] == 1)
        #                ) & (filtered_mdf['Дата закрытия'] >= params["date_begin"]))

        # Классификация запросов
        filtered_mdf["П2С"] = "П"
        filtered_mdf.loc[filtered_mdf["Тип запроса"] == 'Нестандартное', "П2С"] = "СДОП"
        filtered_mdf.loc[filtered_mdf["Тип запроса"] == 'Стандартное без согласования', "П2С"] = "С"

        if params['support']:
            filtered_mdf = filtered_mdf[filtered_mdf["П2С"] == "П"]
        else:
            filtered_mdf = filtered_mdf[(filtered_mdf["П2С"] == "С") | (filtered_mdf["П2С"] == "СДОП")]

        # Фильтрация по датам
        date_filter = ((filtered_mdf['Дата регистрации'] >= ds) &
                       (filtered_mdf['Дата регистрации'] <= dp))
        data_reg = filtered_mdf[date_filter]

        # Фильтрация просроченных запросов
        overdue_filter = ((filtered_mdf['Статус'] != 'Закрыто') & (filtered_mdf['Просрочено в период'] > 0)
                          & (filtered_mdf['Дата закрытия'] <= params["date_end"]) &
                          (filtered_mdf['Дата закрытия'] >= params["date_begin"]))
        data_prosr = filtered_mdf[overdue_filter]

        # Столбцы для вывода
        columns = [
            {"name": "№ п/п", "width": 100, "anchor": 'center'},
            {"name": "Наименование", "width": 600, "anchor": 'w'},
            {"name": "Значение", "width": 100, "anchor": 'center'}
        ]

        # Обработка нарушений SLA  убиваем Nan - пустые строки
        errsla = filtered_mdf[["Номер запроса", "Исполнитель",
                               "Комментарий к нарушению SLA"]].dropna(subset=["Комментарий к нарушению SLA"])

        data = errsla.apply(lambda x: (
            f"ERR SLA: запрос {x['Номер запроса']} / {x['Исполнитель']} :{x['Комментарий к нарушению SLA']}", 1),
                            axis=1).tolist()

        # Добавление поддержки SLA
        # if params['support']:
        #     # params["vid"] = 1
        #     ss = sla_support(mdf=filtered_mdf, data_reg=data_reg, data_prosr=data_prosr, columns=columns,
        #                      date_end=dp, date_begin=ds)
        #     for key, val in ss.items():
        #         data.append((key, val))
        # else:
        params["vid"] = 0
        if params['support']:
            params["vid"] = 1

        params["data_reg"] = data_reg
        data = get_data_sla_3(**params)

        return {"columns": columns, 'data': data}

    except Exception as e:
        print(f"Ошибка вычисления {e}")


def sla_support(**params):
    """
     Данные по SLA Поддержка
    """

    par = get_data_sla(**params)
    mdf = params['data_reg']

    sum1 = par["sum1"].get('П', 0)
    sum2 = par["sum2"].get('П', 0)

    inzindent = mdf[['П2С', 'Дата регистрации', 'Статус',
                     'Просрочено в период', 'Дата закрытия']].loc[mdf['Тип запроса'] == 'Инцидент']

    data_inzindent = inzindent[(inzindent['Дата регистрации'] <= params['date_end'])
                               & (inzindent['Дата регистрации'] >= params['date_begin'])]
    d = data_inzindent[['Дата регистрации', 'П2С']].groupby('П2С').count()
    sum3 = d['Дата регистрации'].get('П', 0)

    data_inzindent = inzindent[(inzindent['Статус'] == 'Закрыто') & (inzindent['Просрочено в период'] > 0)
                               & (inzindent['Дата закрытия'] <= params['date_end'])
                               & (inzindent['Дата закрытия'] >= params['date_begin'])]

    d = data_inzindent[['Просрочено в период', 'П2С']].groupby(['П2С']).count()
    sum4 = d['Просрочено в период'].get('П', 0)

    sum5 = par["sum5"].get('П', 0)

    tur1 = par["sum6"]['Выполнено с просрочкой в период'].get('П', 0)
    tur2 = par["sum7"]['Просрочено в период'].get('П', 0)
    tur3 = par["sum8"]["Открыто на начало периода"].get('П', 0)
    tur4 = sum1

    slap = round((1 - (tur1 + tur2) / (tur3 + sum1)) * 100, 2)

    ss = {
        "  SLA по поддержке ": slap,
        "1. Общее количество зарегистрированных заявок": sum1,
        "2. Общее количество выполненных заявок": sum2,
        "3. Общее количество зарегистрированных заявок за отчетный период, имеющих категорию «Инцидент»": sum3,
        "4. Количество заявок за период с превышением срока выполнения, имеющих категорию «Инцидент»": sum4,
        "5. Количество заявок за период с превышением времени реакции по поддержке": sum5,
        "6. (TUR1)  Количество закрытых заявок на поддержку с нарушением сроков заявок": tur1,
        "7. (TUR2)  Количество незакрытых заявок, с нарушением срока": tur2,
        "8. (TUR3)  Количество перешедших с прошлого периода заявок на поддержку": tur3,
        "9. (TUR4)  Количество зарегистрированных заявок по поддержке": tur4
    }

    return ss


def get_data_sla(**par):
    data_reg = par["data_reg"]

    # tur4 Количество заявок, зарегистрированных в отчетном периоде
    sum1 = data_reg.groupby(['П2С'])["Зарегистрировано в период"].sum()

    sum2 = data_reg.groupby(['П2С'])['Выполнено в период'].sum()
    sum5 = data_reg.groupby(['П2С'])['Просроченное время реакции, часов'].count()
    sum6 = data_reg[['Выполнено с просрочкой в период', 'П2С']].groupby(['П2С']).sum()
    # tur2
    sum7 = par["data_prosr"][['Просрочено в период', 'П2С']].groupby(['П2С']).sum()
    # tur3 = "Открыто на начало периода"
    sum8 = par["mdf"][["Открыто на начало периода", 'П2С']].groupby(['П2С']).sum()
    return {"sum1": sum1, "sum2": sum2, "sum5": sum5, "sum6": sum6, "sum7": sum7, "sum8": sum8}


def get_data_sla_3(**par):
    # Указываем значение `vid`, определяющее тип данных: 0 для Сопровождение, 1 для Поддержка
    vid = par.get("vid", 0)

    data_reg = par["data_reg"]
    # full_data = par["df"]

    # Создаем пустой список для записи результатов
    data_list = []

    # Обрабатываем каждый SLA point
    for name, sla_vid, dote_column, order in SLA_POINTS:
        if vid == sla_vid:
            # Проверка, существует ли колонка с именем `dote_column` в `data_reg`
            if name in data_reg.columns:
                # Проверка, является ли столбец числовым
                if data_reg[name].dtype in ['int64', 'float64']:
                    value = data_reg[name].sum().astype(float)
                else:
                    value = 0
            else:
                value = -1  # Если столбец не найден

            # Добавляем данные в список
            data_list.append({"Порядок": order, "Наименование": name, "Значение": value})

    # Подсчитываем количество значений "Инцидент" в столбце "Тип запроса" для `vid == 1`
    if vid == 1:
        # Фильтруем строки, где "Тип запроса" равен "Инцидент"
        incident_count = (data_reg["Тип запроса"] == "Инцидент").sum()
        # Добавляем информацию о количестве инцидентов
        data_list.append({
            "Порядок": 3,  # Указываем высокий порядок, чтобы информация об инцидентах была в конце
            "Наименование": "Общее количество зарегистрированных заявок за отчетный период, имеющих категорию «Инцидент»",
            "Значение": incident_count
        })
        # todo Проверить
        # incident_data = full_data[full_data["Тип запроса"] == "Инцидент"
        #                           and full_data['Статус']=='Закрыт'
        #                           and full_data['"Просрочено в период"']=='1']
        # # Подсчитываем количество инцидентов
        # if incident_data:
        #     incident_count = len(incident_data)
        # else:
        #     incident_count=0
        #
        #
        # data_list.append({
        #     "Порядок": 4,  # Указываем высокий порядок, чтобы информация об инцидентах была в конце
        #     "Наименование": "Количество заявок за период с превышением срока выполнения,имеющих категорию «Инцидент»",
        #     "Значение": incident_count
        # })

    # Преобразуем список словарей в DataFrame
    result_df = pd.DataFrame(data_list)

    # Сортируем DataFrame по столбцу "Порядок"
    result_df = result_df.sort_values(by="Порядок").reset_index(drop=True)

    # Возвращаем результат в виде списка записей
    return result_df.to_records(index=False).tolist()


# def get_data_sla_2(**par):
#     # Указываем значение `vid`, определяющее тип данных: 0 для Сопровождение, 1 для Поддержка
#     vid = par.get("vid", 0)  # Используем `par.get` для более гибкого определения `vid`
#
#     # SLA points с условием выбора словаря в зависимости от значения `vid`
#     sla_points = [
#         {"Открыто на начало периода": {
#             0: ("Открыто на начало периода",
#                 "( tur3 ) Общее количество незакрытых заявок по сопровождению на начало периода ", 1),
#             1: (
#                 "Открыто на начало периода",
#                 "( tur3 ) Количество перешедших с прошлого периода заявок на поддержку ", 8)
#         }},
#         {"Зарегистрировано в период": {
#             1: ("Зарегистрировано в период", " Общее количество зарегистрированных заявок по поддержке", 1)
#         }},
#         {"Зарегистрировано в период": {
#             0: ("Зарегистрировано в период", "(tur4) Общее количество зарегистрированных заявок по сопровождению ", 2),
#             1: ("Зарегистрировано в период", "(tur4) Количество зарегистрированных заявок по поддержке", 9)
#         }},
#         {'Выполнено в период': {
#             0: ("Выполнено в период", "3 п / п Общее количество закрытых за период заявок по сопровождению ", 3),
#             1: ("Выполнено в период", " Общее количество выполненных заявок по поддержке", 2)
#         }},
#         {'Выполнено с просрочкой в период': {
#             0: ("Выполнено с просрочкой в период",
#                 "4 (tur1) Общее количество закрытых за период заявок по сопровождению c нарушением SLA", 4),
#             1: ("Выполнено с просрочкой в период",
#                 "(tur1) Количество закрытых заявок на поддержку с нарушением сроков заявок", 4)
#         }},
#         {"Открыто на конец периода": {
#             0: (
#                 "Открыто на конец периода",
#                 "5 п/п Общее количество незакрытых заявок по сопровождению на конец периода ",
#                 5)
#         }},
#         {"Открыто на конец периода с просрочкой": {
#             0: ("Открыто на конец периода с просрочкой",
#                 "6 (tur2) Количество заявок за период с превышением времени реакции по сопровождению ", 6),
#             1: ("Открыто на конец периода с просрочкой", "(tur2) Количество незакрытых заявок, с нарушением срока", 6)
#         }},
#         {'Просрочено в период': {
#             0: (
#                 "Просрочено в период",
#                 "7 п/п Количество заявок за период с превышением срока выполнения по сопровождению ",
#                 7)
#         }},
#         {"Комментарий к нарушению SLA": {
#             0: (
#                 "Комментарий к нарушению SLA",
#                 "8 п/п Количество заявок за период с превышением времени диспетчеризации ",
#                 8)
#         }}
#     ]
#
#     # Загрузка данных из `par`
#     data_reg = par["data_reg"]
#
#     # Создаем пустой список для записи результатов
#     data_list = []
#
#     # Обрабатываем каждый SLA point
#     for point in sla_points:
#         for dote, subpoints in point.items():
#             if vid in subpoints:
#                 # Получаем name (для отображения), dote_column (для вычислений) и порядок
#                 name, dote_column, order = subpoints[vid]
#
#                 # Проверка, существует ли колонка с именем `dote_column` в `data_reg`
#                 if name in data_reg.columns:
#                     # Проверка, является ли столбец числовым
#                     if data_reg[name].dtype in ['int64', 'float64']:
#                         value = data_reg[name].sum().astype(float)
#                     else:
#                         value = 0
#                 else:
#                     value = -1  # Используем значение -1, если столбец не найден
#
#                 # Добавляем данные в список
#                 data_list.append({"Порядок": order, "Наименование": name, "Значение": value})
#
#     # Подсчитываем количество значений "Инцидент" в столбце "Тип запроса"
#     if "Тип запроса" in data_reg.columns and vid == 1:
#         # Фильтруем строки, где "Тип запроса" равен "Инцидент"
#         incident_data = data_reg[data_reg["Тип запроса"] == "Инцидент"]
#
#         # Подсчитываем количество инцидентов
#         incident_count = len(incident_data)
#
#         # Добавляем данные о количестве и сумме инцидентов в список
#         data_list.append(
#             {"Порядок": 3, "Наименование": "Общее количество зарегистрированных заявок за отчетный период, "
#                                            "имеющих категорию «Инцидент»", "Значение": incident_count})
#
#     # Преобразуем список словарей в DataFrame
#     result_df = pd.DataFrame(data_list)
#
#     # Сортируем DataFrame по столбцу "Порядок"
#     result_df = result_df.sort_values(by="Порядок").reset_index(drop=True)
#
#     # Возвращаем результат в виде списка записей
#     return result_df.to_records(index=False).tolist()


# def sla_sopr(**params):
#     """
#       Данные по SLA Сопровождение
#      """
#     # mdf = params["mdf"]
#     # overdue_filter = ((mdf['Дата закрытия'] <= params["date_end"]) &
#     #                   (mdf['Дата закрытия'] >= params["date_begin"]))
#     # params["data_reg"] = mdf[overdue_filter]
#
#     # params["data_reg"] = params["mdf"]
#
#     # data_reg = params["data_reg"]
#     # data_reg = params["mdf"]
#     params["vid"]=0
#     par = get_data_sla(**params)
#     sum_reg_in_period = par['sum_reg_in_period']
#     sum_vip_v_period = par['sum_vip_v_period']
#     sum_prosr_time_reaction = par['sum_prosr_time_reaction']
#     sum_vip_v_period_prosr = par['sum_vip_v_period_prosr']
#     sum_prosr_v_peeriod = par['sum_prosr_v_peeriod']
#     sum_open_on_begin_period = par['sum_open_on_begin_period']
#     sum_open_on_begin_period_prosr = par['sum_open_on_begin_period_prosr']
#
#     sum1 = par["sum1"]
#     sum2 = par["sum2"]
#
#     # p5 = "Открыто на конец периода с просрочкой"
#     # sum51 = data_reg.groupby(['П2С'])[p5].sum()
#
#     tur3 = "Открыто на начало периода"
#     sum8 = par["sum8"]
#
#     tur1 = 'Выполнено с просрочкой в период'
#     sum6 = par["sum6"]
#
#     tur2 = 'Просрочено в период'
#     sum7 = par["sum7"]
#
#     slac = calc(sum1, sum6[tur1], sum7[tur2], sum8[tur3], 'С')
#
#     # ss = {
#     #     "9 Уровень исполнения SLA общий": round(100 * (sum2.sum() - sum6[tur1].sum()) / sum2.sum(), 2),
#     #     "       поддержки": slap,
#     #     "       сопровождения": slac,
#     #
#     #     "1 ( tur3 ) Общее количество незакрытых заявок по сопровождению/поддержке на начало периода ": sum8[
#     #         tur3].sum(),
#     #     "       tur3 поддержка": sum8[tur3].get('П', 0),
#     #     "       tur3 сопровождение": sum8[tur3].get('С', 0),
#     #
#     #     "2 (tur4) Общее количество зарегистрированных заявок по сопровождению/поддержке ": sum1.sum(),
#     #     "       tur4 поддержка": sum1.get('П', 0),
#     #     "       tur4 сопровождение": sum1.get('С', 0),
#     #
#     #     "3 п/п Общее количество закрытых за период заявок по сопровождению/поддержке ": sum2.sum(),
#     #     "       p3 поддержка": sum2.get('П', 0),
#     #     "       p3 сопровождение": sum2.get('С', 0),
#     #
#     #     "4 (tur1) Общее количество закрытых за период заявок по сопровождению/поддержке c нарушением SLA": sum6[
#     #         tur1].sum(),
#     #     "       tur1 поддержка": sum6[tur1].get('П', 0),
#     #     "       tur1 сопровождение": sum6[tur1].get('С', 0),
#     #
#     #     "5 п/п Общее количество незакрытых заявок по сопровождению/поддержке на конец периода": sum51.sum(),
#     #     "       p5 поддержка": sum51.get('П', 0),
#     #     "       p5 сопровождение": sum51.get('С', 0),
#     #
#     #     "7 (tur2) Количество заявок за период с превышением срока выполнения ":
#     #         sum7[tur2].sum(),
#     #     "       по поддержке": sum7[tur2].get('П', 0),
#     #     "       по сопровождению": sum7[tur2].get('С', 0)
#     # }
#
#     ss = {
#         "9 Уровень исполнения SLA общий": round(100 * (sum2.sum() - sum6[tur1].sum()) / sum2.sum(), 2),
#         "       сопровождения": slac,
#
#         "1 ( tur3 ) Общее количество незакрытых заявок по сопровождению/поддержке на начало периода ":
#             sum8[tur3].get('С', 0),
#         # "       tur3 сопровождение": sum8[tur3].get('С', 0),
#
#         "2 (tur4) Общее количество зарегистрированных заявок по сопровождению/поддержке ": sum1.get('С', 0),
#         # "       tur4 сопровождение": sum1.get('С', 0),
#
#         "3 п/п Общее количество закрытых за период заявок по сопровождению/поддержке ": sum2.get('С', 0),
#         # "       p3 сопровождение": sum2.get('С', 0),
#
#         "4 (tur1) Общее количество закрытых за период заявок по сопровождению/поддержке c нарушением SLA":
#             sum6[tur1].get('С', 0),
#         # "       tur1 сопровождение": sum6[tur1].get('С', 0),
#
#         "5 п/п Общее количество незакрытых заявок по сопровождению/поддержке на конец периода": sum51.get('С', 0),
#         # "       p5 сопровождение": sum51.get('С', 0),
#
#         "7 (tur2) Количество заявок за период с превышением срока выполнения ":
#             sum7[tur2].get('С', 0),
#         # "       по сопровождению": sum7[tur2].get('С', 0)
#     }
#
#     return ss


def get_data_report(**params):
    report_data = get_report(**params)
    reportnumber = report_data[0]

    params['df'] = report_data[1]
    params['date_begin'] = Univunit.convert_date(params['date_begin'])
    params['date_end'] = Univunit.convert_date(params['date_end'])

    report = reports[reportnumber - 1]
    params['support'] = bool(report.get('support'))
    # Если ключа статус нет
    if report.get('status'):
        params['status'] = report.get('status', 'статус не указан')
    # if report.get('support'):
    #     params['support'] = bool(report.get('support', False))
    if report.get('headers'):
        params['headers'] = report.get('headers')
    if report.get('order by'):
        params['order by'] = report.get('order by')

    report_functions = {
        1: report1,
        2: report2,
        3: report_sla,
        4: report_sla,
        5: report_lukoil,
    }

    frm = report_functions.get(reportnumber)
    # if frm is None:
    #     raise ValueError("Некорректный номер отчета")
    return frm(**params)


def get_report(**params):
    for items in reports:
        if params['name_report'] == items['name']:
            result = (items['reportnumber'],
                      pd.read_excel(params['filename'], header=items['header_row'], parse_dates=items['data_columns'],
                                    date_format='%d.%m.%Y'))
            return result


def names_reports():
    return [items["name"] for items in reports]


def report_lukoil(**params):
    # data_=("rec1", "rec2")

    df = params['df']
    # 'Исполнитель  по задаче', 'Трудозатраты по задаче (десят. часа)'
    # promeg = df.groupby(['Исполнитель  по задаче']).agg({'Трудозатраты по задаче (десят. часа)': 'sum'})
    # fr = df.groupby(['Исполнитель  по задаче']).agg({'Дата': 'max',  'Трудозатраты по задаче (десят. часа)': 'sum'})
    # columns = [{"name": "Код"}, {"name": "Дата"}]
    # data = df[params['headers']].to_records(index=False)
    # df[params['headers']].to_csv('data.csv', header=True, index=False, sep=';')

    # save_to_json(data=df[params['headers']], filename='my.json')

    # Добавляем столбец во фрэйм
    df['vid'] = np.where(df['Исполнитель  по задаче'] != 'Тапехин Алексей Александрович', "Поддержка", 'Сопровождение')

    # Копируем в отдельный фрейм данные
    kk = df.groupby(['vid', 'Исполнитель  по задаче']).agg(
        {
            'ID инцидента/ЗИ': "count",
            'Трудозатраты по задаче (десят. часа)': "sum",
        }
    ).copy()

    kk['sum_all'] = (kk['Трудозатраты по задаче (десят. часа)'].sum())
    kk['sum_vid'] = (kk.groupby("vid")['Трудозатраты по задаче (десят. часа)'].transform('sum'))
    kk['count_vid'] = (kk.groupby("vid")['Трудозатраты по задаче (десят. часа)'].transform('count'))

    kk['sum_ispol'] = (kk.groupby('Исполнитель  по задаче')['Трудозатраты по задаче (десят. часа)'].transform('sum'))
    # kk['count_vid'] = (kk.groupby('Исполнитель  по задаче')['Трудозатраты по задаче (десят. часа)'].transform('count'))
    data = kk.sort_values(["vid", 'Исполнитель  по задаче'])
    columns_ = data.columns.to_list()

    columns = [{"name": col} for col in columns_]

    return {'columns': columns, 'data': data.to_records(index=False)}


def get_data_lukoil(data_fromsql):
    # if not data_fromsql:
    #     print("Нет данных для обработки.")
    #     return pd.DataFrame(columns=['Месяц', 'Неделя', 'Часы', 'fte'])  # Возвращаем пустой DataFrame

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

    values = []
    for index, row in weekly_summary.iterrows():
        month_week = index
        values.append((month_week[0], month_week[1], row['Часы'], row['fte']))

    return values
