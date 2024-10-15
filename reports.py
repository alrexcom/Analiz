import pandas as pd
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
        filtered_mdf = params["df"]

        # status переменная нужна в запросе
        status = params['status']
        # Объединённая фильтрация
        # filtered_mdf = mdf.query("Услуга == 'КИС \"Производственный учет и отчетность\"' and Статус in @status")
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
        columns = [{"name": "Наименование", "width": 600, "anchor": 'w'},
                   {"name": "Значение", "width": 100, "anchor": 'center'}]

        # Обработка нарушений SLA  убиваем Nan - пустые строки
        errsla = filtered_mdf[["Номер запроса", "Исполнитель",
                               "Комментарий к нарушению SLA"]].dropna(subset=["Комментарий к нарушению SLA"])

        data = errsla.apply(lambda x: (
            f"ERR SLA: запрос {x['Номер запроса']} / {x['Исполнитель']} :{x['Комментарий к нарушению SLA']}", 1),
                            axis=1).tolist()

        # Добавление поддержки SLA
        if params['support']:
            ss = sla_support(mdf=filtered_mdf, data_reg=data_reg, data_prosr=data_prosr, columns=columns,
                             date_end=dp, date_begin=ds)
        else:
            ss = sla_sopr(mdf=filtered_mdf, data_reg=data_reg, data_prosr=data_prosr, columns=columns,
                          date_end=dp, date_begin=ds)

        for key, val in ss.items():
            data.append((key, val))

        return {"columns": columns, "data": data}
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


def sla_sopr(**params):
    """
      Данные по SLA Сопровождение
     """
    # mdf = params["mdf"]
    # overdue_filter = ((mdf['Дата закрытия'] <= params["date_end"]) &
    #                   (mdf['Дата закрытия'] >= params["date_begin"]))
    # params["data_reg"] = mdf[overdue_filter]

    # params["data_reg"] = params["mdf"]

    data_reg = params["data_reg"]
    # data_reg = params["mdf"]

    par = get_data_sla(**params)
    sum1 = par["sum1"]
    sum2 = par["sum2"]

    p5 = "Открыто на конец периода с просрочкой"
    sum51 = data_reg.groupby(['П2С'])[p5].sum()

    tur3 = "Открыто на начало периода"
    sum8 = par["sum8"]

    tur1 = 'Выполнено с просрочкой в период'
    sum6 = par["sum6"]

    tur2 = 'Просрочено в период'
    sum7 = par["sum7"]

    slap = calc(sum1, sum6[tur1], sum7[tur2], sum8[tur3], 'П')
    slac = calc(sum1, sum6[tur1], sum7[tur2], sum8[tur3], 'С')

    # ss = {
    #     "9 Уровень исполнения SLA общий": round(100 * (sum2.sum() - sum6[tur1].sum()) / sum2.sum(), 2),
    #     "       поддержки": slap,
    #     "       сопровождения": slac,
    #
    #     "1 ( tur3 ) Общее количество незакрытых заявок по сопровождению/поддержке на начало периода ": sum8[
    #         tur3].sum(),
    #     "       tur3 поддержка": sum8[tur3].get('П', 0),
    #     "       tur3 сопровождение": sum8[tur3].get('С', 0),
    #
    #     "2 (tur4) Общее количество зарегистрированных заявок по сопровождению/поддержке ": sum1.sum(),
    #     "       tur4 поддержка": sum1.get('П', 0),
    #     "       tur4 сопровождение": sum1.get('С', 0),
    #
    #     "3 п/п Общее количество закрытых за период заявок по сопровождению/поддержке ": sum2.sum(),
    #     "       p3 поддержка": sum2.get('П', 0),
    #     "       p3 сопровождение": sum2.get('С', 0),
    #
    #     "4 (tur1) Общее количество закрытых за период заявок по сопровождению/поддержке c нарушением SLA": sum6[
    #         tur1].sum(),
    #     "       tur1 поддержка": sum6[tur1].get('П', 0),
    #     "       tur1 сопровождение": sum6[tur1].get('С', 0),
    #
    #     "5 п/п Общее количество незакрытых заявок по сопровождению/поддержке на конец периода": sum51.sum(),
    #     "       p5 поддержка": sum51.get('П', 0),
    #     "       p5 сопровождение": sum51.get('С', 0),
    #
    #     "7 (tur2) Количество заявок за период с превышением срока выполнения ":
    #         sum7[tur2].sum(),
    #     "       по поддержке": sum7[tur2].get('П', 0),
    #     "       по сопровождению": sum7[tur2].get('С', 0)
    # }

    ss = {
        "9 Уровень исполнения SLA общий": round(100 * (sum2.sum() - sum6[tur1].sum()) / sum2.sum(), 2),
        "       сопровождения": slac,

        "1 ( tur3 ) Общее количество незакрытых заявок по сопровождению/поддержке на начало периода ":
            sum8[tur3].get('С', 0),
        # "       tur3 сопровождение": sum8[tur3].get('С', 0),

        "2 (tur4) Общее количество зарегистрированных заявок по сопровождению/поддержке ": sum1.get('С', 0),
        # "       tur4 сопровождение": sum1.get('С', 0),

        "3 п/п Общее количество закрытых за период заявок по сопровождению/поддержке ": sum2.get('С', 0),
        # "       p3 сопровождение": sum2.get('С', 0),

        "4 (tur1) Общее количество закрытых за период заявок по сопровождению/поддержке c нарушением SLA":
            sum6[tur1].get('С', 0),
        # "       tur1 сопровождение": sum6[tur1].get('С', 0),

        "5 п/п Общее количество незакрытых заявок по сопровождению/поддержке на конец периода": sum51.get('С', 0),
        # "       p5 сопровождение": sum51.get('С', 0),

        "7 (tur2) Количество заявок за период с превышением срока выполнения ":
            sum7[tur2].get('С', 0),
        # "       по сопровождению": sum7[tur2].get('С', 0)
    }

    return ss


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
    col = ['Заявка', 'Подзадача', 'Часы', 'Регистрация', 'Квартал', 'Месяц', 'Содержание']
    df = pd.DataFrame(data_fromsql, columns=col)
    # Преобразуем столбец с датами в datetime формат
    df['Регистрация'] = pd.to_datetime(df['Регистрация'])
    df['fte'] = round(df['Часы'] / 164, 2)
    # Найдем начало месяца для каждой даты
    df['month_start'] = df['Регистрация'].values.astype('datetime64[M]')

    # Рассчитаем, сколько недель прошло с начала месяца
    df['Неделя'] = ((df['Регистрация'] - df['month_start']).dt.days // 7) + 1

    # Группируем по номеру недели в месяце и считаем среднее значение
    weekly_summary = df.groupby(['Месяц', 'Неделя'])[['Часы', 'fte']].sum()

    print(weekly_summary)
    return weekly_summary
