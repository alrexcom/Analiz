import pandas as pd
from univunit import Univunit

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
        "name": "Отчет по запросам и задачам (долго открывается)",
        "reportnumber": 4,
        "header_row": 4,
        "data_columns": ["ID инцидента/ЗИ", "Исполнитель  по задаче", "Трудозатраты по задаче (десят. часа)",
                         "Содержание задачи","Статус","Дата Выполнения работ"]
    },

]


def report1(**param):
    """
    Отчёт Данные по ресурсным планам и списанию трудозатрат сотрудников за период

    """
    df = param['df']
    fte = param['fte']
    fact_sum = 'Фактические трудозатраты (час.) (Сумма)'
    headers = ['Проект', 'План, FTE', 'Пользователь', 'Фактические трудозатраты (час.) (Сумма)',
               'Кол-во штатных единиц']
    user = 'Тапехин Алексей Александрович'

    # fr = df[headers].loc[(df['Менеджер проекта'] == user) | (df['Пользователь'] == user)]

    fr = df.loc[(df['Менеджер проекта'] == user) | (df['Пользователь'] == user), headers]

    fact_fte = round((fr[fact_sum] / fte), 2)
    fr['Факт, FTE'] = fact_fte
    hours_plan = fte * fr['План, FTE']
    fr['Часы план'] = round(hours_plan, 2)
    hours_remain = round(hours_plan - fr[fact_sum], 2)
    fr['Остаток часов'] = hours_remain

    fr = fr[fr['Факт, FTE'] > 0]

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
        mdf = params["df"]

        # status переменная нужна в запросе
        status = params['status']
        # Объединённая фильтрация
        filtered_mdf = mdf.query("Услуга == 'КИС \"Производственный учет и отчетность\"' and Статус in @status")

        # Классификация запросов
        filtered_mdf["П2С"] = "П"
        filtered_mdf.loc[filtered_mdf["Тип запроса"] == 'Нестандартное', "П2С"] = "СДОП"
        filtered_mdf.loc[filtered_mdf["Тип запроса"] == 'Стандартное без согласования', "П2С"] = "С"

        # Фильтрация по датам
        date_filter = ((filtered_mdf['Дата регистрации'] >= params["date_begin"]) &
                       (filtered_mdf['Дата регистрации'] <= params["date_end"]))
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
                             date_end=params["date_end"], date_begin=params["date_begin"])
        else:
            ss = sla_sopr(mdf=filtered_mdf, data_reg=data_reg, data_prosr=data_prosr, columns=columns,
                          date_end=params["date_end"], date_begin=params["date_begin"])

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
    mdf = params['mdf']
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
    sum8 = par['mdf'][["Открыто на начало периода", 'П2С']].groupby(['П2С']).sum()
    return {"sum1": sum1, "sum2": sum2, "sum5": sum5, "sum6": sum6, "sum7": sum7, "sum8": sum8}


def sla_sopr(**params):
    """
      Данные по SLA Сопровождение
     """

    data_reg = params["data_reg"]

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

    ss = {
        "9 Уровень исполнения SLA общий": round(100 * (sum2.sum() - sum6[tur1].sum()) / sum2.sum(), 2),
        "       поддержки": slap,
        "       сопровождения": slac,

        "1 ( tur3 ) Общее количество незакрытых заявок по сопровождению/поддержке на начало периода ": sum8[
            tur3].sum(),
        "       tur3 поддержка": sum8[tur3].get('П', 0),
        "       tur3 сопровождение": sum8[tur3].get('С', 0),

        "2 (tur4) Общее количество зарегистрированных заявок по сопровождению/поддержке ": sum1.sum(),
        "       tur4 поддержка": sum1.get('П', 0),
        "       tur4 сопровождение": sum1.get('С', 0),

        "3 п/п Общее количество закрытых за период заявок по сопровождению/поддержке ": sum2.sum(),
        "       p3 поддержка": sum2.get('П', 0),
        "       p3 сопровождение": sum2.get('С', 0),

        "4 (tur1) Общее количество закрытых за период заявок по сопровождению/поддержке c нарушением SLA": sum6[
            tur1].sum(),
        "       tur1 поддержка": sum6[tur1].get('П', 0),
        "       tur1 сопровождение": sum6[tur1].get('С', 0),

        "5 п/п Общее количество незакрытых заявок по сопровождению/поддержке на конец периода": sum51.sum(),
        "       p5 поддержка": sum51.get('П', 0),
        "       p5 сопровождение": sum51.get('С', 0),

        "7 (tur2) Количество заявок за период с превышением срока выполнения ":
            sum7[tur2].sum(),
        "       по поддержке": sum7[tur2].get('П', 0),
        "       по сопровождению": sum7[tur2].get('С', 0)
    }
    return ss


def get_data_report(**params):
    report_data = get_reports(**params)
    reportnumber = report_data[0]

    params['df'] = report_data[1]
    params['date_begin'] = Univunit.convert_date(params['date_begin'])
    params['date_end'] = Univunit.convert_date(params['date_end'])

    # Если ключа статус нет
    params['status'] = reports[reportnumber - 1].get('status', 'статус не указан')
    params['support'] = bool(reports[reportnumber - 1].get('support', False))

    report_functions = {
        1: report1,
        2: report2,
        3: report_sla,
        4: report_sla,
    }

    frm = report_functions.get(reportnumber)
    if frm is None:
        raise ValueError("Некорректный номер отчета")

    return frm(**params)


def get_reports(**params):
    for items in reports:
        if params['name_report'] == items['name']:
            result = (items['reportnumber'],
                      pd.read_excel(params['filename'], header=items['header_row'], parse_dates=items['data_columns'],
                                    date_format='%d.%m.%Y'))
            return result


def names_reports():
    return [items["name"] for items in reports]
