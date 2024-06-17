from univunit import (pd, convert_date)

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
    }
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
    # fr = df[headers][[df['Пользователь'] == user | df['Менеджер проекта'] == user]]
    fr = df[headers].loc[(df['Менеджер проекта'] == user) | (df['Пользователь'] == user)]

    fact_fte = round((fr[fact_sum] / fte), 2)
    fr['Факт, FTE'] = fact_fte
    hours_plan = fte * fr['План, FTE']
    fr['Часы план'] = round(hours_plan, 2)
    hours_remain = round(hours_plan - fr[fact_sum], 2)
    fr['Остаток часов'] = hours_remain

    fr = fr[fr['Факт, FTE'] > 0]
    data = []
    for id_ in fr.index:
        data.append(tuple(fr.loc[id_]))
    columns = [{"name": items} for items in fr.columns]

    return {'columns': columns, 'data': data}


def report2(**params):
    """
    Контроль заполнения факта за период

    """

    df = params['df']
    frm = df[(df['ФИО'] == 'Тапехин Алексей Александрович') |
             (df['Проект'] == 'Т0133-КИС "Производственный учет и отчетность"')][
        ['Проект', 'ФИО', 'Дата', 'Трудозатрады за день']]

    date_range = pd.date_range(start=params['date_begin'], end=params['date_end'], freq='D')
    data_reg = frm[frm['Дата'].isin(date_range)]
    result = data_reg.sort_values(['ФИО', 'Проект', 'Дата'])

    columns = []
    for items in result.columns:
        if items is not None:
            columns.append({"name": items})

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
    Сводный список запросов  для SLA
    """

    status = reports[2]['status']
    reportnumber = int(params["reportnumber"]) - 1
    support = bool(reports[reportnumber]['support'])
    try:
        mdf = params["df"]
        mdf = mdf.loc[mdf["Услуга"] == "КИС \"Производственный учет и отчетность\""]
        mdf = mdf.loc[mdf["Статус"].isin(status)]
        mdf["П2С"] = "П"
        mdf.loc[mdf["Тип запроса"] == 'Нестандартное', "П2С"] = "СДОП"
        mdf.loc[mdf["Тип запроса"] == 'Стандартное без согласования', "П2С"] = "С"
        data_reg = mdf[
            (mdf['Дата регистрации'] >= params["date_begin"]) & (mdf['Дата регистрации'] <= params["date_end"])]

        data_prosr = mdf[(mdf['Статус'] != 'Закрыто') & (mdf['Просрочено в период'] > 0)
                         & (mdf['Дата закрытия'] <= params["date_end"])
                         & (mdf['Дата закрытия'] >= params["date_begin"])]
        columns = [{
            "name": "Наименование",
            "width": 600,
            "anchor": 'w'
        },
            {
                "name": "Значение",
                "width": 100,
                "anchor": 'center'
            }]

        errsla = mdf[["Номер запроса", "Исполнитель", "Комментарий к нарушению SLA"]][
            mdf["Комментарий к нарушению SLA"].notna()]
        mu = errsla.to_dict('index')

        data = []
        for item in mu.items():
            key, val = item
            data.append((
                f"ERR SLA: запрос {val['Номер запроса']} / {val["Исполнитель"]} :{val['Комментарий к нарушению SLA']}",
                "1"))

        if support:
            ss = sla_support(mdf=mdf, data_reg=data_reg, data_prosr=data_prosr, columns=columns,
                             date_end=params["date_end"], date_begin=params["date_begin"])
        else:
            ss = sla_sopr(mdf=mdf, data_reg=data_reg, data_prosr=data_prosr, columns=columns,
                          date_end=params["date_end"], date_begin=params["date_begin"])

        for item in ss.items():
            key, val = item
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
    reportnumber = params['reportnumber']

    params['date_begin'] = convert_date(params['date_begin'])
    params['date_end'] = convert_date(params['date_end'])

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



    # frm = ''
    # if reportnumber == 1:
    #     frm = report1(**params)
    # elif reportnumber in (3, 4):
    #     frm = report_sla(**params)
    # elif reportnumber == 2:
    #     frm = report2(**params)
    # return frm


def get_reports(name_report, filename):
    for items in reports:
        if name_report == items['name']:
            result = (items['reportnumber'],
                      pd.read_excel(filename, header=items['header_row'], parse_dates=items['data_columns'],
                                    date_format='%d.%m.%Y'))
            return result




