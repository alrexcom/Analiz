import pandas as pd

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
        "name": "Сводный список запросов  для SLA",
        "header_row": 2,
        "reportnumber": 3,
        "data_columns": ["Дата регистрации", "Крайний срок решения", "Дата решения", "Дата закрытия",
                         "Дата последнего назначения в группу"],
        "status": ["В ожидании", "Выполнено", "Закрыто", "Проект изменения", "Решен", "Назначен", "Выполняется",
                   "Планирование изменения", "Выполнение изменения", "Экспертиза решения", "Согласование изменения",
                   "Автроизация изменения"]  # Отмена  убрано
    }
]


def report1(df, fte):
    """
    Отчёт Данные по ресурсным планам и списанию трудозатрат сотрудников за период

    """
    headers = ['Проект', 'План, FTE', 'Пользователь', 'Фактические трудозатраты (час.) (Сумма)',
               'Кол-во штатных единиц']

    fr = df[headers].loc[df['Менеджер проекта'] == 'Тапехин Алексей Александрович']
    fr['Факт, FTE'] = round(fr['Фактические трудозатраты (час.) (Сумма)'] / fte, 2)
    fr['Часы план'] = fte * fr['План, FTE']
    fr['Остаток часов'] = fr['Часы план'] - fr['Фактические трудозатраты (час.) (Сумма)']
    fr = fr.groupby(['Проект', 'Пользователь', 'Кол-во штатных единиц', 'План, FTE', 'Часы план',
                     'Факт, FTE', 'Остаток часов'])['Фактические трудозатраты (час.) (Сумма)'].sum()
    return fr


def report1_test(df, fte):
    """
    Отчёт Данные по ресурсным планам и списанию трудозатрат сотрудников за период

    """
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
    # fr = fr.groupby(['Проект', 'Пользователь', 'Кол-во штатных единиц', 'План, FTE', 'Часы план',
    #                  'Факт, FTE', 'Остаток часов'])[fact_sum].sum()
    # fr['columnsheader'] = set(fr)
    # data={'Факт, FTE':fact_fte}
    fr = fr[fr['Факт, FTE'] > 0]
    mydic = {}
    for id_ in fr.index:
        mydic[id_] = list(fr.loc[id_])

    return {'columnsname': list(fr.columns), 'values': mydic}

    # return mydic


def report2(df, date_begin, date_end, export_to_excell):
    """
    Контроль заполнения факта за период
    """

    frm = df[(df['Проект'] == 'Т0133-КИС "Производственный учет и отчетность"') |
             (df['Проект'] == 'С0134-КИС "Производственный учет и отчетность"')][
        ['Проект', 'ФИО', 'Дата', 'Трудозатрады за день']]
    if not export_to_excell:
        result = frm.groupby(['Проект', 'ФИО']).agg({'Дата': 'max', 'Трудозатрады за день': 'sum'})
    else:
        date_range = pd.date_range(start=date_begin, end=date_end, freq='D')
        data_reg = frm[frm['Дата'].isin(date_range)]
        result = data_reg.sort_values(['Проект', 'ФИО', 'Дата'])
    return result


def calc(sum1, sum6, sum7, sum8, prefix):
    s7 = sum7.get(prefix, 0)
    s8 = sum8.get(prefix, 0)
    s6 = sum6.get(prefix, 0)
    s1 = sum1.get(prefix, 0)

    if s8 + s1 == 0:
        s8 = 1

    slap = round((1 - (s6 + s7) / (s8 + s1)) * 100, 2)
    return slap


def report3(df, date_end, date_begin):
    """
    Сводный список запросов  для SLA
    """
    status = reports[2]['status']
    try:
        mdf = df
        mdf = mdf.loc[mdf["Услуга"] == "КИС \"Производственный учет и отчетность\""]
        # mdf = mdf.loc[mdf["Статус"].isin(status)]
        mdf = mdf.loc[mdf["Статус"] != 'Отменено']
        mdf["П2С"] = "П"
        mdf.loc[mdf["Тип запроса"] == 'Нестандартное', "П2С"] = "СДОП"
        mdf.loc[mdf["Тип запроса"] == 'Стандартное без согласования', "П2С"] = "С"

        # date_range = pd.date_range(start=date_begin, end=date_end, freq='D')

        # data_reg = mdf[mdf['Дата регистрации'].isin(date_range)]
        data_reg = mdf[(mdf['Дата регистрации'] >= date_begin and mdf['Дата регистрации'] <= date_end)]

        sum1 = data_reg.groupby(['П2С'])["Зарегистрировано в период"].sum()
        sum2 = mdf.groupby(['П2С'])['Выполнено в период'].sum()
        # sum3 = df.loc[df['Тип запроса']=='Инцидент'].groupby(['П2С']).count()
        inzindent = mdf[['П2С', 'Дата регистрации', 'Статус',
                         'Просрочено в период', 'Дата закрытия']].loc[mdf['Тип запроса'] == 'Инцидент']

        data_inzindent = inzindent[(inzindent['Дата регистрации'] <= date_end)
                                   & (inzindent['Дата регистрации'] >= date_begin)]

        sum3 = data_inzindent[['Дата регистрации', 'П2С']].groupby('П2С').count()

        data_inzindent = inzindent[(inzindent['Статус'] == 'Закрыто') & (inzindent['Просрочено в период'] > 0)
                                   & (inzindent['Дата закрытия'] <= date_end)
                                   & (inzindent['Дата закрытия'] >= date_begin)]

        sum4 = data_inzindent[['Просрочено в период', 'П2С']].groupby(['П2С']).count()

        sum6 = data_reg[['Просрочено в период', 'П2С']].groupby(['П2С']).sum()

        sum7 = mdf[['Открыто на конец периода с просрочкой', 'П2С']].groupby(['П2С']).sum()

        sum8 = mdf.groupby(['П2С'])["Открыто на начало периода"].sum()

        slap = calc(sum1, sum6['Просрочено в период'], sum7['Открыто на конец периода с просрочкой'], sum8, 'П')
        slac = calc(sum1, sum6['Просрочено в период'], sum7['Открыто на конец периода с просрочкой'], sum8, 'С')

        ss = (f"SLA для поддержки = {slap} "
              f"SLA для сопровождения = {slac} "
              f"\n----------------------------------\n"
              f"1 Общее количество зарегистрированных заявок : {sum1}"
              f"\n-Итого:{sum1.sum()}\n\n"
              f"2 Общее количество выполненных заявок : {sum2}"
              f"\n-Итого:{sum2.sum()}\n\n"
              f"3 Общее количество зарегистрированных заявок за "
              f" отчетный период, имеющих категорию «Инцидент»: {sum3}"
              f"\n-Итого:{sum3.sum()}\n\n"
              f"4 Количество заявок за период с превышением срока выполнения, имеющих категорию «Инцидент» : {sum4}"
              f"\n-Итого:{sum4.sum()}\n\n"
              f"5 Количество заявок за период с превышением времени реакции по поддержке: 0 \n\n"
              f"6 (TUR1) Количество закрытых заявок на поддержку с нарушением сроков заявок:{sum6}"
              f"\n-Итого:{sum6.sum()}\n\n"
              f"7 (TUR2) Количество незакрытых заявок, с нарушением срока: {sum7}"
              f"\n-Итого:{sum7.sum()}\n\n"
              f"8 (TUR3) Количество перешедших с прошлого периода заявок на поддержку: {sum8}"
              f"\n-Итого:{sum8.sum()}\n\n"
              f"9 (TUR4) Количество зарегистрированных заявок по поддержке: {sum1}"
              f"\n-Итого:{sum1.sum()}"

              )
        return ss

    except Exception as e:
        return f"Ошибка вычисления {e}"


def report3_test(df, date_end, date_begin):
    """
    Сводный список запросов  для SLA
    """
    status = reports[2]['status']
    try:
        mdf = df
        mdf = mdf.loc[mdf["Услуга"] == "КИС \"Производственный учет и отчетность\""]
        mdf = mdf.loc[mdf["Статус"].isin(status)]
        mdf["П2С"] = "П"
        mdf.loc[mdf["Тип запроса"] == 'Нестандартное', "П2С"] = "СДОП"
        mdf.loc[mdf["Тип запроса"] == 'Стандартное без согласования', "П2С"] = "С"

        # date_range = pd.date_range(start=date_begin, end=date_end, freq='D')

        # data_reg = mdf[mdf['Дата регистрации'].isin(date_range)]

        data_reg = mdf[(mdf['Дата регистрации'] >= date_begin)
                       & (mdf['Дата регистрации'] <= date_end)]

        tur4 = "Зарегистрировано в период"
        sum1 = data_reg.groupby(['П2С'])[tur4].sum()

        p5 = "Открыто на конец периода с просрочкой"
        sum51 = data_reg.groupby(['П2С'])[p5].sum()

        p3 = 'Выполнено в период'
        sum2 = mdf.groupby(['П2С'])[p3].sum()

        tur3 = "Открыто на начало периода"
        sum8 = mdf[[tur3, 'П2С']].groupby(['П2С']).sum()

        # sum3 = df.loc[df['Тип запроса']=='Инцидент'].groupby(['П2С']).count()
        inzindent = mdf[['П2С', 'Дата регистрации', 'Статус',
                         'Просрочено в период', 'Дата закрытия']].loc[mdf['Тип запроса'] == 'Инцидент']

        data_inzindent = inzindent[(inzindent['Дата регистрации'] <= date_end)
                                   & (inzindent['Дата регистрации'] >= date_begin)]

        # s3 = 'Дата регистрации'
        # sum3 = data_inzindent[s3, 'П2С'].groupby('П2С').count()

        data_inzindent = inzindent[(inzindent['Статус'] == 'Закрыто') & (inzindent['Просрочено в период'] > 0)
                                   & (inzindent['Дата закрытия'] <= date_end)
                                   & (inzindent['Дата закрытия'] >= date_begin)]

        # sum4 = data_inzindent[['Просрочено в период', 'П2С']].groupby(['П2С']).count()
        tur1 = 'Выполнено с просрочкой в период'
        # sum6 = data_reg[['Просрочено в период', 'П2С']].groupby(['П2С']).sum()
        sum6 = data_reg[[tur1, 'П2С']].groupby(['П2С']).sum()
        tur2 = 'Открыто на конец периода с просрочкой'
        sum7 = mdf[[tur2, 'П2С']].groupby(['П2С']).sum()

        slap = calc(sum1, sum6[tur1], sum7[tur2], sum8[tur3], 'П')
        slac = calc(sum1, sum6[tur1], sum7[tur2], sum8[tur3], 'С')
        # slap = calc(sum1, sum6['Просрочено в период'], sum7['Открыто на конец периода с просрочкой'], sum8, 'П')
        #  slac = calc(sum1, sum6['Просрочено в период'], sum7['Открыто на конец периода с просрочкой'], sum8, 'С')
        # columnsname = ["Наименование", "Значение"]

        ss = {
            "SLA для поддержки": slap,
            "SLA для сопровождения": slac,

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

            "7 (tur2) Количество заявок за период с превышением времени реакции по сопровождению/поддержке":
                sum7[tur2].sum(),
            "       tur2 поддержка": sum7[tur2].get('П', 0),
            "       tur2 сопровождение": sum7[tur2].get('С', 0)

            # " Общее количество зарегистрированных заявок за отчетный период, имеющих категорию «Инцидент»":sum3.sum(),
            # "       поддержка": sum3.get('П', 0),
            # "       сопровождение": sum3.get('С', 0),
            #
            # "4 Количество заявок за период с превышением срока выполнения, имеющих категорию «Инцидент»":sum4.sum(),
            # "       поддержка ": sum4.get('П', 0),
            # "       сопровождение ": sum4.get('С', 0)
        }
        return ss



    except Exception as e:
        print(f"Ошибка вычисления {e}")


def get_data(reportnumber, date_end, date_begin, df, fte, export_excell):
    frm = ''
    if reportnumber == 1:
        frm = report1(df, fte)
    elif reportnumber == 3:
        frm = report3(df, date_end=date_end, date_begin=date_begin)
    elif reportnumber == 2:
        frm = report2(df=df, date_begin=date_begin, date_end=date_end, export_to_excell=export_excell)
    return frm


def get_data_test(reportnumber, date_end, date_begin, df, fte, export_excell):
    frm = ''
    if reportnumber == 1:
        frm = report1_test(df, fte)
    elif reportnumber == 3:
        frm = report3_test(df, date_end=date_end, date_begin=date_begin)
    elif reportnumber == 2:
        frm = report2(df=df, date_begin=date_begin, date_end=date_end, export_to_excell=export_excell)
    return frm


def get_report(num_report, filename):
    for items in reports:
        if num_report == items['reportnumber']:
            result = pd.read_excel(filename, header=items['header_row'], parse_dates=items['data_columns'],
                                   date_format='%d.%m.%Y')
            return result


def get_report_test(name_report, filename):
    for items in reports:
        if name_report == items['name']:
            result = (items['reportnumber'],
                      pd.read_excel(filename, header=items['header_row'], parse_dates=items['data_columns'],
                                    date_format='%d.%m.%Y'))
            return result
