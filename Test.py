import pandas as pd
import numpy as np
from reports import reports

rep = reports[4]
cols = rep['headers']

data = pd.read_csv('data.csv', sep=';', encoding='utf-8',
                   parse_dates=[], header=0, date_format='%Y%m%d')

df = pd.DataFrame(data, columns=cols)

# Добавляем столбец во фрэйм
df['vid'] = np.where(df['Исполнитель  по задаче'] != 'Тапехин Алексей Александрович', "Поддержка", 'Сопровождение')
# df['vsego'] = '10'
# Копируем в отдельный фрейм данные
kk = df.groupby(['vid', 'Исполнитель  по задаче']).agg(
    {
        'ID инцидента/ЗИ': "count",
        'Трудозатраты по задаче (десят. часа)': "sum",
    }
).copy()


def window_():
    # print(kk)
    ww1 = kk
    ww2 = kk.rolling(window=2).sum()
    ww3 = kk.expanding(2).sum()
    print(ww1)
    print(ww2)
    print(ww3)


# df2["ma_28_day"] = (
# df2.sort_values("Date")
# .groupby("stocks")["closing_price"]
# .transform(lambda x: x.rolling(28, min_periods=1).mean())

def nwe_():
    # cols1 = ['Исполнитель  по задаче', 'Трудозатраты по задаче (десят. часа)']
    # df = pd.DataFrame(data, columns=cols1)
    # df1=df.groupby('Исполнитель  по задаче').agg({'Трудозатраты по задаче (десят. часа)':'sum'})
    # df1['vid'] = np.where(df['Исполнитель  по задаче'] != 'Тапехин Алексей Александрович', "Поддержка", 'Сопровождение')

    kk['sum_all'] = (kk['Трудозатраты по задаче (десят. часа)'].sum())
    kk['sum_vid'] = (kk.groupby("vid")['Трудозатраты по задаче (десят. часа)'].transform('sum'))
    kk['count_vid'] = (kk.groupby("vid")['Трудозатраты по задаче (десят. часа)'].transform('count'))

    kk['sum_ispol'] = (kk.groupby('Исполнитель  по задаче')['Трудозатраты по задаче (десят. часа)'].transform('sum'))
    kk['count_vid'] = (kk.groupby('Исполнитель  по задаче')['Трудозатраты по задаче (десят. часа)'].transform('count'))
    out_=kk.sort_values(["vid", 'Исполнитель  по задаче'])
    print(out_)


def one():
    # Считаем разные группировки
    my = kk.groupby(['vid']).transform("sum")
    # vs = kk.agg({'vsego':"sum"}).transform("sum")
    my1 = kk.groupby(['Исполнитель  по задаче']).transform("sum")
    # itog=kk.sum()
    # print(itog)
    # Переименуем стобцы полученных значений
    my1 = my1.rename(
        columns={'Трудозатраты по задаче (десят. часа)': 'Трудозатраты', 'ID инцидента/ЗИ': 'Количество заявок'})
    my = my.rename(
        columns={'Трудозатраты по задаче (десят. часа)': 'Трудозатраты по виду', 'ID инцидента/ЗИ': 'Общее количество'})

    # Соединяем по столбцам 2 dataframe
    tab = pd.concat([my, my1], axis=1).reset_index()
    # tab = pd.concat([tab, itog], axis=1)

    # print(tab)
    # Преобразование полученного набора данных в лист
    # print(tab.values.tolist())

    # pivot

    pv = tab.pivot(index=['Исполнитель  по задаче', 'Трудозатраты'],
                   columns=['vid', 'Трудозатраты по виду', 'Общее количество'],
                   values=['Количество заявок'])

    print(pv)


# one()
# window_()

nwe_()
