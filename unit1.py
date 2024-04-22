import pandas as pd


# отчеты
def get_data(fr, file_name, one_hour_fte=164, rprt=1):
    # Данные по ресурсным планам и списанию трудозатрат сотрудников за период

    if rprt == 1:
        df = pd.read_excel(file_name, header=0)
        fr = df[df['Менеджер проекта'] == 'Тапехин Алексей Александрович'][['Проект', 'План, FTE', 'Пользователь',
                                                                            'Фактические трудозатраты (час.) (Сумма)',
                                                                            'Кол-во штатных единиц']]
        fr['Факт, FTE'] = round(fr['Фактические трудозатраты (час.) (Сумма)'] / one_hour_fte, 2)
        fr['Часы план'] = one_hour_fte * fr['План, FTE']
        fr['Остаток часов'] = fr['Часы план'] - fr['Фактические трудозатраты (час.) (Сумма)']
        fr = fr.groupby(['Проект', 'Пользователь', 'Кол-во штатных единиц', 'План, FTE', 'Часы план',
                         'Факт, FTE', 'Остаток часов'])['Фактические трудозатраты (час.) (Сумма)'].sum()
    # Контроль заполнения факта за период
    elif rprt == 2:
        df = pd.read_excel(file_name, header=1)
        fr = df[(df['Проект'] == 'Т0133-КИС "Производственный учет и отчетность"') |
                (df['Проект'] == 'С0134-КИС "Производственный учет и отчетность"')][
            ['Проект', 'ФИО', 'Дата', 'Трудозатрады за день']]
        if rprt == 3:
            fr = fr.sort_values(['Проект', 'ФИО', 'Дата'])
        else:
            fr = round(fr.groupby(['Проект', 'ФИО'])['Трудозатрады за день'].sum(), 2)
    return fr

#
# file_name = 'Контроль заполнения факта за период.xlsx'
# fr = get_data(file_name, 159,2)
# print(f'Контроль заполнения факта за период: \n {fr}')
#
# if input("Нужно экспортировать в excel? Y/N: ") == 'Y':
#     fr.to_excel('output.xlsx', index=True)

# file_name = 'C:\Users\user\Downloads\Данные по ресурсным планам и списанию трудозатрат сотрудников за период.xlsx"'
# fr = get_data(file_name, 159)
# print(f'Данные по ресурсным планам и списанию трудозатрат сотрудников за период: \n {fr}')
# if input("Нужно экспортировать в excel? Y/N: ") == 'Y':
#     fr.to_excel('output.xlsx', index=True)
