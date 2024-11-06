import unittest
import pandas as pd
from datetime import datetime

# Путь к Excel-файлу
file_path = ("C:\\Яндекс\\pro\\PythonPro\\Iteco\\Analiz\\Analiz\\Sources\\2 квартал\\"
             "Сводный список запросов - 2024 апрель.xlsx")

# Для получения минимальной и максимальной даты
date_column = "Дата регистрации"
# Название столбца, в котором нужно искать CRQ
check_column = "Номер запроса"
status_column = "Статус"
request_type_column = "Тип запроса"
registered_column = "Зарегистрировано в период"
complete_period ="Выполнено в период"

class TestExcelData(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        # Загрузка данных из файла Excel с учетом строки заголовка
        cls.df = pd.read_excel(file_path, header=2)  # Заголовок начинается с третьей строки

    # def test_date_column_exists(self):
    #     """Проверка, что столбец с датами существует."""
    #     self.assertIn(date_column, self.df.columns, f"Столбец '{date_column}' не найден в файле.")

    def test_min_max_date(self):
        """Проверка минимальной и максимальной даты в столбце дат."""
        # Преобразование столбца к типу даты
        df_dates = pd.to_datetime(self.df[date_column], errors='coerce').dropna()

        # Нахождение минимальной и максимальной даты
        min_date = df_dates.min()
        max_date = df_dates.max()

        # Проверка, что минимальная дата меньше или равна максимальной
        self.assertLessEqual(min_date, max_date, "Минимальная дата не должна быть больше максимальной")
        print(f"Минимальная дата: {min_date}")
        print(f"Максимальная дата: {max_date}")

    def test_crq_count(self):
        """Проверка, что количество строк, содержащих 'CRQ' (без учета строк со статусом 'отмена'), равно 10."""
        # Проверка наличия столбцов
        self.assertIn(check_column, self.df.columns, f"Столбец '{check_column}' не найден в файле.")
        self.assertIn(status_column, self.df.columns, f"Столбец '{status_column}' не найден в файле.")

        # Фильтрация строк: исключаем строки со статусом "отмена"
        filtered_df = self.df[~self.df[status_column].astype(str).str.contains("Отменено", case=False, na=False)]

        # Подсчет строк, содержащих "CRQ" в отфильтрованных данных
        crq_count = filtered_df[check_column].astype(str).str.contains("CRQ", case=False, na=False).sum()
        inc_data = filtered_df[check_column].astype(str).str.contains("INC", case=False, na=False)

        inc_registered_sum = filtered_df[inc_data][registered_column].sum()
        inc_complete_sum = filtered_df[inc_data][complete_period].sum()
        # Подсчет строк, содержащих "INC" в отфильтрованных данных
        inc_count = inc_data.sum()

        # Дополнительный фильтр для "INC" с учетом столбца "Тип запроса" со значением "Инцидент"
        inc_incident_count = filtered_df[
            (filtered_df[check_column].astype(str).str.contains("INC", case=False, na=False)) &
            (filtered_df[request_type_column] == "Инцидент")
            ].shape[0]

        # Проверка количества "CRQ", общего количества "INC" и количества "INC" с типом "Инцидент"
        # Проверка количества "CRQ", общего количества "INC", количества "INC" с типом "Инцидент" и суммы для "INC"
        self.assertEqual(crq_count, 10, f"Ожидалось 10 строк с 'CRQ' (без учета 'отмена'), но найдено {crq_count}")
        self.assertEqual(inc_count, 51, f"Ожидалось 51 строка с 'INC' (без учета 'отмена'), но найдено {inc_count}")
        self.assertEqual(inc_incident_count, 6,
                         f"Ожидалось 6 строк с 'INC' и 'Инцидент' (без учета 'отмена'), но найдено {inc_incident_count}")
        self.assertEqual(inc_registered_sum, 48,
                         f"Ожидалось, что сумма значений 'Зарегистрировано в период' для 'INC' "
                         f"(без учета 'отмена') будет равна 48, но получилось {inc_registered_sum}")
        self.assertEqual(inc_complete_sum, 42,
                         f"Ожидалось, что сумма значений {complete_period} для 'INC' "
                         f"(без учета 'отмена') будет равна 42, но получилось {inc_complete_sum}")

        print(f"Количество строк с 'CRQ' (без учета 'отмена'): {crq_count}")
        print(f"Количество строк с 'INC' (без учета 'отмена'): {inc_count}")
        print(f"Количество строк с 'INC' и 'Инцидент' (без учета 'отмена'): {inc_incident_count}")
        print(
            f"Сумма значений в столбце 'Зарегистрировано в период' для 'INC' (без учета 'отмена'): {inc_registered_sum}")


if __name__ == "__main__":
    unittest.main()
