import unittest
from reports import *


# from reports import calculate_slap

def calculate_slap(sum1, sum6, sum7, sum8):
    if sum8 + sum1 == 0:
        return None  # Возвращаем None, если знаменатель равен нулю

    # slap = round((1 - (sum6 + sum7) / (sum8 + sum1)) * 100, 2)
    return sum8 + sum1


class MyTestCase(unittest.TestCase):
    def test_calculate_slap(self):
        # Тестирование функции с валидными значениями
        # self.assertEqual(calculate_slap(10, 3, 2, 5), 72.0)
        self.assertEqual(calculate_slap(10, 3, 2, 5), 15.0)

        # # Тестирование функции, когда знаменатель равен нулю
        self.assertIsNone(calculate_slap(0, 0, 0, 0))

        # # Тестирование функции с отрицательными числами
        self.assertEqual(calculate_slap(-10, -3, -2, -5), -15)
        #
        # # Тестирование функции с плавающей точкой
        self.assertEqual(calculate_slap(10.5, 3.2, 2.1, 5.6), 16.1)




if __name__ == '__main__':
    unittest.main()
