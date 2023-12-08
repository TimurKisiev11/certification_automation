import unittest
from unittest import TestCase
from certification_script import read_from_file, convert_to_float, compare, certificate, create_and_write_to_xlsx
import pytest


class Test(TestCase):
    """
    Тесты для certification_automation.py
    """

    @pytest.fixture(autouse=True)
    def _pass_fixtures(self, capsys):
        self.capsys = capsys

    def test_read_error(self):
        file_path = 'А_нет_такого_файла.xlsx'
        sheet_name = 'Оценка компетенций'
        result = read_from_file(file_path, sheet_name, 11, 30, 10)
        captured = self.capsys.readouterr()
        self.assertEqual((
            'Error while reading from file: [Errno 2] No such file or '"directory: 'А_нет_такого_файла.xlsx'\n"),
            captured.out)

    def test_convert_error(self):
        result = convert_to_float(1)  # если на вход подан вообще не список
        captured = self.capsys.readouterr()
        self.assertEqual("Error TWO while converting to float: 'int' object is not iterable\n", captured.out)

    def test_convert_error_2(self):
        result = convert_to_float("A")  # если на вход подано не числовое значение
        captured = self.capsys.readouterr()
        self.assertEqual("Error ONE while converting to float: could not convert string to float: 'A'\n", captured.out)

    def test_both(self):
        file_path = 'Test_file_1.xlsx'
        sheet_name = 'Оценка компетенций'
        result = read_from_file(file_path, sheet_name, 11, 30, 10)
        expected_values = [1.1, 2.1, 3.1, 4.0, 5.0, 6.1, 7.2, 8.6, 9.0, 0.0, 11.0, 12.0, 13.0, 0.0, 15.2, 16.0, 0.0,
                           0.0, 0.0, 0.0]
        self.assertEqual(convert_to_float(result), expected_values)

    def test_empty_cells(self):
        file_path = 'Test_file_1.xlsx'
        sheet_name = 'Оценка компетенций'
        result = read_from_file(file_path, sheet_name, 11, 12, 11)
        expected_values = [0.0, 0.0]
        self.assertEqual(convert_to_float(result), expected_values)

    def test_bad_file(self):
        file_path = 'А_нет_такого_файла.xlsx'
        sheet_name = 'Оценка компетенций'
        result = read_from_file(file_path, sheet_name, 11, 30, 10)
        self.assertEqual(result, [])

    def test_convert_to_float(self):
        test_values = ['1.1', 4, 2.1, '', None]  # еще раз тестируем на все возможные вариации ввода
        result = convert_to_float(test_values)
        self.assertEqual(result, [1.1, 4.0, 2.1, 0.0, 0.0])

    def test_convert_to_float_2(self):
        result = convert_to_float(1)
        self.assertEqual(result, [])

    def test_compare(self):
        test_average = [1.1, 1.4, 1.5, 1.6, 1.0, 1.0, 1.3, 1.0, 2.1, 0.0, 1.0, 1.0, 2.0, 2.1, 1.0, 1.0]
        test_target = [1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 1.0, 2.0, 1.0, 1.0, 2.0, 1.0, 1.0, 1.0]
        self.assertTrue(compare(test_average, test_target, 0.8), True)

    def test_certificate(self):
        file_path = 'Test_file_2.xlsx'
        sheet_name = 'Оценка компетенций'
        file_path_2 = 'Компетенции_по_шкале_DE.xlsx'
        sheet_name_2 = 'Целевые значения'
        res = certificate(file_path, sheet_name, file_path_2, sheet_name_2, False)
        self.assertEqual({'Junior I': (False, 0.78), 'Junior II': (False, 0.78), 'Middle I': (False, 0.72),
                          'Middle II': (False, 0.72), 'Senior I': (False, 0.67), 'Senior II ': (False, 0.28), 'Expert ': (False, 0.11)}, res)

    def test_create_and_write_to_xlsx(self):
        save_to = 'C:/Users/user/PycharmProjects/certification_automation'
        res = {'Junior I': (True, 1.0), 'Junior II': (True, 0.94), 'Middle I': (True, 0.89), 'Middle II': (True, 0.83),
               'Senior I': (True, 0.83), 'Senior II ': (False, 0.78), 'Expert ': (False, 0.78)}
        create_and_write_to_xlsx("Меня нет я тест", res, save_to)
        captured = self.capsys.readouterr()
        self.assertEqual('Файл '"'C:/Users/user/PycharmProjects/certification_automation/certificate_Меня нет "
                         "я тест.xlsx' создан, данные занесены.\n", captured.out)


if __name__ == '__main__':
    unittest.main()
