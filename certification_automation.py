from my_script import certificate


# ФАЙЛ ДЛЯ ОТЛАДКИ, не для пользователя

file_path = 'Test_file_2.xlsx'
sheet_name = 'Оценка компетенций'
file_path_2 = 'Компетенции_по_шкале_DE.xlsx'
sheet_name_2 = 'Целевые значения'
save_to = 'C:/Users/user/PycharmProjects/certification_automation'
res = certificate(file_path, sheet_name, file_path_2, sheet_name_2, save_to)
