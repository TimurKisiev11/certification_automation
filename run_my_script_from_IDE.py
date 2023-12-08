import os
import subprocess

# ФАЙЛ ДЛЯ ОТЛАДКИ, не для пользователя
# Чтобы запускать my_script на всех файлах из определенной директории из IDE

script_path = 'C:/Users/user/PycharmProjects/certification_automation/my_script.py'
file_path_2 = 'C:/Users/user/PycharmProjects/certification_automation/Компетенции_по_шкале_DE.xlsx'
directory_path = 'C:/Users/user/PycharmProjects/certification_automation/certification_results'
save_to = 'C:/Users/user/PycharmProjects/certification_automation/results'

for filename in os.listdir(directory_path):
    if filename.endswith(".xlsx"):
        file_path = os.path.join(directory_path, filename)
        subprocess.call(['python', script_path, file_path, file_path_2, save_to])
