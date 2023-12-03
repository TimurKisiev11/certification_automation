import os
import subprocess
import sys

# ФАЙЛ ДЛЯ ОТЛАДКИ, не для пользователя
# Чтобы запускать my_script на всех файлах из определенной директории из командной строки

if __name__ == "__main__":
    if len(sys.argv) != 5:
        print(
            "Как использовать: \путь\к\ run_my_script_from_cmd.py \путь\к\my_script.py \путь\к\Компетенции_по_шкале_DE.xlsx \путь\к\директории с файлами \куда\сохранить")
    else:
        script_path = sys.argv[1]
        file_path_2 = sys.argv[2]
        directory_path = sys.argv[3]
        save_to = sys.argv[4]

        for filename in os.listdir(directory_path):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(directory_path, filename)
                subprocess.call(['python', script_path, file_path, file_path_2, save_to])
