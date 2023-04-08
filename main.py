import os
import time
import shutil
from datetime import datetime

PATH_TO_DIR = 'C:/Users/admin/Desktop/OCR/'
EXCEL_DIR = PATH_TO_DIR + 'Excel/'
WORD_DIR = PATH_TO_DIR + 'Word/'
OUTPUT_DIR = PATH_TO_DIR + 'YourFiles/'
TRASH_DIR = PATH_TO_DIR + 'Корзина/'
ABBYY_PATH = 'C:/Program Files (x86)/ABBYY FineReader 15/FineCmd.exe'
TYPE_LIST = ['pdf', 'jpg', 'jpeg', 'bmp', 'png', 'gif', 'djvu', 'djv']


def move_file(initial_path, final_path):
    try:
        shutil.move(initial_path, final_path)
        print('Файл успешно перемещен!')
    except FileNotFoundError:
        print('Не удалось переместить файл! Файл не найден!')
    except PermissionError:
        print('Не удалось переместить файл! Нет доступа!')


def main():
    print(EXCEL_DIR)
    print(WORD_DIR)
    print(OUTPUT_DIR)
    print(TRASH_DIR)

    while True:
        timestamp = datetime.now().strftime("%Y_%m_%d_%H_%M_%S_")
        for file in os.scandir(EXCEL_DIR):
            file_name, file_ext = os.path.splitext(file.name)[0], os.path.splitext(file.name)[1]
            print(file_name, file_ext)
            time.sleep(3)

            if file_ext.replace('.', '') not in TYPE_LIST:
                print('файл не в списке разрешенных расширений')
                move_file(EXCEL_DIR + file.name, TRASH_DIR + timestamp + file.name)
            else:
                print('Запускаем распознавание!')
                cmd = fr'chcp 1251 & "{ABBYY_PATH}" "{EXCEL_DIR}{file.name}" /lang English Russian ' \
                      fr'/out "{OUTPUT_DIR}{timestamp}{file.name}.xlsx"'
                print(cmd)
                exit_code = os.system(cmd)
                print("`FineCmd.exe ~` run with exit code %d" % exit_code)
                move_file(EXCEL_DIR + file.name, TRASH_DIR + timestamp + file.name)


if __name__ == "__main__":
    main()
    time.sleep(3)
