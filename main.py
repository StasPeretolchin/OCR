import os
import time
import shutil
from datetime import datetime

PATH_TO_DIR = "/home/defeval/Downloads/OСR/"

EXCEL_DIR = PATH_TO_DIR + "Excel/"
WORD_DIR = PATH_TO_DIR + "Word/"
OUTPUT_DIR = PATH_TO_DIR + "YouFiles/"
TRASH_DIR = PATH_TO_DIR + "Корзина/"
ABBYY_PATH = "\"C:\Program Files (x86)\ABBYY FineReader 15\FineCmd.exe\""

TYPE_LIST = ['pdf', 'jpg', 'jpeg', 'bmp', 'png', 'gif', 'djvu', 'djv']


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
            time.sleep(1)

            if file_ext.replace('.', '') not in TYPE_LIST:
                print('файл не в списке разрешеннных расширений')
                try:
                    shutil.move(EXCEL_DIR + file.name, TRASH_DIR + timestamp + file.name)
                    print('Файл успешно перемещен!')
                except FileNotFoundError:
                    print('Не удалось переместить файл! Файл не найден!')
                except PermissionError:
                    print('Не удалось переместить файл! Нет доступа!')
            else:
                print('запускаем распознование!')
                print(ABBYY_PATH + " " + "\"" + EXCEL_DIR + file.name + "\"" + " /lang Mixed " + "/out " + "\"" + OUTPUT_DIR + timestamp + file.name + ".xlsx" + "\"")
                home_dir = os.system("chcp 1251 & chcp & " + ABBYY_PATH + " " + "\"" + EXCEL_DIR + file.name + "\"" + " /lang Russian " + "/out " + "\"" + OUTPUT_DIR + timestamp + file.name + ".xlsx" + "\"")
                print("`cd ~` ran with exit code %d" % home_dir)

                try:
                    shutil.move(EXCEL_DIR + file.name, TRASH_DIR + timestamp + file.name)
                    print('Файл успешно перемещен!')
                except FileNotFoundError:
                    print('Не удалось переместить файл! Файл не найден!')
                except PermissionError:
                    print('Не удалось переместить файл! Нет доступа!')

        for file in os.scandir(WORD_DIR):
            file_name, file_ext = os.path.splitext(file.name)[0], os.path.splitext(file.name)[1]
            print(file_name, file_ext)
            time.sleep(1)

            if file_ext.replace('.', '') not in TYPE_LIST:
                print('файл не в списке разрешеннных расширений')
                try:
                    shutil.move(WORD_DIR + file.name, TRASH_DIR + timestamp + file.name)
                    print('Файл успешно перемещен!')
                except FileNotFoundError:
                    print('Не удалось переместить файл! Файл не найден!')
                except PermissionError:
                    print('Не удалось переместить файл! Нет доступа!')
            else:
                print('запускаем распознование!')
                print(ABBYY_PATH + " " + "\"" + WORD_DIR + file.name + "\"" + " /lang Mixed " + "/out " + "\"" + OUTPUT_DIR + timestamp + file.name + ".xlsx" + "\"")
                home_dir = os.system("chcp 1251 & chcp & " + ABBYY_PATH + " " + "\"" + WORD_DIR + file.name + "\"" + " /lang Russian " + "/out " + "\"" + OUTPUT_DIR + timestamp + file.name + ".docx" + "\"")
                print("`cd ~` ran with exit code %d" % home_dir)

                try:
                    shutil.move(WORD_DIR + file.name, TRASH_DIR + timestamp + file.name)
                    print('Файл успешно перемещен!')
                except FileNotFoundError:
                    print('Не удалось переместить файл! Файл не найден!')
                except PermissionError:
                    print('Не удалось переместить файл! Нет доступа!')


if __name__ == "__main__":
    main()
    time.sleep(3)
