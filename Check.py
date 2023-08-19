import os
import psutil


def create_file(lineedit_data_file, lineedit_finish_folder, lineedit_file_name):

    for proc in psutil.process_iter():
        if proc.name() == 'WINWORD.EXE':
            return ['УПС!', 'Закройте все файлы Word!']
    path_file = lineedit_data_file.text().strip()
    if not path_file:
        return ['УПС!', 'Путь к файлу выгрузок пуст']
    if os.path.exists(path_file) is False:
        return ['УПС!', 'Не удается найти указанный файл выгрузок']
    if os.path.isdir(path_file):
        return ['УПС!', 'Указанный путь к файлу выгрузок является директорией']
    if path_file.endswith('.txt'):
        pass
    else:
        return ['УПС!', 'Загружаемый файл не формата ".txt"']
    finish_path = lineedit_finish_folder.text().strip()
    if not finish_path:
        return ['УПС!', 'Путь к конечной папке пуст']
    if os.path.isfile(finish_path):
        return ['УПС!', 'Путь к конечной папке является файлом']
    if os.path.exists(finish_path) is False:
        return ['УПС!', 'Не удается найти указанную конечную директорию']
    file_name = lineedit_file_name.text().strip()
    if not file_name:
        return ['УПС!', 'Нет имени файла']

    return {'path_file': path_file, 'finish_path': finish_path, 'file_name': file_name}

