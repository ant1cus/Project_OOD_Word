import os


def create_file(lineedit_data_file, lineedit_finish_folder, lineedit_file_name):

    path_file = lineedit_data_file.text().strip()
    if not path_file:
        return ['УПС!', 'Путь к файлу выгрузок пуст']
    if os.path.isdir(path_file):
        return ['УПС!', 'Указанный путь к файлу выгрузок является директорией']
    finish_path = lineedit_finish_folder.text().strip()
    if not finish_path:
        return ['УПС!', 'Путь к конечной папке пуст']
    if os.path.isfile(finish_path):
        return ['УПС!', 'Путь к конечной папке является файлом']
    file_name = lineedit_file_name.text().strip()
    if not file_name:
        return ['УПС!', 'Нет имени файла']

    return {'path_file': path_file, 'finish_path': finish_path, 'file_name': file_name}

