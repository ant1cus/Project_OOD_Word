import os
import pathlib
import math

import numpy as np
import pandas as pd
import docx
from docx.enum.section import WD_ORIENTATION
from docxtpl import DocxTemplate
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import _Cell
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Cm
from docx.enum.table import WD_ROW_HEIGHT_RULE

import datetime
import traceback

from PyQt5.QtCore import QThread, pyqtSignal


class CreateTable(QThread):  # Если требуется вставить колонтитулы
    progress = pyqtSignal(int)  # Сигнал для прогресс бара
    status = pyqtSignal(str)  # Сигнал для статус бара
    messageChanged = pyqtSignal(str, str)
    errors = pyqtSignal()

    def __init__(self, incoming_data):  # Список переданных элементов.
        QThread.__init__(self)
        self.path_file = incoming_data['path_file']
        self.finish_path = incoming_data['finish_path']
        self.file_name = incoming_data['file_name']
        self.queue = incoming_data['queue']
        self.logging = incoming_data['logging']

    def set_vertical_cell_direction(self, cell: _Cell, direction: str):
        # direction: tbRl -- top to bottom, btLr -- bottom to top
        assert direction in ("tbRl", "btLr")
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        textDirection = OxmlElement('w:textDirection')
        textDirection.set(qn('w:val'), direction)  # btLr tbRl
        tcPr.append(textDirection)

    def run(self):
        try:
            progress = 0
            self.logging.info("\n*******************************************************************************\n")
            self.logging.info('Начинаем создание таблицы')
            self.status.emit('Считываем значения из файла')
            self.progress.emit(progress)
            time_start = datetime.datetime.now()
            index = ['numSet', 'numTs', 'nameTs', 'firm', 'model', 'sn', 'quant',
                     'secTs', 'secRom', 'first', 'second', 'secondK', 'mod', 'dat', 'snMni',
                     'textMni', 'plat', 'textPlat']
            name_col = ['№ комплекта', '№ ТС в комплекте', 'Наименование ТС', 'Фирма-производитель', 'Модель (Тип)',
                        'Заводской номер', 'Кол-во, ТС  шт.', 'Степень секретности информации, обрабатываемой ТС',
                        'Категория помещений, в котором установлено ТС', 'СЗЗ-1', 'СЗЗ-2', 'СЗЗ-2к',
                        'Наименование и модель удалённого модуля позволяющего организовать канал обмена информацией',
                        'Тип, наименование датчиков физических полей в составе ТС',
                        'Тип, модель, заводской  номер подлежащего регистрации несъёмного МНИ',
                        'Краткое описание места размещения несъёмного МНИ',
                        'Платы содержащие элементы накопления и хранения информации',
                        'Краткое описание места размещения плат, содержащих элементы накопления и хранения информации']
            name_1_col = ['Тип и кол-во СЗЗ, нанесенных на каждое ТС', 'Несъёмные МНИ в составе ТС',
                          'Элемент накопления и хранения информации в составе ТС']
            df = pd.read_csv(self.path_file, delimiter='|', encoding='ANSI', header=None)
            serial_number = ''
            incoming_errors = []
            if df[0].isnull().any():
                for ind, (first, second) in enumerate(zip(df[0].to_numpy(), df[1].to_numpy())):
                    if second == 1:
                        serial_number = first if pd.isna(first) is False else ''  # как вернуться к предыдущей?
                    if pd.isna(first):
                        incoming_errors.append(str(ind + 1))
                        df.loc[ind, 0] = serial_number
            if (int(math.log10(df.loc[0, 0]))+1) == 8:  # Для сверки номеров. Если это серийник - преобразование к int
                df = df.astype({0: np.int})
                df = df.astype({0: np.str})
                df[0] = ['00' + element for element in df[0]]
            table_contents = []
            len_rows = {}
            number_set = []
            number_set_val = 1000000000
            progress += 15
            self.logging.info('Преобразовываем df и формируем таблицу высоты, если требуется')
            self.status.emit('Преобразование данных')
            self.progress.emit(progress)
            previous_val = ''
            for ind_row, row in enumerate(df.itertuples()):
                len_string = 0
                if row[1] != number_set_val:
                    number_set_val = row[1]
                    number_set.append(ind_row + 2)
                list_val = [j for i, j in enumerate(row) if i > 0]
                dict_val = {}
                for ind, val in enumerate(list_val):
                    if ind == 1:
                        if pd.isna(val):
                            dict_val[index[ind]] = previous_val
                        else:
                            dict_val[index[ind]] = int(val) if val % 1 == 0 else val
                            previous_val = dict_val[index[ind]]
                    elif 8 < ind < 12 and (pd.isna(val) or val == '-'):
                        dict_val[index[ind]] = 0
                    elif ind >= 12 and (pd.isna(val) or val == '-'):
                        dict_val[index[ind]] = 'Отсутствуют'
                    else:
                        if ind == 6 or 8 < ind < 12:
                            dict_val[index[ind]] = int(val)
                        else:
                            dict_val[index[ind]] = val
                    if ind >= 12 and isinstance(val, str) and len(val) > len_string:
                        len_string = len(val)
                table_contents.append(dict_val)
                index_str = ind_row + 2
                if 60 < len_string <= 80:
                    len_rows[index_str] = 7
                elif 80 < len_string <= 100:
                    len_rows[index_str] = 8
                elif 100 < len_string <= 120:
                    len_rows[index_str] = 9
                elif 120 < len_string <= 150:
                    len_rows[index_str] = 10
                elif 150 < len_string <= 170:
                    len_rows[index_str] = 12
                elif 170 < len_string <= 200:
                    len_rows[index_str] = 14
                elif 200 < len_string <= 250:
                    len_rows[index_str] = 16
                elif 250 < len_string <= 300:
                    len_rows[index_str] = 18
                elif len_string > 300:
                    len_rows[index_str] = 20
            number_set.append(len(df) + 2)
            self.logging.info('Удаляем отчёт с таким же именем, если он есть')
            try:
                os.remove(str(pathlib.Path(self.finish_path, str(self.file_name) + '.docx')))
            except FileNotFoundError:
                pass
            progress += 15
            self.logging.info('Создаем шаблон таблицы')
            self.status.emit('Создаем шаблон таблицы')
            self.progress.emit(progress)
            document = docx.Document()  # Открываем
            section = document.sections[0]
            section.orientation = WD_ORIENTATION.LANDSCAPE  # Альбомная ориентация
            section.page_width = Cm(42)
            section.page_height = Cm(29.7)
            table = document.add_table(rows=5, cols=18, style='Table Grid')
            table.autofit = True
            table.cell(2, 0).merge(table.cell(2, 17))
            table.cell(4, 0).merge(table.cell(4, 17))
            for i in range(18):
                if i < 9 or i in [12, 13]:
                    table.cell(0, i).merge(table.cell(1, i))
                    table.cell(0, i).text = name_col[i]
                elif i == 9:
                    table.cell(0, i).merge(table.cell(0, 11))
                    table.cell(1, i).text = name_col[i]
                    table.cell(0, i).text = name_1_col[0]
                elif i in [10, 11]:
                    table.cell(1, i).text = name_col[i]
                elif i == 14:
                    table.cell(0, i).merge(table.cell(0, 15))
                    table.cell(1, i).text = name_col[i]
                    table.cell(0, i).text = name_1_col[1]
                elif i == 16:
                    table.cell(0, i).merge(table.cell(0, 17))
                    table.cell(1, i).text = name_col[i]
                    table.cell(0, i).text = name_1_col[2]
                else:
                    table.cell(1, i).text = name_col[i]
                if i in [0, 1, 6, 7, 8, 12, 13]:
                    self.set_vertical_cell_direction(table.cell(0, i), "btLr")
                elif 8 < i < 12 or i > 13:
                    self.set_vertical_cell_direction(table.cell(1, i), "btLr")
                table.cell(1, i).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                table.cell(0, i).vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                table.cell(1, i).vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            table.cell(0, 9).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            table.cell(0, 14).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            table.cell(0, 16).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
            table.cell(2, 0).text = '{%tr for item in table_contents %}'

            # table.rows[3].height_rule = WD_ROW_HEIGHT_RULE.AUTO
            table.rows[1].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            table.rows[1].height = Cm(7)
            table.rows[3].height_rule = WD_ROW_HEIGHT_RULE.EXACTLY
            table.rows[3].height = Cm(5)
            for i, j in enumerate(index):
                table.cell(3, i).text = '{{item.' + j + '}}'
                table.cell(3, i).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                table.cell(3, i).vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                if i > 11:
                    self.set_vertical_cell_direction(table.cell(3, i), "btLr")

            table.cell(4, 0).text = '{%tr endfor %}'
            context = {'table_contents': table_contents}
            document.save(str(pathlib.Path(self.finish_path, str(self.file_name) + '.docx')))
            progress += 5
            self.logging.info('Заносим данные в таблицу')
            self.status.emit('Заносим данные в таблицу')
            self.progress.emit(progress)
            template = DocxTemplate(str(pathlib.Path(self.finish_path, str(self.file_name) + '.docx')))
            template.render(context)
            template.save(str(pathlib.Path(self.finish_path, str(self.file_name) + '.docx')))
            progress += 5
            self.progress.emit(progress)
            self.logging.info('Объединяем номера комплектов, проверяем высоту столбцов и изменяем, если нужно')
            self.status.emit('Изменяем форматирование')
            number_start = False
            number_finish = False
            document = docx.Document(str(pathlib.Path(self.finish_path, str(self.file_name) + '.docx')))
            table = document.tables[0]
            percent = 50 / len(number_set)
            for number in number_set:
                if number_start is False:
                    number_start = number
                elif number_finish is False:
                    number_finish = number - 1
                    self.status.emit(f'Изменяем форматирование строк с {number_start} по {number_finish}')
                    value = table.cell(number_start, 0).text
                    table.cell(number_start, 0).merge(table.cell(number_finish, 0))
                    table.cell(number_start, 0).text = value
                    table.cell(number_start, 0).vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    table.cell(number_start, 0).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    self.set_vertical_cell_direction(table.cell(number_start, 0), "btLr")
                    value_second = table.cell(number_start, 1).text
                    number_start_sec = number_start
                    for numb in range(number_start + 1, number_finish + 1):
                        if table.cell(numb, 1).text == value_second:
                            table.cell(number_start_sec, 1).merge(table.cell(numb, 1))
                            table.cell(number_start_sec, 1).text = value_second
                            table.cell(number_start_sec, 1).vertical_alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                            table.cell(number_start_sec, 1).paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                        value_second = table.cell(numb, 1).text
                        number_start_sec = numb
                    number_start = number
                    number_finish = False
                if len(number_set) > 100 and number_set.index(number) % 50 == 0:
                    document.save(str(pathlib.Path(self.finish_path, str(self.file_name) + '.docx')))
                    document = docx.Document(str(pathlib.Path(self.finish_path, str(self.file_name) + '.docx')))
                    table = document.tables[0]
                progress += percent
                self.progress.emit(int(progress))
            # progress += 50
            # self.progress.emit(progress)
            if len_rows:
                percent = 10/len(len_rows)
                # document = docx.Document(str(pathlib.Path(self.finish_path, str(self.file_name) + '.docx')))
                # table = document.tables[0]
                for key in len_rows:
                    table.rows[key].height = Cm(len_rows[key])
                    progress += percent
                    self.progress.emit(progress)
            document.save(str(pathlib.Path(self.finish_path, str(self.file_name) + '.docx')))
            os.startfile(str(pathlib.Path(self.finish_path, str(self.file_name) + '.docx')))
            self.logging.info("Конец программы, время работы: " + str(datetime.datetime.now() - time_start))
            self.logging.info("\n*******************************************************************************\n")
            if incoming_errors:
                self.logging.info("Выводим ошибки")
                self.status.emit('Готово, есть ошибки.')  # Посылаем значние если готово
                self.queue.put({'errors': incoming_errors})
                self.errors.emit()
            else:
                self.status.emit('Готово!')  # Посылаем значние если готово
                self.progress.emit(100)  # Завершаем прогресс бар
        except BaseException as e:  # Если ошибка
            self.status.emit('Ошибка')  # Сообщение в статус бар
            self.logging.error("Ошибка:\n " + str(e) + '\n' + traceback.format_exc())
            self.progress.emit(0)
