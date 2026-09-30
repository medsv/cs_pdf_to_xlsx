#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Программа для распознавания кабельных ведомостей из PDF-файлов
и формирования сводного Excel-файла по образцу Output.xlsx.

Использование:
    python process_pdfs.py <папка_с_PDF> [путь_к_выходному_Excel]

Пример:
    python process_pdfs.py /home/z/my-project/test_pdfs /home/z/my-project/download/Output.xlsx

Входные данные:  PDF-файлы с именами Input1.pdf, Input2.pdf, ... в указанной папке
Выходные данные: Excel-файл со структурой, идентичной Output.xlsx
"""

import os
import re
import sys
import glob
import openpyxl
import pdfplumber


# ============================================================================
# Парсинг вертикального (повёрнутого) текста
# ============================================================================
def parse_vertical_text(text):
    """
    Восстанавливает текст, записанный вертикально в PDF.

    pdfplumber при извлечении повёрнутого на 90° текста часто выдаёт символы
    в порядке, обратном читаемому. Алгоритм:
      1. Разбить по переводам строк
      2. Реверсировать символы в каждой строке
      3. Реверсировать порядок строк
      4. Склеить без разделителя
    Применяется только когда значение похоже на вертикальный текст
    (несколько коротких строк по 1-4 символа).
    """
    if not text:
        return text
    lines = text.split('\n')
    non_empty = [l for l in lines if l.strip()]
    # Критерий "вертикального" текста: более 2 строк, каждая не длиннее 4 символов
    if len(non_empty) > 2 and all(len(l.strip()) <= 4 for l in non_empty):
        reversed_lines = [line[::-1] for line in lines]
        reversed_lines.reverse()
        return ''.join(reversed_lines)
    return text


# ============================================================================
# Извлечение констант из титульного блока PDF
# ============================================================================
def extract_title_info(full_text):
    """
    Извлекает код здания и шифр РД из титульного блока.

    Формат блока: "1102-10UHJ-3148-ED/1102-00010-3148-ЭМ"
    Возвращает: (здание, шифр_РД), например ("10UHJ", "3148-ED").
    Если не найдено — возвращает ("?", "?").
    """
    # Шаблон: digits-WORD-digits-WORD/  (группа 1 — здание, группа 2 — РД)
    #match = re.search(r'\d+-([A-Z0-9]+)-(\d+-[A-Z]+)/', full_text)
    match = re.search(r'\d+-([A-Z0-9]+)-(\d+-[A-ZА-Я]+)\/', full_text)
    
    if match:
        return match.group(1), match.group(2)
    return '?', '?'


def extract_change_number(full_text):
    """
    Извлекает номер изменения из блока изменений.

    Блок имеет вид: "1 - Зам 4506-26 10.08.26"
    Возвращает: "изм.N" где N — номер изменения, или "?" если не найдено.
    """
    # Шаблон: число, дефис, "Зам" (с возможными пробелами/переводами строк)
    match = re.search(r'(\d+)\s*\-\s*Зам', full_text)
    if match:
        return f'изм.{match.group(1)}'
    return '?'


# ============================================================================
# Определение типа кабеля по маркировке
# ============================================================================
def determine_cable_type(cable_marking):
    """
    Определяет тип кабеля по первой букве маркировки.

    По ГОСТ на маркировку кабелей:
      - "К" в начале = Контрольный кабель (например КВВГЭнг, КВБбШв)
      - прочие буквы (А, В, С, П...) = Силовой кабель
    Если маркировка отсутствует — "?".
    """
    if not cable_marking or cable_marking == '?':
        return '?'
    first_char = cable_marking[0].upper()
    if first_char == 'К':
        return 'Контрольный'
    # Для других типов не делаем предположений
    return '?'


# ============================================================================
# Разделение информации об устройстве на KKS и наименование
# ============================================================================
def split_device_info(device_text):
    """
    Разделяет текст ячейки устройства на KKS-код и наименование.

    Формат в PDF: "KKS_код\nНаименование" или "KKS_код Наименование".
    Возвращает: (kks, name). Если данных нет — ("?", "?").
    """
    if not device_text:
        return '?', '?'
    text = device_text.replace('\n', ' ').strip()
    if not text:
        return '?', '?'
    parts = text.split(' ', 1)
    if len(parts) >= 2:
        return parts[0], parts[1]
    if len(parts) == 1 and parts[0]:
        return parts[0], '?'
    return '?', '?'

import all_templates as at


def extract_cable_mark_sec(text):
    # Паттерн для марки кабеля: буквы, цифры, скобки, дефисы
    cable_mark_pattern = r"[А-Яа-я]{3,10}(?:нг)?(\([АA]{1}\))?(?:-[A-Z]{1,4})?(?:-[А-Я]{2})?"
    
    # Паттерн для сечения и напряжения: 10х1,5-0,66 (кириллическая «х» или латинская x)
    section_voltage_pattern = r"\b(?:\d{1,2}\s*[xXхХ]{1}\s*){1,2}(?:\d{1,3}(?:[.,]\d{1,2})?)?(?:\s*-\s*\d{1,3}(?:[.,]\d{1,2})?)?"
    
    # Ищем первое вхождение каждого фрагмента
    cable_match = re.search(cable_mark_pattern, text)
    section_match = re.search(section_voltage_pattern, text)
    
    cable_mark = cable_match.group() if cable_match else None
    section_voltage = section_match.group() if section_match else None
    
    #print("Марка кабеля:", cable_mark)
    #print("Сечение и напряжение:", section_voltage)
    
    # Удаляем найденные фрагменты из строки (по одному)
    cleaned_text = text
    if cable_mark:
        cleaned_text = cleaned_text.replace(cable_mark, '', 1)
    if section_voltage:
        cleaned_text = cleaned_text.replace(section_voltage, '', 1)
    
    cleaned_text = ' '.join(cleaned_text.split())  # убираем лишние пробелы и переносы
    #print("Очищенная строка:", cleaned_text)
    
    return cable_mark, section_voltage, cleaned_text
    

# ============================================================================
# Обработка одной таблицы
# ============================================================================
def process_table(table, building, rd_code, change_num):
    """
    Обрабатывает таблицу и извлекает записи о кабелях.

    Каждая запись о кабеле занимает 2 строки:
      - строка N: содержит номер кабеля, маркировку, тип, сечение, длину,
                  координаты "Откуда"/"Куда", трассировку
      - строка N+1: содержит наименования устройств "Откуда" и "Куда"
                  (col 7 и col 11)

    Возвращает: список словарей с полями кабеля.
    """

    cs_type, cable_recs = at.process_tab(table)
    #print(cable_recs)

 


    #print(col_numbers)
    

    
    cables = []
    for i, cr in enumerate(cable_recs):
        #print(row)
        # Длина — целое число, если возможно
        length = cr.project.length_m
        if length != '?':
            try:
                length = int(length)
            except (ValueError, TypeError):
                pass

        cable = {
            'Шифр РД': rd_code,
            'Актуальный ИЗМ': change_num,
            'Здание': building,
            'Тип кабеля': determine_cable_type(cr.project.cable_mark),
            '№ нитки': parse_vertical_text(cr.cable_no),
            'KKS нитки': cr.cable_marking,
            'Марки кабеля': cr.project.cable_mark,
            'Сечение кабеля': cr.project.cores,
            'Напряжение, кВ':  cr.project.voltage_kv,   # в PDF нет данных — оставляем пусто
            'OTKУДAKKS': '?',
            'OTKУДA Наименование': cr.src.device_name,
            'KУДA KKS': '?',
            'KУДA Наименование': cr.dst.device_name,
            'Длина кабельной линии КЖ': length,
            'Длина': None,           # фактическая длина — в PDF нет
            'Дельта': None,          # дельта — в PDF нет
            'Трассировка':  cr.laying_route,
            'Дата': None,
            'Организация': None,
            'ФИО исполнителя': None,
            'Объем по ИД': None,
            'Подписание ИНЖ': None,
            'Примечания': None,
            'максимально': None,
            'ТОМ ИД': None,
            'Участок': None,
        }

        if cs_type == 1:
            from_kks, from_name = split_device_info(cr.src.device_name)
            to_kks, to_name = split_device_info(cr.dst.device_name)
            if any(ch.isdigit() for ch in from_kks):
                cable['OTKУДAKKS'] = from_kks
                cable['OTKУДA Наименование'] = from_name
            if any(ch.isdigit() for ch in to_kks):            
                cable['KУДA KKS'] = to_kks
                cable['KУДA Наименование'] = to_name


            
        if cs_type == 2:
            cable['Трассировка'] = cr.note  # Трассировка пмшется в Примечании
            
        if cs_type == 4:
            # Марка и жилы кабеля 
            cable_mark, section_voltage, cleaned_text = extract_cable_mark_sec(cr.cable_marking)
            cable['Марки кабеля'] = cable_mark
            cable['Сечение кабеля'] = section_voltage
            cable['KKS нитки'] = cleaned_text
            
        cables.append(cable)
 
    return cables


# ============================================================================
# Обработка одного PDF
# ============================================================================
def process_pdf(pdf_path):
    """
    Обрабатывает один PDF-файл и возвращает список записей о кабелях.
    Извлекает константы (Шифр РД, Здание, ИЗМ) из титульного блока,
    затем проходит по всем таблицам во всех страницах PDF.
    """
    cables = []
    full_text = ''

    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = pdf.pages[0].extract_text()
            # Извлекаем константы из титульного блока
            building, rd_code = extract_title_info(full_text)
            change_num = extract_change_number(full_text)
            # Проходим по таблицам на каждой странице
            for page in pdf.pages:
                if change_num == "?":
                    change_num = extract_change_number(page.extract_text())

                tables = page.find_tables()
                for tab in tables:
                    table = tab.extract()
                    if len(table[0])<20: continue
                    cables.extend(process_table(table, building, rd_code, change_num))
    except Exception as e:
        print(f"  ОШИБКА при обработке {os.path.basename(pdf_path)}: {e}")
        return []

    return cables


# ============================================================================
# Запись в Excel по образцу Output.xlsx
# ============================================================================
def write_excel(cables, output_path):
    """
    Записывает список кабелей в Excel-файл.
    Структура файла идентична образцу Output.xlsx:
      - строка 3: заголовки столбцов
      - строка 4: номера столбцов (1..28)
      - строка 5+: данные
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Лист1'

    # Заголовки в строке 3 (точное соответствие образцу Output.xlsx)
    headers = [
        ('B',  'Шифр РД'),
        ('C',  'Актуальный ИЗМ'),
        ('D',  'Здание'),
        ('E',  'Тип кабеля'),
        ('F',  '№ нитки'),
        ('G',  'KKS нитки'),
        ('H',  'Марки кабеля'),
        ('I',  'Сечение кабеля'),
        ('J',  'Напряжение, кВ'),
        ('K',  'OTKУДAKKS'),
        ('L',  'OTKУДA Наименование оборудования, наименование помещений'),
        ('M',  'KУДA KKS'),
        ('N',  'KУДA Наименование оборудования, наименовоание помещений'),
        ('O',  'Длина кабельной линии КЖ'),
        ('P',  'Длина'),
        ('Q',  'Дельта'),
        ('R',  'Трассировка'),
        ('S',  'Дата'),
        ('T',  'Организация'),
        ('U',  'ФИО исполнителя'),
        ('V',  'Объем по ИД'),
        ('W',  'Подписание ИНЖ'),
        ('X',  'Примечания'),
        ('Y',  'максимально'),
        ('Z',  'ТОМ ИД'),
        ('AA', None),    # В образце заголовок пуст
        ('AB', None),    # В образце заголовок пуст
        ('AC', 'Участок'),
    ]

    col_numbers = ['B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',
                   'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y',
                   'Z', 'AA', 'AB', 'AC']

    # Заголовки
    for col, header in headers:
        if header is not None:
            ws[f'{col}3'] = header

    # Номера столбцов (1..28)
    for idx, col in enumerate(col_numbers, start=1):
        ws[f'{col}4'] = idx

    # Маппинг: буква столбца -> ключ в словаре кабеля
    field_map = {
        'B':  'Шифр РД',
        'C':  'Актуальный ИЗМ',
        'D':  'Здание',
        'E':  'Тип кабеля',
        'F':  '№ нитки',
        'G':  'KKS нитки',
        'H':  'Марки кабеля',
        'I':  'Сечение кабеля',
        'J':  'Напряжение, кВ',
        'K':  'OTKУДAKKS',
        'L':  'OTKУДA Наименование',
        'M':  'KУДA KKS',
        'N':  'KУДA Наименование',
        'O':  'Длина кабельной линии КЖ',
        'P':  'Длина',
        'Q':  'Дельта',
        'R':  'Трассировка',
        'S':  'Дата',
        'T':  'Организация',
        'U':  'ФИО исполнителя',
        'V':  'Объем по ИД',
        'W':  'Подписание ИНЖ',
        'X':  'Примечания',
        'Y':  'максимально',
        'Z':  'ТОМ ИД',
        'AC': 'Участок',
    }

    # Данные начиная со строки 5
    for idx, cable in enumerate(cables):
        row_num = 5 + idx
        for col, field in field_map.items():
            value = cable.get(field)
            if value is not None:
                ws[f'{col}{row_num}'] = value

    # Ширина столбцов как в образце
    ws.column_dimensions['L'].width = 23.54296875
    ws.column_dimensions['N'].width = 25.36328125
    ws.column_dimensions['R'].width = 20.54296875

    # Директория выхода
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    wb.save(output_path)


# ============================================================================
# Точка входа
# ============================================================================
def main():
    if len(sys.argv) < 2:
        print("Использование: python process_pdfs.py <папка_с_PDF> [путь_к_Excel]")
        print("  папка_с_PDF: папка, содержащая Input1.pdf, Input2.pdf, ...")
        print("  путь_к_Excel: путь к выходному файлу (по умолчанию: Output.xlsx)")
        #sys.exit(1)

    #input_folder = sys.argv[1]
    input_folder = 'C:/Users/sv.medvedev/Documents/00 Задачи/091 Кабельный журнал/Тип_таблицы/test/'
    output_path = sys.argv[2] if len(sys.argv) > 2 else 'Output.xlsx'
    pdf_files = []
    # Найти все PDF по шаблону Input*.pdf (включая Input.pdf без номера)
    #pdf_files = ['Тип_1.pdf', 'Тип_2.pdf', 'Тип_3.pdf', 'Тип_4.pdf']
    #pdf_files = [ 'Тип_4.pdf']
    for pattern in ['*.pdf']:
        pdf_files.extend(glob.glob(os.path.join(input_folder, pattern)))
    pdf_files = sorted(set(pdf_files))

    if not pdf_files:
        print(f"В папке {input_folder} не найдено PDF-файлов по шаблону Input*.pdf")
        sys.exit(1)

    print(f"Найдено PDF-файлов: {len(pdf_files)}")
    for f in pdf_files:
        print(f"  - {os.path.basename(f)}")

    all_cables = []
    for pdf_file in pdf_files:
        print(f"\nОбработка {os.path.basename(pdf_file)} ...")
        cables = process_pdf(pdf_file)
        print(f"  Извлечено записей о кабелях: {len(cables)}")
        all_cables.extend(cables)

    write_excel(all_cables, output_path)
    print(f"\nГотово. Выходной файл: {output_path}")
    print(f"Всего записей о кабелях: {len(all_cables)}")


if __name__ == '__main__':
    main()
