import pdfplumber
import re
from dataclasses import dataclass, field, asdict
from typing import Optional

@dataclass
class Endpoint:
    """Точка подключения (начало или конец кабеля)."""
    room_code:   Optional[str] = '?'   # шифр помещения
    device_name: Optional[str] = '?'   # наименование устройства/оборудования
    x: Optional[str] = '?'
    y: Optional[str] = '?'
    z: Optional[str] = '?'

@dataclass
class CableSpec:
    """Характеристики кабеля одного варианта (проект / факт)."""
#    brand:             Optional[str] = None   # марка кабеля
    cable_mark:        Optional[str] = '?'   # тип кабеля
    voltage_kv:        Optional[str] = '?' # напряжение, кВ
    cores:             Optional[int] = '?'   # число жил
    cross_section_mm2: Optional[str] = '?' # сечение жилы, мм²
    reserve_cores:     Optional[int] = '?'   # число резервных жил
    length_m:          Optional[int] = None # длина, м

@dataclass
class CableRecord:
    """Универсальная запись кабельного журнала любого типа."""
    cable_no:       Optional[str] = '?'
    cable_marking:  Optional[str] = '?'
    install_unit:   Optional[str] = '?'
    mutual_reserve: Optional[str] = '?'
    src: Endpoint = field(default_factory=Endpoint)
    dst: Endpoint = field(default_factory=Endpoint)
    laying_route:   Optional[str] = '?'
    project: CableSpec = field(default_factory=CableSpec)
    actual:  CableSpec = field(default_factory=CableSpec)  # пустой = факта нет
    note:           Optional[str] = '?'

ROW_MAP = {
    1: [
        # физическая строка 1 блока
        (0, 0,  "mutual_reserve"),   (0, 1,  "cable_no"),
        (0, 2,  "install_unit"),     (0, 3,  "cable_marking"),
        (0, 4,  "src.room_code"),    (0, 5,  "src.x"), (0, 6,  "src.y"), (0, 7,  "src.z"),
        (0, 8,  "dst.room_code"),    (0, 9,  "dst.x"), (0, 10, "dst.y"), (0, 11, "dst.z"),
        (0, 12, "project.cable_mark"),           # ячейка «тип + напряжение»
        (0, 13, "project.cores"),                # ячейка «число жил + сечение»
        (0, 14, "project.reserve_cores"),
        (0, 15, "project.length_m"), (0, 16, "laying_route"), (0, 17, "note"),
        # физическая строка 2 блока: наименования устройств
        (1, 4,  "src.device_name"),   # объединённая ячейка под группой «откуда»
        (1, 8,  "dst.device_name"),   # объединённая ячейка под группой «куда»
    ],
    2: [
        (0, 0, "cable_no"),
        (0, 1, "src.device_name"), (0, 2, "dst.device_name"),
        (0, 3, "project.cable_mark"), (0, 4, "project.cores"), (0, 5, "project.length_m"),
        (0, 6, "actual.cable_mark"),  (0, 7, "actual.cores"),  (0, 8, "actual.length_m"),
        (0, 9, "note"),
    ],
    3: [
        (0, 0, "cable_no"), (0, 1, "src.device_name"), (0, 2, "dst.device_name"),
        (0, 3, "laying_route"), (0, 4, "project.cable_mark"), (0, 5, "project.cores"), 
        (0, 6, "project.length_m"), (0, 7, "actual.cable_mark"),  (0, 8, "actual.cores"),  
        (0, 9, "actual.length_m"),
        ],
    4: [
        (0, 0, "cable_no"), (0, 1,  "install_unit"), (0, 2,  "cable_marking"),
        (0, 3, "src.device_name"), (0, 4,  "src.x"), (0, 5,  "src.y"), (0, 6,  "src.z"),
        (0, 7, "dst.device_name"), (0, 8,  "dst.x"), (0, 9, "dst.y"), (0, 10, "dst.z"),
        (0, 11, "project.length_m"), (0, 12, "laying_route"), (0, 13, "note"),
        
        ]
}

def fill(record: CableRecord, path: str, value):
    """record, 'src.x', 12.5 -> record.src.x = 12.5"""
    target = record
    *parents, attr = path.split(".")
    for p in parents:
        target = getattr(target, p)
    setattr(target, attr, value)

def parse_block(block, mapping):
    rec = CableRecord()
    for row_i, col_i, path in mapping:
        if row_i < len(block) and col_i < len(block[row_i]):
            value = (block[row_i][col_i] or "").strip()
            if value:
                fill(rec, path, value)   # fill из прошлого ответа
    return rec

def flatten(rec: CableRecord) -> dict:
    # здесь не используется
    flat = {}
    for k, v in asdict(rec).items():
        if isinstance(v, dict):
            flat.update({f"{k}_{k2}": v2 for k2, v2 in v.items()})
        else:
            flat[k] = v
    return flat

def get_raw_value(block, field_path, row_map):
    """Достаёт сырое значение поля из блока, не запуская полный парсинг."""
    for row_i, col_i, path in row_map:
        if path == field_path:
            if row_i < len(block) and col_i < len(block[row_i]):
                val = block[row_i][col_i]
                return val.strip() if val else ""
    return ""

# Кол-во строк в шапке в зависимости от cs_type
TAB_TITLE_ROWS_COUNT = {
    1: 5,
    2: 3,
    3: 3,
    4: 2}

def prune_empty_columns(raw_table, cs_type):
    # Удаляет пустые столбцы (у которых нет заголовка)
    header = raw_table[:TAB_TITLE_ROWS_COUNT[cs_type]]
    keep = [j for j in range(max(len(r) for r in header))
            if any((r[j] or "").strip() for r in header
                   if j < len(r))]                             # ваш шаг 3
    new_table = [[(row[j] if j < len(row) else "") for j in keep]
           for row in raw_table]
    assert all(len(r) == len(keep) for r in new_table)
    return new_table#, header 

TITLES = {
    1: ['Направление', 'Характеристика'],
    2: ['Трасса', 'Кабель', 'Примечание'],
    3: ['Трасса', 'Участок', 'Кабель'],
    4: ['Марка', 'Координаты']    
        }

def get_cs_type(row):
    for key, words in TITLES.items():
        for word in words:
            is_in_title = False
            for col in row:
                if col and word in col: 
                    is_in_title = True
                    continue
            if not is_in_title: continue
        if is_in_title: return key
    raise ValueError('Неизвестный тип кабельного журнала')
            
    

def main():
    pdf_path = 'Тип_4.pdf'
    #pdf_path = '1102-01UBR-2415-ED_1102-00133-2415-ЭМ_r1.pdf'
    with pdfplumber.open(pdf_path) as pdf:
        print(pdf_path)
        print(f'Количество страниц в файле: {len(pdf.pages)}')
        for i, page in enumerate(pdf.pages):
            print(f'Страница {i}')
            tabs = page.find_tables()
            #tabs = page.extract_tables()
            print(f'Количество таблиц на странице: {len(tabs)}')
            for j, table in enumerate(tabs):
                tab = table.extract()
                cols_count = len(tab[0])
                if cols_count <= 14: continue
                print(process_tab(tab))
                   
def process_tab(tab):
    cable_recs = []
    cs_type = get_cs_type(tab[0])
    # Удаляем лишние столбцы
    new_tab = prune_empty_columns(tab, cs_type)
    # Удаляем пустые строки
    new_tab = [row for row in new_tab if any(cell is not None and cell != "" for cell in row)]
    
    #cs_type = 1  # тип представления кабельного журнала
    title_rows_count = TAB_TITLE_ROWS_COUNT[cs_type]
    block_size = 2 if cs_type == 1 else 1

    #print(f'Таблица {j}')
    #print(f'Количество строк в таблице: {len(tab)}')
    #print(f'Количество столбцов в строке: {len(row0)}')
    #for k, row in enumerate(new_tab, start = title_rows_count):
        #if all([not col for col in row[1:]]): continue
        #print(row)
    k = title_rows_count
    while k < len(new_tab):
        length_m = get_raw_value(new_tab[k:], 'project.length_m', ROW_MAP[cs_type])
            
        if not length_m or not re.search(r'\b[0-9]{1,5}\b', length_m):
            k += 1
            continue
        if k + block_size <= len(new_tab):
            rec = parse_block(new_tab[k:k+block_size], ROW_MAP[cs_type])
            #print(rec)
            cable_recs.append(rec)
        k += block_size    
    return cs_type, cable_recs            

if __name__ == '__main__':
    main()
