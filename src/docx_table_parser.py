"""
Модуль для извлечения таблиц из docx и их структурирования
"""

import logging
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

logger = logging.getLogger(__name__)


def extract_tables_from_docx(file_path: str) -> list:
    """
    Извлекает таблицы из docx файла и преобразует их в структурированный формат
    """
    doc = Document(file_path)
    tables_data = []
    
    for table_idx, table in enumerate(doc.tables):
        table_info = {
            "table_number": table_idx + 1,
            "headers": [],
            "rows": []
        }
        
        # Извлекаем заголовки (первая строка)
        if len(table.rows) > 0:
            header_row = table.rows[0]
            headers = []
            for cell in header_row.cells:
                text = cell.text.strip().replace('\n', ' ').replace('\r', ' ')
                # Убираем маркдаун форматирование
                text = text.replace('**', '').replace('*', '')
                if text:
                    headers.append(text)
                else:
                    headers.append(f"col_{len(headers)}")
            table_info["headers"] = headers
        
        # Извлекаем данные строк
        for row_idx, row in enumerate(table.rows[1:], 1):
            # Проверяем, не пустая ли строка
            row_data = {}
            is_empty = True
            
            for col_idx, cell in enumerate(row.cells):
                text = cell.text.strip().replace('\n', ' ').replace('\r', ' ')
                text = ' '.join(text.split())
                
                if text and text not in ['', ' ', '—', '-']:
                    is_empty = False
                
                header = table_info["headers"][col_idx] if col_idx < len(table_info["headers"]) else f"col_{col_idx}"
                row_data[header] = text
            
            # Добавляем только непустые строки
            if not is_empty:
                table_info["rows"].append(row_data)
        
        # Добавляем только таблицы с данными
        if table_info["rows"]:
            tables_data.append(table_info)
            logger.info(f"Таблица {table_idx + 1}: {len(table_info['rows'])} строк, заголовки: {table_info['headers'][:3]}...")
    
    return tables_data


def tables_to_json_format(tables_data: list, source_name: str) -> list:
    """
    Преобразует извлеченные таблицы в формат JSON для спецификации
    """
    result = []
    pos_counter = 1
    
    # Маппинг возможных названий столбцов
    column_mapping = {
        'поз': 'pos',
        '№': 'pos',
        'номер': 'pos',
        'наименование': 'name',
        'наименование и техническая характеристика': 'name',
        'характеристика': 'name',
        'тип': 'type_mark',
        'марка': 'type_mark',
        'тип, марка': 'type_mark',
        'обозначение': 'type_mark',
        'код продукции': 'product_code',
        'код': 'product_code',
        'поставщик': 'supplier',
        'ед. изм.': 'unit',
        'единица измерения': 'unit',
        'кол': 'quantity',
        'количество': 'quantity',
        'кол.': 'quantity',
        'масса': 'mass',
        'вес': 'mass',
        'масса 1 ед.': 'mass',
        'примечание': 'note',
        'прим.': 'note'
    }
    
    for table in tables_data:
        headers = [h.lower() for h in table["headers"]]
        
        for row in table["rows"]:
            item = {
                "pos": str(pos_counter),
                "name": "",
                "type_mark": "",
                "product_code": "",
                "supplier": "",
                "unit": "",
                "quantity": "",
                "mass": "",
                "note": "",
                "source": source_name
            }
            
            # Заполняем поля из строки таблицы
            for original_header, value in row.items():
                if not value or value in ['', ' ', '—', '-']:
                    continue
                    
                header_lower = original_header.lower()
                
                # Ищем соответствие
                for key, field in column_mapping.items():
                    if key in header_lower:
                        if field == "quantity" and value.replace('.', '').replace(',', '').isdigit():
                            item[field] = value
                        elif field == "mass" and value.replace('.', '').replace(',', '').isdigit():
                            item[field] = value
                        elif field in ["name", "type_mark", "product_code", "supplier", "unit", "note"]:
                            if item[field]:
                                item[field] += f" {value}"
                            else:
                                item[field] = value
                        break
            
            # Если есть имя, добавляем запись
            if item["name"]:
                result.append(item)
                pos_counter += 1
            else:
                # Пробуем определить имя из первого непустого поля
                for field in ["type_mark", "product_code", "note"]:
                    if row.get(field):
                        item["name"] = row[field]
                        result.append(item)
                        pos_counter += 1
                        break
    
    return result


def extract_all_specifications(file_path: str, source_name: str) -> list:
    """
    Основная функция: извлекает все спецификации из docx файла
    """
    logger.info(f"Извлечение таблиц из {source_name}")
    
    try:
        tables = extract_tables_from_docx(file_path)
        
        if not tables:
            logger.warning(f"В файле {source_name} не найдено таблиц с данными")
            return []
        
        result = tables_to_json_format(tables, source_name)
        logger.info(f"Извлечено {len(result)} записей из таблиц файла {source_name}")
        
        return result
        
    except Exception as e:
        logger.error(f"Ошибка при извлечении таблиц из {source_name}: {e}")
        return []