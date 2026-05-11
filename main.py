"""
Главный скрипт парсинга таблиц из DOCX
"""

import os
import sys
import json
import logging
import warnings
from datetime import datetime
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))
warnings.filterwarnings('ignore')
os.environ['PYTHONWARNINGS'] = 'ignore'

# === Определяем пути локально ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(BASE_DIR, "input_docs")
OUTPUT_JSON_DIR = os.path.join(BASE_DIR, "output_json")
OUTPUT_EXCEL_DIR = os.path.join(BASE_DIR, "output_excel")
LOGS_DIR = os.path.join(BASE_DIR, "logs")

# Создаём папки
os.makedirs(INPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_JSON_DIR, exist_ok=True)
os.makedirs(OUTPUT_EXCEL_DIR, exist_ok=True)
os.makedirs(LOGS_DIR, exist_ok=True)

# Импорты модулей проекта
from word_reader import get_all_docx_files
from docx_table_parser import extract_all_specifications
from excel_writer import create_excel_from_json


def setup_logging():
    """Настройка логирования"""
    log_file = os.path.join(LOGS_DIR, f"parsing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger(__name__)


def save_to_json(data: list, filename_prefix: str) -> str:
    """Сохраняет данные в JSON файл"""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_prefix = filename_prefix.replace('.docx', '').replace('.', '_')
    json_filename = f"{safe_prefix}_{timestamp}.json"
    json_path = os.path.join(OUTPUT_JSON_DIR, json_filename)
    
    with open(json_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    
    return json_path


def load_json_files() -> list:
    """Загружает все JSON файлы из папки output_json"""
    all_data = []
    
    if not os.path.exists(OUTPUT_JSON_DIR):
        return all_data
    
    for filename in os.listdir(OUTPUT_JSON_DIR):
        if filename.endswith('.json'):
            filepath = os.path.join(OUTPUT_JSON_DIR, filename)
            try:
                with open(filepath, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        all_data.extend(data)
                    elif isinstance(data, dict):
                        all_data.append(data)
            except Exception as e:
                logging.error(f"Ошибка загрузки {filename}: {e}")
    
    return all_data


def merge_json_files() -> list:
    """Объединяет все JSON файлы в один список"""
    all_data = load_json_files()
    
    if all_data:
        save_to_json(all_data, "COMBINED_ALL")
    
    return all_data


def process_documents():
    """Основной процесс обработки документов"""
    logger = setup_logging()
    logger.info("=" * 60)
    logger.info("ЗАПУСК ПАРСИНГА ДОКУМЕНТОВ")
    logger.info("=" * 60)
    
    docx_files = get_all_docx_files(INPUT_DIR)
    
    if not docx_files:
        logger.error(f"Не найдено .docx файлов в папке: {INPUT_DIR}")
        return
    
    logger.info(f"Найдено файлов: {len(docx_files)}")
    for f in docx_files:
        logger.info(f"  - {os.path.basename(f)}")
    
    all_data = []
    
    for file_idx, file_path in enumerate(docx_files, 1):
        filename = os.path.basename(file_path)
        logger.info(f"\n--- Обработка {file_idx}/{len(docx_files)}: {filename} ---")
        
        try:
            file_data = extract_all_specifications(file_path, filename)
            
            if file_data:
                json_path = save_to_json(file_data, filename)
                logger.info(f"  📁 JSON: {json_path} ({len(file_data)} записей)")
                all_data.extend(file_data)
                logger.info(f"  ✅ {len(file_data)} записей")
            else:
                logger.warning(f"  ⚠️ Данные не извлечены")
                
        except Exception as e:
            logger.error(f"Ошибка: {e}")
    
    if not all_data:
        logger.error("Не удалось извлечь данные ни из одного документа")
        return
    
    # Сохраняем объединённый JSON
    save_to_json(all_data, "COMBINED_ALL")
    logger.info(f"\n📊 Всего записей: {len(all_data)}")
    
    # Создаём Excel
    excel_path = create_excel_from_json(all_data, "Спецификация_итоговая")
    logger.info(f"📊 Excel: {excel_path}")
    
    logger.info("=" * 60)
    logger.info("ОБРАБОТКА ЗАВЕРШЕНА")
    logger.info("=" * 60)


def json_to_excel_only():
    """Режим: только JSON -> Excel"""
    logger = setup_logging()
    logger.info("Режим: JSON -> Excel")
    
    all_data = merge_json_files()
    
    if not all_data:
        logger.error("Нет данных в JSON файлах")
        return
    
    excel_path = create_excel_from_json(all_data, "Спецификация_из_JSON")
    logger.info(f"Excel создан: {excel_path}")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--json-to-excel":
        json_to_excel_only()
    else:
        process_documents()