import os
import shutil
import pandas as pd
import xlwings as xw
import unicodedata
from pathlib import Path

# --- КОНФИГУРАЦИЯ ---
BASE_DIR = Path(__file__).parent
MAIN_FILE_NAME = 'ВБ нарезать.xlsx' # Или 'Каталог ОЗОН.xlsx'
CATEGORY_COLUMN = 'Категория продавца' # Тип* для Озона
OUTPUT_FOLDER = BASE_DIR / "_готово"
LOG_FILE = BASE_DIR / "process_log.txt"

def normalize_text(s):
    """Удаляет лишние пробелы и нормализует символы Unicode."""
    return unicodedata.normalize("NFC", str(s).strip())

def write_data_to_excel(template_path, df_to_append, start_row=5):
    """Безопасная запись данных в Excel через xlwings."""
    app = xw.App(visible=False)
    try:
        wb = app.books.open(template_path)
        sheet = wb.sheets[0]
        
        # Автоматический маппинг колонок
        header_range = sheet.range("A3").expand("right") # Настройка под ваш шаблон
        header_map = {h: xw.utils.col_name(i + 1) for i, h in enumerate(header_range.value) if h}
        
        # Поиск первой пустой строки
        last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
        current_row = max(last_row + 1, start_row)

        for header, col_letter in header_map.items():
            if header in df_to_append.columns:
                values = df_to_append[header].tolist()
                rng = sheet.range(f"{col_letter}{current_row}").resize(len(values), 1)
                rng.value = [[v] for v in values]
        
        wb.save()
        wb.close()
    finally:
        app.quit() # Гарантированно закрываем процесс Excel

def main():
    if not OUTPUT_FOLDER.exists():
        OUTPUT_FOLDER.mkdir(parents=True)

    main_file_path = BASE_DIR / MAIN_FILE_NAME
    if not main_file_path.exists():
        print(f"[!] Файл {MAIN_FILE_NAME} не найден в {BASE_DIR}")
        return

    df = pd.read_excel(main_file_path)
    
    for category in df[CATEGORY_COLUMN].dropna().unique():
        cat_name = normalize_text(category)
        target_file = BASE_DIR / f"{cat_name}.xlsx"
        
        if not target_file.exists():
            print(f"[Пропуск] Шаблон для '{cat_name}' не найден.")
            continue

        filtered_df = df[df[CATEGORY_COLUMN] == category]
        
        try:
            write_data_to_excel(target_file, filtered_df)
            
            # Перемещение в папку готово
            dest_path = OUTPUT_FOLDER / f"{cat_name}.xlsx"
            shutil.move(str(target_file), str(dest_path))
            print(f"[OK] Категория '{cat_name}' обработана.")
        except Exception as e:
            print(f"[Ошибка] Ошибка в {cat_name}: {e}")

if __name__ == '__main__':
    main()