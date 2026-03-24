import os
import time
import random
import pandas as pd
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- КОНФИГУРАЦИЯ ---
BASE_DIR = Path(__file__).parent
# Драйвер теперь ищем в папке со скриптом
CHROMEDRIVER_PATH = BASE_DIR / "chromedriver.exe" 
EXCEL_INPUT = BASE_DIR / "keywords_input.xlsx"
EXCEL_OUTPUT = BASE_DIR / "keywords_result.xlsx"

MAX_KEYS = 40
MIN_DELAY, MAX_DELAY = 6, 12

def setup_driver():
    options = Options()
    options.add_argument("--start-maximized")
    # Можно добавить headless режим, чтобы браузер не мешал работать
    # options.add_argument("--headless") 
    service = Service(str(CHROMEDRIVER_PATH))
    return webdriver.Chrome(service=service, options=options)

def main():
    if not EXCEL_INPUT.exists():
        print(f"❌ Файл {EXCEL_INPUT.name} не найден!")
        return

    driver = setup_driver()
    df = pd.read_excel(EXCEL_INPUT)
    df["Ключевые слова"] = ""

    try:
        for index, row in df.iterrows():
            query = row.get("Первые 3 слова")
            if not query: continue
            
            print(f"\n🔍 Парсинг: {query}")
            
            # Рандомная пауза "человека"
            time.sleep(random.uniform(MIN_DELAY, MAX_DELAY))
            
            driver.get("https://sellego.com/podbor-klyuchej-wb/")
            
            try:
                search_box = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.ID, "keywordsform-query"))
                )
                search_box.clear()
                search_box.send_keys(query)
                
                driver.find_element(By.CSS_SELECTOR, "button[type='submit']").click()
                
                # Ожидание загрузки таблицы
                time.sleep(8) 
                
                keywords = driver.find_elements(By.XPATH, "//table//tbody/tr/td[1]")
                key_list = [kw.text.strip() for kw in keywords if kw.text.strip()][:MAX_KEYS]
                
                df.at[index, "Ключевые слова"] = ", ".join(key_list)
                print(f"✅ Найдено ключей: {len(key_list)}")
                
                # Промежуточное сохранение, чтобы не потерять данные при сбое
                df.to_excel(EXCEL_OUTPUT, index=False)

            except Exception as e:
                print(f"⚠️ Ошибка на запросе {query}: {e}")
                
    finally:
        driver.quit()
        print("🏁 Работа завершена.")

if __name__ == "__main__":
    main()