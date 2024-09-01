from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from bs4 import BeautifulSoup
import pandas as pd
import time
import re
import os

chrome_options = Options()
chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36")
driver_path = 'chromedriver-win64\\chromedriver.exe'
chrome_service = Service(driver_path)
driver = webdriver.Chrome(service=chrome_service, options=chrome_options)
try:
    driver.get('https://trading.finam.ru/profile/MOEX-LKOH')
    print("Страница загружена...")
    try:
        markets_tab_button = WebDriverWait(driver, 35).until(
            EC.presence_of_element_located((By.ID, 'markets-tab-button'))
        )
        markets_tab_button.click()
        print("Кнопка 'Маркет' нажата.")
    except TimeoutException:
        print("Время ожидания истекло. Кнопка 'Маркет' не найдена.")
    try:
        market_group_dropdown = WebDriverWait(driver, 35).until(
            EC.presence_of_element_located((By.ID, 'market-group-dropdown'))
        )
        market_group_dropdown.click()
        print("Кнопка 'Группа рынка' нажата.")
    except TimeoutException:
        print("Время ожидания истекло. Кнопка 'Группа рынка' не найдена.")
    try:
        time.sleep(1)
        stock_and_funds_element = WebDriverWait(driver, 35).until(
            EC.presence_of_element_located((By.XPATH, "//li[contains(text(), 'Акции и фонды')]"))
        )
        stock_and_funds_element.click()
        print("Нажата кнопка 'Акции и фонды'.")
    except TimeoutException:
        print("Время ожидания истекло. Элемент 'Акции и фонды' не найден.")
    try:
        element = WebDriverWait(driver, 35).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'div[data-testid="virtuoso-item-list"]'))
        )
        print("Элемент найден. Ждем 10 секунд до полной загрузки...")
        driver.execute_script("""
            const targetNode = document.querySelector('div[data-testid="virtuoso-item-list"]');
            const config = { childList: true, subtree: true };
            const callback = function(mutationsList) {
                for(const mutation of mutationsList) {
                    if (mutation.type === 'childList') {
                        mutation.addedNodes.forEach(node => {
                            if (node.nodeType === 1) {
                                node.style.height = '0';
                            }
                        });
                    }
                }
            };
            const observer = new MutationObserver(callback);
            observer.observe(targetNode, config);
            const existingItems = targetNode.children;
            for (let i = 0; i < existingItems.length; i++) {
                existingItems[i].style.height = '0';
                existingItems[i].style.overflow = 'hidden';
            }
        """)
        time.sleep(10)
    except TimeoutException:
        print("Время ожидания истекло. Элемент 'virtuoso-item-list' не найден.")

    html_content = driver.page_source
    soup = BeautifulSoup(html_content, 'html.parser')
    virtuoso_item_lists = soup.select('div[data-testid="virtuoso-item-list"]')
    formatted_text = ""
    if virtuoso_item_lists:
        print("Элементы найдены.")
        for virtuoso_item_list in virtuoso_item_lists:
            button_elements = virtuoso_item_list.find_all('div', role='button')
            for button in button_elements:
                p_elements = button.find_all('p')
                structured_texts = []
                for p in p_elements:
                    text = p.get_text(strip=True)
                    structured_texts.append(text)
                if structured_texts:
                    formatted_text += ' / '.join(structured_texts) + ' /\n'
        #  print(formatted_text)
    else:
        print("Элементы virtuoso-item-list не найдены.")

    lines = formatted_text.strip().split("\n")
    data = []
    for line in lines:
        line = line.strip().strip('/')
        line = line.replace('\xa0', '')
        line = re.sub(r'\s+', '', line)
        if '(' in line:
            left_part = line.split('(')[0]
            percentage = line.split('(')[1].rstrip(')')
        else:
            left_part = line
            percentage = ''
        fields = [field.strip() for field in left_part.split('/') if field.strip()]
        if percentage:
            fields.append(percentage)
        if len(fields) >= 4:
            fields[2] = fields[2].rstrip('₽')
        if len(fields) >= 5:
            fields[3] = fields[3].rstrip('₽')
        data.append(fields)
        # print("Обрабатываемая строка:", line)
        # print("Поля:", fields)

    # export данных в Excel
    if data:
        df = pd.DataFrame(data, columns=['Компания', 'Код', 'Цена', 'Изменение', 'Процент'])
        output_file = 'output.xlsx'

        if os.path.exists(output_file):
            try:
                with open(output_file, 'a'):
                    pass
            except PermissionError:
                print("Закройте таблицу >:(")
            else:
                df.to_excel(output_file, index=False)
                print("Данные успешно сохранены в файл output.xlsx.")
        else:
            df.to_excel(output_file, index=False)
            print("Данные успешно сохранены в файл output.xlsx.")
    else:
        print("Нет данных для сохранения.")
finally:
    driver.quit()