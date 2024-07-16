import pandas as pd
from openpyxl import load_workbook, Workbook
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
import time
import psutil
import os

def init_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("start-maximized")
    chrome_options.add_argument("disable-infobars")
    chrome_options.add_argument("--disable-extensions")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("window-size=1920,1080")
    chrome_options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36")

    return webdriver.Chrome(options=chrome_options)

def get_current_chrome_processes():
    chrome_pids = []
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] in ('chrome.exe', 'chromedriver.exe'):
            chrome_pids.append(proc.info['pid'])
    return set(chrome_pids)

def kill_new_chrome_processes(initial_pids):
    current_pids = get_current_chrome_processes()
    new_pids = current_pids - initial_pids
    for pid in new_pids:
        try:
            proc = psutil.Process(pid)
            proc.kill()
        except psutil.NoSuchProcess:
            pass

def scrape_xo_gr(profession, location, page_num):
    driver = init_driver()
    url = f'https://www.xo.gr/search/?what={profession}&where={location}&locId=B2&page={page_num}'
    driver.get(url)
    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'listingWhiteArea')))
    time.sleep(2)

    soup = BeautifulSoup(driver.page_source, 'html.parser')
    listings = soup.find_all('div', class_='listingWhiteArea')
    data = []

    for listing in listings:
        name = listing.find('span', itemprop='name').text
        address = listing.find('span', id=lambda x: x and x.startswith('listingAddress')).text.replace('\n', ', ')
        
        phone, mobile = '', ''
        phone_container = listing.find('div', class_='phoneContainer')
        if phone_container:
            phone_entries = phone_container.find_all('li', class_='phoneEntry')
            phone = phone_entries[0].text.replace(' ', '') if len(phone_entries) > 0 else ''
            mobile = phone_entries[1].text.replace(' ', '') if len(phone_entries) > 1 else ''

        data.append({
            'Name': name,
            'Address': address,
            'Profession': profession,
            'Phone': phone,
            'Mobile': mobile,
            'Email': '',
            'Ωρα': ''
        })

    driver.quit()
    return data

def check_entry_exists(df, phone, mobile):
    phone = str(phone).strip()
    mobile = str(mobile).strip()
    return ((df['Phone'].astype(str).str.strip() == phone) | (df['Mobile'].astype(str).str.strip() == mobile)).any()

def ensure_headers(file_path):
    headers = ['Name', 'Address', 'Profession', 'Phone', 'Mobile', 'Email', 'Ωρα']
    if not os.path.exists(file_path) or os.stat(file_path).st_size == 0:
        # Create a new file with headers
        df = pd.DataFrame(columns=headers)
        df.to_excel(file_path, index=False)
    else:
        # Load existing workbook and check headers
        df = pd.read_excel(file_path)
        missing_headers = [header for header in headers if header not in df.columns]
        if missing_headers:
            for header in missing_headers:
                df[header] = ''
            df.to_excel(file_path, index=False)

def update_excel(file_path, new_data):
    ensure_headers(file_path)

    # Load existing workbook
    df = pd.read_excel(file_path)
    
    # Check for duplicates and append new data
    new_entries = []
    for entry in new_data:
        if not check_entry_exists(df, entry['Phone'], entry['Mobile']):
            new_entries.append(entry)
    
    if new_entries:
        new_df = pd.DataFrame(new_entries)
        df = pd.concat([df, new_df], ignore_index=True)
        df.to_excel(file_path, index=False)

def main():
    profession = input('Enter the profession: ')
    location = input('Enter the location: ')
    page_num = input('Enter the page number: ')
    excel_file = 'info.xlsx'

    initial_chrome_pids = get_current_chrome_processes()
    
    new_data = scrape_xo_gr(profession, location, page_num)
    update_excel(excel_file, new_data)
    
    kill_new_chrome_processes(initial_chrome_pids)
    print('Data has been successfully appended to the Excel file.')

if __name__ == "__main__":
    main()
