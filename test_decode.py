"""Test script to explore VINMatchPro decode flow."""
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time

options = webdriver.ChromeOptions()
options.add_argument('--headless=new')
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')
options.add_argument('--window-size=1920,1080')
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)

TEST_VIN = 'JT3HN86R0X0197825'

try:
    # Login
    driver.get('https://www.vinmatchpro.com/users/sign_in')
    time.sleep(3)
    driver.find_element(By.ID, 'user_email').send_keys('sales@gfs1.com')
    pf = driver.find_element(By.ID, 'user_password')
    pf.send_keys('Gerald2!')
    pf.send_keys(Keys.RETURN)

    WebDriverWait(driver, 15).until(
        EC.url_changes('https://www.vinmatchpro.com/users/sign_in')
    )
    time.sleep(2)
    print(f'Logged in at: {driver.current_url}')

    # Navigate to Single VIN Decode
    driver.get('https://www.vinmatchpro.com/decode')
    time.sleep(3)
    print(f'Decode page: {driver.current_url}')

    # Find VIN input
    inputs = driver.find_elements(By.TAG_NAME, 'input')
    vin_input = None
    for inp in inputs:
        ph = (inp.get_attribute('placeholder') or '').lower()
        name = (inp.get_attribute('name') or '').lower()
        if 'vin' in ph or 'vin' in name:
            if inp.is_displayed():
                vin_input = inp
                print(f'Found VIN input: placeholder="{inp.get_attribute("placeholder")}" name="{inp.get_attribute("name")}"')
                break

    if not vin_input:
        print('No VIN input found. Page text:')
        print(driver.find_element(By.TAG_NAME, 'body').text[:2000])
        driver.quit()
        exit(1)

    vin_input.clear()
    vin_input.send_keys(TEST_VIN)

    # Find submit button
    buttons = driver.find_elements(By.TAG_NAME, 'button')
    for b in buttons:
        if b.is_displayed():
            txt = b.text.strip()
            btype = b.get_attribute('type')
            print(f'  button: text="{txt}"  type={btype}  displayed={b.is_displayed()}')

    # Click the DECODE VIN button specifically
    for b in buttons:
        if b.is_displayed() and 'decode' in b.text.strip().lower():
            print(f'Clicking: "{b.text.strip()}"')
            driver.execute_script('arguments[0].click()', b)
            break

    time.sleep(10)
    print(f'After submit: {driver.current_url}')

    # Parse results
    tables = driver.find_elements(By.TAG_NAME, 'table')
    print(f'Tables found: {len(tables)}')
    for i, t in enumerate(tables):
        rows = t.find_elements(By.TAG_NAME, 'tr')
        print(f'\n  Table {i} ({len(rows)} rows):')
        for r in rows[:25]:
            tds = r.find_elements(By.TAG_NAME, 'td')
            ths = r.find_elements(By.TAG_NAME, 'th')
            cells = ths + tds
            row_text = ' | '.join(c.text.strip() for c in cells if c.text.strip())
            if row_text:
                print(f'    {row_text}')

    # Also grab body text for keywords
    body_text = driver.find_element(By.TAG_NAME, 'body').text
    keywords = ['year', 'make', 'model', 'engine', 'toyota', '4runner',
                'trim', 'body', 'drive', 'transmission', 'fuel']
    print('\nRelevant lines from page body:')
    for line in body_text.split('\n'):
        line = line.strip()
        if line and any(kw in line.lower() for kw in keywords):
            print(f'  >> {line}')

finally:
    driver.quit()
