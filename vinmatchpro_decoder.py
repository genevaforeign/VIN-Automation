"""
vinmatchpro_decoder.py - Automates VINMatchPro web portal via Selenium.

Logs into the VINMatchPro website, submits a VIN, and parses the decoded
vehicle information (year, make, model, engine, trim, etc.).
"""

import configparser
import re
import sys
import time

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager


DECODE_TIMEOUT = 20  # seconds to wait for page loads


def load_config(config_path: str = 'config.ini') -> dict:
    """Read VINMatchPro credentials from config file."""
    cfg = configparser.ConfigParser()
    cfg.read(config_path)
    return {
        'url': cfg.get('vinmatchpro', 'url'),
        'username': cfg.get('vinmatchpro', 'username'),
        'password': cfg.get('vinmatchpro', 'password'),
    }


def create_driver() -> webdriver.Chrome:
    """Create a Chrome WebDriver instance with standard options."""
    options = webdriver.ChromeOptions()
    options.add_argument('--disable-gpu')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('--window-size=1920,1080')
    # Run headless by default; remove this line to watch the browser
    options.add_argument('--headless=new')

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def login(driver: webdriver.Chrome, url: str, username: str, password: str):
    """Log into the VINMatchPro portal."""
    sign_in_url = url.rstrip('/') + '/users/sign_in'
    driver.get(sign_in_url)
    time.sleep(3)

    driver.find_element(By.ID, 'user_email').send_keys(username)
    pass_field = driver.find_element(By.ID, 'user_password')
    pass_field.send_keys(password)
    pass_field.send_keys(Keys.RETURN)

    # Wait for redirect away from sign-in page
    WebDriverWait(driver, DECODE_TIMEOUT).until(EC.url_changes(sign_in_url))
    time.sleep(2)


def decode_vin(driver: webdriver.Chrome, url: str, vin: str) -> dict:
    """Navigate to the decode page for a VIN and parse the results.

    Returns a dict with keys: vin, year, make, model, body_style,
    drive_type, engine, transmission, trim, fuel_type, and the full
    specs dict.
    """
    decode_url = url.rstrip('/') + f'/decode/{vin}'
    driver.get(decode_url)
    time.sleep(5)

    body_text = driver.find_element(By.TAG_NAME, 'body').text
    lines = body_text.split('\n')

    vehicle = {'vin': vin}

    # Parse the title line like "1999 Toyota 4Runner"
    # It appears on its own line after "Download PDF"
    for line in lines:
        line = line.strip()
        title_match = re.match(r'^(\d{4})\s+(\S+)\s+(.+)$', line)
        if title_match:
            year_candidate = title_match.group(1)
            if 1900 <= int(year_candidate) <= 2030:
                vehicle['year'] = year_candidate
                vehicle['make'] = title_match.group(2)
                vehicle['model'] = title_match.group(3).strip()
                break

    # Parse "Style:4dr SR5 3.4L Auto 4WD" line
    style_match = re.search(r'Style:\s*(.+)', body_text)
    if style_match:
        vehicle['trim'] = style_match.group(1).strip()

    # Parse key:value pairs from the page text
    field_patterns = {
        'body_style': r'Body Type:\s*(.+)',
        'drive_type': r'Drive Type:\s*(.+)',
        'fuel_type': r'Fuel Type:\s*(.+)',
        'fuel_economy': r'Fuel Economy \(City/Highway/Combined\):\s*(.+)',
        'fuel_tank': r'Fuel Tank Capacity:\s*(.+)',
        'engine_cylinders': r'Engine Cylinders:\s*(\d+)',
        'exterior_color': r'Exterior Color:\s*\n?\s*(.+)',
        'interior_color': r'Interior Color:\s*\n?\s*(.+)',
    }

    for key, pattern in field_patterns.items():
        match = re.search(pattern, body_text)
        if match:
            val = match.group(1).strip()
            if val and val != 'Copy':
                vehicle[key] = val

    # Parse engine from dedicated lines: "Engine:\n3.4L V-6 DOHC..."
    for i, line in enumerate(lines):
        if line.strip() == 'Engine:' and i + 1 < len(lines):
            engine_val = lines[i + 1].strip()
            if engine_val and engine_val != 'Copy':
                vehicle['engine'] = engine_val
            break

    # Parse transmission from dedicated lines
    for i, line in enumerate(lines):
        if line.strip() == 'Transmission:' and i + 1 < len(lines):
            trans_val = lines[i + 1].strip()
            if trans_val and trans_val != 'Copy':
                vehicle['transmission'] = trans_val
            break

    # Parse axle ratio
    for i, line in enumerate(lines):
        if line.strip() == 'Axle Ratio:' and i + 1 < len(lines):
            axle_val = lines[i + 1].strip()
            if axle_val and axle_val != 'Copy':
                vehicle['axle_ratio'] = axle_val
            break

    if not vehicle.get('year'):
        raise RuntimeError(
            f'Failed to parse VIN decode results for {vin}. '
            'The page may not have loaded correctly.'
        )

    return vehicle


def decode(vin: str, config_path: str = 'config.ini') -> dict:
    """High-level helper: log in, decode a VIN, close the browser."""
    cfg = load_config(config_path)
    driver = create_driver()
    try:
        login(driver, cfg['url'], cfg['username'], cfg['password'])
        return decode_vin(driver, cfg['url'], vin)
    finally:
        driver.quit()


def decode_batch(vins: list[str], config_path: str = 'config.ini') -> list[dict]:
    """Decode multiple VINs in a single browser session."""
    cfg = load_config(config_path)
    driver = create_driver()
    results = []
    try:
        login(driver, cfg['url'], cfg['username'], cfg['password'])
        for vin in vins:
            try:
                info = decode_vin(driver, cfg['url'], vin)
                results.append(info)
            except Exception as exc:
                print(f'  ERROR decoding {vin}: {exc}')
                results.append({'vin': vin, 'error': str(exc)})
    finally:
        driver.quit()
    return results


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print('Usage: python vinmatchpro_decoder.py <VIN> [VIN2 ...]', file=sys.stderr)
        sys.exit(1)

    vins = [v.upper().strip() for v in sys.argv[1:]]

    if len(vins) == 1:
        try:
            info = decode(vins[0])
            for k, v in info.items():
                print(f'{k}: {v}')
        except Exception as exc:
            print(f'Error: {exc}', file=sys.stderr)
            sys.exit(1)
    else:
        results = decode_batch(vins)
        for info in results:
            print(f"\n--- {info.get('vin', '?')} ---")
            for k, v in info.items():
                print(f'  {k}: {v}')
