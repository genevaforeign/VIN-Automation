"""
carpart_scraper.py - Scrapes Car-Part.com for parts pricing.

Submits a search on Car-Part.com using decoded vehicle info (year, make,
model) along with a zip code, then parses the resulting parts listings.

Car-Part.com uses a two-step flow:
  1. POST initial search -> may return an interchange selection page
  2. Select the first interchange option and POST again -> results page
"""

import configparser
import re
import sys

import requests
from bs4 import BeautifulSoup


BASE_URL = 'https://www.car-part.com'
SEARCH_URL = f'{BASE_URL}/cgi-bin/search.cgi'

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/120.0.0.0 Safari/537.36'
    ),
}

# Common parts to search for
DEFAULT_PARTS = [
    'Engine',
    'Transmission',
    'Hood',
    'Fender',
    'Headlight Assembly',
    'Tail Light',
    'Radiator',
    'Front Bumper Assembly (includes cover)',
    'Rear Bumper Assembly (includes cover)',
    'Front Door (see also Door Shell, Front)',
    'A/C Compressor',
]


def load_config(config_path: str = 'config.ini') -> dict:
    """Read Car-Part search settings from config file."""
    cfg = configparser.ConfigParser()
    cfg.read(config_path)
    return {
        'zip_code': cfg.get('carpart', 'zip_code'),
        'search_radius': cfg.get('carpart', 'search_radius', fallback='200'),
    }


def _get_homepage_hidden_fields(session: requests.Session) -> dict:
    """Fetch Car-Part.com homepage and extract hidden form fields."""
    resp = session.get(BASE_URL, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, 'html.parser')
    form = soup.find('form')
    fields = {}
    if form:
        for inp in form.find_all('input', {'type': 'hidden'}):
            name = inp.get('name', '')
            if name:
                fields[name] = inp.get('value', '')
    return fields


def _handle_interchange(session: requests.Session, soup: BeautifulSoup) -> BeautifulSoup:
    """If the response is an interchange selection page, pick the first option and submit."""
    radios = soup.find_all('input', {'type': 'radio'})
    if not radios:
        return soup  # Already on results page

    form = soup.find('form')
    if not form:
        return soup

    data = {}
    for inp in form.find_all('input', {'type': 'hidden'}):
        name = inp.get('name', '')
        if name:
            data[name] = inp.get('value', '')

    # Select the first non-None radio value
    for r in radios:
        val = r.get('value', '')
        if val and val != 'None':
            data['userInterchange'] = val
            break

    resp = session.post(SEARCH_URL, data=data, timeout=30)
    resp.raise_for_status()
    return BeautifulSoup(resp.text, 'html.parser')


def _parse_results_table(soup: BeautifulSoup, search_part: str) -> list[dict]:
    """Parse the results table from a Car-Part.com search results page.

    The results table has 8 columns:
      [0] Year/Part/Model, [1] Description, [2] Damage Code,
      [3] Part Grade, [4] Stock#, [5] US Price,
      [6] Dealer Info, [7] Distance (miles)
    """
    results = []

    # Find the table that has the header row with "YearPartModel"
    for table in soup.find_all('table'):
        rows = table.find_all('tr')
        if len(rows) < 2:
            continue

        # Check if first row is the header
        first_cells = rows[0].find_all(['td', 'th'])
        header_text = ''.join(c.get_text(strip=True) for c in first_cells).lower()
        if 'yearpartmodel' not in header_text and 'usprice' not in header_text:
            continue

        # Parse data rows (skip header row)
        for row in rows[1:]:
            cells = row.find_all('td')
            if len(cells) < 6:
                continue

            cell_texts = [c.get_text(strip=True) for c in cells]

            # Column 0: "2000HoodToyota 4Runner" -> parse year, part, model
            col0 = cell_texts[0]
            year_match = re.match(r'^(\d{4})', col0)
            year = year_match.group(1) if year_match else ''

            # Column 5: "$250actual" or "$Call"
            price_text = cell_texts[5] if len(cell_texts) > 5 else ''
            price_match = re.search(r'(\$[\d,.]+)', price_text)
            price = price_match.group(1) if price_match else price_text

            # Column 6: Dealer info - extract yard name and phone
            dealer_text = cell_texts[6] if len(cell_texts) > 6 else ''
            # Try to extract yard name (text before "USA-" or "Can-")
            yard_match = re.match(r'^(.+?)(?:USA-|Can-)', dealer_text)
            vendor = yard_match.group(1).strip() if yard_match else dealer_text[:60]
            # Clean up vendor name - remove trailing suffixes
            for suffix in [' - PRP Freight, ARA, CDC', ' - PRP Freight, CDC', ' - CDC', ' - ARA']:
                vendor = vendor.replace(suffix, '')

            # Extract location
            loc_match = re.search(r'(?:USA|Can)-(\w+)\(([^)]+)\)', dealer_text)
            location = f'{loc_match.group(2)}, {loc_match.group(1)}' if loc_match else ''

            # Extract phone
            phone_match = re.search(r'(\d[\d\-]{9,})', dealer_text)
            phone = phone_match.group(1) if phone_match else ''

            # Column 7: Distance
            distance = cell_texts[7] if len(cell_texts) > 7 else ''

            # Column 1: Description
            description = cell_texts[1] if len(cell_texts) > 1 else ''
            # Remove CO2e savings text
            description = re.sub(r'Estimated CO2e Savings:\s*\d+kg', '', description).strip()

            # Column 3: Grade
            grade = cell_texts[3] if len(cell_texts) > 3 else ''

            listing = {
                'part_name': search_part,
                'year': year,
                'description': description,
                'grade': grade,
                'price': price,
                'vendor': vendor,
                'location': location,
                'phone': phone,
                'distance_miles': distance,
            }
            results.append(listing)

        break  # Found the results table, stop looking

    return results


def search_parts(
    year: str,
    make: str,
    model: str,
    zip_code: str,
    parts: list[str] = None,
) -> list[dict]:
    """Query Car-Part.com for multiple part types and return all listings."""
    if parts is None:
        parts = DEFAULT_PARTS

    session = requests.Session()
    session.headers.update(HEADERS)

    # Get homepage hidden fields (needed for valid session)
    hidden_fields = _get_homepage_hidden_fields(session)

    model_value = f'{make} {model}'
    all_results = []

    for part_name in parts:
        data = {
            'userDate': str(year),
            'userModel': model_value,
            'userPart': part_name,
            'userLocation': 'All States',
            'userPreference': 'zip',
            'userZip': zip_code,
        }
        data.update(hidden_fields)

        try:
            resp = session.post(SEARCH_URL, data=data, timeout=30)
            resp.raise_for_status()
        except Exception as exc:
            print(f'    Warning: search failed for "{part_name}": {exc}')
            continue

        soup = BeautifulSoup(resp.text, 'html.parser')

        # Check for "INVALID" messages
        page_text = soup.get_text()
        if 'INVALID' in page_text:
            continue

        # Handle interchange selection if needed
        soup = _handle_interchange(session, soup)

        listings = _parse_results_table(soup, part_name)
        all_results.extend(listings)

    return all_results


def search(vehicle: dict, config_path: str = 'config.ini') -> list[dict]:
    """High-level helper: read config and search for parts."""
    cfg = load_config(config_path)
    return search_parts(
        year=vehicle.get('year', ''),
        make=vehicle.get('make', ''),
        model=vehicle.get('model', ''),
        zip_code=cfg['zip_code'],
    )


if __name__ == '__main__':
    if len(sys.argv) < 4:
        print(
            'Usage: python carpart_scraper.py <year> <make> <model>',
            file=sys.stderr,
        )
        sys.exit(1)

    year_arg, make_arg, model_arg = sys.argv[1], sys.argv[2], sys.argv[3]
    cfg = load_config()
    results = search_parts(
        year=year_arg,
        make=make_arg,
        model=model_arg,
        zip_code=cfg['zip_code'],
    )

    if not results:
        print('No parts found.')
    else:
        print(f'Found {len(results)} listing(s):')
        for p in results:
            print(
                f"  {p['part_name']:<40}  {p['price']:<12}  "
                f"{p['vendor']:<35}  {p['location']:<20}  {p['distance_miles']}mi"
            )
