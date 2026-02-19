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

# Make aliases: VINMatchPro name → Car-Part.com name
_MAKE_ALIASES = {
    'Mercedes-Benz': 'Mercedes',
}

# Mercedes models that don't follow the "{prefix} Class" pattern on Car-Part.com
_MERCEDES_MODEL_OVERRIDES = {
    'ML':     'ML Series',
    'CLK':    'CLK',
    'CLS':    'CLS',
    'SLK':    'SLK',
    'SLR':    'SLR',
    'SLS':    'SLS',
    'AMG GT': 'AMG GT',
    'Metris': 'Metris',
}


def _normalize_for_carpart(make: str, model: str) -> str:
    """Return the Car-Part.com userModel string for a given VINMatchPro make/model."""
    norm_make = _MAKE_ALIASES.get(make, make)

    if make == 'Mercedes-Benz':
        # Model is like "A 220", "GLE 450", "AMG GT 63" — extract letter prefix
        m = re.match(r'^([A-Za-z]+(?:\s+[A-Za-z]+)?)', model)
        if m:
            prefix = m.group(1).strip()
            suffix = _MERCEDES_MODEL_OVERRIDES.get(prefix, f'{prefix} Class')
            return f'{norm_make} {suffix}'

    return f'{norm_make} {model}'


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


def _handle_interchange(
    session: requests.Session, soup: BeautifulSoup, original_data: dict = None
) -> BeautifulSoup:
    """If the response is an interchange selection page, pick the first option and submit."""
    radios = soup.find_all('input', {'type': 'radio'})
    if not radios:
        return soup  # Already on results page

    form = soup.find('form')
    if not form:
        return soup

    # Start with original data so fields like userDate aren't lost
    data = dict(original_data) if original_data else {}

    for inp in form.find_all('input', {'type': 'hidden'}):
        name = inp.get('name', '')
        if name:
            data[name] = inp.get('value', '')

    # Select the first non-None radio value using the actual field name
    for r in radios:
        val = r.get('value', '')
        name = r.get('name', 'userInterchange')
        if val and val != 'None':
            data[name] = val
            break

    resp = session.post(SEARCH_URL, data=data, timeout=30)
    resp.raise_for_status()
    return BeautifulSoup(resp.text, 'html.parser')


def _parse_results_table(soup: BeautifulSoup, search_part: str) -> list[dict]:
    """Parse the results table from a Car-Part.com search results page.

    Car-Part.com uses either a 7-column or 8-column layout depending on the part type.
    7-col: YearPartModel | Description | PartGrade | Stock# | USPrice | Dealer Info | Distmile
    8-col: YearPartModel | Description | DamageCode | PartGrade | Stock# | USPrice | Dealer Info | Distmile

    Column positions are determined dynamically by reading the header row so that
    layout changes don't break parsing.
    """
    results = []

    # Canonical header names -> listing dict key
    _HEADER_MAP = {
        'description': 'description',
        'partgrade':   'grade',
        'stock#':      'stock_num',
        'usprice':     'price',
        'dealer info': 'dealer',
        'dealerinfo':  'dealer',
        'distmile':    'distance_miles',
    }

    for table in soup.find_all('table'):
        rows = table.find_all('tr')
        if len(rows) < 2:
            continue

        # Identify the header row
        first_cells = rows[0].find_all(['td', 'th'])
        header_texts = [c.get_text(strip=True) for c in first_cells]
        header_joined = ''.join(header_texts).lower()
        if 'yearpartmodel' not in header_joined and 'usprice' not in header_joined:
            continue

        # Build col-index lookup from header
        col_idx = {}
        for i, h in enumerate(header_texts):
            key = h.lower().strip()
            if key in _HEADER_MAP:
                col_idx[_HEADER_MAP[key]] = i

        # Require at minimum price and dealer columns
        if 'price' not in col_idx or 'dealer' not in col_idx:
            continue

        # Parse data rows (skip header row)
        for row in rows[1:]:
            cells = row.find_all('td')
            if len(cells) < len(header_texts):
                continue

            cell_texts = [c.get_text(strip=True) for c in cells]

            def _col(key, default=''):
                idx = col_idx.get(key)
                return cell_texts[idx] if idx is not None and idx < len(cell_texts) else default

            # Column 0: "2017Fuel TankAudi A6" -> parse year
            year_match = re.match(r'^(\d{4})', cell_texts[0])
            year = year_match.group(1) if year_match else ''

            # Description
            description = re.sub(
                r'Estimated CO2e Savings:\s*\d+kg', '', _col('description')
            ).strip()

            # Grade and Stock#
            grade = _col('grade')
            stock_num = _col('stock_num')

            # US Price — "$125actual" or "$Call"
            price_text = _col('price')
            price_match = re.search(r'(\$[\d,.]+)', price_text)
            price = price_match.group(1) if price_match else price_text

            # Dealer Info — extract yard name, location, phone
            dealer_text = _col('dealer')
            yard_match = re.match(r'^(.+?)(?:USA-|Can-)', dealer_text)
            vendor = yard_match.group(1).strip() if yard_match else dealer_text[:60]
            for suffix in [' - PRP Freight, ARA, CDC', ' - PRP Freight, CDC', ' - CDC', ' - ARA']:
                vendor = vendor.replace(suffix, '')

            loc_match = re.search(r'(?:USA|Can)-(\w+)\(([^)]+)\)', dealer_text)
            location = f'{loc_match.group(2)}, {loc_match.group(1)}' if loc_match else ''

            phone_match = re.search(r'(\d[\d\-]{9,})', dealer_text)
            phone = phone_match.group(1) if phone_match else ''

            distance = _col('distance_miles')

            listing = {
                'part_name': search_part,
                'year': year,
                'description': description,
                'grade': grade,
                'stock_num': stock_num,
                'price': price,
                'vendor': vendor,
                'location': location,
                'phone': phone,
                'distance_miles': distance,
            }
            results.append(listing)

        break  # Found the results table, stop looking

    return results


def search_single_part(
    part_name: str,
    year: str,
    make: str,
    model: str,
    zip_code: str,
    session: requests.Session = None,
    hidden_fields: dict = None,
) -> dict:
    """Search Car-Part.com for a single part and return deduplicated price stats.

    Deduplication is performed by (vendor, stock#): when the same yard lists the
    same stock number more than once, only the first occurrence is counted so that
    duplicate listings from one yard do not skew the average.

    Args:
        part_name:     The part description to search for (e.g. "Hood").
        year:          Vehicle model year (e.g. "2015").
        make:          Vehicle make (e.g. "Toyota").
        model:         Vehicle model (e.g. "Camry").
        zip_code:      Buyer zip code for proximity search.
        session:       Optional existing requests.Session to reuse.
        hidden_fields: Optional hidden form fields already fetched from the homepage.

    Returns:
        {
            'part_name':      str,
            'avg_price':      float | None,
            'low_price':      float | None,
            'listing_count':  int,
        }
    """
    if session is None:
        session = requests.Session()
        session.headers.update(HEADERS)

    if hidden_fields is None:
        hidden_fields = _get_homepage_hidden_fields(session)

    model_value = _normalize_for_carpart(make, model)

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
        return {'part_name': part_name, 'avg_price': None, 'low_price': None, 'listing_count': 0}

    soup = BeautifulSoup(resp.text, 'html.parser')

    if 'INVALID' in soup.get_text():
        return {'part_name': part_name, 'avg_price': None, 'low_price': None, 'listing_count': 0}

    soup = _handle_interchange(session, soup, data)
    listings = _parse_results_table(soup, part_name)

    # Deduplicate by (vendor, stock#)
    seen = set()
    unique_listings = []
    for listing in listings:
        key = (listing.get('vendor', ''), listing.get('stock_num', ''))
        if key not in seen:
            seen.add(key)
            unique_listings.append(listing)

    # Extract numeric prices
    prices = []
    for listing in unique_listings:
        price_text = listing.get('price', '')
        price_match = re.search(r'([\d,.]+)', price_text.replace('$', ''))
        if price_match:
            try:
                prices.append(float(price_match.group(1).replace(',', '')))
            except ValueError:
                pass

    avg_price = (sum(prices) / len(prices)) if prices else None
    low_price = min(prices) if prices else None

    return {
        'part_name': part_name,
        'avg_price': avg_price,
        'low_price': low_price,
        'listing_count': len(unique_listings),
    }


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

    model_value = _normalize_for_carpart(make, model)
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
        soup = _handle_interchange(session, soup, data)

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
