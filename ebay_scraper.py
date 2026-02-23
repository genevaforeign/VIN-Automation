"""
ebay_scraper.py - Scrapes eBay completed/sold listings for OEM auto parts pricing.

Uses a single shared Selenium Chrome session for all searches to avoid the
overhead of launching a new browser per query.  Call `create_driver()` once,
pass the driver to `search_sold()` for each part, then call `close_driver()`.
"""

import re
import time
import urllib.parse

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


# How many pages of eBay results to scan per part (each page ≈ 60 items)
_MAX_PAGES = 2

_HEADERS = (
    'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
    'AppleWebKit/537.36 (KHTML, like Gecko) '
    'Chrome/120.0.0.0 Safari/537.36'
)


def create_driver() -> webdriver.Chrome:
    """Launch a headless Chrome instance for eBay scraping."""
    opts = Options()
    opts.add_argument('--headless')
    opts.add_argument('--no-sandbox')
    opts.add_argument('--disable-dev-shm-usage')
    opts.add_argument('--disable-blink-features=AutomationControlled')
    opts.add_experimental_option('excludeSwitches', ['enable-automation'])
    opts.add_argument(f'user-agent={_HEADERS}')
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()), options=opts
    )
    return driver


def close_driver(driver: webdriver.Chrome):
    """Quit the Chrome driver."""
    try:
        driver.quit()
    except Exception:
        pass


def _build_url(query: str, page: int = 1) -> str:
    """Build an eBay completed/sold listings search URL."""
    params = {
        '_nkw':         query,
        'LH_Sold':      '1',
        'LH_Complete':  '1',
        '_sacat':       '0',
        '_pgn':         str(page),
    }
    return 'https://www.ebay.com/sch/i.html?' + urllib.parse.urlencode(params)


def _parse_prices(html: str) -> list[float]:
    """Extract sold prices from an eBay search results page."""
    soup = BeautifulSoup(html, 'html.parser')
    prices = []

    # Each sold listing has aria-label="Sold Item" on its date span.
    # The price lives in span.s-card__price within the same card container.
    sold_spans = soup.find_all('span', attrs={'aria-label': 'Sold Item'})
    for span in sold_spans:
        # Walk up to the card container (su-card-container__content or similar)
        card = span
        price_span = None
        for _ in range(15):
            card = card.parent
            if card is None:
                break
            price_span = card.find('span', class_='s-card__price')
            if price_span:
                break

        if not price_span:
            continue

        price_text = price_span.get_text(strip=True)
        # Handle ranges like "$50.00 to $75.00" — take the lower value
        nums = re.findall(r'[\d,]+\.?\d*', price_text)
        if nums:
            try:
                prices.append(float(nums[0].replace(',', '')))
            except ValueError:
                pass

    return prices


def search_sold(
    part_name: str,
    year: str,
    make: str,
    model: str,
    driver: webdriver.Chrome,
    oem_only: bool = True,
) -> dict:
    """Search eBay sold listings for a part and return price stats.

    Args:
        part_name:  Part description (e.g. "Hood").
        year:       Vehicle year (e.g. "2014").
        make:       Vehicle make (e.g. "Ford").
        model:      Vehicle model (e.g. "Escape").
        driver:     Shared Selenium Chrome driver.
        oem_only:   If True, appends "OEM" to the query to filter out
                    aftermarket listings.

    Returns:
        {
            'avg_price':     float | None,
            'low_price':     float | None,
            'listing_count': int,
        }
    """
    # Build query: "2014 Ford Escape Hood OEM"
    query_parts = [year, make, model, part_name]
    if oem_only:
        query_parts.append('OEM')
    query = ' '.join(p for p in query_parts if p)

    all_prices = []

    for page in range(1, _MAX_PAGES + 1):
        url = _build_url(query, page)
        try:
            driver.get(url)
            time.sleep(2.5)
        except Exception:
            break

        prices = _parse_prices(driver.page_source)
        if not prices:
            break   # No results on this page — stop paging

        all_prices.extend(prices)

    if not all_prices:
        return {'avg_price': None, 'low_price': None, 'listing_count': 0}

    return {
        'avg_price':     round(sum(all_prices) / len(all_prices), 2),
        'low_price':     round(min(all_prices), 2),
        'listing_count': len(all_prices),
    }
