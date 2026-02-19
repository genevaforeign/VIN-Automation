"""
vin_automation.py - Main orchestrator for the VIN Automation System.

Reads a VIN from Pinnacle Professional, decodes it via VINMatchPro,
pulls parts pricing from Car-Part.com, and exports results to CSV/Excel.
"""

import argparse
import csv
import os
import re
import sys
import time
from datetime import datetime

import openpyxl
import requests

from pinnacle_reader import JABReader, read_vin_from_pinnacle, validate_vin
from vinmatchpro_decoder import decode as decode_vin
from carpart_scraper import (
    load_config as carpart_load_config,
    search as search_parts,
    search_single_part,
    _get_homepage_hidden_fields,
)
from mvr_reader import open_mvr_and_read_parts, get_open_mvr_titles


def _clean_search_term(term: str) -> str:
    """Strip Pinnacle-internal suffixes from a Hollander category name before searching.

    Removes:
    - Everything after a semicolon  ("Chassis Brain Box; on-board…" → "Chassis Brain Box")
    - ", ID XXXX…" suffixes         ("Heat/AC Controller rear, ID 4G0…" → "Heat/AC Controller rear")
    - Trailing 4+ digit numbers     ("Fuel Tank 59452" → "Fuel Tank")
    - Trailing punctuation left over after stripping
    """
    term = term.split(';')[0]
    term = re.sub(r',\s*ID\s+\S+.*$', '', term)
    term = re.sub(r'\s+\d{3,}\s*$', '', term)   # strip trailing Pinnacle part-type codes (3+ digits)
    return term.strip().rstrip(',').strip()


def load_config():
    """Load output settings from config.ini."""
    import configparser
    cfg = configparser.ConfigParser()
    cfg.read('config.ini')
    return {
        'output_dir': cfg.get('output', 'directory', fallback='output'),
        'format': cfg.get('output', 'format', fallback='csv').lower(),
    }


def export_csv(rows: list[dict], filepath: str):
    """Write results to a CSV file."""
    if not rows:
        print('No data to export.')
        return

    fieldnames = list(rows[0].keys())
    with open(filepath, 'w', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)

    print(f'CSV saved to {filepath}')


def export_excel(rows: list[dict], filepath: str):
    """Write results to an Excel (.xlsx) file."""
    if not rows:
        print('No data to export.')
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Parts Results'

    fieldnames = list(rows[0].keys())
    ws.append(fieldnames)
    for row in rows:
        ws.append([row.get(k, '') for k in fieldnames])

    wb.save(filepath)
    print(f'Excel file saved to {filepath}')


def process_vin(vin: str) -> list[dict]:
    """Decode a VIN and fetch parts, returning combined result rows."""
    print(f'\n{"="*60}')
    print(f'Processing VIN: {vin}')
    print(f'{"="*60}')

    # Step 1: Decode VIN
    print('Decoding VIN via VINMatchPro...')
    try:
        vehicle = decode_vin(vin)
        print(f"  Vehicle: {vehicle.get('year', '?')} {vehicle.get('make', '?')} {vehicle.get('model', '?')}")
    except Exception as exc:
        print(f'  ERROR decoding VIN: {exc}')
        return []

    # Step 2: Search for parts
    print('Searching Car-Part.com for parts...')
    try:
        parts = search_parts(vehicle)
        print(f'  Found {len(parts)} part listing(s).')
    except Exception as exc:
        print(f'  ERROR searching parts: {exc}')
        parts = []

    # Combine vehicle info with each part row
    rows = []
    for part in parts:
        row = {
            'vin': vin,
            'year': vehicle.get('year', ''),
            'make': vehicle.get('make', ''),
            'model': vehicle.get('model', ''),
            'trim': vehicle.get('trim', ''),
            'engine': vehicle.get('engine', ''),
        }
        row.update(part)
        rows.append(row)

    # If no parts found, still output the vehicle info
    if not rows:
        rows.append({
            'vin': vin,
            'year': vehicle.get('year', ''),
            'make': vehicle.get('make', ''),
            'model': vehicle.get('model', ''),
            'trim': vehicle.get('trim', ''),
            'engine': vehicle.get('engine', ''),
            'part_name': '',
            'price': '',
            'vendor': '',
            'location': '',
            'grade': '',
        })

    return rows


def main():
    parser = argparse.ArgumentParser(
        description='VIN Automation System for Fitz Auto Parts'
    )
    parser.add_argument(
        '--vin',
        type=str,
        help='Process a specific VIN instead of reading from Pinnacle',
    )
    parser.add_argument(
        '--vins',
        type=str,
        nargs='+',
        help='Process multiple VINs',
    )
    parser.add_argument(
        '--format',
        choices=['csv', 'excel'],
        help='Override output format (csv or excel)',
    )
    parser.add_argument(
        '--unpriced',
        action='store_true',
        help=(
            'Open the selected vehicle\'s MVR, read un-priced parts with Hollander '
            'interchange numbers, look up market prices on Car-Part.com, and export '
            'a suggested-price report.'
        ),
    )
    args = parser.parse_args()

    config = load_config()
    output_dir = config['output_dir']
    output_format = args.format or config['format']

    os.makedirs(output_dir, exist_ok=True)

    # --unpriced workflow: open MVR, read Parts tab, price each un-priced part
    if args.unpriced:
        print('Reading selected VIN from Pinnacle...')
        try:
            reader = JABReader()
            vin = reader.read_selected_vin()
            print(f'  VIN: {vin}')
        except RuntimeError as exc:
            print(f'ERROR: {exc}', file=sys.stderr)
            sys.exit(1)

        print('Decoding VIN...')
        try:
            vehicle = decode_vin(vin)
            print(f"  Vehicle: {vehicle.get('year', '?')} {vehicle.get('make', '?')} {vehicle.get('model', '?')}")
        except Exception as exc:
            print(f'ERROR decoding VIN: {exc}', file=sys.stderr)
            sys.exit(1)

        print('Opening MVR (double-clicking selected row)...')
        pre_click_titles = get_open_mvr_titles()
        try:
            reader.open_selected_vehicle()
        except RuntimeError as exc:
            print(f'ERROR: {exc}', file=sys.stderr)
            sys.exit(1)

        print('Waiting for MVR to load...')
        time.sleep(3)

        print('Reading un-priced parts from MVR Parts tab...')
        try:
            parts = open_mvr_and_read_parts(pre_click_titles=pre_click_titles)
            print(f'  Found {len(parts)} un-priced part(s) with Hollander numbers.')
        except RuntimeError as exc:
            print(f'ERROR reading MVR parts: {exc}', file=sys.stderr)
            sys.exit(1)

        if not parts:
            print('No un-priced parts with Hollander numbers found. Nothing to export.')
            sys.exit(0)

        print('Looking up market prices on Car-Part.com...')
        carpart_cfg = carpart_load_config()
        zip_code = carpart_cfg['zip_code']

        session = requests.Session()
        from carpart_scraper import HEADERS
        session.headers.update(HEADERS)
        hidden = _get_homepage_hidden_fields(session)

        rows = []
        for part in parts:
            # Always prefer the Hollander category name (col 0) — it is the standardised
            # interchange description that maps directly to Car-Part.com search terms.
            # Fall back to the extracted part name only if the category is absent.
            raw_term = part.get('category', '') or part['part_name']
            search_term = _clean_search_term(raw_term) if raw_term else ''
            print(f"  Searching: {search_term or part['description'][:60]}")

            if search_term:
                result = search_single_part(
                    search_term,
                    vehicle.get('year', ''),
                    vehicle.get('make', ''),
                    vehicle.get('model', ''),
                    zip_code,
                    session=session,
                    hidden_fields=hidden,
                )
            else:
                result = {'avg_price': None, 'low_price': None, 'listing_count': 0}

            if result['listing_count']:
                print(
                    f"    avg=${result['avg_price']:.2f}  "
                    f"low=${result['low_price']:.2f}  "
                    f"({result['listing_count']} listings)"
                )
            else:
                print('    No listings found.')

            if not search_term:
                notes = 'Description unclear – manual review'
            elif result['listing_count'] == 0:
                notes = 'No Car-Part listings'
            else:
                notes = ''

            rows.append({
                'vin': vin,
                'year': vehicle.get('year', ''),
                'make': vehicle.get('make', ''),
                'model': vehicle.get('model', ''),
                'stock_num': part['stock_num'],
                'hollander': part['hollander'],
                'part_name': search_term or part['description'][:60],
                'grade': part['grade'],
                'location': part['location'],
                'avg_price': f"${result['avg_price']:.2f}" if result['avg_price'] is not None else '',
                'low_price': f"${result['low_price']:.2f}" if result['low_price'] is not None else '',
                'listing_count': result['listing_count'],
                'notes': notes,
            })

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filepath = os.path.join(output_dir, f'unpriced_{timestamp}.csv')
        export_csv(rows, filepath)
        print(f'\nDone. {len(rows)} part(s) written to {filepath}')
        return

    # Determine which VINs to process
    vins = []

    if args.vins:
        vins = [v.upper().strip() for v in args.vins]
    elif args.vin:
        vins = [args.vin.upper().strip()]
    else:
        # Read from Pinnacle Professional
        print('Reading VIN from Pinnacle Professional...')
        try:
            vin = read_vin_from_pinnacle()
            print(f'  Found VIN: {vin}')
            vins = [vin]
        except RuntimeError as exc:
            print(f'ERROR: {exc}', file=sys.stderr)
            sys.exit(1)

    # Validate all VINs
    for vin in vins:
        if not validate_vin(vin):
            print(f'WARNING: "{vin}" does not look like a valid VIN. Skipping.')
            vins.remove(vin)

    if not vins:
        print('No valid VINs to process.')
        sys.exit(1)

    # Process each VIN
    all_rows = []
    for vin in vins:
        rows = process_vin(vin)
        all_rows.extend(rows)

    # Export results
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    if output_format == 'excel':
        filepath = os.path.join(output_dir, f'parts_{timestamp}.xlsx')
        export_excel(all_rows, filepath)
    else:
        filepath = os.path.join(output_dir, f'parts_{timestamp}.csv')
        export_csv(all_rows, filepath)

    print(f'\nDone. Processed {len(vins)} VIN(s), {len(all_rows)} total row(s).')


if __name__ == '__main__':
    main()
