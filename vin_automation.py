"""
vin_automation.py - Main orchestrator for the VIN Automation System.

Reads a VIN from Pinnacle Professional, decodes it via VINMatchPro,
pulls parts pricing from Car-Part.com, and exports results to CSV/Excel.
"""

import argparse
import csv
import os
import sys
from datetime import datetime

import openpyxl

from pinnacle_reader import read_vin_from_pinnacle, validate_vin
from vinmatchpro_decoder import decode as decode_vin
from carpart_scraper import search as search_parts


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
    args = parser.parse_args()

    config = load_config()
    output_dir = config['output_dir']
    output_format = args.format or config['format']

    os.makedirs(output_dir, exist_ok=True)

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
