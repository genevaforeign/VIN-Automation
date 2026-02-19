# VIN Automation

Automates used-vehicle parts pricing for Fitz Auto Parts. Reads the selected row from Pinnacle Professional's Find Vehicles table, decodes the VIN via VINMatchPro, searches Car-Part.com for part listings, and exports results to CSV or Excel.

## How it works

```
Pinnacle Professional (JAB) → VINMatchPro decoder → Car-Part.com scraper → CSV/Excel
```

1. **pinnacle_reader.py** — Connects to Pinnacle Professional via Java Access Bridge, finds the Find Vehicles table, and reads the VIN from the currently highlighted row.
2. **vinmatchpro_decoder.py** — Logs into VINMatchPro and decodes the VIN into year/make/model/trim/engine.
3. **carpart_scraper.py** — Searches Car-Part.com for matching part listings.
4. **vin_automation.py** — Orchestrates the full pipeline and writes the output file.

## Requirements

- Windows 10/11 (64-bit)
- Java Access Bridge enabled (`jabswitch -enable` in an elevated prompt)
- Pinnacle Professional running with the Find Vehicles screen open
- Python 3.10+

```
pip install -r requirements.txt
```

## Configuration

Copy `config.example.ini` to `config.ini` and fill in your credentials:

```ini
[vinmatchpro]
url = https://www.vinmatchpro.com
username = your_username
password = your_password

[carpart]
zip_code = 12345
search_radius = 200

[output]
directory = output
format = csv        # or excel
```

## Usage

**Read selected row from Pinnacle and run the full pipeline:**
```
python vin_automation.py
```

**Process a specific VIN:**
```
python vin_automation.py --vin 1HGCM82633A123456
```

**Process multiple VINs:**
```
python vin_automation.py --vins 1HGCM82633A123456 WDD3G4FB4KW005653
```

**Override output format:**
```
python vin_automation.py --format excel
```

**Test Pinnacle reader standalone:**
```
python pinnacle_reader.py           # print selected VIN
python pinnacle_reader.py --dump    # list first 10 VINs in table
python pinnacle_reader.py --all     # list all VINs in table
```

## Output

Results are saved to the `output/` directory as `parts_YYYYMMDD_HHMMSS.csv` (or `.xlsx`). Each row contains:

| Column | Description |
|--------|-------------|
| vin | Full 17-character VIN |
| year / make / model / trim / engine | Decoded vehicle info |
| part_name | Part description from Car-Part.com |
| price | Listed price |
| vendor | Selling yard name |
| location | Yard location |
| grade | Part grade |
