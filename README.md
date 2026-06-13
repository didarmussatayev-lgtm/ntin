# Selenium NTIN scraper

Automate SKU search and data extraction from algatop.kz NTIN system.

## Files
- `ntin_scraper.py` - main script
- `requirements.txt` - Python dependencies

## Quick start

1. Create a virtualenv and install deps:

```bash
python -m venv .venv
source .venv/bin/activate   # on Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

2. **Configure SKU list inside the script:**

Open `ntin_scraper.py` and edit the `SKU_LIST` variable (around line 6):

```python
SKU_LIST = [
    "115050528_928421346",
    "SKU_2",
    "SKU_3",
    # Add up to 1000 SKU as needed - one per line
]
```

3. Set environment variables (optional but recommended):

```bash
# Linux/Mac:
export ALGATOP_USER=didar.musataev@gmail.com
export ALGATOP_PASS='1QQQQ2qqqq3#@'

# Windows PowerShell:
$env:ALGATOP_USER="didar.musataev@gmail.com"
$env:ALGATOP_PASS="1QQQQ2qqqq3#@"
```

4. Run the script:

```bash
python ntin_scraper.py
```

## How it works

1. Opens the NTIN website and logs in automatically
2. Clicks "Перейти" button after successful login
3. Processes each SKU from `SKU_LIST` sequentially:
   - Enters SKU into search field
   - Waits 2 seconds
   - Clicks "Создать заявку" button
   - Waits 7 seconds for modal to load
   - Extracts data starting from code "1103-..."
   - Saves each result to `output_<SKU>.xlsx` in Excel format

## Usage options

Run with single SKU (overrides SKU_LIST):
```bash
python ntin_scraper.py "115050528_928421346"
```

Command-line options:
```bash
python ntin_scraper.py --help
```

Available flags:
- `--url` - override target URL (default: https://app.algatop.kz/ntin)
- `--headless` - run Chrome in headless mode (no UI)
- `--outdir` - specify output directory for Excel files
- `--username` - login email (or set ALGATOP_USER env var)
- `--password` - login password (or set ALGATOP_PASS env var)

Example with CLI credentials (less secure):
```bash
python ntin_scraper.py --username "didar.musataev@gmail.com" --password "1QQQQ2qqqq3#@"
```

## Troubleshooting

If the script can't find elements on the page:
1. Open `ntin_scraper.py` and look at the CONFIG section (lines 30-44)
2. Open the website in a browser and use DevTools (F12) to inspect elements
3. Update the XPath selectors:
   - `SEARCH_XPATH` - search input field
   - `CREATE_BUTTON_XPATH` - "Создать заявку" button
   - `GO_BUTTON_XPATH` - "Перейти" button

If login fails:
- Double-check `ALGATOP_USER` and `ALGATOP_PASS` environment variables
- Try running without `--headless` to see what's happening
