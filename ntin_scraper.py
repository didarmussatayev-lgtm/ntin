"""ntin_scraper.py

Usage:
  python ntin_scraper.py [--outdir=path] [--input=ACTIVE.xlsx]

This script opens Chrome, lets the operator manually prepare the session,
then processes SKU values from an Excel file, extracts modal fields, and
periodically saves progress so the run can be resumed after failure.
"""
import time
import re
import sys
import os
import logging
from pathlib import Path
import argparse

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, InvalidSessionIdException, WebDriverException

from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

try:
    import winsound
except ImportError:
    winsound = None

URL = "https://app.algatop.kz/ntin"
INPUT_FILE_NAME = "ACTIVE.xlsx"
INPUT_SHEET_NAME = "offer"
INPUT_SKU_COLUMN = "sku"
OUTPUT_FILE_NAME = "output.xlsx"
PROGRESS_FILE_NAME = "progress.xlsx"
LOG_FILE_NAME = "scraper.log"
CHECKPOINT_EVERY = 100
OUTPUT_DIR = Path.cwd()
WAIT_TIMEOUT = 20
CHROME_DEBUG_PORT = 9222
SEARCH_XPATH = "//*[@id=':r5:']"
SEARCH_CANDIDATES = [
    "//*[@id=':r5:']",
    "//input[contains(@placeholder, 'Поиск')]",
    "//input[contains(@placeholder, 'назв') or contains(@placeholder, 'артикул')]",
    "//input[@type='search']",
    "//input[contains(@class, 'search') or contains(@class, 'Search')]",
    "//input[@role='combobox']",
]
CREATE_BUTTON_XPATH = "//*[@id='root']/div/div[3]/div/div[2]/div[2]/div[4]/div[2]/table/tbody/tr[1]/td[5]/div/span/button"
GO_BUTTON_XPATH = "//button[contains(., 'Перейти')]"
MODAL_FALLBACK_XPATH = "//div[@role='dialog'] | //div[contains(@class,'modal')] | //div[contains(@class,'MuiDialog-root')]"
TABLE_COLUMNS = [
    "SKU",
    "NTIN_CODE",
    "Полное наименование товара (рус)",
    "Полное наименование товара (каз)",
    "Краткое наименование товара (рус)",
    "Страна происхождения",
    "Единица измерения",
    "Количественное значение",
    "ТНВЭД ЕАЭС",
    "Наименование производителя",
    "Категория ОКТРУ (НКТ)",
    "Подобрано AI",
    "Расширенная форма заявки",
    "Raw Text",
    "Status",
    "Error",
]
REQUIRED_FIELDS = [
    "Полное наименование товара (рус)",
    "Полное наименование товара (каз)",
    "Краткое наименование товара (рус)",
    "Страна происхождения",
    "Единица измерения",
    "Количественное значение",
    "ТНВЭД ЕАЭС",
    "Наименование производителя",
]


def setup_logging(outdir=OUTPUT_DIR):
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    log_path = outdir / LOG_FILE_NAME
    logger = logging.getLogger("ntin_scraper")
    logger.setLevel(logging.INFO)
    logger.handlers.clear()
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    return logger


def alert_operator(message, logger=None):
    if logger:
        logger.warning(message)
    else:
        print(message)
    if winsound:
        try:
            for _ in range(3):
                winsound.Beep(1400, 500)
                time.sleep(0.2)
        except Exception:
            pass
    else:
        try:
            print("\a", end="", flush=True)
        except Exception:
            pass


def get_driver(headless=False):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1400,1000")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)


def load_skus_from_excel(input_file, sku_column=INPUT_SKU_COLUMN, sheet_name=INPUT_SHEET_NAME):
    df = pd.read_excel(input_file, sheet_name=sheet_name)
    normalized_columns = {str(col).strip().lower(): col for col in df.columns}
    target_col = normalized_columns.get(sku_column.strip().lower())
    if not target_col:
        raise ValueError(f"Column '{sku_column}' not found in {input_file}. Found columns: {list(df.columns)}")
    series = df[target_col].dropna().astype(str).map(str.strip)
    series = series[series != ""]
    return series.tolist()


def load_existing_progress(progress_file):
    progress_path = Path(progress_file)
    if not progress_path.exists():
        return pd.DataFrame(columns=TABLE_COLUMNS)
    try:
        df = pd.read_excel(progress_path)
        for col in TABLE_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df[TABLE_COLUMNS]
    except Exception:
        return pd.DataFrame(columns=TABLE_COLUMNS)


def get_processed_skus(progress_df, only_success=True):
    if progress_df.empty or "SKU" not in progress_df.columns:
        return set()
    working = progress_df.copy()
    if only_success and "Status" in working.columns:
        working = working[working["Status"].astype(str) == "success"]
    return set(working["SKU"].dropna().astype(str).str.strip())


def save_checkpoint(rows, outdir=OUTPUT_DIR, filename=PROGRESS_FILE_NAME):
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    out_file = outdir / filename
    pd.DataFrame(rows, columns=TABLE_COLUMNS).to_excel(out_file, index=False)
    return out_file


def save_results_to_excel(rows, outdir=OUTPUT_DIR, filename=OUTPUT_FILE_NAME):
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    out_file = outdir / filename
    pd.DataFrame(rows, columns=TABLE_COLUMNS).to_excel(out_file, index=False)
    return out_file


def wait_for_manual_ready(driver, logger):
    alert_operator(
        "Chrome is open. Manually open the target site, log in, pass captcha, reach the working page with the search field, then press Enter here.",
        logger,
    )
    input("After you manually prepare the browser and see the search field, press Enter to continue...")
    return True


def clear_search_input(driver):
    try:
        el = find_search_input(driver)
    except Exception:
        return True
    for action in [
        lambda: el.clear(),
        lambda: (el.send_keys(Keys.CONTROL, 'a'), el.send_keys(Keys.DELETE)),
        lambda: driver.execute_script("arguments[0].value = ''; arguments[0].dispatchEvent(new Event('input'));", el),
    ]:
        try:
            action()
            time.sleep(0.15)
        except Exception:
            pass
    try:
        return (el.get_attribute('value') or '').strip() == ''
    except Exception:
        return False


def find_search_input(driver):
    for xpath in SEARCH_CANDIDATES:
        try:
            el = WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.XPATH, xpath)))
            if el.is_displayed() and el.is_enabled():
                return el
        except Exception:
            continue
    return WebDriverWait(driver, WAIT_TIMEOUT).until(EC.presence_of_element_located((By.XPATH, SEARCH_XPATH)))


def click_create_button(driver):
    candidates = [
        CREATE_BUTTON_XPATH,
        "//button[contains(., 'Создать заявку') or contains(., 'Создать')]",
        "//span[contains(., 'Создать заявку')]/ancestor::button",
    ]
    last_exc = None
    for xp in candidates:
        try:
            btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xp)))
            btn.click()
            return True
        except Exception as e:
            last_exc = e
            continue
    raise last_exc


def get_modal_container(driver):
    try:
        return WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, MODAL_FALLBACK_XPATH)))
    except Exception:
        return None


def extract_ntin_code(modal):
    if modal is None:
        return ""
    patterns = [r'(\d{4}-\d{4}-\d{4}-\d+)', r'(1103[-\d]+)', r'(1070[-\d]+)', r'(8504[-\d]+)']
    try:
        text = modal.text or ""
        for pattern in patterns:
            m = re.search(pattern, text)
            if m:
                return m.group(1)
    except Exception:
        pass
    try:
        elems = modal.find_elements(By.XPATH, ".//*[self::span or self::div or self::p or self::label]")
        for el in elems:
            txt = (el.text or "").strip()
            if not txt:
                continue
            for pattern in patterns:
                m = re.search(pattern, txt)
                if m:
                    return m.group(1)
    except Exception:
        pass
    return ""


def get_field_value_from_modal(modal, labels):
    if modal is None:
        return ""
    for label in labels:
        label_xpaths = [
            f".//label[normalize-space()='{label}']",
            f".//label[contains(normalize-space(), '{label}')]",
            f".//*[self::span or self::div or self::p][normalize-space()='{label}']",
            f".//*[self::span or self::div or self::p][contains(normalize-space(), '{label}')]",
        ]
        for label_xpath in label_xpaths:
            try:
                label_elements = modal.find_elements(By.XPATH, label_xpath)
            except Exception:
                continue
            for label_el in label_elements:
                try:
                    label_for = label_el.get_attribute("for")
                    if label_for:
                        linked = modal.find_element(By.XPATH, f".//*[@id='{label_for}']")
                        tag = linked.tag_name.lower()
                        role = (linked.get_attribute("role") or "").lower()
                        if tag in ["input", "textarea"]:
                            value = linked.get_attribute("value")
                            if value and value.strip():
                                return value.strip()
                        if tag == "div" and role == "combobox":
                            text = (linked.text or "").strip()
                            if text:
                                return text
                        if tag == "input" and role == "combobox":
                            value = linked.get_attribute("value")
                            if value and value.strip():
                                return value.strip()
                        text = (linked.text or "").strip()
                        if text:
                            return text
                except Exception:
                    pass
    return ""


def extract_fields_from_modal(driver):
    modal = get_modal_container(driver)
    text = ""
    if modal:
        try:
            text = modal.text
        except Exception:
            text = ""
    row = {col: "" for col in TABLE_COLUMNS}
    row["Raw Text"] = text.strip()
    row["NTIN_CODE"] = extract_ntin_code(modal)
    row["Полное наименование товара (рус)"] = get_field_value_from_modal(modal, ["Полное наименование товара (рус)", "Полное наименование товара (рус) *", "Полное наименование товара (рус) **"])
    row["Полное наименование товара (каз)"] = get_field_value_from_modal(modal, ["Полное наименование товара (каз)", "Полное наименование товара (каз) *", "Полное наименование товара (каз) **"])
    row["Краткое наименование товара (рус)"] = get_field_value_from_modal(modal, ["Краткое наименование товара (рус)", "Краткое наименование товара (рус) *", "Краткое наименование товара (рус) **"])
    row["Страна происхождения"] = get_field_value_from_modal(modal, ["Страна происхождения", "Страна происхождения *", "Страна происхождения **"])
    row["Единица измерения"] = get_field_value_from_modal(modal, ["Единица измерения", "Единица измерения *", "Единица измерения **"])
    row["Количественное значение"] = get_field_value_from_modal(modal, ["Количество количественное значение", "Количество количественное значение (в [ед. изм.])", "Количественное значение", "Количество значение"])
    row["ТНВЭД ЕАЭС"] = get_field_value_from_modal(modal, ["ТНВЭД ЕАЭС", "ТНВЭД ЕАЭС *", "ТНВЭД ЕАЭС **"])
    row["Наименование производителя"] = get_field_value_from_modal(modal, ["Наименование производителя", "Наименование производителя *", "Наименование производителя **"])
    row["Категория ОКТРУ (НКТ)"] = get_field_value_from_modal(modal, ["Категория ОКТРУ (НКТ)", "Категория ОКТРУ (НКТ) *", "Категория ОКТРУ (НКТ) **"])
    row["Подобрано AI"] = get_field_value_from_modal(modal, ["Подобрано AI", "Подобрано AI *", "Подобрано AI **"])
    row["Расширенная форма заявки"] = get_field_value_from_modal(modal, ["Расширенная форма заявки", "Расширенная форма заявки *", "Расширенная форма заявки **"])
    return row


def are_required_fields_loaded(row):
    return all(str(row.get(field, "")).strip() for field in REQUIRED_FIELDS)


def wait_for_modal_ready(driver, timeout=30):
    wait = WebDriverWait(driver, timeout)
    wait.until(EC.visibility_of_element_located((By.XPATH, MODAL_FALLBACK_XPATH)))

    def modal_is_ready(_driver):
        try:
            modal = get_modal_container(_driver)
            if modal is None or not modal.is_displayed():
                return False
            row = extract_fields_from_modal(_driver)
            return modal if are_required_fields_loaded(row) else False
        except Exception:
            return False

    return wait.until(modal_is_ready)


def process_single_sku(driver, sku, logger, outdir=OUTPUT_DIR):
    try:
        clear_search_input(driver)
        search = find_search_input(driver)
        try:
            search.clear()
        except Exception:
            pass
        search.send_keys(sku)
        time.sleep(2)
        try:
            search.send_keys(Keys.ENTER)
        except Exception:
            pass
        click_create_button(driver)
        wait_for_modal_ready(driver, timeout=30)
        row = extract_fields_from_modal(driver)
        row["SKU"] = sku
        row["Status"] = "success"
        row["Error"] = ""
        logger.info(f"[SKU: {sku}] Extracted successfully")
        try:
            close_modal_cancel(driver)
        except Exception as e:
            logger.warning(f"[SKU: {sku}] Error while closing modal: {e}")
        return row
    except Exception as e:
        logger.error(f"[SKU: {sku}] Error: {e}")
        failed_row = {col: "" for col in TABLE_COLUMNS}
        failed_row["SKU"] = sku
        failed_row["Status"] = "error"
        failed_row["Error"] = str(e)
        return failed_row


def main(sku=None, url=URL, headless=False, outdir=None, username=None, password=None, sku_list=None, input_file=None, sku_column=INPUT_SKU_COLUMN):
    outdir = Path(outdir or OUTPUT_DIR)
    outdir.mkdir(parents=True, exist_ok=True)
    logger = setup_logging(outdir)
    input_path = Path(input_file or INPUT_FILE_NAME)
    progress_path = outdir / PROGRESS_FILE_NAME

    if sku_list is None:
        sku_list = [sku] if sku else load_skus_from_excel(input_path, sku_column=sku_column)
    if not sku_list:
        logger.error("No SKU provided or found in input file.")
        return 1

    existing_progress = load_existing_progress(progress_path)
    rows = existing_progress.to_dict("records") if not existing_progress.empty else []
    processed_skus = get_processed_skus(existing_progress, only_success=True)
    pending_skus = [str(x).strip() for x in sku_list if str(x).strip() and str(x).strip() not in processed_skus]

    logger.info(f"Loaded {len(sku_list)} SKU(s) from source")
    logger.info(f"Already processed successfully: {len(processed_skus)}")
    logger.info(f"Pending: {len(pending_skus)}")

    if not pending_skus:
        final_file = save_results_to_excel(rows, outdir=outdir)
        logger.info(f"Nothing to process. Final file is up to date: {final_file}")
        return 0

    try:
        driver = get_driver(headless=headless)
    except WebDriverException as e:
        logger.error(f"Could not start Chrome automatically: {e}")
        logger.error("Start Chrome manually, open the target page, log in, pass captcha, then rerun the script.")
        return 1

    try:
        logger.info("Chrome opened by Selenium.")
        logger.info(f"Open this URL manually in the browser if needed: {url}")
        driver.get("about:blank")
        wait_for_manual_ready(driver, logger)

        try:
            find_search_input(driver)
        except Exception:
            logger.error("Search field not found after manual preparation. Make sure you are on the working NTIN page.")
            return 1

        successful = sum(1 for r in rows if str(r.get("Status", "")) == "success")
        failed = sum(1 for r in rows if str(r.get("Status", "")) == "error")
        newly_processed = 0

        for index, single_sku in enumerate(pending_skus, start=1):
            logger.info(f"Processing SKU {index}/{len(pending_skus)}: {single_sku}")
            try:
                row = process_single_sku(driver, single_sku, logger, outdir=outdir)
            except (InvalidSessionIdException, WebDriverException) as e:
                logger.error(f"Browser session was lost while processing SKU {single_sku}: {e}")
                logger.error("Save progress, reopen browser manually, and rerun the script to continue from remaining successful items.")
                checkpoint_file = save_checkpoint(rows, outdir=outdir)
                logger.info(f"Emergency checkpoint saved: {checkpoint_file}")
                break

            rows.append(row)
            if row.get("Status") == "success":
                successful += 1
            else:
                failed += 1
            newly_processed += 1

            if newly_processed % CHECKPOINT_EVERY == 0:
                checkpoint_file = save_checkpoint(rows, outdir=outdir)
                logger.info(f"Checkpoint saved after {newly_processed} new SKU(s): {checkpoint_file}")
            time.sleep(1)

        checkpoint_file = save_checkpoint(rows, outdir=outdir)
        output_file = save_results_to_excel(rows, outdir=outdir)
        logger.info(f"Final checkpoint saved: {checkpoint_file}")
        logger.info(f"Saved final results to {output_file}")
        logger.info(f"Completed: {successful} successful, {failed} failed, total stored rows: {len(rows)}")
        return 0 if successful > 0 else 1
    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='NTIN Scraper: Automate SKU search and data extraction')
    parser.add_argument('sku', nargs='?', default=None, help='Optional single SKU to process')
    parser.add_argument('--headless', action='store_true')
    parser.add_argument('--outdir', default=None, help='Output directory for Excel/log files')
    parser.add_argument('--url', default=None, help='Override target URL')
    parser.add_argument('--username', default=None, help='Unused in manual mode; kept for compatibility')
    parser.add_argument('--password', default=None, help='Unused in manual mode; kept for compatibility')
    parser.add_argument('--input', default=INPUT_FILE_NAME, help='Input Excel file with SKU values')
    parser.add_argument('--sku-column', default=INPUT_SKU_COLUMN, help='Column name containing SKU values')
    args = parser.parse_args()
    if args.url:
        URL = args.url
    rc = main(sku=args.sku, url=URL, headless=args.headless, outdir=args.outdir, username=args.username, password=args.password, input_file=args.input, sku_column=args.sku_column)
    sys.exit(rc)
