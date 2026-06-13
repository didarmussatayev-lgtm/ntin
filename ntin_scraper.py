"""ntin_scraper.py

Usage:
  python ntin_scraper.py [--outdir=path] [--input=ACTIVE.xlsx]

This script opens the provided site, logs in, loads SKU values from an Excel file,
processes them sequentially, waits for the modal to fully load, extracts structured
field values, and periodically saves progress so the run can be resumed after failure.
If Cloudflare/captcha appears, the script pauses, emits a sound on Windows, and lets
an operator solve it manually before continuing in the same browser session.
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
from selenium.common.exceptions import TimeoutException

from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

try:
    import winsound
except ImportError:
    winsound = None

# ----- CONFIG (adjust if needed) -----
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
CHROME_USER_DATA_DIR = r"C:\Users\medal\AppData\Local\Google\Chrome\User Data"
CHROME_PROFILE_DIRECTORY = "Profile 5"

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
# --------------------------------------


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
    options.add_argument(f"--user-data-dir={CHROME_USER_DATA_DIR}")
    options.add_argument(f"--profile-directory={CHROME_PROFILE_DIRECTORY}")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--remote-debugging-port=9222")

    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


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


def get_processed_skus(progress_df):
    if progress_df.empty or "SKU" not in progress_df.columns:
        return set()
    return set(progress_df["SKU"].dropna().astype(str).str.strip())


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


def is_cloudflare_page(driver):
    try:
        body_text = (driver.find_element(By.TAG_NAME, "body").text or "").lower()
    except Exception:
        body_text = ""
    page_source = (driver.page_source or "").lower()
    markers = [
        "cloudflare",
        "проверки безопасности",
        "подтвердите, что вы человек",
        "verify you are human",
        "security check",
    ]
    return any(marker in body_text or marker in page_source for marker in markers)


def wait_for_manual_captcha_if_needed(driver, logger, reason="Captcha detected"):
    if not is_cloudflare_page(driver):
        return
    alert_operator(f"{reason}. Solve captcha manually in the browser window.", logger)
    input("Solve captcha manually in the opened browser, then press Enter here to continue...")


def ensure_ready_for_work(driver, logger, timeout=180):
    wait_for_manual_captcha_if_needed(driver, logger, reason="Initial captcha/check page detected")
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            find_search_input(driver)
            return True
        except Exception:
            if is_cloudflare_page(driver):
                wait_for_manual_captcha_if_needed(driver, logger, reason="Captcha reappeared")
            else:
                time.sleep(2)
    return False


def perform_login(driver, username, password):
    email_selectors = [
        (By.NAME, 'email'),
        (By.NAME, 'username'),
        (By.CSS_SELECTOR, 'input[type="email"]'),
        (By.CSS_SELECTOR, 'input[placeholder*="Email"]'),
        (By.CSS_SELECTOR, 'input[placeholder*="Эл. почта"]'),
    ]
    password_selectors = [
        (By.NAME, 'password'),
        (By.CSS_SELECTOR, 'input[type="password"]'),
        (By.CSS_SELECTOR, 'input[placeholder*="Пароль"]'),
    ]
    email_el = None
    pass_el = None
    for by, sel in email_selectors:
        try:
            email_el = driver.find_element(by, sel)
            break
        except Exception:
            continue
    for by, sel in password_selectors:
        try:
            pass_el = driver.find_element(by, sel)
            break
        except Exception:
            continue
    if not email_el or not pass_el:
        return False
    try:
        email_el.clear()
        email_el.send_keys(username)
        pass_el.clear()
        pass_el.send_keys(password)
    except Exception:
        return False
    try:
        btn = pass_el.find_element(By.XPATH, "ancestor::form//button[@type='submit']")
    except Exception:
        btn = None
    if not btn:
        try:
            btn = driver.find_element(By.XPATH, "//button[contains(., 'Войти') or contains(., 'Вход') or contains(., 'Login') or contains(., 'Sign in')]")
        except Exception:
            btn = None
    try:
        if btn:
            btn.click()
        else:
            pass_el.send_keys(Keys.ENTER)
    except Exception:
        pass
    return True


def click_go_button(driver, wait_timeout=20):
    try:
        btn = WebDriverWait(driver, wait_timeout).until(EC.element_to_be_clickable((By.XPATH, GO_BUTTON_XPATH)))
        btn.click()
        time.sleep(1)
        return True
    except Exception:
        return False


def close_modal_cancel(driver, wait_timeout=5):
    candidates = [
        "//button[contains(., 'Отмена') or contains(., 'Отмен')]",
        "//button[@aria-label='close' or @aria-label='Close']",
        "//div[contains(@class,'MuiDialog-root')]//button[contains(., 'Отмена') or contains(@class,'cancel')]",
    ]
    for xp in candidates:
        try:
            btn = WebDriverWait(driver, wait_timeout).until(EC.element_to_be_clickable((By.XPATH, xp)))
            try:
                btn.click()
                time.sleep(0.5)
                return True
            except Exception:
                driver.execute_script("arguments[0].click();", btn)
                time.sleep(0.5)
                return True
        except Exception:
            continue
    try:
        body = driver.find_element(By.TAG_NAME, 'body')
        body.send_keys(Keys.ESCAPE)
        time.sleep(0.5)
        return True
    except Exception:
        return False


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


def wait_for_modal_ready(driver, logger, timeout=30):
    wait = WebDriverWait(driver, timeout)
    wait.until(EC.visibility_of_element_located((By.XPATH, MODAL_FALLBACK_XPATH)))

    def modal_is_ready(_driver):
        try:
            if is_cloudflare_page(_driver):
                return False
            modal = get_modal_container(_driver)
            if modal is None or not modal.is_displayed():
                return False
            row = extract_fields_from_modal(_driver)
            return modal if are_required_fields_loaded(row) else False
        except Exception:
            return False

    try:
        return wait.until(modal_is_ready)
    except TimeoutException:
        if is_cloudflare_page(driver):
            wait_for_manual_captcha_if_needed(driver, logger, reason="Captcha appeared while waiting for modal")
            return wait_for_modal_ready(driver, logger, timeout=timeout)
        raise


def process_single_sku(driver, sku, logger, outdir=OUTPUT_DIR):
    try:
        wait_for_manual_captcha_if_needed(driver, logger, reason=f"Before processing SKU {sku}")
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
        wait_for_manual_captcha_if_needed(driver, logger, reason=f"After entering SKU {sku}")
        click_create_button(driver)
        wait_for_modal_ready(driver, logger, timeout=30)
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
    processed_skus = get_processed_skus(existing_progress)
    pending_skus = [str(x).strip() for x in sku_list if str(x).strip() and str(x).strip() not in processed_skus]

    logger.info(f"Loaded {len(sku_list)} SKU(s) from source")
    logger.info(f"Already processed: {len(processed_skus)}")
    logger.info(f"Pending: {len(pending_skus)}")

    if not pending_skus:
        final_file = save_results_to_excel(rows, outdir=outdir)
        logger.info(f"Nothing to process. Final file is up to date: {final_file}")
        return 0

    driver = get_driver(headless=headless)
    try:
        driver.get(url)
        if not ensure_ready_for_work(driver, logger, timeout=180):
            logger.error("Search field did not become available after manual captcha/login waiting.")
            return 1

        logged_in = False
        try:
            find_search_input(driver)
            logged_in = True
        except Exception:
            user = username or os.environ.get('ALGATOP_USER')
            pwd = password or os.environ.get('ALGATOP_PASS')
            if user and pwd:
                logger.info("Attempting login...")
                perform_login(driver, user, pwd)
                time.sleep(2)
                wait_for_manual_captcha_if_needed(driver, logger, reason="After login attempt")
                logger.info("Clicking 'Перейти' button...")
                if click_go_button(driver):
                    time.sleep(2)
                ensure_ready_for_work(driver, logger, timeout=180)
                logged_in = True

        if not logged_in:
            logger.error("Could not find search field or login. Check credentials/page layout.")
            return 1

        successful = sum(1 for r in rows if str(r.get("Status", "")) == "success")
        failed = sum(1 for r in rows if str(r.get("Status", "")) == "error")
        newly_processed = 0

        for index, single_sku in enumerate(pending_skus, start=1):
            logger.info(f"Processing SKU {index}/{len(pending_skus)}: {single_sku}")
            row = process_single_sku(driver, single_sku, logger, outdir=outdir)
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
        driver.quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='NTIN Scraper: Automate SKU search and data extraction')
    parser.add_argument('sku', nargs='?', default=None, help='Optional single SKU to process')
    parser.add_argument('--headless', action='store_true')
    parser.add_argument('--outdir', default=None, help='Output directory for Excel/log files')
    parser.add_argument('--url', default=None, help='Override target URL')
    parser.add_argument('--username', default=None, help='Login username/email (or set ALGATOP_USER env)')
    parser.add_argument('--password', default=None, help='Login password (or set ALGATOP_PASS env)')
    parser.add_argument('--input', default=INPUT_FILE_NAME, help='Input Excel file with SKU values')
    parser.add_argument('--sku-column', default=INPUT_SKU_COLUMN, help='Column name containing SKU values')
    args = parser.parse_args()
    if args.url:
        URL = args.url
    rc = main(sku=args.sku, url=URL, headless=args.headless, outdir=args.outdir, username=args.username, password=args.password, input_file=args.input, sku_column=args.sku_column)
    sys.exit(rc)
