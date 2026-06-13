"""ntin_scraper.py

Usage:
  python ntin_scraper.py [--headless] [--outdir=path]

This script opens the provided site, logs in, enters each SKU from SKU_LIST,
waits, clicks "Создать заявку", extracts structured field values from the modal,
and saves everything into a single Excel file with one row per SKU.

Selectors are pre-configured from user input. Adjust XPATHs below if needed.
"""
import time
import re
import sys
import os
from pathlib import Path
import argparse

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd

# ----- CONFIG (adjust if needed) -----
URL = "https://app.algatop.kz/ntin"

# LIST OF SKU TO PROCESS - add/remove/edit as needed
SKU_LIST = [
    "115050528_928421346",
    "104841221_113781",
    "189893632"

    # Add more SKU here, one per line, like:
    # "SKU_2",
    # "SKU_3",
]

# User-provided selectors
# Primary XPATH may be dynamic (contains colons). Use a list of candidates for robustness.
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
# Кнопка "Перейти" - appears after login
GO_BUTTON_XPATH = "//button[contains(., 'Перейти')]"
MODAL_TITLE_TEXT = "Создание заявки"
MODAL_FALLBACK_XPATH = "//div[@role='dialog'] | //div[contains(@class,'modal')] | //div[contains(@class,'MuiDialog-root')]"
OUTPUT_FILE_NAME = "output.xlsx"
OUTPUT_DIR = Path.cwd()
WAIT_TIMEOUT = 20
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
]
# --------------------------------------


def get_driver(headless=False):
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1400,1000")
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)


def perform_login(driver, username, password, wait_timeout=20):
    """Try several common selectors to locate login form, fill credentials and submit.
    Returns True if login action likely performed, False otherwise.
    """
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

    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, SEARCH_XPATH)))
        return True
    except Exception:
        return True


def click_go_button(driver, wait_timeout=20):
    """Click the 'Перейти' button that appears after login."""
    try:
        btn = WebDriverWait(driver, wait_timeout).until(
            EC.element_to_be_clickable((By.XPATH, GO_BUTTON_XPATH))
        )
        btn.click()
        time.sleep(1)
        return True
    except Exception as e:
        print(f"Warning: Could not find/click 'Перейти' button: {e}")
        return False


def close_modal_cancel(driver, wait_timeout=5):
    """Close the open modal by clicking 'Отмена' button, or the close icon, or sending ESC."""
    candidates = [
        "//button[contains(., 'Отмена') or contains(., 'Отмен')]",
        "//button[@aria-label='close' or @aria-label='Close']",
        "//div[contains(@class,'MuiDialog-root')]//button[contains(., 'Отмена') or contains(@class,'cancel')]",
    ]
    for xp in candidates:
        try:
            btn = WebDriverWait(driver, wait_timeout).until(
                EC.element_to_be_clickable((By.XPATH, xp))
            )
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
        close_x = driver.find_element(By.XPATH, "//button[contains(@aria-label,'close') or contains(., '×') or contains(., '✖')]")
        try:
            close_x.click()
            time.sleep(0.5)
            return True
        except Exception:
            driver.execute_script("arguments[0].click();", close_x)
            time.sleep(0.5)
            return True
    except Exception:
        pass

    try:
        body = driver.find_element(By.TAG_NAME, 'body')
        body.send_keys(Keys.ESCAPE)
        time.sleep(0.5)
        return True
    except Exception:
        print("Warning: could not close modal (no cancel/close found)")
        return False


def clear_search_input(driver, wait_timeout=3):
    """Attempt multiple ways to clear the search input or remove selected tokens.
    Returns True if cleared or no input found, False otherwise.
    """
    try:
        el = find_search_input(driver)
    except Exception:
        return True

    try:
        el.clear()
        time.sleep(0.15)
    except Exception:
        pass

    try:
        el.send_keys(Keys.CONTROL, 'a')
        el.send_keys(Keys.DELETE)
        time.sleep(0.15)
    except Exception:
        try:
            el.send_keys(Keys.COMMAND, 'a')
            el.send_keys(Keys.DELETE)
            time.sleep(0.15)
        except Exception:
            pass

    try:
        driver.execute_script("arguments[0].value = ''; arguments[0].dispatchEvent(new Event('input'));", el)
        time.sleep(0.15)
    except Exception:
        pass

    try:
        val = el.get_attribute('value') or el.get_attribute('text') or ''
        if val.strip() == '':
            return True
    except Exception:
        pass

    token_xpaths = [
        "//div[contains(@class,'MuiChip-root')]//button",
        "//button[contains(@class,'chip-remove') or contains(@aria-label,'remove') or contains(@aria-label,'Удалить')]",
        "//span[contains(@class,'token')]/button",
    ]
    removed_any = False
    for xp in token_xpaths:
        try:
            elems = driver.find_elements(By.XPATH, xp)
            for e in elems:
                try:
                    e.click()
                    removed_any = True
                    time.sleep(0.05)
                except Exception:
                    try:
                        driver.execute_script("arguments[0].click();", e)
                        removed_any = True
                        time.sleep(0.05)
                    except Exception:
                        continue
        except Exception:
            continue

    if removed_any:
        try:
            el = find_search_input(driver)
            val = el.get_attribute('value') or ''
            return val.strip() == ''
        except Exception:
            return True

    try:
        el.click()
        for _ in range(10):
            el.send_keys(Keys.BACKSPACE)
        time.sleep(0.1)
        val = el.get_attribute('value') or ''
        return val.strip() == ''
    except Exception:
        return False


def find_search_input(driver):
    for xpath in SEARCH_CANDIDATES:
        try:
            el = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.XPATH, xpath))
            )
            if el.is_displayed() and el.is_enabled():
                return el
        except Exception:
            continue
    return WebDriverWait(driver, WAIT_TIMEOUT).until(
        EC.presence_of_element_located((By.XPATH, SEARCH_XPATH))
    )


def click_create_button(driver):
    candidates = [
        CREATE_BUTTON_XPATH,
        "//button[contains(., 'Создать заявку') or contains(., 'Создать')]",
        "//span[contains(., 'Создать заявку')]/ancestor::button",
    ]
    last_exc = None
    for xp in candidates:
        try:
            btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.XPATH, xp))
            )
            btn.click()
            return True
        except Exception as e:
            last_exc = e
            continue
    raise last_exc


def find_modal_text(driver):
    try:
        title_el = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{MODAL_TITLE_TEXT}')]") )
        )
        ancestor = title_el.find_element(By.XPATH, "ancestor::div[@role='dialog'] | ancestor::div[contains(@class,'modal')] | ancestor::div[contains(@class,'MuiDialog-root')] | ..")
        text = ancestor.text
        if text.strip():
            return text
    except Exception:
        pass

    try:
        modal = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, MODAL_FALLBACK_XPATH))
        )
        time.sleep(0.3)
        return modal.text
    except Exception:
        pass

    return driver.page_source


def extract_from_1103(text):
    m = re.search(r'(1103[-\d]+)', text)
    if m:
        return m.group(1)

    lines = text.splitlines()
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("1103-"):
            m = re.match(r'(1103[-\d]+)', stripped)
            if m:
                return m.group(1)
    return None


def save_results_to_excel(rows, outdir=OUTPUT_DIR, filename=OUTPUT_FILE_NAME):
    outdir = Path(outdir)
    outdir.mkdir(parents=True, exist_ok=True)
    out_file = outdir / filename
    df = pd.DataFrame(rows, columns=TABLE_COLUMNS)
    df.to_excel(out_file, index=False)
    return out_file


def get_next_value(lines, index):
    if index < len(lines):
        line = lines[index]
        parts = re.split(r"\s*[:\-]\s*", line, maxsplit=1)
        if len(parts) == 2 and parts[0].strip() != line.strip():
            return parts[1].strip()
        if index + 1 < len(lines):
            return lines[index + 1].strip()
    return ""


def find_label_value(lines, labels):
    for i, line in enumerate(lines):
        for label in labels:
            if line.startswith(label):
                value = line[len(label):].strip()
                if value:
                    return value
                return get_next_value(lines, i + 1)
    return ""


def get_modal_container(driver):
    try:
        return WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, MODAL_FALLBACK_XPATH))
        )
    except Exception:
        return None


def extract_ntin_code(modal):
    if modal is None:
        return ""

    patterns = [
        r'(\d{4}-\d{4}-\d{4}-\d+)',
        r'(1103[-\d]+)',
        r'(1070[-\d]+)',
        r'(8504[-\d]+)',
    ]

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
            try:
                txt = (el.text or "").strip()
                if not txt:
                    continue
                for pattern in patterns:
                    m = re.search(pattern, txt)
                    if m:
                        return m.group(1)
            except Exception:
                continue
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
                        linked = modal.find_element(By.ID, label_for)
                        tag = linked.tag_name.lower()

                        if tag in ["input", "textarea"]:
                            value = linked.get_attribute("value")
                            if value and value.strip():
                                return value.strip()

                        text = (linked.text or "").strip()
                        if text:
                            return text
                except Exception:
                    pass

                near_input_xpaths = [
                    "./following-sibling::*//input[1]",
                    "./following-sibling::*//textarea[1]",
                    "./ancestor::div[1]//input[1]",
                    "./ancestor::div[1]//textarea[1]",
                    "./ancestor::div[2]//input[1]",
                    "./ancestor::div[2]//textarea[1]",
                    "./parent::*//input[1]",
                    "./parent::*//textarea[1]",
                ]

                for xp in near_input_xpaths:
                    try:
                        found = label_el.find_elements(By.XPATH, xp)
                        for el in found:
                            tag = el.tag_name.lower()
                            if tag in ["input", "textarea"]:
                                value = el.get_attribute("value")
                                if value and value.strip():
                                    return value.strip()
                    except Exception:
                        continue

                near_select_xpaths = [
                    "./following-sibling::*//*[contains(@class,'MuiSelect')][1]",
                    "./ancestor::div[1]//*[contains(@class,'MuiSelect')][1]",
                    "./ancestor::div[2]//*[contains(@class,'MuiSelect')][1]",
                    "./parent::*//*[contains(@class,'MuiSelect')][1]",
                ]

                for xp in near_select_xpaths:
                    try:
                        found = label_el.find_elements(By.XPATH, xp)
                        for el in found:
                            text = (el.text or "").strip()
                            if text:
                                return text
                    except Exception:
                        continue

                text_xpaths = [
                    "./following-sibling::*[1]",
                    "./parent::*",
                    "./ancestor::div[1]",
                    "./ancestor::div[2]",
                ]

                for xp in text_xpaths:
                    try:
                        found = label_el.find_elements(By.XPATH, xp)
                        for el in found:
                            try:
                                inner_inputs = el.find_elements(By.XPATH, ".//input | .//textarea")
                                for inner in inner_inputs:
                                    value = inner.get_attribute("value")
                                    if value and value.strip():
                                        return value.strip()
                            except Exception:
                                pass

                            txt = (el.text or "").strip()
                            if txt and txt != label and label not in txt:
                                return txt
                    except Exception:
                        continue

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

    row["Полное наименование товара (рус)"] = get_field_value_from_modal(modal, [
        "Полное наименование товара (рус)",
        "Полное наименование товара (рус) *",
        "Полное наименование товара (рус) **",
    ])

    row["Полное наименование товара (каз)"] = get_field_value_from_modal(modal, [
        "Полное наименование товара (каз)",
        "Полное наименование товара (каз) *",
        "Полное наименование товара (каз) **",
    ])

    row["Краткое наименование товара (рус)"] = get_field_value_from_modal(modal, [
        "Краткое наименование товара (рус)",
        "Краткое наименование товара (рус) *",
        "Краткое наименование товара (рус) **",
    ])

    row["Страна происхождения"] = get_field_value_from_modal(modal, [
        "Страна происхождения",
        "Страна происхождения *",
        "Страна происхождения **",
    ])

    row["Единица измерения"] = get_field_value_from_modal(modal, [
        "Единица измерения",
        "Единица измерения *",
        "Единица измерения **",
    ])

    row["Количественное значение"] = get_field_value_from_modal(modal, [
        "Количество количественное значение",
        "Количество количественное значение (в [ед. изм.])",
        "Количественное значение",
        "Количество значение",
    ])

    row["ТНВЭД ЕАЭС"] = get_field_value_from_modal(modal, [
        "ТНВЭД ЕАЭС",
        "ТНВЭД ЕАЭС *",
        "ТНВЭД ЕАЭС **",
    ])

    row["Наименование производителя"] = get_field_value_from_modal(modal, [
        "Наименование производителя",
        "Наименование производителя *",
        "Наименование производителя **",
    ])

    row["Категория ОКТРУ (НКТ)"] = get_field_value_from_modal(modal, [
        "Категория ОКТРУ (НКТ)",
        "Категория ОКТРУ (НКТ) *",
        "Категория ОКТРУ (НКТ) **",
    ])

    row["Подобрано AI"] = get_field_value_from_modal(modal, [
        "Подобрано AI",
        "Подобрано AI *",
        "Подобрано AI **",
    ])

    row["Расширенная форма заявки"] = get_field_value_from_modal(modal, [
        "Расширенная форма заявки",
        "Расширенная форма заявки *",
        "Расширенная форма заявки **",
    ])

    return row


def process_single_sku(driver, sku, outdir=OUTPUT_DIR):
    """Process a single SKU: input it, click button, extract data, return row dict."""
    try:
        try:
            clear_search_input(driver)
        except Exception:
            pass

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
        time.sleep(7)

        row = extract_fields_from_modal(driver)
        row["SKU"] = sku

        print(f"  [SKU: {sku}] Extracted row values")
        try:
            closed = close_modal_cancel(driver)
            if not closed:
                print(f"  [SKU: {sku}] Warning: could not automatically close modal")
        except Exception as e:
            print(f"  [SKU: {sku}] Warning: error while closing modal: {e}")
        return row
    except Exception as e:
        print(f"  [SKU: {sku}] Error: {e}")
        return None


def main(sku=None, url=URL, headless=False, outdir=None, username=None, password=None, sku_list=None):
    """Main function: login once, click 'Перейти', then process all SKUs from sku_list."""
    if sku_list is None:
        if sku:
            sku_list = [sku]
        else:
            sku_list = SKU_LIST

    if not sku_list:
        print("Error: No SKU provided. Add SKU to SKU_LIST in config or pass --sku argument.")
        return 1

    driver = get_driver(headless=headless)
    try:
        driver.get(url)

        logged_in = False
        try:
            find_search_input(driver)
            logged_in = True
        except Exception:
            user = username or os.environ.get('ALGATOP_USER')
            pwd = password or os.environ.get('ALGATOP_PASS')
            if user and pwd:
                print("Attempting login...")
                perform_login(driver, user, pwd)
                time.sleep(2)
                logged_in = True
                print("Clicking 'Перейти' button...")
                if click_go_button(driver):
                    time.sleep(2)
                else:
                    print("Warning: 'Перейти' button not found, continuing anyway...")

        if not logged_in:
            print("Error: Could not find search field or login. Check credentials/page layout.")
            return 1

        rows = []
        successful = 0
        failed = 0
        print(f"\nProcessing {len(sku_list)} SKU(s)...\n")
        for single_sku in sku_list:
            print(f"Processing SKU: {single_sku}")
            row = process_single_sku(driver, single_sku, outdir=outdir or OUTPUT_DIR)
            if row:
                rows.append(row)
                successful += 1
            else:
                failed += 1
            time.sleep(1)

        if rows:
            output_file = save_results_to_excel(rows, outdir=outdir or OUTPUT_DIR)
            print(f"\nSaved combined results to {output_file}")
        else:
            print("\nNo rows were extracted. No output file created.")

        print(f"\n✓ Completed: {successful} successful, {failed} failed out of {len(sku_list)}")
        return 0 if successful > 0 else 1
    finally:
        driver.quit()


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='NTIN Scraper: Automate SKU search and data extraction')
    parser.add_argument('sku', nargs='?', default=None, help='Optional single SKU to process (if not provided, uses SKU_LIST from config)')
    parser.add_argument('--headless', action='store_true')
    parser.add_argument('--outdir', default=None, help='Output directory for Excel')
    parser.add_argument('--url', default=None, help='Override target URL')
    parser.add_argument('--username', default=None, help='Login username/email (or set ALGATOP_USER env)')
    parser.add_argument('--password', default=None, help='Login password (or set ALGATOP_PASS env)')
    args = parser.parse_args()
    if args.url:
        URL = args.url
    rc = main(sku=args.sku, url=URL, headless=args.headless, outdir=args.outdir, username=args.username, password=args.password)
    sys.exit(rc)
