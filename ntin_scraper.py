import pandas as pd
from playwright.sync_api import sync_playwright
import time
import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

PRODUCTS_URL = "https://kaspi.kz/mc/#/products?status=active&search"

def run():
    try:
        df = pd.read_excel("personal-products-2026-06-10_2200.xlsx")
        df['NTIN'] = df['NTIN'].astype(str).str.strip()
    except Exception as e:
        print(f"Ошибка чтения файла: {e}")
        return

    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=False,
            executable_path=r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            args=[
                "--remote-debugging-port=9222",
                "--user-data-dir=C:\\ChromeData",
                "--profile-directory=Profile 4"
            ]
        )

        page = browser.contexts[0].new_page()
        page.goto(PRODUCTS_URL)
        input("Нажмите Enter когда страница загружена...")

        for index, row in df.iterrows():
            name_ru = str(row['Наименование на русском']).strip()
            ntin    = str(row['NTIN']).strip()

            print(f"\n[{index+1}] {name_ru} | NTIN: {ntin}")

            try:
                # ── 1. Список товаров ────────────────────────────────────────
                page.goto(PRODUCTS_URL)
                page.wait_for_load_state("networkidle")
                time.sleep(2)

                # ── 2. Поиск ─────────────────────────────────────────────────
                search_input = page.locator('input[type="text"]').first
                search_input.wait_for(state="visible", timeout=15000)
                search_input.click()
                search_input.fill(name_ru)
                search_input.press("Enter")
                time.sleep(3)

                # ── 3. Клик по товару ────────────────────────────────────────
                product_link = page.locator(f'text="{name_ru}"').first
                product_link.wait_for(state="visible", timeout=15000)
                product_link.click()
                page.wait_for_load_state("networkidle")
                time.sleep(2)

                # ── 4. Добавить штрихкод ─────────────────────────────────────
                add_btn = page.locator('button:has-text("Добавить штрихкод")')
                if add_btn.count() == 0:
                    print(f"  -> ПРОПУСК: штрихкод уже есть")
                    continue

                add_btn.first.scroll_into_view_if_needed()
                add_btn.first.click()

                # ── 5. Ждём диалог ───────────────────────────────────────────
                page.wait_for_selector('[role="dialog"]', timeout=10000)
                time.sleep(1.5)

                # ── 6. Берём поле по классу — он точно есть на скриншоте ─────
                # input.ds-input__input.body-500.ds-input__l
                barcode_input = page.locator('input.ds-input__input').last
                barcode_input.wait_for(state="visible", timeout=8000)

                # Печатаем что сейчас в поле
                before = barcode_input.input_value()
                print(f"  -> Поле до ввода: '{before}'")

                # ── 7. Клик по полю и посимвольный ввод ──────────────────────
                barcode_input.click()
                time.sleep(0.5)
                barcode_input.press("Control+a")
                barcode_input.press("Delete")
                time.sleep(0.3)

                # .type() эмулирует нажатие каждой клавиши отдельно
                barcode_input.type(ntin, delay=100)
                time.sleep(0.5)

                after = barcode_input.input_value()
                print(f"  -> Поле после ввода: '{after}'")

                # ── 8. Если всё ещё пусто — пробуем через page.keyboard ──────
                if after != ntin:
                    print(f"  -> Пробуем через page.keyboard...")
                    barcode_input.focus()
                    time.sleep(0.3)
                    page.keyboard.press("Control+a")
                    page.keyboard.press("Delete")
                    time.sleep(0.2)
                    page.keyboard.type(ntin, delay=100)
                    time.sleep(0.5)
                    after = barcode_input.input_value()
                    print(f"  -> Поле после keyboard: '{after}'")

                # ── 9. Tab для валидации ─────────────────────────────────────
                barcode_input.press("Tab")
                time.sleep(0.8)

                # ── 10. Сохранить ─────────────────────────────────────────────
                save_btn = page.locator('[role="dialog"] button:has-text("Сохранить")')
                save_btn.wait_for(state="visible", timeout=5000)

                enabled = False
                for _ in range(20):
                    if save_btn.is_enabled():
                        enabled = True
                        break
                    time.sleep(0.3)

                if not enabled:
                    print(f"  -> ОШИБКА: кнопка неактивна! Значение: '{after}'")
                    page.keyboard.press("Escape")
                    continue

                save_btn.click()
                print(f"  -> ✅ Сохранено!")
                time.sleep(1.5)

            except Exception as e:
                print(f"  -> ОШИБКА: {e}")
                try:
                    page.goto(PRODUCTS_URL)
                    page.wait_for_load_state("networkidle")
                    time.sleep(2)
                except Exception:
                    page = browser.contexts[0].new_page()

        print("\n✅ Готово!")
        browser.close()

if __name__ == "__main__":
    run()
