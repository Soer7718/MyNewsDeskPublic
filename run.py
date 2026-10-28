import time
import json
import os

from src.selenium_driver import get_driver
from src.login import login
from src.report_list_scraper import scrape_report_list, get_report_buttons, _scroll_page
from src.report_metrics_scraper import scrape_metrics

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def main():
    config = json.load(open("config.json", encoding="utf-8"))

    output_dir = config.get("output_folder", "output")
    os.makedirs(output_dir, exist_ok=True)
    excel_path = os.path.join(output_dir, config["excel_filename"])

    # PPTX-konfiguration (tilføj disse til config.json – se kommentarer nedenfor)
    pptx_filename = config.get("pptx_filename", "Impact report - standard report sample.pptx")
    pptx_path = os.path.join(output_dir, pptx_filename)

    driver = get_driver()

    try:
        # ── 1. Log ind ───────────────────────────────────────────────────────
        login(driver)

        # ── 2. Hent liste over rapport-navne ─────────────────────────────────
        print("\nHenter rapport-liste...")
        report_names = scrape_report_list(driver, config["reports_list_url"])
        total = len(report_names)

        if total == 0:
            print("Ingen rapporter fundet – afslutter.")
            return

        print(f"\nStarter scraping af {total} rapporter...\n")

        results = []

        # ── 3. Iterer over hvert rapport-index ───────────────────────────────
        for index in range(total):
            name = report_names[index]
            print(f"== {index + 1}/{total}: {name} ==")

            driver.get(config["reports_list_url"])
            WebDriverWait(driver, 15).until(EC.presence_of_element_located((
                By.XPATH,
                "//h2[contains(normalize-space(), 'Alle rapporter') "
                "or contains(normalize-space(), 'All reports')]"
            )))

            _scroll_page(driver, pause=0.5)

            buttons = get_report_buttons(driver)

            if index >= len(buttons):
                print(f"  ✗ Knap index {index} ikke fundet (kun {len(buttons)} knapper)")
                results.append([name, None, None, None, None, None, None])
                continue

            btn = buttons[index]

            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(0.3)
            try:
                btn.click()
            except Exception:
                driver.execute_script("arguments[0].click();", btn)

            try:
                WebDriverWait(driver, 12).until(
                    lambda d: "/reports/" in d.current_url
                    and d.current_url != config["reports_list_url"]
                )
            except Exception:
                print(f"  ✗ Navigation til rapport fejlede for '{name}'")
                results.append([name, None, None, None, None, None, None])
                continue

            view_url = driver.current_url.replace("/edit", "").split("?")[0]
            if view_url != driver.current_url:
                driver.get(view_url)

            print(f"   URL: {view_url}")

            try:
                metrics = scrape_metrics(driver, view_url)
                results.append([
                    name,
                    metrics.get("coverage"),
                    metrics.get("article_impressions"),
                    metrics.get("ad_value_dkk"),
                    metrics.get("sentiment_positive"),
                    metrics.get("sentiment_neutral"),
                    metrics.get("sentiment_negative"),
                ])
                print(f"   ✓ coverage={metrics.get('coverage')}, "
                      f"impressions={metrics.get('article_impressions')}, "
                      f"ad_value={metrics.get('ad_value_dkk')}")
            except Exception as e:
                print(f"   ✗ Scrape fejlede: {e}")
                results.append([name, None, None, None, None, None, None])

            time.sleep(0.5)

    finally:
        driver.quit()

    # ── 4. Skriv og formater Excel ────────────────────────────────────────────
    print(f"\nSkriver Excel til: {excel_path}")
    from src.excel_writer import write_final_excel
    write_final_excel(excel_path, results)

    from src.excel_formatter import format_excel
    format_excel(excel_path)
    print(f"Excel gemt og formateret: {excel_path}")

    # ── 5. Hent PPTX fra SharePoint, opdater og re-upload ────────────────────
    try:
        from src.sharepoint_uploader import download_from_sharepoint, upload_to_sharepoint
        from src.pptx_updater import update_pptx

        # 5a. Download den eksisterende PPTX fra SharePoint
        pptx_sp_folder = config.get("pptx_sp_folder", "5. Related Minds Group/A. Related Insights/7. Insights skabeloner/Impact report")
        print(f"\nHenter PPTX fra SharePoint: {pptx_filename}")
        download_from_sharepoint(pptx_filename, pptx_path, folder=pptx_sp_folder)

        # 5b. Opdater PPTX med nye tal fra Excel
        print("Opdaterer PPTX med nye tal...")
        update_pptx(excel_path=excel_path, pptx_path=pptx_path)

        # 5c. Upload den opdaterede PPTX tilbage til SharePoint
        web_url = upload_to_sharepoint(pptx_path, folder=pptx_sp_folder)
        print(f"PPTX uploadet til SharePoint: {web_url}")

    except Exception as e:
        print(f"PPTX-opdatering fejlede (ikke kritisk): {e}")
        import traceback
        traceback.print_exc()

    # ── 6. Upload Excel til SharePoint ───────────────────────────────────────
    try:
        from src.sharepoint_uploader import upload_to_sharepoint
        web_url = upload_to_sharepoint(excel_path)
        print(f"Excel uploadet til SharePoint: {web_url}")
    except Exception as e:
        print(f"SharePoint Excel-upload fejlede (ikke kritisk): {e}")

    print("\nFærdig! ✓")


if __name__ == "__main__":
    main()
