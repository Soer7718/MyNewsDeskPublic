import time
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def _scroll_page(driver, pause=0.8, max_loops=25):
    """Scroller hele siden ned til bunden er nået."""
    last = 0
    for _ in range(max_loops):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(pause)
        new = driver.execute_script("return document.body.scrollHeight")
        if new == last:
            break
        last = new


def _get_alle_rapporter_section(driver):
    """Find sektionen med h2 = 'Alle rapporter' (ikke 'Redigeret for nyligt')."""
    sections = driver.find_elements(By.XPATH, "//section[.//h2]")
    for sec in sections:
        try:
            h2 = sec.find_element(By.XPATH, ".//h2")
            txt = (h2.text or "").strip().lower()
            if "alle rapporter" in txt or "all reports" in txt:
                return sec
        except Exception:
            pass
    return None


def scrape_report_list(driver, list_url):
    """
    Returnerer liste af rapport-navne (strings).

    Verificeret HTML-struktur:
      Hvert kort har en <button aria-label="Rediger: <NAVN>"> –
      dette er den mest stabile måde at hente navne på.
    """
    wait = WebDriverWait(driver, 20)
    driver.get(list_url)

    try:
        wait.until(EC.presence_of_element_located((
            By.XPATH,
            "//h2[contains(normalize-space(), 'Alle rapporter') "
            "or contains(normalize-space(), 'All reports')]"
        )))
    except Exception:
        print("Advarsel: Overskrift ikke fundet – fortsætter alligevel")

    _scroll_page(driver)

    target = _get_alle_rapporter_section(driver)
    search_root = target if target is not None else driver

    rediger_buttons = search_root.find_elements(
        By.XPATH,
        ".//button[starts-with(@aria-label, 'Rediger:') "
        "or starts-with(@aria-label, 'Edit:')]"
    )

    seen = set()
    labels = []
    for btn in rediger_buttons:
        aria = (btn.get_attribute("aria-label") or "").strip()
        if aria.startswith("Rediger: "):
            name = aria[len("Rediger: "):]
        elif aria.startswith("Edit: "):
            name = aria[len("Edit: "):]
        else:
            name = aria
        name = name.strip()
        if name and name not in seen:
            seen.add(name)
            labels.append(name)

    print(f"FOUND {len(labels)} REPORTS")
    for l in labels[:5]:
        print(f"  {l!r}")
    if len(labels) > 5:
        print(f"  ... og {len(labels) - 5} mere")

    return labels


def get_report_buttons(driver):
    """
    Returnerer kortknapperne der navigerer ind i en rapport ved klik.
    Det er <button> inde i <li> der har en <h3> men IKKE har aria-label.
    """
    target = _get_alle_rapporter_section(driver)
    if target is None:
        return []

    return target.find_elements(
        By.XPATH,
        ".//li//button[not(@aria-label) and .//h3]"
    )