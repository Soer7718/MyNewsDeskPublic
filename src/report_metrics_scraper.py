import time
import re
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def _clean_number(val: str):
    if not val:
        return None

    val = val.strip().lower()

    # tag alt før første parentes
    if "(" in val:
        val = val.split("(")[0].strip()

    # håndter "m"
    if val.endswith("m"):
        base = re.sub(r"[^\d.,]", "", val[:-1])
        base = base.replace(".", "").replace(",", "")
        if base.isdigit():
            return int(base) * 1_000_000

    # håndter "k"
    if val.endswith("k"):
        base = re.sub(r"[^\d.,]", "", val[:-1])
        base = base.replace(".", "").replace(",", "")
        if base.isdigit():
            return int(base) * 1000

    # nu fjerner vi ALT andet og normaliserer
    val = re.sub(r"[^0-9,.\s]", "", val)
    val = val.replace(" ", "")
    val = val.replace(".", "").replace(",", "")

    try:
        return int(val)
    except:
        return None


def _find_value(driver, labels):
    """
    Find tallet lige efter en label. Bruges til coverage, impressions, ad value.
    """
    for lbl in labels:
        try:
            element = driver.find_element(
                By.XPATH,
                f"//*[contains(., '{lbl}')]/following::div[contains(@class,'sc-2855e9d8-4')][1]"
            )
            return _clean_number(element.text)
        except:
            pass
    return None


def _find_sentiment(driver, label_text):
    try:
        dd = driver.find_element(
            By.XPATH,
            f"//*[contains(., '{label_text}')]/ancestor::dt/following-sibling::dd[1]"
        )
        return _clean_number(dd.text)
    except:
        return None


def scrape_metrics(driver, url):
    wait = WebDriverWait(driver, 15)

    # Vent på metrics overskrift
    try:
        wait.until(EC.presence_of_element_located((By.XPATH, "//*[contains(text(),'Metrics')]")))
    except:
        pass

    time.sleep(0.4)

    # -----------------------------
    #  Labels for robustness (EN + DK)
    # -----------------------------

    coverage_labels = [
        "Coverage",
        "Dækning",
        "Dækningsgrad",
        "Medieomtale",
        "Dækningsrapport"
    ]

    impressions_labels = [
        "Article impressions",
        "Artikel impressions",
        "Artikelvisninger",
        "Artiklens visninger",
        "Impressions"
    ]

    advalue_labels = [
        "Ad value",
        "Ad value (DKK)",
        "Annonceværdi",
        "Annonce værdi",
        "Værdi"
    ]

    # -----------------------------
    # Extract values
    # -----------------------------

    coverage = _find_value(driver, coverage_labels)
    article_impressions = _find_value(driver, impressions_labels)
    ad_value_dkk = _find_value(driver, advalue_labels)

    sentiment_positive = _find_sentiment(driver, "Positive")
    sentiment_neutral = _find_sentiment(driver, "Neutral")
    sentiment_negative = _find_sentiment(driver, "Negative")

    metrics = {
        "coverage": coverage,
        "article_impressions": article_impressions,
        "ad_value_dkk": ad_value_dkk,
        "sentiment_positive": sentiment_positive,
        "sentiment_neutral": sentiment_neutral,
        "sentiment_negative": sentiment_negative,
    }

    print("[metrics] Parsed:", metrics)
    return metrics