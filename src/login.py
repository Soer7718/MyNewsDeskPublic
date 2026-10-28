import os
from dotenv import load_dotenv
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

load_dotenv()

def login(driver):
    email = os.getenv("MYNEWS_EMAIL")
    password = os.getenv("MYNEWS_PASSWORD")

    wait = WebDriverWait(driver, 30)

    # 1) Gå til login
    driver.get("https://publish.mynewsdesk.com/users/sign_in")

    # 2) Vent på felter
    email_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='email']")))
    email_input.send_keys(email)

    password_input = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input[type='password']")))
    password_input.send_keys(password)

    password_input.send_keys(Keys.RETURN)

    # 3) Vent på at du ER logget ind:
    # Vi venter på at React loader top-navigationen
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "nav")))

    # 4) Navigér ind i workspace direkte
    driver.get("https://publish.mynewsdesk.com/127067")

    # 5) Vent på at workspace loader
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "nav")))

    # 6) Gå nu til rapport-listen
    driver.get("https://publish.mynewsdesk.com/127067/reports")

    # 7) Vent på "Alle rapporter"-overskriften
    wait.until(
        EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'Alle rapporter')]"))
    )

    print("Login + workspace navigation complete")