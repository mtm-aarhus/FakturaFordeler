import os
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
import time
import random 
import string

#   ---- Henter Assets ----
orchestrator_connection = OrchestratorConnection("Henter Assets", os.getenv('OpenOrchestratorSQL'),os.getenv('OpenOrchestratorKey'), None)
OpusLogin = orchestrator_connection.get_credential("OpusLoginFaktura")
OpusUserName = OpusLogin.username
OpusPassword = OpusLogin.password
OpusURL = orchestrator_connection.get_constant("OpusAdgangUrl").value
downloads_folder = os.path.join(os.path.expanduser("~"), "Downloads")

# Configure Chrome options
print("Initializing Chrome Driver...")

options = Options()
# options.add_argument("--headless=new")
options.add_argument("--window-size=1920,900")
options.add_argument("--start-maximized")
#options.add_argument("force-device-scale-factor=0.5")
options.add_argument("--disable-extensions")
options.add_argument("--remote-debugging-pipe")
options.add_argument("--disable-features=WebUsb")
#options.add_argument("--incognito")
options.add_experimental_option("prefs", {
    "download.default_directory": downloads_folder,
    "download.prompt_for_download": False,
    "download.directory_upgrade": True,
})
chrome_service = Service()
max_retries = 3


for attempt in range(1, max_retries + 1):
    try:
        print(f"Attempt {attempt}: Initializing ChromeDriver...")
        time.sleep(1)
        driver = webdriver.Chrome(service=chrome_service, options=options)
        break
    except Exception as e:
        print(f"Attempt {attempt} failed: {e}")
        if attempt == max_retries:
            raise
        time.sleep(1)

wait = WebDriverWait(driver, 100)
print("ChromeDriver initialized successfully.")

try:
    orchestrator_connection.log_info("Navigating to Opus login page")
    driver.get(OpusURL)
    driver.maximize_window()
    wait.until(EC.visibility_of_element_located((By.ID, "logonuidfield"))).send_keys(OpusUserName)
    wait.until(EC.visibility_of_element_located((By.ID, "logonpassfield"))).send_keys(OpusPassword)
    wait.until(EC.element_to_be_clickable((By.ID, "buttonLogon"))).click()

   
    try: 
        print("Navigerer til min økonomi")
        wait.until(EC.element_to_be_clickable((By.ID, "tabIndex3"))).click()
    except Exception as e:      
        orchestrator_connection.log_info(f'Fejl ved at finde knap, {e}')
        orchestrator_connection.log_info('Trying to find change button')
        
        wait.until(EC.element_to_be_clickable((By.ID, "changeButton"))).click()
        
        lower = string.ascii_lowercase
        upper = string.ascii_uppercase
        digits = string.digits
        special = "!@#&%"

        password_chars = []
        password_chars += random.choices(lower, k=2)
        password_chars += random.choices(upper, k=2)
        password_chars += random.choices(digits, k=4)
        password_chars += random.choices(special, k=2)

        random.shuffle(password_chars)
        password = ''.join(password_chars)

        wait.until(EC.visibility_of_element_located((By.ID, "inputUsername"))).send_keys(OpusPassword)
        wait.until(EC.visibility_of_element_located((By.NAME, "j_sap_password"))).send_keys(password)
        wait.until(EC.visibility_of_element_located((By.NAME, "j_sap_again"))).send_keys(password)
        wait.until(EC.element_to_be_clickable((By.ID, "changeButton"))).click()

        #Opdaterer credentials i OpenOrchestrator
        orchestrator_connection.update_credential('OpusLoginFaktura', OpusUserName, password)
        orchestrator_connection.log_info('Password changed and credential updated')
        time.sleep(2)

        #Forsøger at trykke knappen igen: 
        wait.until(EC.element_to_be_clickable((By.ID, "tabIndex3"))).click()
    
    try:
        print("Trying to click using ID...")
        wait.until(EC.element_to_be_clickable((By.ID, "subTabIndex2"))).click()
    except TimeoutException:
        print("ID failed, trying fallback XPath...")
        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Bilag og fakturaer')]"))).click()
    

    # STEP 1: Wait and switch to the outer iframe
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "contentAreaFrame")))
    print("Switched to outer iframe: contentAreaFrame")

    # STEP 2: Now inside the first iframe, wait and switch to the inner iframe
    wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Bilagsindbakke")))
    print("Switched to inner iframe: Bilagsindbakke")

    ean_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@title='EAN Nr']")))
    ean_input.clear()
    ean_input.send_keys("5798005770220")
    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Fremfinder de bilag, som opfylder dine søgekriterier (Ctrl+F8)']"))).click()

    time.sleep(10)

    wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@title='Load view']"))).click()
    time.sleep(3)
    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='lsListbox__value' and normalize-space(text())='Fuld view']"))).click()

    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Eksport' and contains(@class, 'lsButton--popupmenu')]"))).click()
    wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[@role='menuitem' and .//span[text()='Eksport til Excel']]"))).click()



    time.sleep(10)


   

except Exception as e:
    orchestrator_connection.log_error(f"An error occurred: {e}")
    print(f"An error occurred: {e}")
    driver.quit()
    raise e