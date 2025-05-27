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
OpusLogin = orchestrator_connection.get_credential("OpusLoginGustav")
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
options.add_argument("force-device-scale-factor=0.5")
options.add_argument("--disable-extensions")
options.add_argument("--remote-debugging-pipe")

chrome_service = Service()
max_retries = 3
driver = None

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
    WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "logonuidfield")))
    

    wait.until(EC.visibility_of_element_located((By.ID, "logonuidfield"))).send_keys(OpusUserName)
    wait.until(EC.visibility_of_element_located((By.ID, "logonpassfield"))).send_keys(OpusPassword)
    wait.until(EC.element_to_be_clickable((By.ID, "buttonLogon"))).click()

    time.sleep(10)

except Exception as e:
    orchestrator_connection.log_error(f"An error occurred: {e}")
    print(f"An error occurred: {e}")
    driver.quit()
    raise e