import os
import time
import random
import string
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import TimeoutException
from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
import pyodbc
import re

# ---- Henter Assets ----
orchestrator_connection = OrchestratorConnection("Henter Assets", os.getenv('OpenOrchestratorSQL'), os.getenv('OpenOrchestratorKey'), None)
OpusLogin = orchestrator_connection.get_credential("OpusLoginFaktura")
OpusUserName = OpusLogin.username
OpusPassword = OpusLogin.password
OpusURL = orchestrator_connection.get_constant("OpusAdgangUrl").value
EAN_Naturafdelingen = orchestrator_connection.get_constant("EAN_Naturafdelingen").value
EAN_Vejafdelingen = orchestrator_connection.get_constant("EAN_Vejafdelingen").value
BilagsDato = "05-06-2025"

# Original date string in DD-MM-YYYY format
date_str = "03-06-2025"

# Convert to datetime object
BilagsDato = datetime.strptime(date_str, "%d-%m-%Y")

# Format it as a timestamp string
#timestamp_str = date_obj.strftime("%Y-%m-%d %H:%M:%S")

downloads_folder = os.path.join("C:\\Users", os.getlogin(), "Downloads")

# ---- Configure Chrome options ----
options = Options()
options.add_argument("--headless=new")
options.add_argument("--window-size=1920,900")
options.add_argument("--start-maximized")
options.add_argument("--disable-extensions")
options.add_argument("--remote-debugging-pipe")
options.add_argument("--disable-features=WebUsb")

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
        password_chars = random.choices(lower, k=2) + random.choices(upper, k=2) + random.choices(digits, k=4) + random.choices(special, k=2)
        random.shuffle(password_chars)
        password = ''.join(password_chars)

        wait.until(EC.visibility_of_element_located((By.ID, "inputUsername"))).send_keys(OpusPassword)
        wait.until(EC.visibility_of_element_located((By.NAME, "j_sap_password"))).send_keys(password)
        wait.until(EC.visibility_of_element_located((By.NAME, "j_sap_again"))).send_keys(password)
        wait.until(EC.element_to_be_clickable((By.ID, "changeButton"))).click()

        orchestrator_connection.update_credential('OpusLoginFaktura', OpusUserName, password)
        orchestrator_connection.log_info('Password changed and credential updated')
        time.sleep(2)

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

    # ---- Function to download and rename Excel file ----
    def download_excel_for_ean(ean_number: str, label: str, set_view=True):

        ean_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@title='EAN Nr']")))
        ean_input.clear()
        ean_input.send_keys(ean_number)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Fremfinder de bilag, som opfylder dine søgekriterier (Ctrl+F8)']"))).click()
        time.sleep(3)

        if set_view:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@title='Load view']"))).click()
            time.sleep(2)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='lsListbox__value' and normalize-space(text())='Fuld view']"))).click()
        time.sleep(2)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Eksport']"))).click()
        wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[@role='menuitem' and .//span[text()='Eksport til Excel']]"))).click()

        before_files = set(os.listdir(downloads_folder))
        timeout = 90
        polling_interval = 1
        elapsed = 0
        downloaded_file_path = None

        print(f"Waiting for download of {label}...")

        while elapsed < timeout:
            time.sleep(polling_interval)
            after_files = set(os.listdir(downloads_folder))
            new_files = {f for f in (after_files - before_files) if f.lower().endswith(".xlsx")}
            if new_files:
                newest_file = max(
                    new_files,
                    key=lambda fn: os.path.getmtime(os.path.join(downloads_folder, fn))
                )
                downloaded_file_path = os.path.join(downloads_folder, newest_file)
                try:
                    df = pd.read_excel(downloaded_file_path, engine='openpyxl')
                    break
                except Exception as e:
                    print(f"Waiting for file to unlock: {e}")
            elapsed += polling_interval

        if not downloaded_file_path or not os.path.exists(downloaded_file_path):
            raise FileNotFoundError(f"Download failed or file locked for {label}")

        today_str = datetime.today().strftime("%d.%m.%Y")
        final_name = f"Bilagsliste_{label}_{today_str}.xlsx"
        final_path = os.path.join(downloads_folder, final_name)
        if os.path.exists(final_path):
            os.remove(final_path)
        os.rename(downloaded_file_path, final_path)
        print(f"Downloaded and saved as: {final_path}")

        return df  # <-- Return DataFrame to use later

    # ---- Download both files and store DataFrames ----
    df_Naturafdelingen = download_excel_for_ean("5798005770220", "Naturafdelingen", set_view=True) #omkostni
    df_Vejafdelingen = download_excel_for_ean("5798005770213", "Vejafdelingen", set_view=False)

    # ---- 1. Setup and Read SQL from file ----
    conn_str = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=faellessql;"
        "DATABASE=Opus;"
        "Trusted_Connection=yes"
    )

    sql_file_path = 'GET_AZIDENT.sql'

    with open(sql_file_path, 'r', encoding='cp1252') as file:
        sql_query = file.read()

    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    cursor.execute(sql_query)

    columns = [col[0] for col in cursor.description]
    rows = cursor.fetchall()
    df_ident = pd.DataFrame.from_records(rows, columns=columns)

    cursor.close()
    conn.close()

    print("✅ SQL data loaded. Rows:", len(df_ident))

    # ---- 3. Set BilagsDato cutoff ----
    BilagsDato = datetime.strptime("03-06-2025", "%d-%m-%Y")

    # ---- 4. Helper to extract first name ----
    DISALLOWED_TERMS = {
        "anden", "diverse", "entreprenørenheden", "entreprenørafdelingen",
        "adm", "adm.", "økonomi", "vejafdelingen", "naturafdelingen",
        "ikke angivet", "ukendt", "leverandør", "ref.", "faktura", "x"
    }

    def extract_first_name(text):
        if not isinstance(text, str):
            return None

        cleaned = re.sub(r"(?i)(lev\.?\s*ref\.?\s*nr\.?:?|att:)", "", text).strip()
        match = re.match(r"([A-ZÆØÅa-zæøå]+)", cleaned)
        if match:
            name_candidate = match.group(1)
            if len(name_candidate) > 1 and name_candidate.lower() not in DISALLOWED_TERMS:
                return name_candidate
        return None

    # ---- 5. Loop over both DataFrames ----
    dataframes = [
        ("Naturafdelingen", df_Naturafdelingen),
        ("Vejafdelingen", df_Vejafdelingen)
    ]

    def process_dataframe(label, df):
        if df.empty:
            print(f"DataFrame for {label} is empty. Skipping.")
            return

        print(f"🔍 Processing DataFrame for {label}. Rows: {len(df)}")

        for idx, row in df.iterrows():
            reg_dato = row.get("Reg.dato")
            ref_navn = str(row.get("Ref.navn")).strip()
            faktura_nummer = row.get("Fakturabilag")
")

            if pd.notnull(reg_dato) and pd.to_datetime(reg_dato, errors='coerce') > BilagsDato:
                if ref_navn.lower() not in ("n/a", "nan") and re.match(r'^[\w\s\-\.,:]+$', ref_navn):
                    MedarbejderNavn = extract_first_name(ref_navn)
                    if not MedarbejderNavn:
                        continue

                    print(f"[{label}] Looking for: {MedarbejderNavn}")
                    match = df_ident[df_ident["Fornavn"].str.lower() == MedarbejderNavn.lower()]
                    if match.empty:
                        match = df_ident[df_ident["Fornavn"].str.contains(MedarbejderNavn, case=False, na=False)]

                    if match.empty:
                        print(f"[{label}] Medarbejder not found in Opus: {MedarbejderNavn}")
                    else:
                        azident = match.iloc[0]["Ident"]
                        print(f"[{label}] AZIDENT for {MedarbejderNavn}: {azident}")
                        wait.until(EC.element_to_be_clickable((By.XPATH,"//div[.//span[text()='Bilagsforespørgsel']]"))).click()
                        wait.until(EC.element_to_be_clickable((By.XPATH,"//div[.//span[text()='Søg omposteringsbilag']]"))).click()
                        
                        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "contentAreaFrame")))
                        print("Switched to outer iframe: contentAreaFrame")
                     
                        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Søg andre bilag")))
                        print("Switched to inner iframe: Søg andre bilag")
                        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Fakturabilagsnummer']"))).send_keys(faktura_nummer)

                        




    # Execute for both
    for label, df in dataframes:
        process_dataframe(label, df)


    #6. hvis der er et match så videresend bilag via brugergrænsefladen
    # 7. sæt ny bilagsdato, som robotten kan hente ved næste kørsel



except Exception as e:
    orchestrator_connection.log_error(f"An error occurred: {e}")
    print(f"An error occurred: {e}")
    driver.quit()
    raise e
