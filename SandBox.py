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
import Levenshtein


# ---- Henter Assets ----
orchestrator_connection = OrchestratorConnection("Henter Assets", os.getenv('OpenOrchestratorSQL'), os.getenv('OpenOrchestratorKey'), None)
#OpusLogin = orchestrator_connection.get_credential("OpusRobotBruger")
OpusLogin = orchestrator_connection.get_credential("OpusLoginFaktura")
OpusUserName = OpusLogin.username
OpusPassword = OpusLogin.password
OpusURL = orchestrator_connection.get_constant("OpusAdgangUrl").value
EAN_Naturafdelingen = orchestrator_connection.get_constant("EAN_Naturafdelingen").value
EAN_Vejafdelingen = orchestrator_connection.get_constant("EAN_Vejafdelingen").value
date_str = orchestrator_connection.get_constant("Bilagsdato").value
# Convert to datetime object
BilagsDato = datetime.strptime(date_str, "%d-%m-%Y")
az_pattern = re.compile(r'^AZ\d{5}$')

downloads_folder = os.path.join("C:\\Users", os.getlogin(), "Downloads")

# ---- Configure Chrome options ----
options = Options()
#options.add_argument("--headless=new")
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
        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[title='Min Økonomi']"))).click()
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

        orchestrator_connection.update_credential('OpusRobotBruger', OpusUserName, password)
        orchestrator_connection.log_info('Password changed and credential updated')
        print(password)
        time.sleep(2)

        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[title='Min Økonomi']"))).click()

    
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

    def download_excel_for_ean(ean_number: str, label: str, set_view=True):
        try:
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

            # Wrap export in try-except
            try:
                wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Eksport']"))).click()
                wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[@role='menuitem' and .//span[text()='Eksport til Excel']]"))).click()
            except TimeoutException:
                print(f"[{label}] Export button not found. Skipping this table.")
                return None, None

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

            return df, final_path

        except Exception as e:
            print(f"[{label}] Error during export: {e}")
            return None, None

    # ---- Download both files and store DataFrames ----
    df_Naturafdelingen, path_Natur = download_excel_for_ean(EAN_Naturafdelingen, "Naturafdelingen", set_view=True)
    df_Vejafdelingen, path_Vej = download_excel_for_ean(EAN_Vejafdelingen, "Vejafdelingen", set_view=False)

    dataframes = []
    excel_paths = []

    if df_Naturafdelingen is not None:
        dataframes.append(("Naturafdelingen", df_Naturafdelingen))
        if path_Natur:
            excel_paths.append(path_Natur)
    else:
        print("Naturafdelingen eksport blev sprunget over.")

    if df_Vejafdelingen is not None:
        dataframes.append(("Vejafdelingen", df_Vejafdelingen))
        if path_Vej:
            excel_paths.append(path_Vej)
    else:
        print("Vejafdelingen eksport blev sprunget over.")


    # --- SQL setup ---
    conn_str = (
        "DRIVER={ODBC Driver 17 for SQL Server};"
        "SERVER=faellessql;"
        "DATABASE=Opus;"
        "Trusted_Connection=yes"
    )
    sql_file_path = 'Get_AZIDENT_NEW.sql'

    def fetch_azident_for_ean(ean_number: str):
        with open(sql_file_path, 'r', encoding='cp1252') as file:
            sql_query = file.read()
        sql_query = sql_query.replace("'EANNummer'", f"'{ean_number}'")
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(sql_query)
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
        df_result = pd.DataFrame.from_records(rows, columns=columns)
        cursor.close()
        conn.close()
        return df_result

    ean_dict = {
        "Naturafdelingen": EAN_Naturafdelingen,
        "Vejafdelingen": EAN_Vejafdelingen
    }

    df_ident_map = {}
    for label, ean_number in ean_dict.items():
        print(f"Henter AZIdent data for {label}")
        df_ident = fetch_azident_for_ean(ean_number)
        print(f"{label}: {len(df_ident)} rækker hentet fra SQL.")
        df_ident_map[label] = df_ident

    # ---- 4. Helper to extract first name ----
    DISALLOWED_TERMS = {
        "anden", "diverse", "entreprenørenheden", "entreprenørafdelingen",
        "adm", "adm.", "økonomi", "vejafdelingen", "naturafdelingen",
        "ikke angivet", "ukendt", "leverandør", "ref.", "faktura", "x", "aarhus", "kommune"
    }

    
    def extract_first_name(text):
        if not isinstance(text, str):
            return None

        cleaned = re.sub(r"(?i)(lev\.?\s*ref\.?\s*nr\.?:?|att:)", "", text).strip()

        words = re.findall(r"[A-ZÆØÅa-zæøå]+", cleaned)
        for word in words:
            if len(word) > 1 and word.lower() not in DISALLOWED_TERMS:
                return word
        return None  


    def extract_first_from_kaldenavn(kaldenavn):
        if not isinstance(kaldenavn, str):
            return None
        parts = kaldenavn.strip().split()
        return parts[0] if parts else None
    
    #Sætter bilagsdato list 
    handled_bilagsdatoer = []
    

    # --- main processing function updated to accept df_ident ---
    def process_dataframe(label, df, df_ident):
        if df is None or df.empty:
            print(f"DataFrame for {label} is empty. Skipping.")
            return

        print(f"Processing DataFrame for {label}. Rows: {len(df)}")

        for idx, row in df.iterrows():
            reg_dato = row.get("Reg.dato")
            ref_navn = str(row.get("Ref.navn")).strip()
            faktura_nummer = row.get("Fakturabilag")
            AktueltBilagsDato = row.get("Bilagsdato")
            formatted_date = AktueltBilagsDato.strftime("%d.%m.%Y") if pd.notnull(AktueltBilagsDato) else ""

            if pd.notnull(reg_dato) and pd.to_datetime(reg_dato, errors='coerce') > BilagsDato:
                if ref_navn.lower() not in ("n/a", "nan") and re.match(r'^[\w\s\-\.,:]+$', ref_navn):
                    
                    # First: Try to find an AZ-ident directly in the full ref_navn
                    azid_search = re.search(r'AZ\d{5}', ref_navn)
                    if azid_search:
                        MedarbejderNavn = azid_search.group(0)  # Extract the full AZ-ident from anywhere in the string
                        print(f"[{label}] Found AZIdent in Ref.navn: {MedarbejderNavn}")
                        match = df_ident[df_ident["AZIdent"] == MedarbejderNavn]
                    else:
                        MedarbejderNavn = extract_first_name(ref_navn)
                        if not MedarbejderNavn:
                            continue

                        print(f"[{label}] Looking for Kaldenavn match: {MedarbejderNavn}")
                        
                        # Fuzzy match using Levenshtein on first names
                        def is_first_name_match(medarbejder_navn, kaldenavn):
                            kalde_first = extract_first_from_kaldenavn(kaldenavn)
                            #print(f"Full Kaldenavn: '{kaldenavn}', Extracted First Name: '{kalde_first}'")
                            if not kalde_first:
                                return False
                            return Levenshtein.distance(medarbejder_navn.lower(), kalde_first.lower()) <= 1


                        fuzzy_matches = df_ident[
                            df_ident["Kaldenavn"].apply(lambda x: is_first_name_match(MedarbejderNavn, x))
                        ]

                        if fuzzy_matches.empty:
                            print(f"[{label}] No fuzzy match for {MedarbejderNavn}")
                            continue

                        elif len(fuzzy_matches) == 1:
                            # One match – proceed normally
                            match = fuzzy_matches

                        else:
                            # Multiple matches – apply Levenshtein to refine
                            distances = []
                            for _, ident_row in fuzzy_matches.iterrows():
                                kaldenavn = ident_row["Kaldenavn"]
                                kalde_first = extract_first_from_kaldenavn(kaldenavn)
                                if kalde_first:
                                    distance = Levenshtein.distance(MedarbejderNavn.lower(), kalde_first.lower())
                                    distances.append((ident_row, distance))

                            min_distance = min(d[1] for d in distances)
                            best_matches = [row for row, dist in distances if dist == min_distance]

                            if len(best_matches) == 1:
                                match = pd.DataFrame([best_matches[0]])
                            else:
                                # Compare full cleaned names
                                def cleaned(s): return re.sub(r'\s+', '', s.lower()) if isinstance(s, str) else ''
                                ref_clean = cleaned(ref_navn)
                                full_name_matches = [
                                    row for row in best_matches
                                    if cleaned(row["Kaldenavn"]) == ref_clean
                                ]
                                if len(full_name_matches) == 1:
                                    match = pd.DataFrame([full_name_matches[0]])
                                else:
                                    print(f"[{label}] Ambiguous matches for '{MedarbejderNavn}', skipping invoice.")
                                    continue


                    if match.empty:
                        print(f"[{label}] Medarbejder not found in Opus: {MedarbejderNavn}")
                    else:
                        azident = match.iloc[0]["AZIdent"]
                        print(f"[{label}] AZIDENT for {MedarbejderNavn}: {azident}")
                            
                        driver.refresh()
                        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[title='Min Økonomi']"))).click()

                        try:
                            wait.until(EC.element_to_be_clickable((By.ID, "subTabIndex2"))).click()
                        except TimeoutException:
                            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Bilag og fakturaer')]"))).click()
                        
                        
                        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[title='Bilagsforespørgsel']"))).click()

                        wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "div[title='Søg andre bilag']"))).click()
                        
                        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "contentAreaFrame")))
                        
                        wait.until(EC.frame_to_be_available_and_switch_to_it((By.NAME, "Søg andre bilag")))

                        fakturabilag_input = wait.until(EC.element_to_be_clickable((By.XPATH,"//label[contains(., 'Fakturabilag')]/following::input[1]")))
                        fakturabilag_input.clear()
                        fakturabilag_input.send_keys(faktura_nummer)

                        input_azident = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(., 'Brugerid')]/ancestor::td/following-sibling::td//input[@type='text']")))
                        input_azident.clear()
                        #input_azident.send_keys(azident) # Sender ikke keys forci AZID ikke skal indgå i søgningen

                        date_input = wait.until(EC.element_to_be_clickable((By.XPATH,"//span[contains(., 'Registreringsdato')]/ancestor::td/following-sibling::td//input[@type='text']")))
                        date_input.clear()
                        date_input.send_keys(formatted_date)  # or whatever date format is expected
                        
                        #click søg: 
                        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Klik for at søge (Ctrl+F8)']"))).click()
                
                        # --- Wait and check for 'ingen data' message ---
                        time.sleep(5)
                        if driver.find_elements(By.XPATH, "//span[text()='Tabel indeholder ingen data']"):
                            print("Kunne ikke finde et bilag")
                            continue
                        
                        #Click videresend: 
                        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Videresend']"))).click()
                        time.sleep(2)

                        # --- Exit both iframes ---
                        driver.switch_to.default_content()

                        # Now switch to the popup iframe
                        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "URLSPW-0")))

                        # Continue with interaction in URLSPW-0
                        next_agent_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Næste agent')]/ancestor::td//following::input[@type='text' and contains(@class, 'lsField__input')]")))
            
                        next_agent_input.clear()
                        next_agent_input.send_keys(azident)
                        
                        textarea = wait.until(EC.element_to_be_clickable((By.XPATH, "//textarea[@title='Comments for Forwarding']")))
                        textarea.clear()
                        textarea.send_keys("Videresendt af Robot")

                        
                        #wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='OK']/ancestor::div[contains(@class, 'lsButton')]"))).click()
                        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='Annuller']/ancestor::div[contains(@class, 'lsButton')]"))).click()


                        time.sleep(2)
                    
                        # Add AktueltBilagsDato to handled list
                        if pd.notnull(AktueltBilagsDato):
                            handled_bilagsdatoer.append(AktueltBilagsDato)
            else:
                print("Datoen er overskredet")            


    # --- execute processing with correct df_ident ---
    for label, df in dataframes:
        df_ident = df_ident_map.get(label)
        process_dataframe(label, df, df_ident)

    # 6. Slet excel filer
    for file_path in excel_paths:
        try:
            if os.path.exists(file_path):
                os.remove(file_path)
                print(f"Deleted: {file_path}")
            else:
                print(f"File not found (already removed or never downloaded): {file_path}")
        except Exception as e:
            print(f"Error deleting {file_path}: {e}")


# 7. Sæt ny bilagsdato hvis der er håndterede bilag – og find max dato i begge Excel-filer
    if handled_bilagsdatoer:
        all_bilagsdatoer = []

        for label, df in dataframes:
            if 'Bilagsdato' in df.columns:
                df_dates = pd.to_datetime(df['Bilagsdato'], errors='coerce')
                valid_dates = df_dates.dropna()
                if not valid_dates.empty:
                    max_date = valid_dates.max()
                    all_bilagsdatoer.append(max_date)
                    print(f"Maks bilagsdato for {label}: {max_date.strftime('%d-%m-%Y')}")
                else:
                    print(f"Ingen gyldige 'Bilagsdato' fundet i {label}")
            else:
                print(f"'Bilagsdato' kolonne ikke fundet i {label}")

        if all_bilagsdatoer:
            ny_bilagsdato = max(all_bilagsdatoer)
            orchestrator_connection.update_constant("Bilagsdato", ny_bilagsdato.strftime("%d-%m-%Y"))
            print(f"Opdateret Bilagsdato til: {ny_bilagsdato.strftime('%d-%m-%Y')} baseret på Excel-data.")
        else:
            print("Excel-filerne indeholdt ingen gyldige bilagsdatoer – Bilagsdato ikke opdateret.")
    else:
        print("Ingen bilag blev behandlet – Bilagsdato forbliver uændret.")

  

except Exception as e:
    orchestrator_connection.log_error(f"An error occurred: {e}")
    print(f"An error occurred: {e}")
    driver.quit()
    raise e
