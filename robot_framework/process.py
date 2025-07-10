"""This module contains the main process of the robot."""

from OpenOrchestrator.orchestrator_connection.connection import OrchestratorConnection
from OpenOrchestrator.database.queues import QueueElement
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
from selenium.common.exceptions import StaleElementReferenceException
import pyodbc
import re
import Levenshtein

# pylint: disable-next=unused-argument
def process(orchestrator_connection: OrchestratorConnection, queue_element: QueueElement | None = None) -> None:
    orchestrator_connection.log_trace("Running process.")
    
    # === 2. Global Configuration ===
    az_pattern = re.compile(r'^AZ\d{5}$')
    downloads_folder = os.path.join("C:\\Users", os.getlogin(), "Downloads")
    DISALLOWED_TERMS = {
        "anden", "diverse", "entreprenørenheden", "entreprenørafdelingen",
        "adm", "adm.", "økonomi", "vejafdelingen", "naturafdelingen",
        "ikke angivet", "ukendt", "leverandør", "ref.", "faktura", "x", "aarhus", "kommune"
    }
    max_retries = 3

    OpusLogin = orchestrator_connection.get_credential("OpusRobotBruger")
    #OpusLogin = orchestrator_connection.get_credential("OpusLoginFaktura")
    OpusUserName = OpusLogin.username
    OpusPassword = OpusLogin.password
    OpusURL = orchestrator_connection.get_constant("OpusAdgangUrl").value
    EAN_Naturafdelingen = orchestrator_connection.get_constant("EAN_Naturafdelingen").value
    EAN_Vejafdelingen = orchestrator_connection.get_constant("EAN_Vejafdelingen").value
    date_str = orchestrator_connection.get_constant("Bilagsdato").value
    BilagsDato = datetime.strptime(date_str, "%d-%m-%Y")


    # === 4. Setup Chrome Driver ===
    options = Options()
    #options.add_argument("--headless=new")
    options.add_argument("--window-size=1920,900")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-extensions")
    options.add_argument("--remote-debugging-pipe")
    options.add_argument("--disable-features=WebUsb")
    options.add_argument("--log-level=3")

    chrome_service = Service()
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


    # === 5. Initialize Working Variables ===
    dataframes = []
    excel_paths = []
    handled_bilagsdatoer = []

    def safe_click(wait, by, value, retries=3):
        for attempt in range(retries):
            try:
                elem = wait.until(EC.element_to_be_clickable((by, value)))
                elem.click()
                return True
            except StaleElementReferenceException:
                print(f"Stale element encountered. Retrying ({attempt + 1}/{retries})...")
                time.sleep(1)
        print("Failed to click element after retries.")
        return False



    # === 6. Define Utility and Processing Functions ===
    def download_excel_for_ean(ean_number: str, label: str, set_view=True):
        try:
            time.sleep(2)
            ean_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@title='EAN Nr']")))
            ean_input.clear()
            print(f"Indtaster ean: {ean_number}")
            # Wait until the input is really empty (not auto-filled again)
            wait.until(lambda driver: ean_input.get_attribute("value") == "")
            ean_input.send_keys(ean_number)
            time.sleep(3)
            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Fremfinder de bilag, som opfylder dine søgekriterier (Ctrl+F8)']"))).click()
            time.sleep(3)


            if set_view:
                # Find the "Load view" input
                current_view = wait.until(EC.visibility_of_element_located((By.XPATH, "//input[@title='Load view']")))
                # Read its current value
                current_view_value = current_view.get_attribute("value").strip()
                print("Current view:", current_view_value)
            
                if current_view_value != "Fuld view":
                    # Only switch if not already set
                    current_view.click()
                    time.sleep(2)
            
                    # Wait for and click "Fuld view"
                    wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='lsListbox__value' and normalize-space(text())='Fuld view']"))).click()
                    time.sleep(2)
                else:
                    print("Fuld view already active, skipping view change.")
            
            # Always wait a little to let the view load
            time.sleep(2)




            

            # Wrap export in try-except
            try:
                if not safe_click(wait, By.XPATH, "//div[@title='Eksport']"):
                    raise TimeoutException("Failed to click 'Eksport'")
                if not safe_click(wait, By.XPATH, "//tr[@role='menuitem' and .//span[text()='Eksport til Excel']]"):
                    raise TimeoutException("Failed to click 'Eksport til Excel'")
            except TimeoutException:
                print(f"[{label}] Export button not found or stale. Skipping this table.")
                return None, None

            before_files = set(os.listdir(downloads_folder))
            timeout = 90
            polling_interval = 1
            elapsed = 0
            downloaded_file_path = None

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

                    # Check if a .crdownload file exists with the same name prefix
                    partial_file = downloaded_file_path + ".crdownload"
                    if not os.path.exists(partial_file):
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
        
    def check_and_extract_references(df):
        kreditor_value = "55828415"

        # Safety: Ensure df is not None and not empty
        if df is None or df.empty:
            print("DataFrame is None or empty.")
            return [], None

        # Safety: Ensure required columns exist
        required_cols = ["Kreditornr.", "Bilagsdato", "Fakturanr./Reference."]
        for col in required_cols:
            if col not in df.columns:
                print(f"Required column missing: {col}")
                return [], None

        # Filter by kreditor
        filtered_df = df[df["Kreditornr."].astype(str) == kreditor_value].copy()
        if filtered_df.empty:
            return [], None

        # Parse dates
        filtered_df["Bilagsdato"] = pd.to_datetime(filtered_df["Bilagsdato"], errors='coerce')
        filtered_df = filtered_df.dropna(subset=["Bilagsdato"])
        if filtered_df.empty:
            print("No valid 'Bilagsdato' for filtered rows.")
            return [], None

        oldest_date = filtered_df["Bilagsdato"].min()
        faktura_refs = filtered_df["Fakturanr./Reference."].dropna().astype(str).tolist()
        
        print(f"Extracted Faktura references: {faktura_refs}")

        return faktura_refs, oldest_date

    def get_filtered_department_data(driver, wait, dept_name, refs, oldest_date, ean):
        if not refs:
            return None
        driver.switch_to.default_content()
        print(f"[{dept_name}] Refreshing and navigating to search page...")
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

        # Fill form
        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(., 'Brugerid')]/ancestor::td/following-sibling::td//input[@type='text']"))).clear()

        date_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(., 'Registreringsdato')]/ancestor::td/following-sibling::td//input[@type='text']")))
        date_input.clear()
        date_input.send_keys(oldest_date.strftime("%d.%m.%Y"))

        ean_input = wait.until(EC.element_to_be_clickable((By.XPATH,"//label[.//span[text()='EAN Nummer']]/following::input[1]")))
        ean_input.clear()
        ean_input.send_keys(ean)

        kreditor_input = wait.until(EC.element_to_be_clickable((By.XPATH,"//label[.//span[contains(text(), 'Kreditor')]]/following::input[1]")))
        kreditor_input.clear()
        kreditor_input.send_keys("55828415")  # Update if other creditors should be supported

        # Click search
        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Klik for at søge (Ctrl+F8)']"))).click()
        time.sleep(2)

        # Load view
        wait.until(EC.element_to_be_clickable((By.XPATH, "//input[@title='Load view']"))).click()
        time.sleep(2)
        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@class='lsListbox__value' and normalize-space(text())='Fuld view']"))).click()
        time.sleep(2)

        # Export to Excel
        try:
            wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Eksport']"))).click()
            wait.until(EC.element_to_be_clickable((By.XPATH, "//tr[@role='menuitem' and .//span[text()='Eksport til Excel']]"))).click()
        except TimeoutException:
            print(f"[{dept_name}] Export button not found.")
            return None

        # Detect and read downloaded Excel
        timeout = 60
        polling_interval = 1
        downloaded_file_path = None
        before_files = set(os.listdir(downloads_folder))

        for _ in range(timeout):
            time.sleep(polling_interval)
            after_files = set(os.listdir(downloads_folder))
            new_files = {f for f in (after_files - before_files) if f.lower().endswith(".xlsx")}
            if new_files:
                newest_file = max(new_files, key=lambda fn: os.path.getmtime(os.path.join(downloads_folder, fn)))
                downloaded_file_path = os.path.join(downloads_folder, newest_file)
                break

        if not downloaded_file_path or not os.path.exists(downloaded_file_path):
            print(f"[{dept_name}] Download failed or file not found.")
            return None

        df_downloaded = pd.read_excel(downloaded_file_path, engine="openpyxl")
        os.remove(downloaded_file_path)  # Optional: clean up after read

        # Determine correct reference column
        if "Reference" in df_downloaded.columns:
            ref_col = "Reference"  # For Stark files
        elif "Fakturanr./Reference." in df_downloaded.columns:
            ref_col = "Fakturanr./Reference."  # For original files
        else:
            print(f"[{dept_name}] No known reference column found in DataFrame.")
            return None

        # Filter based on references
        df_filtered = filter_rows_by_references(df_downloaded, refs, column_name=ref_col)

        if df_filtered.empty:
            print(f"[{dept_name}] No matching rows found after filtering.")
        else:
            print(f"[{dept_name}] Filtered down to {len(df_filtered)} rows.")

        return df_filtered

    def register_dataframe(label, df, path):
        if df is not None:
            dataframes.append((label, df))
            if path:
                excel_paths.append(path)
        else:
            print(f"{label} eksport blev sprunget over.")

    def fetch_combined_azident(ean_vej: str, ean_natur: str):
        with open(sql_file_path, 'r', encoding='cp1252') as file:
            sql_query = file.read()

        sql_query = sql_query.replace("'EANVEJ'", f"'{ean_vej}'")
        sql_query = sql_query.replace("'EANNATUR'", f"'{ean_natur}'")

        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        cursor.execute(sql_query)
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
        df_result = pd.DataFrame.from_records(rows, columns=columns)
        cursor.close()
        conn.close()
        return df_result

    def process_dataframe(label, df, df_ident, col_ref_navn="Ref.navn", col_bilagsdato="Reg.dato"):
        if df is None or df.empty:
            print(f"DataFrame for {label} is empty. Skipping.")
            return

        print(f"Processing DataFrame for {label}. Rows: {len(df)}")

        for idx, row in df.iterrows():
            reg_dato = row.get("Reg.dato")
            ref_navn = str(row.get(col_ref_navn)).strip()
            faktura_nummer = row.get("Fakturabilag")
            AktueltBilagsDato = row.get(col_bilagsdato)
            formatted_date = AktueltBilagsDato.strftime("%d.%m.%Y") if pd.notnull(AktueltBilagsDato) else ""

            if pd.notnull(reg_dato) and pd.to_datetime(reg_dato, errors='coerce') > BilagsDato:
                if ref_navn.lower() not in ("n/a", "nan") and re.match(r'^[\w\s\-\.,:]+$', ref_navn):
                    
                    azid_search = re.search(r'AZ\d{5}', ref_navn)
                    if azid_search:
                        MedarbejderNavn = azid_search.group(0)
                        print(f"[{label}] Found AZIdent in Ref.navn: {MedarbejderNavn}")
                        match = df_ident[df_ident["AZIdent"] == MedarbejderNavn]
                    else:
                        MedarbejderNavn = extract_first_name(ref_navn)
                        if not MedarbejderNavn:
                            continue

                        print(f"[{label}] Looking for Kaldenavn match: {MedarbejderNavn}")
                        
                        def is_first_name_match(medarbejder_navn, kaldenavn):
                            kalde_first = extract_first_from_kaldenavn(kaldenavn)
                            if not kalde_first:
                                return False
                            dist = Levenshtein.distance(medarbejder_navn.lower(), kalde_first.lower())
                            return dist <= 2

                        fuzzy_matches = df_ident[
                            df_ident["Kaldenavn"].apply(lambda x: is_first_name_match(MedarbejderNavn, x))
                        ]

                        if fuzzy_matches.empty:
                            print(f"[{label}] No match for {MedarbejderNavn}")
                            continue
                        elif len(fuzzy_matches) == 1:
                            match = fuzzy_matches
                        else:
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

                        date_input = wait.until(EC.element_to_be_clickable((By.XPATH,"//span[contains(., 'Registreringsdato')]/ancestor::td/following-sibling::td//input[@type='text']")))
                        date_input.clear()
                        date_input.send_keys(formatted_date)

                        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Klik for at søge (Ctrl+F8)']"))).click()

                        time.sleep(5)
                        if driver.find_elements(By.XPATH, "//span[text()='Tabel indeholder ingen data']"):
                            print("Kunne ikke finde et bilag")
                            continue
                        
                        wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@title='Videresend']"))).click()
                        time.sleep(2)
                        driver.switch_to.default_content()
                        wait.until(EC.frame_to_be_available_and_switch_to_it((By.ID, "URLSPW-0")))

                        next_agent_input = wait.until(EC.element_to_be_clickable((By.XPATH, "//span[contains(text(), 'Næste agent')]/ancestor::td//following::input[@type='text']")))
                        next_agent_input.clear()
                        next_agent_input.send_keys(azident)

                        textarea = wait.until(EC.element_to_be_clickable((By.XPATH, "//textarea[@title='Comments for Forwarding']")))
                        textarea.clear()
                        textarea.send_keys("Videresendt af Robot")

                        wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='Annuller']/ancestor::div[contains(@class, 'lsButton')]"))).click()

                        #Mangler at trykke send
                        #wait.until(EC.element_to_be_clickable((By.XPATH, "//span[normalize-space(text())='OK']/ancestor::div[contains(@class, 'lsButton')]"))).click()

                        time.sleep(2)
                        if pd.notnull(AktueltBilagsDato):
                            handled_bilagsdatoer.append(AktueltBilagsDato)
            else:
                print("Datoen er overskredet")

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

    def filter_rows_by_references(df, ref_list, column_name="Fakturanr./Reference."):
        if column_name not in df.columns:
            print(f"'{column_name}' column not found.")
            return pd.DataFrame()

        filtered = df[df[column_name].astype(str).isin(ref_list)]
        print(f"Found {len(filtered)} matching rows based on references in '{column_name}'.")
        return filtered


    # === 7. Main Automation Flow ===
    try:
        orchestrator_connection.log_info("Navigating to Opus login page")
        driver.get(OpusURL)
        driver.maximize_window()

        wait.until(EC.visibility_of_element_located((By.ID, "logonuidfield"))).send_keys(OpusUserName)
        wait.until(EC.visibility_of_element_located((By.ID, "logonpassfield"))).send_keys(OpusPassword)
        wait.until(EC.element_to_be_clickable((By.ID, "buttonLogon"))).click()

        #Login
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


        # ---- Download both files and store DataFrames ----
        df_Naturafdelingen, path_Natur = download_excel_for_ean(EAN_Naturafdelingen, "Naturafdelingen", set_view=True)
        df_Vejafdelingen, path_Vej = download_excel_for_ean(EAN_Vejafdelingen, "Vejafdelingen", set_view=True)


        try:
            refs_to_find_natur, natur_oldest = check_and_extract_references(df_Naturafdelingen)
        except Exception as e:
            refs_to_find_natur, natur_oldest = [], None
            print(f"Error processing Naturafdelingen references: {e}")

        try:
            refs_to_find_vej, vej_oldest = check_and_extract_references(df_Vejafdelingen)
        except Exception as e:
            refs_to_find_vej, vej_oldest = [], None
            print(f"Error processing Vejafdelingen references: {e}")


        # Combine data into a list of tuples for iteration
        departments = [
            ("Naturafdelingen", refs_to_find_natur, natur_oldest,EAN_Naturafdelingen),
            ("Vejafdelingen", refs_to_find_vej, vej_oldest,EAN_Vejafdelingen)
        ]

        df_Stark_Naturafdelingen = None
        df_Stark_Vejafdelingen = None

        for dept_name, refs, oldest_date, ean in departments:
            df_filtered = get_filtered_department_data(driver, wait, dept_name, refs, oldest_date, ean)
            if df_filtered is not None and not df_filtered.empty:
                if dept_name == "Naturafdelingen":
                    df_Stark_Naturafdelingen = df_filtered
                elif dept_name == "Vejafdelingen":
                    df_Stark_Vejafdelingen = df_filtered

        print(f"\nNaturafdelingen - {len(df_Stark_Naturafdelingen)} rows:" if df_Stark_Naturafdelingen is not None else "\nNaturafdelingen - No data.")
        if df_Stark_Naturafdelingen is not None:
            print(df_Stark_Naturafdelingen.head())

        print(f"\nVejafdelingen - {len(df_Stark_Vejafdelingen)} rows:" if df_Stark_Vejafdelingen is not None else "\nVejafdelingen - No data.")
        if df_Stark_Vejafdelingen is not None:
            print(df_Stark_Vejafdelingen.head())
        
        # Register both raw and filtered Stark data
        register_dataframe("Naturafdelingen", df_Naturafdelingen, path_Natur)
        register_dataframe("Vejafdelingen", df_Vejafdelingen, path_Vej)

        register_dataframe("Stark Naturafdelingen", df_Stark_Naturafdelingen, None)
        register_dataframe("Stark Vejafdelingen", df_Stark_Vejafdelingen, None)


        # Fetch AZIdent data
        conn_str = (
            "DRIVER={ODBC Driver 17 for SQL Server};"
            "SERVER=faellessql;"
            "DATABASE=Opus;"
            "Trusted_Connection=yes"
        )
        sql_file_path = 'Get_AZIDENT_NEW.sql'

        # Only fetch SQL data if any of the Excel dataframes have rows
        has_dataframes_with_rows = any(df is not None and not df.empty for _, df in dataframes)

        df_ident_all = None
        if has_dataframes_with_rows:
            df_ident_all = fetch_combined_azident(EAN_Vejafdelingen, EAN_Naturafdelingen)
            print(f"{len(df_ident_all)} rækker hentet fra SQL.")
        else:
            print("Ingen data i nogen af Excel-filerne – springer SQL-opslag over.")
        


        # Process standard (non-Stark) departments
        standard_labels = ["Naturafdelingen", "Vejafdelingen"]

        for label, df in dataframes:
            if label in standard_labels and df_ident_all is not None:
                process_dataframe(label, df, df_ident_all)

        # Collect and process Stark DataFrames separately
        stark_dataframes = []

        if df_Stark_Naturafdelingen is not None:
            stark_dataframes.append(("Stark Naturafdelingen", df_Stark_Naturafdelingen))

        if df_Stark_Vejafdelingen is not None:
            stark_dataframes.append(("Stark Vejafdelingen", df_Stark_Vejafdelingen))

        for label, df in stark_dataframes:
            if df_ident_all is not None:
                process_dataframe(label, df, df_ident_all, col_ref_navn="Købers ordrenr", col_bilagsdato="Reg.dato")

        #Cleanup
        for file_path in excel_paths:
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                else:
                    print(f"File not found (already removed or never downloaded): {file_path}")
            except Exception as e:
                print(f"Error deleting {file_path}: {e}")

        # Update Bilagsdato
        if handled_bilagsdatoer:
            all_bilagsdatoer = []

            for label, df in dataframes:
                if 'Reg.dato' in df.columns:
                    df_dates = pd.to_datetime(df['Reg.dato'], errors='coerce')
                    valid_dates = df_dates.dropna()
                    if not valid_dates.empty:
                        max_date = valid_dates.max()
                        all_bilagsdatoer.append(max_date)
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






