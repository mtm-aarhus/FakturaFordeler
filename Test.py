import os
import pandas as pd
from datetime import datetime
import re
import pyodbc  # Optional: Only if you want to include the DB lookup


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

# ---- 2. Load Excel File ----
file_name = "Bilagsliste_Naturafdelingen_10.06.2025.xlsx"
downloads_folder = os.path.join("C:\\Users", os.getlogin(), "Downloads")
file_path = os.path.join(downloads_folder, file_name)

df_Naturafdelingen = pd.read_excel(file_path, engine="openpyxl")
print("✅ Excel file loaded. Rows:", len(df_Naturafdelingen))

# ---- 3. Set BilagsDato cutoff ----
BilagsDato = datetime.strptime("03-06-2025", "%d-%m-%Y")



# ---- 4. Helper to extract first name ----
# ---- Disallowed non-name values ----
DISALLOWED_TERMS = {
    "anden", "diverse", "entreprenørenheden", "entreprenørafdelingen",
    "adm", "adm.", "økonomi", "vejafdelingen", "naturafdelingen",
    "ikke angivet", "ukendt", "leverandør", "ref.", "faktura", "x"
}

# ---- Extract cleaned name from text ----
def extract_first_name(text):
    if not isinstance(text, str):
        return None

    # Remove known prefixes
    cleaned = re.sub(r"(?i)(lev\.?\s*ref\.?\s*nr\.?:?|att:)", "", text)
    cleaned = cleaned.strip()

    # Extract first word
    match = re.match(r"([A-ZÆØÅa-zæøå]+)", cleaned)
    if match:
        name_candidate = match.group(1)
        if len(name_candidate) > 1 and name_candidate.lower() not in DISALLOWED_TERMS:
            return name_candidate
    return None

# ---- Process each bilag row ----
for idx, row in df_Naturafdelingen.iterrows():
    reg_dato = row.get("Reg.dato")
    ref_navn = str(row.get("Ref.navn")).strip()

    # Check if Reg.dato is a valid date and greater than Oldbilagsdato
    if pd.notnull(reg_dato) and pd.to_datetime(reg_dato) > BilagsDato:
        ref_navn = str(row.get("Ref.navn")).strip()

    # Allow ".", "-", ",", and space in names
    if ref_navn.lower() not in ("n/a", "nan") and re.match(r'^[\w\s\-\.,:]+$', ref_navn):
        MedarbejderNavn = extract_first_name(ref_navn)


        if not MedarbejderNavn:
            continue

        print(f"Looking for: {MedarbejderNavn}")

        # First: exact match
        match = df_ident[df_ident["Fornavn"].str.lower() == MedarbejderNavn.lower()]

        # Fallback: partial match
        if match.empty:
            match = df_ident[df_ident["Fornavn"].str.contains(MedarbejderNavn, case=False, na=False)]

        if match.empty:
            print("Medarbejder not found in Opus")
        else:
            azident = match.iloc[0]["Ident"]
            print("AZIDENT:", azident)