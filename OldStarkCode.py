        
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


        # #Tjekker om stark findes
        # try:
        #     refs_to_find_natur, natur_oldest = check_and_extract_references(df_Naturafdelingen)
        # except Exception as e:
        #     refs_to_find_natur, natur_oldest = [], None
        #     print(f"Error processing Naturafdelingen references: {e}")
        # #Tjekker om stark findes
        # try:
        #     refs_to_find_vej, vej_oldest = check_and_extract_references(df_Vejafdelingen)
        # except Exception as e:
        #     refs_to_find_vej, vej_oldest = [], None
        #     print(f"Error processing Vejafdelingen references: {e}")


        # # Combine data into a list of tuples for iteration
        # departments = [
        #     ("Naturafdelingen", refs_to_find_natur, natur_oldest,EAN_Naturafdelingen),
        #     ("Vejafdelingen", refs_to_find_vej, vej_oldest,EAN_Vejafdelingen)
        # ]

        # df_Stark_Naturafdelingen = None
        # df_Stark_Vejafdelingen = None

        # #Henter stark data hvis der er noget
        # for dept_name, refs, oldest_date, ean in departments:
        #     print(f"Processing department: {dept_name}, using EAN: {ean}")
        #     df_filtered = get_filtered_department_data(driver, wait, dept_name, refs, oldest_date, ean)
        #     if df_filtered is not None and not df_filtered.empty:
        #         if dept_name == "Naturafdelingen":
        #             df_Stark_Naturafdelingen = df_filtered
        #         elif dept_name == "Vejafdelingen":
        #             df_Stark_Vejafdelingen = df_filtered

        # print(f"\nNaturafdelingen - {len(df_Stark_Naturafdelingen)} rows:" if df_Stark_Naturafdelingen is not None else "\nNaturafdelingen - No data.")
        # if df_Stark_Naturafdelingen is not None:
        #     print(df_Stark_Naturafdelingen.head())

        # print(f"\nVejafdelingen - {len(df_Stark_Vejafdelingen)} rows:" if df_Stark_Vejafdelingen is not None else "\nVejafdelingen - No data.")
        # if df_Stark_Vejafdelingen is not None:
        #     print(df_Stark_Vejafdelingen.head())


        # register_dataframe("Stark Naturafdelingen", df_Stark_Naturafdelingen, None)
        # register_dataframe("Stark Vejafdelingen", df_Stark_Vejafdelingen, None)


                # # Collect and process Stark DataFrames separately
        # stark_dataframes = []

        # if df_Stark_Naturafdelingen is not None:
        #     stark_dataframes.append(("Stark Naturafdelingen", df_Stark_Naturafdelingen))

        # if df_Stark_Vejafdelingen is not None:
        #     stark_dataframes.append(("Stark Vejafdelingen", df_Stark_Vejafdelingen))

        # for label, df in stark_dataframes:
        #     if df_ident_all is not None:
        #         process_dataframe(label, df, df_ident_all, col_ref_navn="Købers ordrenr", col_bilagsdato="Reg.dato")