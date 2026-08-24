import os
import glob
import pandas as pd

# File Paths & Configuration
MASTER_DB_PATH = "VAPAHCS CPET DB.xlsx"
MASTER_SHEET_NAME = "baseline (new)"
BXB_FILES_DIRECTORY = "./CPET results (new)"  # Folder containing breath-by-breath Excel files


def extract_header_metadata(df_raw):
    """Scans the header section (rows 0-14) to extract patient demographics and test details."""
    key_mapping = {
        "ID1": "Study ID #",
        "Last Name": "Last Name",
        "First Name": "First Name",
        "Gender": "Gender",
        "Age": "Age at ETT",
        "Height (in)": "Height\n (in.)",
        "Weight (lbs)": "Weight \n(lbs.)",
        "Test date": "Date of ETT",
        "Reason for Stopping Test": "Angina (Ex or R) reason for stopping?",
        "Reason for Test": "Angina (Ex or R) reason for stopping?",
    }

    metadata = {}
    for r in range(min(15, len(df_raw))):
        for c in range(df_raw.shape[1]):
            val = df_raw.iloc[r, c]
            if pd.notna(val):
                key = str(val).strip()
                if key in key_mapping:
                    target_col = key_mapping[key]
                    if target_col in metadata:
                        continue  # Avoid overwriting primary reason for stopping

                    # Look right in the row for the associated value
                    for nc in range(c + 1, df_raw.shape[1]):
                        next_val = df_raw.iloc[r, nc]
                        if pd.notna(next_val):
                            next_str = str(next_val).strip()
                            if next_str in key_mapping:
                                break  # Stop if scanning into another metadata header label
                            if next_str != "":
                                metadata[target_col] = next_val
                                break
    return metadata


def process_cpet_file(file_path):
    """Parses a single CPET BxB Excel file and isolates peak exercise metrics."""
    xl = pd.ExcelFile(file_path)
    sheet_name = "Data" if "Data" in xl.sheet_names else xl.sheet_names[0]
    df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    # 1. Extract Metadata Demographics
    meta = extract_header_metadata(df_raw)

    # Combine First/Last Name into "First Last" format
    first_name = str(meta.pop("First Name", "")).strip()
    last_name = str(meta.pop("Last Name", "")).strip()
    full_name = f"{first_name} {last_name}".strip()

    # Format Date
    date_val = meta.get("Date of ETT")
    formatted_date = (
        pd.to_datetime(date_val).strftime("%Y-%m-%d")
        if pd.notna(date_val)
        else None
    )

    # 2. Extract Breath-by-Breath Table Data (Header at Row 0, Col 9+)
    table_headers = list(df_raw.iloc[0, 9:].values)
    df_table = df_raw.iloc[3:, 9:].copy()
    df_table.columns = table_headers

    # 3. Locate Peak VO2 Row Safely
    vo2_col = "VO2/kg" if "VO2/kg" in df_table.columns else "VO2"
    df_table[vo2_col] = pd.to_numeric(df_table[vo2_col], errors="coerce")
    df_table = df_table.dropna(subset=[vo2_col]).reset_index(drop=True)

    max_idx = df_table[vo2_col].idxmax()
    max_row = df_table.iloc[max_idx].to_dict()

    # 4. Map Extracted Data
    extracted_data = {
        "Name": full_name if full_name else None,
        "Study ID #": meta.get("Study ID #"),
        "Date of ETT": formatted_date,
        "Gender": meta.get("Gender"),
        "Age at ETT": meta.get("Age at ETT"),
        "Height\n (in.)": meta.get("Height\n (in.)"),
        "Weight \n(lbs.)": meta.get("Weight \n(lbs.)"),
        "VO2 ml/Kg/min (max)": max_row.get("VO2/kg"),
        "VO2 in ml (max)": max_row.get("VO2"),
        "VCO2 (max)": max_row.get("VCO2"),
        "VE (max)": max_row.get("VE"),
        "RER max": max_row.get("RQ")
        if pd.notna(max_row.get("RQ"))
        else max_row.get("RER"),
        "Measured METs (Max VO2)": max_row.get("METS"),
        "VE/VC02 Slope": max_row.get("VE/VCO2"),
        "Angina (Ex or R) reason for stopping?": meta.get(
            "Angina (Ex or R) reason for stopping?"
        ),
    }

    return extracted_data


def append_to_master_db(new_records):
    """Appends structured records to the Master DB matching its schema perfectly."""
    # Load original master sheet column layout
    df_master = pd.read_excel(MASTER_DB_PATH, sheet_name=MASTER_SHEET_NAME)
    master_columns = list(df_master.columns)

    # Re-index new entries to align with all 87 columns in the exact master order
    df_new = pd.DataFrame(new_records)
    df_new = df_new.reindex(columns=master_columns)

    # Append to master workbook using openpyxl overlay
    with pd.ExcelWriter(
        MASTER_DB_PATH, engine="openpyxl", mode="a", if_sheet_exists="overlay"
    ) as writer:
        start_row = len(df_master) + 1
        df_new.to_excel(
            writer,
            sheet_name=MASTER_SHEET_NAME,
            index=False,
            header=False,
            startrow=start_row,
        )

    print(
        f"Successfully appended {len(new_records)} record(s) to '{MASTER_DB_PATH}'!"
    )


# Execution Pipeline
if __name__ == "__main__":
    # Find all BxB Excel files matching pattern
    bxb_files = glob.glob(
        os.path.join(BXB_FILES_DIRECTORY, "*BxB*.xlsx")
    ) + glob.glob(os.path.join(BXB_FILES_DIRECTORY, "*BxB*.xls"))

    records_to_append = []
    for file in bxb_files:
        if file.startswith("~$"):  # Skip Excel temporary lock files
            continue
        try:
            record = process_cpet_file(file)
            records_to_append.append(record)
            print(f"Processed: {file}")
        except Exception as e:
            print(f"Error processing {file}: {e}")

    if records_to_append:
        append_to_master_db(records_to_append)
    else:
        print("No new BxB files found to process.")