import os
import glob
import pandas as pd
import openpyxl

# ==========================================
# CONFIGURATION
# ==========================================
INPUT_FOLDER = "./CPET results (new)"          # Directory containing individual CPET spreadsheets
OUTPUT_FILE = "VAPAHCS CPET DB.xlsx"    # Target master spreadsheet
OUTPUT_SHEET = "baseline (new)"        # Target sheet name inside master file

# Master Column Mapping: {Master_Column_Name: [Possible_Source_Column_Names]}
COLUMN_MAPPING = {
    # Identifying Information
    "Name": ["Name", "Subject Name"],
    "Study ID #": ["ID1", "Study ID #", "Study ID"],
    "Gender": ["Gender", "Sex"],
    "Age at ETT": ["Age"],
    "Height\n (in.)": ["Height (in)", "Height (in.)", "Height (in) "],
    "Weight \n(lbs.)": ["Weight (lbs)", "Weight (lbs.)", "Weight (lbs) "],
    "Date of ETT": ["Test date", "Date"],

    # Peak CPET Metrics (Captured at Peak VO2)
    "VO2 ml/Kg/min (max)": ["VO2/kg", "VO2/Kg"],
    "VO2 in ml (max)": ["VO2"],
    "VCO2 (max)": ["VCO2"],
    "VE (max)": ["VE"],
    "RER max": ["RQ", "RER", "RER max"],  # Treats RQ and RER as equivalent
    "Measured METs (Max VO2)": ["METS", "METs"],

    # Custom Mappings
    "VE/VC02 Slope": ["VE/VCO2", "VE/VCO2 Slope", "VE/VC02 Slope"],
    "Angina (Ex or R) reason for stopping?": ["Reason for Stopping Test", "Reason for Test", "Angina reason for stopping"]
}


def extract_header_metadata(df_raw):
    """
    Scans the header rows and extracts key-value pairs, automatically looking
    to the right past empty/merged cells to find the matching value for each key (e.g., ID1).
    """
    metadata = {}
    for row_idx in range(min(15, len(df_raw))):
        row = df_raw.iloc[row_idx]
        for col_idx in range(len(row)):
            cell_val = str(row[col_idx]).strip() if pd.notna(row[col_idx]) else ""
            if cell_val in ["ID1", "Last Name", "First Name", "Gender", "Age", "Height (in)", "Weight (lbs)", "Test date", "Reason for Stopping Test", "Reason for Test"]:
                # Look to the right for the first non-empty value
                for next_col in range(col_idx + 1, len(row)):
                    val = row[next_col]
                    if pd.notna(val) and str(val).strip() != "":
                        metadata[cell_val] = val
                        break
    return metadata


def process_cpet_file(file_path):
    try:
        xl = pd.ExcelFile(file_path)
        sheet = "Data" if "Data" in xl.sheet_names else xl.sheet_names[0]

        # Parse metadata from header section
        df_raw = pd.read_excel(file_path, sheet_name=sheet, header=None)
        metadata = extract_header_metadata(df_raw)

        # Parse breath-by-breath data table
        df_data = pd.read_excel(file_path, sheet_name=sheet, header=0)

        # Strip out unit row if present under header
        if df_data.iloc[0].astype(str).str.contains("mL|L/min|s|%").any():
            df_data = df_data.iloc[1:].reset_index(drop=True)

        # Identify VO2 column for finding peak effort
        vo2_col = None
        for col in ["VO2/kg", "VO2", "VO2 ml/Kg/min"]:
            if col in df_data.columns:
                vo2_col = col
                break

        if not vo2_col:
            return None

        # Find row corresponding to Peak VO2
        df_data[vo2_col] = pd.to_numeric(df_data[vo2_col], errors="coerce")
        max_vo2_idx = df_data[vo2_col].idxmax()
        max_row = df_data.loc[max_vo2_idx].to_dict()

        # Combine header metadata and max VO2 metrics
        combined_source = {**metadata, **max_row}

        # Map to Master schema
        extracted_data = {}
        for master_col, source_aliases in COLUMN_MAPPING.items():
            value = None
            for alias in source_aliases:
                if alias in combined_source and pd.notna(combined_source[alias]):
                    value = combined_source[alias]
                    break
            extracted_data[master_col] = value

        # Construct full name if First/Last Name are separately parsed
        if not extracted_data.get("Name"):
            last = combined_source.get("Last Name", "")
            first = combined_source.get("First Name", "")
            if last or first:
                extracted_data["Name"] = f"{last}, {first}".strip(", ")

        return extracted_data

    except Exception as e:
        print(f"Error processing {os.path.basename(file_path)}: {e}")
        return None


def main():
    file_pattern = os.path.join(INPUT_FOLDER, "*.xlsx")
    files = glob.glob(file_pattern)

    if not files:
        print(f"No Excel files found in folder '{INPUT_FOLDER}'.")
        return

    results = []
    for file_path in files:
        if os.path.basename(file_path) == os.path.basename(OUTPUT_FILE):
            continue

        data = process_cpet_file(file_path)
        if data:
            results.append(data)

    if not results:
        print("No valid data extracted.")
        return

    df_new = pd.DataFrame(results)

    # Append records to the master file
    if os.path.exists(OUTPUT_FILE):
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            try:
                df_existing = pd.read_excel(OUTPUT_FILE, sheet_name=OUTPUT_SHEET)
                start_row = len(df_existing) + 1
                df_new.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False, header=False, startrow=start_row)
            except ValueError:
                df_new.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)
    else:
        df_new.to_excel(OUTPUT_FILE, sheet_name=OUTPUT_SHEET, index=False)

    print(f"Successfully appended {len(results)} records to '{OUTPUT_FILE}'.")


if __name__ == "__main__":
    main()