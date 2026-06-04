import os
import glob
import argparse
import pandas as pd

# 1. Define the targets and product descriptions
TARGET_NUMBERS = [
    344143, 344168, 344382, 225300, 6207476, 
    6397293, 6397384, 6397400, 6526529, 6275374, 
    862565, 254607
]

PRODUCT_MAP = {
    '344143': '344143 - SMPLY TOSTITS ORGNC SCOOP',
    '344168': '344168 - SIMPL TOSTITS ORGNC YLLOW GF',
    '344382': '344382 - SIMPLY TOSTITS ORGNC BLUE GF',
    '225300': '225300 - SIMPLY DORITS WHITE CHDR GF',
    '6207476': '6207476 - MISS VICKIES VARIETY PCK 10 CT',
    '6397293': '6397293 - CHESTERS CORN TWISTS',
    '6397384': '6397384 - HOSTESS HICKORY S/VINEGAR',
    '6397400': '6397400 - MUNCHOS',
    '6526529': '6526529 - DORITOS COLLS PZA CL RNCH -P',
    '6275374': '6275374 - LAYS SWEET CHILI HEAT',
    '862565': '862565 - SPITZ PUMPKIN SEEDS DILL',
    '254607': '254607 - SPITZ SNFLWR SEEDS SALTED'
}

target_strings = [str(num) for num in TARGET_NUMBERS]
results = {num: [] for num in target_strings}
searched_files = []

# 2. Set up Command Line Argument Parsing
parser = argparse.ArgumentParser(description="Scan Excel files in a specific directory using Column C.")
parser.add_argument("folder_path", type=str, help="The path to the folder containing your Excel files")
args = parser.parse_args()

if not os.path.isdir(args.folder_path):
    print(f"Error: The directory '{args.folder_path}' does not exist.")
    exit(1)

# 3. Gather all Excel files
search_path = os.path.join(args.folder_path, "*.xl*")
excel_files = glob.glob(search_path)

if not excel_files:
    print(f"No Excel files found in the directory: '{args.folder_path}'")
    exit()

print(f"Scanning {len(excel_files)} Excel file(s) targeting Column C...\n")

# 4. Process each file
for file_path in excel_files:
    file_name = os.path.basename(file_path)
    searched_files.append(file_name)
    try:
        excel_workbook = pd.ExcelFile(file_path)
        for sheet_name in excel_workbook.sheet_names:
            df = pd.read_excel(excel_workbook, sheet_name=sheet_name, dtype=str)
            
            if len(df.columns) >= 3:
                column_c_data = df.iloc[:, 2]
                
                found_values = (
                    column_c_data
                    .dropna()
                    .str.strip()
                    .str.split('.')
                    .str[0]
                    .tolist()
                )
                
                if column_c_data.name:
                    header_clean = str(column_c_data.name).strip().split('.')[0]
                    found_values.append(header_clean)

                for num in target_strings:
                    if num in found_values:
                        if file_name not in results[num]:
                            results[num].append(file_name)
                            
    except Exception as e:
        print(f"Could not read file {file_name} due to an error: {e}")

print("\nProcessing complete. Generating text report...")

# 5. Generate the simplified text file output
output_txt_file = "product_search_results.txt"
with open(output_txt_file, 'w') as f:
    f.write("==================================================\n")
    f.write("FILES SEARCHED:\n")
    for file_name in sorted(searched_files):
        f.write(f"  - {file_name}\n")
    f.write("==================================================\n\n")
    
    for num in target_strings:
        f.write(f"Product: {PRODUCT_MAP[num]}\n")
        if results[num]:
            for filename in results[num]:
                f.write(f"  - {filename}\n")
        else:
            f.write("  -> NOT FOUND\n")
        f.write("-" * 50 + "\n")

print(f"Done! Results successfully saved to '{output_txt_file}'")