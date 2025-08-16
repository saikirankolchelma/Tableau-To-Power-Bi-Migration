# import zipfile
# import os
# import pandas as pd
# import xml.etree.ElementTree as ET
# import time
# import subprocess
# import pyautogui
# import re
# from tableauhyperapi import HyperProcess, Connection, Telemetry, HyperException

# # ✅ **Dynamically Set Directories**
# BASE_DIR = os.getcwd()  # Get the directory where the script is running
# OUTPUT_DIR = os.path.join(BASE_DIR, "output")
# EXTRACT_DIR = os.path.join(OUTPUT_DIR, "extracted")  # Extracted `.twbx` content
# CSV_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "csv_output")  # Extracted datasets

# # ✅ **Ensure Required Directories Exist**
# os.makedirs(EXTRACT_DIR, exist_ok=True)
# os.makedirs(CSV_OUTPUT_DIR, exist_ok=True)

# def extract_twbx(twbx_file):
#     """Extracts a `.twbx` file into `output/extracted/`."""
#     try:
#         with zipfile.ZipFile(twbx_file, 'r') as zip_ref:
#             zip_ref.extractall(EXTRACT_DIR)
#         print(f"✅ Extracted {twbx_file} to {EXTRACT_DIR}")
#     except zipfile.BadZipFile:
#         print(f"❌ Error: {twbx_file} is not a valid ZIP file.")
#     except Exception as e:
#         print(f"❌ Error extracting {twbx_file}: {e}")

# def find_hyper_files():
#     """Finds `.hyper` files inside `output/extracted/`."""
#     hyper_files = []
#     for root, _, files in os.walk(EXTRACT_DIR):
#         for file in files:
#             if file.endswith('.hyper'):
#                 hyper_files.append(os.path.join(root, file))
#     return hyper_files

# def find_tableau_csv_names():
#     """Extracts the original Tableau CSV names from `.twb` or `.tds` files."""
#     tableau_csv_names = {}
#     for root, _, files in os.walk(EXTRACT_DIR):
#         for file in files:
#             if file.endswith('.twb') or file.endswith('.tds'):
#                 file_path = os.path.join(root, file)
#                 try:
#                     tree = ET.parse(file_path)
#                     xml_root = tree.getroot()
#                     for ds in xml_root.findall(".//connection"):
#                         hyper_file = ds.get("dbname", "").replace(".hyper", "").strip()
#                         csv_name = ds.get("name", "").strip()
#                         if hyper_file and csv_name:
#                             tableau_csv_names[hyper_file] = csv_name + ".csv"
#                 except ET.ParseError as e:
#                     print(f"❌ Could not parse {file_path}: {e}")
#                 except Exception as e:
#                     print(f"❌ Error processing {file_path}: {e}")
#     return tableau_csv_names

# def extract_hyper_to_csv(hyper_file, csv_name_mapping):
#     """Extracts data from a `.hyper` file and saves it as CSV in `output/csv_output/`."""
#     try:
#         csv_filename = "File1.csv"  # Fixed filename for extracted CSVs
#         csv_filepath = os.path.join(CSV_OUTPUT_DIR, csv_filename)

#         with HyperProcess(telemetry=Telemetry.SEND_USAGE_DATA_TO_TABLEAU) as hyper:
#             with Connection(endpoint=hyper.endpoint, database=hyper_file) as connection:
#                 schema_name = "Extract"
#                 tables = connection.catalog.get_table_names(schema_name)

#                 if not tables:
#                     print(f"❌ No tables found in `{hyper_file}`.")
#                     return None

#                 for table in tables:
#                     print(f"📊 Extracting data from table: {table.name} in `{hyper_file}`")
#                     columns = connection.catalog.get_table_definition(table).columns
#                     column_names = [str(col.name).replace('"', '') for col in columns]

#                     query = f"SELECT * FROM {table}"
#                     rows = connection.execute_query(query)
#                     data = pd.DataFrame(rows, columns=column_names)

#                     if data.empty:
#                         print(f"⚠️ Table `{table.name}` is empty. Skipping...")
#                         continue

#                     data.to_csv(csv_filepath, index=False)
#                     print(f"✅ Data saved to `{csv_filepath}`")

#         return csv_filepath
#     except HyperException as e:
#         print(f"❌ Hyper API error processing `{hyper_file}`: {e}")
#     except Exception as e:
#         print(f"❌ Error extracting data from `{hyper_file}`: {e}")
#     return None

# def process_twbx_file(twbx_file):
#     """Processes the `.twbx` file: extracts data and saves to CSV in `output/csv_output/`."""
    
#     # ✅ Step 1: Extract `.twbx`
#     extract_twbx(twbx_file)
    
#     # ✅ Step 2: Extract CSV names from `.twb/.tds`
#     csv_name_mapping = find_tableau_csv_names()
    
#     # ✅ Step 3: Find `.hyper` files
#     hyper_files = find_hyper_files()
#     if not hyper_files:
#         print(f"❌ No `.hyper` files found in `{twbx_file}`. Skipping...")
#         return
    
#     # ✅ Step 4: Extract data from `.hyper` files and save CSV
#     for hyper_file in hyper_files:
#         extract_hyper_to_csv(hyper_file, csv_name_mapping)
    
#     print(f"✅ Dataset extraction completed! CSV files saved in `{CSV_OUTPUT_DIR}`")

# # ✅ **Execute when run directly**
# if __name__ == "__main__":
#     twbx_file = input("🔹 Enter the path to the Tableau `.twbx` file: ").strip()
    
#     if not os.path.exists(twbx_file):
#         print("❌ Error: The provided `.twbx` file does not exist.")
#     else:
#         process_twbx_file(twbx_file)




import zipfile
import os
import xml.etree.ElementTree as ET
import pandas as pd
import re
from tableauhyperapi import HyperProcess, Connection, Telemetry, HyperException
import json
import numpy as np  # ✅ Required for CASE evaluation

# Set Directories
BASE_DIR = os.getcwd()
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
EXTRACT_DIR = os.path.join(OUTPUT_DIR, "extracted")
CSV_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "csv_output")
EXCEL_OUTPUT_FILE = os.path.join(OUTPUT_DIR, "combined_datasets.xlsx")
MSCRIPT_FILE = os.path.join(OUTPUT_DIR, "powerbi_mscript.txt")

# Ensure Required Directories Exist
os.makedirs(EXTRACT_DIR, exist_ok=True)
os.makedirs(CSV_OUTPUT_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

def extract_twbx(twbx_file):
    """Extracts a .twbx file into output/extracted/."""
    try:
        with zipfile.ZipFile(twbx_file, 'r') as zip_ref:
            zip_ref.extractall(EXTRACT_DIR)
        print(f"✅ Extracted {twbx_file} to {EXTRACT_DIR}")
    except zipfile.BadZipFile:
        print(f"❌ Error: {twbx_file} is not a valid ZIP file.")
    except Exception as e:
        print(f"❌ Error extracting {twbx_file}: {e}")

def find_hyper_files():
    """Finds .hyper files inside output/extracted/."""
    hyper_files = {}
    for root, _, files in os.walk(EXTRACT_DIR):
        for file in files:
            if file.endswith('.hyper'):
                hyper_files[file] = os.path.join(root, file)
    return hyper_files

# def find_table_names():
#     """Extracts dataset names and table names from the .twb file."""
#     table_mapping = {}
#     table_names = {}
#     for root, _, files in os.walk(EXTRACT_DIR):
#         for file in files:
#             if file.endswith('.twb'):
#                 file_path = os.path.join(root, file)
#                 try:
#                     tree = ET.parse(file_path)
#                     xml_root = tree.getroot()
#                     for datasource in xml_root.findall(".//datasource"):
#                         caption = datasource.get("caption", "").strip()
#                         name = datasource.get("name", "").strip()
#                         for connection in datasource.findall(".//connection"):
#                             dbname = connection.get("dbname", "").strip()
#                             if dbname and dbname.endswith(".hyper"):
#                                 hyper_filename = os.path.basename(dbname)
#                                 table_mapping[hyper_filename] = caption
#                         for relation in datasource.findall(".//relation"):
#                             table_name = relation.get("name", "").strip()
#                             if table_name:
#                                 table_names[table_name] = caption
#                 except ET.ParseError as e:
#                     print(f"❌ XML Parsing Error in {file_path}: {e}")
#                 except Exception as e:
#                     print(f"❌ Error processing {file_path}: {e}")
#     return table_mapping, table_names

def find_table_names():
    """Extracts dataset names and table names from the .twb file."""
    table_mapping = {}
    table_names = {}
    
    for root, _, files in os.walk(EXTRACT_DIR):
        for file in files:
            if file.endswith('.twb'):
                file_path = os.path.join(root, file)
                try:
                    # Parse the XML file
                    tree = ET.parse(file_path)
                    xml_root = tree.getroot()
                    
                    # Get workbook name as fallback
                    workbook_name = os.path.splitext(file)[0]
                    
                    # Find all datasources
                    datasources = xml_root.findall(".//datasource")
                    if not datasources:
                        # For single datasource files, try alternative path
                        datasources = xml_root.findall(".//datasources/datasource")
                    
                    # If still no datasources found, create a dummy entry
                    if not datasources:
                        # Find any hyper files referenced directly
                        connections = xml_root.findall(".//connection")
                        for connection in connections:
                            dbname = connection.get('dbname', '')
                            if dbname and dbname.endswith('.hyper'):
                                hyper_filename = os.path.basename(dbname)
                                table_mapping[hyper_filename] = workbook_name
                        continue
                    
                    for datasource in datasources:
                        # Get datasource name (try caption first, then name)
                        caption = datasource.get('caption', '').strip()
                        name = datasource.get('name', '').strip()
                        ds_name = caption if caption else name if name else workbook_name
                        
                        # Find connections to hyper files
                        connections = datasource.findall(".//connection")
                        for connection in connections:
                            dbname = connection.get('dbname', '').strip()
                            if dbname and dbname.endswith('.hyper'):
                                hyper_filename = os.path.basename(dbname)
                                table_mapping[hyper_filename] = ds_name
                        
                        # Find relations (tables)
                        relations = datasource.findall(".//relation") + datasource.findall(".//relation-table")
                        for relation in relations:
                            table_name = relation.get('name', '').strip() or relation.get('table', '').strip()
                            if table_name:
                                table_names[table_name] = ds_name
                        
                        # If no relations found, use default "Extract" name
                        if not relations:
                            table_names["Extract"] = ds_name
                
                except ET.ParseError as e:
                    print(f"❌ XML Parsing Error in {file_path}: {e}")
                except Exception as e:
                    print(f"❌ Error processing {file_path}: {e}")
    
    return table_mapping, table_names

# def extract_hyper_to_csv(hyper_file, hyper_filename, calculations_json=None):
#     """Extracts data from .hyper files with clean table names without hash suffixes."""
#     try:
#         extracted_files = []
#         with HyperProcess(telemetry=Telemetry.SEND_USAGE_DATA_TO_TABLEAU) as hyper:
#             with Connection(endpoint=hyper.endpoint, database=hyper_file) as connection:
#                 schema_name = "Extract"
#                 tables = connection.catalog.get_table_names(schema_name)
#                 if not tables:
#                     print(f"❌ No tables found in {hyper_file}.")
#                     return None

#                 for table in tables:
#                     # Get original table name
#                     original_name = str(table.name).replace('"', '')
                    
#                     # Clean the table name - remove everything after the last underscore
#                     # This handles both the !_HASH pattern and _HASH pattern
#                     clean_name = original_name.split('!')[0]  # Remove !_HASH if present
#                     clean_name = clean_name.split('_')[0]     # Remove _HASH if present
                    
#                     # Further clean special characters
#                     clean_name = clean_name.lower()  # Convert to lowercase
#                     clean_name = re.sub(r'[^a-zA-Z0-9_]', '_', clean_name)  # Replace special chars
#                     clean_name = re.sub(r'_+', '_', clean_name).strip('_')  # Clean multiple underscores
                    
#                     # Create CSV filename
#                     csv_filename = f"{clean_name}.csv"
#                     csv_filepath = os.path.join(CSV_OUTPUT_DIR, csv_filename)
                    
#                     # Handle duplicates
#                     counter = 1
#                     while os.path.exists(csv_filepath):
#                         csv_filepath = os.path.join(CSV_OUTPUT_DIR, f"{clean_name}_{counter}.csv")
#                         counter += 1
                    
#                     # Extract data
#                     columns = connection.catalog.get_table_definition(table).columns
#                     column_names = [str(col.name).replace('"', '') for col in columns]
#                     query = f"SELECT * FROM {table}"
#                     rows = connection.execute_query(query)
#                     data = pd.DataFrame(rows, columns=column_names)
                    
#                     if data.empty:
#                         print(f"⚠ Table {clean_name} is empty. Skipping...")
#                         continue
                        

#                     if calculations_json:
#                         for calc in calculations_json.values():
#                             try:
#                                 field_name = calc.get("field_name")
#                                 formula = calc.get("formula")
#                                 target_table = calc.get("table", "").replace("!", "_")
                                
#                                 if target_table == clean_name:
#                                     # Handle CASE statements differently
#                                     if formula.strip().lower().startswith("case"):
#                                         # [Previous CASE handling logic...]
#                                         pass
#                                     else:
#                                         # Replace column references
#                                         column_refs = re.findall(r'\[([^\]]+)\]', formula)
#                                         for col in column_refs:
#                                             formula = formula.replace(f"[{col}]", f"data['{col}']")
#                                         data[field_name] = eval(formula)
#                                         print(f"➕ Added calculated field: {field_name}")
#                             except Exception as e:
#                                 print(f"⚠️ Failed to compute {field_name}: {e}")
                    
#                     data.to_csv(csv_filepath, index=False)
#                     extracted_files.append(csv_filepath)
#                     print(f"✅ Data saved to {csv_filepath}")
                    
#         return extracted_files
#     except HyperException as e:
#         print(f"❌ Hyper API error processing {hyper_file}: {e}")
#     except Exception as e:
#         print(f"❌ Error extracting data from {hyper_file}: {e}")
#     return None

def extract_hyper_to_csv(hyper_file, hyper_filename, calculations_json=None):
    """Extracts data from a .hyper file and saves each table as a CSV."""
    try:
        extracted_files = []
        with HyperProcess(telemetry=Telemetry.SEND_USAGE_DATA_TO_TABLEAU) as hyper:
            with Connection(endpoint=hyper.endpoint, database=hyper_file) as connection:
                schema_name = "Extract"
                tables = connection.catalog.get_table_names(schema_name)
                if not tables:
                    print(f"❌ No tables found in {hyper_file}.")
                    return None
                
                # Get the proper table name from our dynamic mapping
                table_mapping, _ = find_table_names()
                ds_name = table_mapping.get(hyper_filename, os.path.splitext(hyper_filename)[0])
                
                for table in tables:
                    table_name_str = str(table.name).replace('"', '')  # e.g. 'Sales Order!data_FE11F2FA...'

    # Extract base table name (e.g., 'Sales Order') and clean it
                    if len(tables) > 1:
                        base_name = table_name_str.split('!')[0].strip()  # Extract 'Sales Order'
                        clean_table_name = base_name.replace(" ", "_") + "_data"  # Result: 'Sales_Order_data'
                        print(f"Using extracted table name for multiple sheets: {clean_table_name}")
                    else:
        # Single sheet logic (keep existing)
                        clean_table_name = ds_name if ds_name else re.sub(r'_[A-F0-9]{32}$', '', table_name_str).replace("!", "_")
                        print(f"Using cleaned table name for single sheet: {clean_table_name}")
                    # Use the actual table name when there are multiple sheets
                    # if len(tables) > 1:
                    #     # Use the extracted table name (or clean it)
                    #     clean_table_name = re.sub(r'[\\/*?:"<>|]', "_", table_name_str)
                    #     print(f"Using extracted table name for multiple sheets: {clean_table_name}")
                    # else:
                    #     # For single sheet, keep the previous logic
                    #     print(f"Using cleaned table name for single sheet: {clean_table_name}")
                    
                    # Final CSV filename
                    csv_filename = clean_table_name + ".csv"
                    csv_filepath = os.path.join(CSV_OUTPUT_DIR, csv_filename)
                    
                    # Handle duplicate filenames by appending a number
                    counter = 1
                    while os.path.exists(csv_filepath):
                        csv_filepath = os.path.join(CSV_OUTPUT_DIR, f"{clean_table_name}_{counter}.csv")
                        counter += 1
                    
                    # Rest of your extraction logic remains the same...
                    columns = connection.catalog.get_table_definition(table).columns
                    column_names = [str(col.name).replace('"', '') for col in columns]
                    query = f"SELECT * FROM {table}"
                    rows = connection.execute_query(query)
                    data = pd.DataFrame(rows, columns=column_names)
                    
                    if data.empty:
                        print(f"⚠ Table {clean_table_name} is empty. Skipping...")
                        continue
                        
                    # Handle calculations if provided
                    if calculations_json:
                        # Your existing calculations logic...
                        pass
                    
                    data.to_csv(csv_filepath, index=False)
                    extracted_files.append(csv_filepath)
                    print(f"✅ Data saved to {csv_filepath}")
                    
        return extracted_files
    except HyperException as e:
        print(f"❌ Hyper API error processing {hyper_file}: {e}")
    except Exception as e:
        print(f"❌ Error extracting data from {hyper_file}: {e}")
    return None


# def extract_hyper_to_csv(hyper_file, hyper_filename, calculations_json=None):
#     """Extracts data from .hyper files with proper naming for both single and multiple sheets."""
#     try:
#         extracted_files = []
#         with HyperProcess(telemetry=Telemetry.SEND_USAGE_DATA_TO_TABLEAU) as hyper:
#             with Connection(endpoint=hyper.endpoint, database=hyper_file) as connection:
#                 schema_name = "Extract"
#                 tables = connection.catalog.get_table_names(schema_name)
#                 if not tables:
#                     print(f"❌ No tables found in {hyper_file}.")
#                     return None

#                 # Get table mapping to find original names
#                 table_mapping, _ = find_table_names()
                
#                 for table in tables:
#                     # Get original table name
#                     original_name = str(table.name).replace('"', '')
                    
#                     # Clean the table name (remove hash suffixes, special chars)
#                     clean_name = re.sub(r'!_[A-F0-9]{32}$', '', original_name)  # Remove hash
#                     clean_name = clean_name.replace("!", "_")  # Replace special chars
#                     clean_name = re.sub(r'[^a-zA-Z0-9_]', '_', clean_name).lower()  # Sanitize
#                     clean_name = re.sub(r'^extract_|^table_', '', clean_name)  # Remove prefixes
#                     clean_name = re.sub(r'_+', '_', clean_name).strip('_')  # Clean underscores
                    
#                     # For single table workbooks, check if name is generic
#                     if len(tables) == 1 and clean_name in ['extract', 'table', 'data', 'sheet']:
#                         # Try to get a better name from the table mapping
#                         ds_name = table_mapping.get(hyper_filename, '')
#                         if ds_name:
#                             clean_name = re.sub(r'[^a-zA-Z0-9_]', '_', ds_name).lower()
                    
#                     # Final filename
#                     csv_filename = f"{clean_name}.csv"
#                     csv_filepath = os.path.join(CSV_OUTPUT_DIR, csv_filename)
                    
#                     # Handle duplicates
#                     counter = 1
#                     while os.path.exists(csv_filepath):
#                         csv_filepath = os.path.join(CSV_OUTPUT_DIR, f"{clean_name}_{counter}.csv")
#                         counter += 1
                    
#                     # Extract data
#                     columns = connection.catalog.get_table_definition(table).columns
#                     column_names = [str(col.name).replace('"', '') for col in columns]
#                     query = f"SELECT * FROM {table}"
#                     rows = connection.execute_query(query)
#                     data = pd.DataFrame(rows, columns=column_names)
                    
#                     if data.empty:
#                         print(f"⚠ Table {clean_name} is empty. Skipping...")
#                         continue
                        
#                     # Apply calculations if provided
#                     if calculations_json:
#                         for calc in calculations_json.values():
#                             try:
#                                 field_name = calc.get("field_name")
#                                 formula = calc.get("formula")
#                                 target_table = calc.get("table", "").replace("!", "_")
                                
#                                 if target_table == clean_name:
#                                     # Handle CASE statements differently
#                                     if formula.strip().lower().startswith("case"):
#                                         # [Previous CASE handling logic...]
#                                         pass
#                                     else:
#                                         # Replace column references
#                                         column_refs = re.findall(r'\[([^\]]+)\]', formula)
#                                         for col in column_refs:
#                                             formula = formula.replace(f"[{col}]", f"data['{col}']")
#                                         data[field_name] = eval(formula)
#                                         print(f"➕ Added calculated field: {field_name}")
#                             except Exception as e:
#                                 print(f"⚠️ Failed to compute {field_name}: {e}")
                    
#                     data.to_csv(csv_filepath, index=False)
#                     extracted_files.append(csv_filepath)
#                     print(f"✅ Data saved to {csv_filepath}")
                    
#         return extracted_files
#     except HyperException as e:
#         print(f"❌ Hyper API error processing {hyper_file}: {e}")
#     except Exception as e:
#         print(f"❌ Error extracting data from {hyper_file}: {e}")
#     return None


def list_tables_in_hyper(hyper_file):
    """Lists all tables inside a .hyper file."""
    try:
        with HyperProcess(telemetry=Telemetry.SEND_USAGE_DATA_TO_TABLEAU) as hyper:
            with Connection(endpoint=hyper.endpoint, database=hyper_file) as connection:
                schema_name = "Extract"
                tables = connection.catalog.get_table_names(schema_name)
                if not tables:
                    print(f"⚠ No tables found in {hyper_file}.")
                    return []
                table_list = [table.name for table in tables]
                return table_list
    except HyperException as e:
        print(f"❌ Hyper API error processing {hyper_file}: {e}")
    except Exception as e:
        print(f"❌ Error extracting table names from {hyper_file}: {e}")
    return []

def csv_to_excel(csv_folder, excel_file):
    """Converts all CSV files in the folder to sheets in a single Excel file."""
    csv_files = [f for f in os.listdir(csv_folder) if f.endswith(".csv")]
    if not csv_files:
        print(f"❌ No CSV files found in {csv_folder} to convert to Excel.")
        return []
    
    sheet_names = []
    with pd.ExcelWriter(excel_file, engine='openpyxl') as writer:
        for csv_file in csv_files:
            csv_path = os.path.join(csv_folder, csv_file)
            df = pd.read_csv(csv_path)
            # Use the CSV filename (without .csv) as the sheet name, sanitized
            sheet_name = os.path.splitext(csv_file)[0]
            # Sanitize sheet name: replace invalid chars and truncate to 31 chars
            sheet_name = re.sub(r'[\/:*?"<>|]', '_', sheet_name)
            if len(sheet_name) > 31:
                sheet_name = sheet_name[:31]
            # Ensure uniqueness by appending a counter if necessary
            base_name = sheet_name
            counter = 1
            while sheet_name in sheet_names:
                suffix = f"_{counter}"
                sheet_name = base_name[:31 - len(suffix)] + suffix
                counter += 1
            sheet_names.append(sheet_name)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"✅ Converted {csv_file} to sheet '{sheet_name}' in {excel_file}")
    return sheet_names

def generate_mscript_for_powerbi(excel_file, sheet_names):
    """Generates a Power BI M script for loading data from the Excel file."""
    if not sheet_names:
        return "// Error: No sheets found."
    
    excel_file_path = excel_file.replace("\\", "\\\\")
    dataset_name = "Excel_Dataset"
    selected_sheets_str = ", ".join(f'"{sheet}"' for sheet in sheet_names)
    default_sheet = sheet_names[0] if sheet_names else ""
    
    mscript = f'''
    let
        // Load the Excel file
        Source_{dataset_name} = Excel.Workbook(File.Contents("{excel_file_path}"), null, true),

        // Parameter for sheet selection
        SelectedSheetName = "",

        // Filter sheets of interest
        SelectedSheets = {{{selected_sheets_str}}},
        FilteredSheets = Table.SelectRows(Source_{dataset_name}, each List.Contains(SelectedSheets, [Name])),

        // Validate selected sheet exists
        TargetSheet = Table.SelectRows(FilteredSheets, each [Name] = SelectedSheetName),
        CheckSheet = if Table.IsEmpty(TargetSheet) then 
            error Error.Record(
                "Sheet not found", 
                "Available sheets: " & Text.Combine(FilteredSheets[Name], ", "), 
                [RequestedSheet = SelectedSheetName]
            )
        else TargetSheet,

        // Extract sheet data
        SheetData = try CheckSheet{{0}}[Data] otherwise error Error.Record(
            "Data extraction failed",
            "Verify sheet structure",
            [SheetName = SelectedSheetName, AvailableColumns = Table.ColumnNames(CheckSheet)]
        ),

        // Promote headers
        PromotedHeaders = Table.PromoteHeaders(SheetData, [PromoteAllScalars=true]),
        
        // Detect and apply column types dynamically
       ColumnsToTransform = Table.ColumnNames(PromotedHeaders),
    ChangedTypes = Table.TransformColumnTypes(
        PromotedHeaders,
        List.Transform(
            ColumnsToTransform,
            each {{_, 
                let
                    SampleValue = List.First(Table.Column(PromotedHeaders, _), null),
                    TypeDetect = 
                        if SampleValue = null then type text
                        else if Value.Is(SampleValue, Number.Type) then
                            if Number.Round(SampleValue) = SampleValue then Int64.Type else type number
                        else if try Date.From(SampleValue) is date then type date
                        else if try DateTime.From(SampleValue) is datetime then type datetime
                        else type text
                in
                    TypeDetect
            }}
            )
        ),

        // Clean data
        CleanedData = Table.SelectRows(ChangedTypes, each not List.Contains(Record.FieldValues(_), null)),
        FinalTable_{dataset_name} = Table.Distinct(CleanedData)
    in
        FinalTable_{dataset_name}
    '''
    return mscript

def process_twbx_file(twbx_file):
    """Processes the .twbx file: extracts data, converts CSVs to Excel, and generates M script."""
    # Step 1: Extract .twbx
    extract_twbx(twbx_file)

    # Step 2: Extract dataset names & table names from .twb
    table_mapping, table_names = find_table_names()

    # Step 3: Find .hyper files
    hyper_files = find_hyper_files()
    if not hyper_files:
        print(f"❌ No .hyper files found in {twbx_file}. Skipping extraction...")
        return

    # Step 4: Extract table names from .hyper files
    all_tables = []
    for hyper_filename, hyper_file_path in hyper_files.items():
        table_names_from_hyper = list_tables_in_hyper(hyper_file_path)
        all_tables.extend(table_names_from_hyper)

    print("\n📊 Extracted Table Names:")
    for table in all_tables:
        print(f" - {table}")

    calculations_json = {}
    json_path = os.path.join(OUTPUT_DIR, "tableau_extracted_data.json")
    if os.path.exists(json_path):
        with open(json_path, 'r') as f:
            data = json.load(f)
            calculations_json = data.get("calculations", {})


    # Step 5: Extract data to CSV files
    for hyper_filename, hyper_file_path in hyper_files.items():
        extract_hyper_to_csv(hyper_file_path, hyper_filename, calculations_json)
    print(f"\n✅ Dataset extraction completed! CSV files saved in {CSV_OUTPUT_DIR}")

    # Step 6: Convert CSV files to a single Excel file
    sheet_names = csv_to_excel(CSV_OUTPUT_DIR, EXCEL_OUTPUT_FILE)
    if sheet_names:
        print(f"\n✅ All CSVs combined into {EXCEL_OUTPUT_FILE} with {len(sheet_names)} sheets.")

    # Step 7: Generate Power BI M script using the Excel file
    mscript = generate_mscript_for_powerbi(EXCEL_OUTPUT_FILE, sheet_names)
    with open(MSCRIPT_FILE, "w", encoding="utf-8") as file:
        file.write(mscript)
    print(f"\n✅ Power BI M script saved to: {MSCRIPT_FILE}")

#✅ *Execute when run directly*
if __name__ == "__main__":
    twbx_file = input("🔹 Enter the path to the Tableau .twbx file: ").strip()
    
    if not os.path.exists(twbx_file):
        print("❌ Error: The provided .twbx file does not exist.")
    else:
        process_twbx_file(twbx_file)
