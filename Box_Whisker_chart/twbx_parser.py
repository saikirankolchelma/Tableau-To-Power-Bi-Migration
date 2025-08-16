import os
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import json

def extract_twbx(twbx_file_path, extract_path):
    """Unzips a .twbx file to access the .twb file inside."""
    with zipfile.ZipFile(twbx_file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_path)
    for file in os.listdir(extract_path):
        if file.endswith('.twb'):
            return os.path.join(extract_path, file)
    return None

def save_to_output_folder(data, file_name, output_folder):
    """Save data to a JSON file in the output folder."""
    os.makedirs(output_folder, exist_ok=True)
    file_path = os.path.join(output_folder, file_name)
    with open(file_path, 'w') as file:
        json.dump(data, file, indent=4)
    print(f"Saved: {file_path}")
    return file_path

def extract_calculations_and_references(root, data_source_mapping):
    calculations_mapping = {}
    usage_mapping = []

    # Extract Calculations
    datasources = root.findall('.//datasource')
    for datasource in datasources:
        ds_name = datasource.get('name')
        columns = datasource.findall('.//column')
        for column in columns:
            field_name = column.get('caption') or column.get('name')
            tableau_generated_name = column.get('name')
            calculation = column.find('.//calculation')
            formula = calculation.get('formula') if calculation is not None else None

            # Map calculations to formulas and readable field names
            if 'calculation' in tableau_generated_name.lower():
                calculations_mapping[tableau_generated_name] = {
                    'field_name': field_name,
                    'formula': formula,
                    'data_source': data_source_mapping.get(ds_name, ds_name)
                }

    # Cross-reference these calculations in worksheets and dashboards
    worksheets = root.findall('.//worksheet')
    for worksheet in worksheets:
        ws_name = worksheet.get('name')
        columns_in_use = worksheet.findall('.//column')
        for col in columns_in_use:
            col_name = col.get('name')
            if col_name in calculations_mapping:
                usage_mapping.append({
                    'worksheet': ws_name,
                    'calculation': col_name,
                    'field_name': calculations_mapping[col_name]['field_name'],
                    'data_source': calculations_mapping[col_name]['data_source']
                })

    return calculations_mapping, usage_mapping

def extract_references_and_links(root):
    reference_data = []
    data_source_mapping = {}

    # Extract references to external files, data connections, and cross-file relationships
    datasources = root.findall('.//datasource')
    for datasource in datasources:
        ds_name = datasource.get('name')
        ds_caption = datasource.get('caption') or ds_name

        connections = datasource.findall('.//connection')
        for connection in connections:
            db_class = connection.get('class')
            db_name = connection.get('dbname') or connection.get('server') or 'N/A'
            ext_ref = connection.get('filename')  # Reference to external files (e.g., tde, hyper)

            if ext_ref:
                csv_name = os.path.splitext(os.path.basename(ext_ref))[0]
                data_source_mapping[ds_name] = csv_name
                reference_data.append({
                    'Data Source': ds_caption,
                    'Connection Type': db_class,
                    'Database Name': db_name,
                    'External Reference': ext_ref
                })

    return reference_data, data_source_mapping

def extract_visuals_and_layouts(root):
    visual_data = []

    # Extract Worksheets and their details
    worksheets = root.findall('.//worksheet')
    for worksheet in worksheets:
        ws_name = worksheet.get('name')

        # Initialize lists for rows, columns, and filters
        rows = [row.text for row in worksheet.findall('.//rows') if row.text]
        columns = [col.text for col in worksheet.findall('.//cols') if col.text]
        filters = [filt.get('column') for filt in worksheet.findall('.//filter') if filt.get('column')]

        # Process <panes> for <encodings>
        panes = worksheet.findall('.//pane')
        for pane in panes:
            encodings = pane.find('.//encodings')
            if encodings is not None:
                for encoding in encodings:
                    column = encoding.get('column')
                    if column:
                        columns.append(column)

        # Deduplicate columns
        columns = list(set(columns))

        # Append the worksheet details to the visual data
        visual_data.append({
            'Type': 'Worksheet',
            'Source': ws_name,
            'Rows': ', '.join(rows),
            'Columns': ', '.join(columns),
            'Filters': ', '.join(filters)
        })

    # Extract Dashboards
    dashboards = root.findall('.//dashboard')
    for dashboard in dashboards:
        db_name = dashboard.get('name')
        worksheets_in_dashboard = set()
        zones = dashboard.findall('.//zone')
        for zone in zones:
            view = zone.find('.//view')
            if view is not None:
                worksheet_name = view.get('name')
                if worksheet_name:
                    worksheets_in_dashboard.add(worksheet_name)
        worksheet_list = ', '.join(worksheets_in_dashboard)
        visual_data.append({
            'Type': 'Dashboard',
            'Source': db_name,
            'Worksheets': worksheet_list,
            'Rows': '',
            'Columns': '',
            'Filters': ''
        })

    return visual_data

def parse_workbook(twb_file_path):
    """Parses the .twb file to extract calculations, visuals, and external file references."""
    tree = ET.parse(twb_file_path)
    root = tree.getroot()

    # Extract external file references and data source mapping
    references, data_source_mapping = extract_references_and_links(root)

    # Extract calculations and usage with data_source_mapping applied
    calculations_mapping, usage_mapping = extract_calculations_and_references(root, data_source_mapping)

    # Extract visuals and layout information
    visuals = extract_visuals_and_layouts(root)

    return calculations_mapping, usage_mapping, references, visuals

def find_csv_or_hyper_files(extract_dir):
    csv_files = []
    hyper_files = []

    # Walk through the extracted directory to find CSV or Hyper files
    for root, dirs, files in os.walk(extract_dir):
        for file in files:
            if file.endswith('.csv'):
                csv_files.append(os.path.join(root, file))
            elif file.endswith('.hyper'):
                hyper_files.append(os.path.join(root, file))

    return csv_files, hyper_files

def extract_tableau_workbook(file_path, main_output_dir):
    """Extract Tableau workbook data and save it to the main output folder."""
    # DYNAMIC PATH: Get current script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "output")
    
    extract_dir = os.path.join(output_dir, "input_tableau_output")
    os.makedirs(extract_dir, exist_ok=True)

    # Extract the .twb file from .twbx
    if file_path.endswith('.twbx'):
        twb_file = extract_twbx(file_path, extract_dir)
    else:
        twb_file = file_path

    if not twb_file:
        raise ValueError("No .twb file found in the .twbx package or invalid file path.")

    # Parse the .twb file
    calculations, usage, references, visuals = parse_workbook(twb_file)

    # Find CSV and Hyper files
    csv_files, hyper_files = find_csv_or_hyper_files(extract_dir)

    # Create a CSV preview
    csv_preview = []
    if csv_files:
        for csv_file in csv_files:
            try:
                df = pd.read_csv(csv_file, nrows=3)
                csv_preview.append({
                    'file': csv_file,
                    'preview': df.to_dict(orient='records')
                })
            except Exception as e:
                print(f"⚠️ Error previewing CSV {csv_file}: {str(e)}")
                csv_preview.append({
                    'file': csv_file,
                    'error': str(e)
                })

    # Consolidate all extracted data
    extracted_data = {
        'calculations': calculations,
        'usage': usage,
        'references': references,
        'visuals': visuals,
        'csv_files': csv_files,
        'hyper_files': hyper_files,
        'csv_preview': csv_preview
    }

    # Save extracted data to the output folder
    save_to_output_folder(extracted_data, 'tableau_extracted_data.json', output_dir)

    print(f"✅ Extraction completed. All data saved in: {output_dir}")

if __name__ == "__main__":
    # Ask the user for the .twbx/.twb file path
    file_path = input("Enter the path to the Tableau .twbx or .twb file: ").strip()

    # Validate the input file
    if not os.path.isfile(file_path):
        print("Error: The specified file does not exist.")
    elif not file_path.endswith(('.twbx', '.twb')):
        print("Error: The file must be a .twbx or .twb Tableau file.")
    else:
        # Get current script directory for output
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_dir = os.path.join(script_dir, "output")

        try:
            # Extract Tableau workbook
            extract_tableau_workbook(file_path, output_dir)
        except Exception as e:
            print(f"An error occurred during extraction: {e}")