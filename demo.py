import os
import json
import pandas as pd
import google.generativeai as genai
import re
import uuid
import time
from google.api_core import exceptions as api_exceptions
import sys 

# === Configuration Paths ===
BASE_DIR = os.path.dirname(os.path.abspath(__file__)) 
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
CSV_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "csv_output")
REFERENCE_TEXT_PATH = os.path.join(BASE_DIR, "Reference Chart Configurations.txt")

# Validate directories
if not os.path.exists(CSV_OUTPUT_DIR):
    print(f"❌ CSV folder not found at: {CSV_OUTPUT_DIR}")
    os.makedirs(CSV_OUTPUT_DIR, exist_ok=True)
    print(f"✅ Created CSV folder: {CSV_OUTPUT_DIR}")

# === API Key ===
api_key = "AIzaSyDNKuhdWd5gQHgWVrQ8uABTTaDI6CEoKys" # Replace with your actual key or load securely
if not api_key:
    raise ValueError("Gemini API key not found. Please set the 'GEMINI_API_KEY' environment variable.")
genai.configure(api_key=api_key)
model = genai.GenerativeModel("gemini-1.5-flash")


# === Global Helper Functions ===
# --- General Utilities ---
def deep_copy_json_obj(obj):
    """
    Performs a deep copy of a JSON-serializable Python object using serialization/deserialization.
    """
    return json.loads(json.dumps(obj))

def generate_unique_id():
    """Generates a unique hexadecimal ID."""
    return uuid.uuid4().hex
    
def save_to_json(data, output_path, filename):
    """Save data to a JSON file in the specified output path."""
    os.makedirs(output_path, exist_ok=True)
    full_path = os.path.join(output_path, filename)
    try:
        with open(full_path, 'w', encoding="utf-8") as file:
            json.dump(data, file, indent=2)
        print(f"✅ Successfully saved to {full_path}")
    except Exception as e:
        print(f"❌ Error writing to JSON file: {e}")

def normalize_name(name):
    """Normalize names for flexible matching."""
    if not name:
        return ""
    return re.sub(r'[^a-z0-9]', '', str(name).lower().strip())

# --- Field and Aggregation Helpers ---
def extract_base_field_name(field_full_path: str) -> str:
    """
    Extracts base field name, e.g., 'Profit' from 'SUM(Orders.Profit)' or 'Orders.Profit'.
    """
    if not isinstance(field_full_path, str):
        return None
    
    # First, remove aggregation if present: SUM(Orders.Profit) -> Orders.Profit
    agg_match = re.search(r'\((.*?)\)', field_full_path)
    field_without_agg = agg_match.group(1).strip() if agg_match else field_full_path.strip()
    
    # Then, get the base name without dataset prefix: Orders.Profit -> Profit
    if '.' in field_without_agg:
        return field_without_agg.split('.')[-1]
    return field_without_agg

def get_clean_agg_name(agg_func_str: str) -> str:
    """Converts a common aggregation string to 'Sum', 'Average' etc., for NativeReferenceName."""
    clean_map = {
        "Sum": "Sum", "Average": "Average", "Min": "Min", "Max": "Max", 
        "Count": "Count", "CountD": "Count of Distinct", "DistinctCount": "Count of Distinct"
    }
    match = re.match(r'(\w+)\(.*\)', agg_func_str, re.IGNORECASE)
    if match:
        func_name = match.group(1).capitalize() 
        return clean_map.get(func_name, func_name) 
    return "Value" # For a bare field assumed to be summed or just a value

def get_pbi_agg_function_id(field_str: str) -> int:
    """
    Maps common aggregation strings to their numerical IDs in Power BI PQL.
    Function ID 0 is typically Sum.
    """
    agg_map = {"Sum": 0, "Avg": 1, "Min": 2, "Max": 3, "Count": 4, "CountD": 5, "DistinctCount": 5}
    match = re.match(r'(\w+)\((.*?)\)', field_str, re.IGNORECASE) 
    if match:
        agg_name = match.group(1).capitalize() 
        return agg_map.get(agg_name, 0) 
    return 0 

# --- Dataframe/CSV specific Helpers ---
def clean_dataset_filename_for_reference(original_name: str) -> str:
    """
    Cleans dataset names extracted from filenames by removing UUID-like strings and '_data' suffix.
    e.g., 'Orders_3759F66AE19340B5A44DC7B40426AAA0_data' -> 'Orders'
    """
    pattern = r"_[0-9a-fA-F]{32}_data$"
    match = re.search(pattern, original_name)
    if match:
        cleaned_name = original_name[:match.start()]
        print(f"DEBUG: Cleaned dataset name from '{original_name}' to '{cleaned_name}'")
        return cleaned_name
    return original_name 

def is_likely_identifier_or_category_column(col_name: str) -> bool:
    """
    Heuristic to check if a column name implies it should remain string/object,
    even if it contains digits. Prevents aggressive numeric coercion for IDs, codes, etc.
    """
    lower_col = col_name.lower()
    common_identifier_keywords = ['id', 'code', 'zip', 'postal', 'number', 'identifier', 'sku', 'isbn', 'account', 'phone', 'key']
    common_metric_keywords = ['sale', 'profit', 'quantity', 'amount', 'cost', 'revenue', 'value'] # These should be numeric

    is_id_like = any(keyword in lower_col for keyword in common_identifier_keywords)
    is_metric_like = any(metric_word in lower_col for metric_word in common_metric_keywords)
    
    return is_id_like and not is_metric_like

# --- Visual Config Structure / Styling Helpers ---
def extract_first_json_object_from_string(text_fragment: str) -> dict or None:
    """
    Attempts to extract the first well-formed JSON object from a given string fragment.
    """
    start_brace_index = -1
    balance = 0
    in_string = False
    escaped = False

    for i, char in enumerate(text_fragment):
        if char == '{':
            start_brace_index = i
            break
    
    if start_brace_index == -1:
        return None

    for i in range(start_brace_index, len(text_fragment)):
        char = text_fragment[i]

        if char == '\\':
            escaped = not escaped
        elif char == '"' and not escaped:
            in_string = not in_string
        elif not in_string:
            if char == '{':
                balance += 1
            elif char == '}':
                balance -= 1
        
        if escaped and char != '\\':
            escaped = False

        if balance == 0 and start_brace_index != -1 and not in_string:
            json_candidate_str = text_fragment[start_brace_index : i + 1]
            try:
                parsed_json = json.loads(json_candidate_str)
                if isinstance(parsed_json, dict):
                    return parsed_json
            except json.JSONDecodeError:
                pass
            
    return None

def find_bullet_chart_style_sample(reference_text_path: str) -> dict or None:
    """
    Reads the reference text file, finds the 'Bullet Chart Examples:' marker,
    and extracts the first JSON object immediately following it for styling.
    """
    BULLET_CHART_MARKER = "Bullet Chart  Examples:" # As per your file content (two spaces)
    
    try:
        with open(reference_text_path, 'r', encoding="utf-8") as f:
            content = f.read()
        
        print(f"DEBUG: Searching for '{BULLET_CHART_MARKER}' in '{reference_text_path}'...")
        marker_start_index = content.find(BULLET_CHART_MARKER)

        if marker_start_index == -1:
            print(f"DEBUG: Marker '{BULLET_CHART_MARKER}' not found in file.")
            return None

        search_start_pos = marker_start_index + len(BULLET_CHART_MARKER)
        text_after_marker = content[search_start_pos:]
        print(f"DEBUG: Marker found. Searching for JSON from char {search_start_pos}. Fragment starts with (first 100 chars):\n'{text_after_marker[:100]}'")

        bullet_style_sample = extract_first_json_object_from_string(text_after_marker)

        if bullet_style_sample:
            print(f"DEBUG: Potentially found bullet style sample. Outer keys: {bullet_style_sample.keys()}")
            
            if 'config' in bullet_style_sample:
                try:
                    inner_config = json.loads(bullet_style_sample['config'])
                    if 'singleVisual' in inner_config:
                        single_visual = inner_config['singleVisual']
                        if single_visual.get('visualType') == 'barChart' and \
                           'objects' in single_visual and 'dataPoint' in single_visual['objects']:
                            print("DEBUG: Verified as a Bar Chart (Bullet Style Candidate) object with 'objects' and 'dataPoint'.")
                            return bullet_style_sample 
                except json.JSONDecodeError as e:
                    print(f"DEBUG: Error parsing inner 'config' of identified JSON object: {e}. Skipping.")
        
        print("DEBUG: No suitable Bar Chart (Bullet Style) sample found after marker.")
        return None

    except FileNotFoundError:
        print(f"❌ '{reference_text_path}' not found for style extraction.")
        return None
    except Exception as e:
        print(f"❌ An error occurred during style sample lookup: {e}")
        return None


def extract_style_objects_from_sample(sample_config_obj: dict) -> tuple[dict, dict]:
    """
    Extracts the 'objects' and 'vcObjects' dictionaries from the singleVisual part
    of a sample Power BI visual configuration. Also extracts 'hasDefaultSort' if present.
    """
    print("\n--- DEBUG: Inside extract_style_objects_from_sample (after finding candidate) ---")
    
    inner_config = {}
    try:
        inner_config = json.loads(sample_config_obj['config'])
        print(f"DEBUG: Successfully parsed inner config. Keys: {inner_config.keys()}")
    except json.JSONDecodeError as e:
        print(f"❌ DEBUG Error parsing inner 'config' string from sample_config_obj: {e}")
        print("--- DEBUG: Exiting extract_style_objects_from_sample (inner parse failed) ---\n")
        return {}, {} 

    single_visual_config = inner_config.get('singleVisual', {})
    print(f"DEBUG: single_visual_config keys: {single_visual_config.keys()}")

    extracted_objects = single_visual_config.get('objects', {})
    extracted_vc_objects = single_visual_config.get('vcObjects', {})
    
    if 'hasDefaultSort' in single_visual_config:
        extracted_vc_objects['_hasDefaultSort_flag'] = single_visual_config['hasDefaultSort']
    
    print(f"DEBUG: Extracted objects (before copy): {len(extracted_objects)} entries. Keys: {list(extracted_objects.keys())}")
    print(f"DEBUG: Extracted vcObjects (before copy): {len(extracted_vc_objects)} entries. Keys: {list(extracted_vc_objects.keys())}")
    print("--- DEBUG: Exiting extract_style_objects_from_sample (successful) ---\n")
    
    return deep_copy_json_obj(extracted_objects), deep_copy_json_obj(extracted_vc_objects)


def integrate_style_objects(
    generated_visual_config: dict, 
    extracted_objects: dict, 
    extracted_vc_objects: dict
) -> dict:
    """
    Integrates extracted style objects into a generated Power BI visual configuration.
    """
    config_dict_inner = json.loads(generated_visual_config['config'])
    single_visual = config_dict_inner.get('singleVisual', {})

    if '_hasDefaultSort_flag' in extracted_vc_objects: 
        single_visual['hasDefaultSort'] = extracted_vc_objects.pop('_hasDefaultSort_flag')

    single_visual['objects'] = single_visual.get('objects', {})

    generated_measures = [proj['queryRef'] for proj in single_visual['projections'].get('Y', [])]
    
    sample_data_points = extracted_objects.get('dataPoint', [])
    new_data_points_for_generated = []
    
    default_datapoint_settings = next((dp for dp in sample_data_points if 'selector' not in dp), None)
    if default_datapoint_settings:
        new_data_points_for_generated.append(deep_copy_json_obj(default_datapoint_settings))

    sample_measure_datapoints = [dp for dp in sample_data_points if 'selector' in dp and 'metadata' in dp['selector']]
    
    for i, gen_measure in enumerate(generated_measures):
        if i < len(sample_measure_datapoints):
            dp_style_template = deep_copy_json_obj(sample_measure_datapoints[i])
            dp_style_template['selector']['metadata'] = gen_measure
            new_data_points_for_generated.append(dp_style_template)
        else:
            print(f"  ⚠ Warning: No specific dataPoint style found in sample for measure '{gen_measure}'. "
                  "It will fall back to default or inherited theme colors.")

    if new_data_points_for_generated:
        single_visual['objects']['dataPoint'] = new_data_points_for_generated

    for key, value in extracted_objects.items():
        if key != 'dataPoint': 
            single_visual['objects'][key] = deep_copy_json_obj(value)

    single_visual['vcObjects'] = single_visual.get('vcObjects', {})
    
    generated_title_obj = single_visual['vcObjects'].get('title')
    
    for key, value_from_extracted in extracted_vc_objects.items():
        if key.startswith('_'): 
            continue 

        if isinstance(value_from_extracted, list):
            single_visual['vcObjects'][key] = deep_copy_json_obj(value_from_extracted)
        elif isinstance(value_from_extracted, dict):
            if key not in single_visual['vcObjects']:
                single_visual['vcObjects'][key] = {}
            single_visual['vcObjects'][key].update(deep_copy_json_obj(value_from_extracted))
        else:
            single_visual['vcObjects'][key] = value_from_extracted

    if generated_title_obj:
        if 'title' not in single_visual['vcObjects'] or not isinstance(single_visual['vcObjects']['title'], list) or not single_visual['vcObjects']['title']:
            single_visual['vcObjects']['title'] = deep_copy_json_obj(generated_title_obj)
        else:
            if len(single_visual['vcObjects']['title']) > 0 and 'properties' in single_visual['vcObjects']['title'][0]:
                single_visual_title_entry = single_visual['vcObjects']['title'][0] 
                if 'properties' not in single_visual_title_entry:
                    single_visual_title_entry['properties'] = {}
                if generated_title_obj[0].get('properties', {}).get('text'):
                    single_visual_title_entry['properties']['text'] = generated_title_obj[0]['properties']['text']
                
    generated_visual_config['config'] = json.dumps(config_dict_inner)
    
    return generated_visual_config


# === GEMINI AI INTEGRATION (Global) ===
def get_gemini_title_suggestion(category_field: str, measures: list[str], dataset_name: str) -> str:
    """
    Uses Gemini AI to suggest a more descriptive title for a chart.
    """
    measures_str = " and ".join(measures)
    prompt = (
        f"Generate a very concise (under 10 words) and professional title "
        f"for a Power BI bar chart that visualizes '{measures_str}' by '{category_field}' "
        f"from the '{dataset_name}' dataset. Make it suitable for a dashboard context. "
        f"Example: 'Profit & Sales by State'. Do not include dataset name in title. "
        f"Only return the title, no extra text or quotes."
    )
    
    try:
        print(f"DEBUG: Querying Gemini for title suggestion...")
        response = model.generate_content(prompt)
        suggested_title = response.text.strip().replace("'", "").replace('"', '') 
        print(f"DEBUG: Gemini suggested title: '{suggested_title}'")
        return suggested_title
    except api_exceptions.ResourceExhausted as e:
        print(f"❌ Error: Gemini API rate limit or quota exceeded for title suggestion: {e}")
        time.sleep(2) 
        return "" 
    except Exception as e:
        print(f"❌ Error querying Gemini AI for title: {e}")
        time.sleep(1) 
        return ""

def get_all_available_columns_with_datasets(dataset_map: dict) -> list[str]:
    """Flattens all dataset.column combinations into a list (e.g., ['Orders.Profit', 'Orders.Category'])."""
    available_cols = []
    for dataset_name, df in dataset_map.items():
        for col_name in df.columns:
            available_cols.append(f"{dataset_name}.{col_name}")
    return available_cols

def get_gemini_field_suggestion(
    requested_category_conceptual_name: str, 
    requested_measures_raw_strings: list[str], 
    available_columns_full_names: list[str], 
    chart_worksheet_title: str,
    table_context_for_category_search: str = None 
) -> str:
    """
    Uses Gemini AI to suggest field mapping corrections or explanations when deterministic matching fails.
    """
    measures_display = " and ".join(requested_measures_raw_strings)
    available_cols_str = ", ".join(available_columns_full_names)
    
    if table_context_for_category_search and requested_category_conceptual_name == "Category" and not requested_measures_raw_strings:
        prompt_prefix = (
            f"For a chart named '{chart_worksheet_title}', the conceptual category is '{requested_category_conceptual_name}' from the '{table_context_for_category_search}' table. "
            f"Please suggest the *most appropriate Power BI column name* (fully qualified like 'Dataset.Column Name') "
            f"from the following available columns that semantically represents this 'Category': [{available_cols_str}]. "
            f"Example suggestions: 'Orders.State or Province', 'Products.Category', 'Orders.Segment'. "
            f"Only return the fully qualified column name, no extra text or quotes. "
            f"If no clear category column is apparent from the '{table_context_for_category_search}' table among available columns, state 'No clear category column found in table'."
        )
    else:
        prompt_prefix = (
            f"I am trying to map fields for a chart named '{chart_worksheet_title}'. "
            f"I need a category '{requested_category_conceptual_name}' and measures '{measures_display}'. "
            f"My available data columns are: [{available_cols_str}]. "
            f"I could not find exact matches for my required fields. "
            f"Can you provide a very concise explanation (under 50 words) for why these fields might not be matching, "
            f"or suggest which of the available columns are the closest semantic matches "
            f"for my required category and measures? Format suggestions as 'Category: [suggestion], Measures: [suggestion1], [suggestion2]'. "
            f"If no good match, clearly clearly state 'No clear semantic match found based on available data.' "
        )
    
    prompt = prompt_prefix + " Do not ask questions, just provide the analysis or suggestions."

    try:
        print(f"\nDEBUG: Querying Gemini for field mapping suggestion for '{chart_worksheet_title}'...")
        response = model.generate_content(prompt)
        suggestion = response.text.strip()
        print(f"DEBUG: Gemini Field Suggestion: \n{suggestion}\n")
        return suggestion
    except api_exceptions.ResourceExhausted as e:
        print(f"❌ Error: Gemini API rate limit or quota exceeded for field suggestion: {e}")
        time.sleep(5) 
        return "Gemini could not provide suggestions due to API limits. Try again later."
    except Exception as e:
        print(f"❌ Error querying Gemini AI for field suggestion: {e}")
        time.sleep(2)
        return "Gemini AI failed to provide a suggestion."


# === Extractors for final_json.json (specific to new format) ===
def extract_bullet_fields(chart_entry: dict) -> dict:
    """
    Extracts chart components from the new abstract final_json.json format.
    """
    conceptual_category_field = None
    category_entity_table = None

    rows_data = chart_entry.get("Rows", {})
    if rows_data and isinstance(rows_data, dict):
        first_key = next(iter(rows_data), None) 
        if first_key:
            conceptual_category_field = first_key 
            category_entity_table = rows_data.get(first_key) 

    measures_aggregation_strings = chart_entry.get("Aggregation_columns", [])
    if not isinstance(measures_aggregation_strings, list):
        measures_aggregation_strings = [measures_aggregation_strings] if measures_aggregation_strings else []

    chart_display_title = chart_entry.get("title", "Untitled Chart")
    
    legend_field_details = None
    legend_data = chart_entry.get("Legend")
    if isinstance(legend_data, dict) and legend_data:
        first_legend_field = next(iter(legend_data), None)
        if first_legend_field:
            legend_field_details = {
                'field': first_legend_field,
                'table': legend_data.get(first_legend_field)
            }
    elif isinstance(legend_data, str):
        if legend_data in chart_entry.get("Columns", {}):
            legend_field_details = {'field': legend_data, 'table': chart_entry["Columns"][legend_data]}
        elif legend_data in chart_entry.get("Rows", {}):
            legend_field_details = {'field': legend_data, 'table': chart_entry["Rows"][legend_data]}
        else:
             print(f"DEBUG: Legend '{legend_data}' found as string but could not deduce its table from Columns/Rows. Skipping Series projection.")


    return {
        "conceptual_category_field": conceptual_category_field,
        "category_entity_table": category_entity_table,
        "measures_aggregation_strings": measures_aggregation_strings,
        "chart_display_title": chart_display_title,
        "legend_field_details": legend_field_details
    }


# === Load Datasets ===
def load_all_datasets():
    """
    Load all CSV datasets from the specified directory and ensure numeric columns are correctly typed.
    Dataset names are cleaned from UUIDs/suffixes.
    """
    if not os.path.exists(CSV_OUTPUT_DIR):
        print(f"❌ CSV folder not found at: {CSV_OUTPUT_DIR}.")
        return {}

    dataset_map = {}
    print("\n📌 Loading Datasets...")
    for file in os.listdir(CSV_OUTPUT_DIR):
        if file.endswith(".csv"):
            raw_dataset_name = os.path.splitext(file)[0].strip()
            cleaned_dataset_name = clean_dataset_filename_for_reference(raw_dataset_name)
            
            file_path = os.path.join(CSV_OUTPUT_DIR, file)
            try:
                df = pd.read_csv(file_path, encoding='utf-8-sig', on_bad_lines='skip')
                
                for col in df.columns:
                    if pd.api.types.is_object_dtype(df[col]):
                        if is_likely_identifier_or_category_column(col):
                            print(f"  ℹ️ Keeping column '{col}' in '{raw_dataset_name}.csv' as string type (likely an identifier/category).")
                            continue 
                        
                        original_dtype = df[col].dtype
                        converted_col = pd.to_numeric(df[col], errors='coerce')
                        
                        if pd.api.types.is_numeric_dtype(converted_col) and not converted_col.isnull().all():
                            if not df[col].equals(converted_col): 
                                df[col] = converted_col
                                if pd.isna(df[col]).sum() > 0:
                                    num_nans = pd.isna(df[col]).sum()
                                    print(f"  ⚠ Warning: Column '{col}' in '{raw_dataset_name}.csv' ({original_dtype}) converted to numeric, with {num_nans} non-numeric values now NaN.")
                                else:
                                    print(f"  ✅ Converted column '{col}' in '{raw_dataset_name}.csv' to numeric.")
                        else:
                             print(f"  ℹ️ Column '{col}' in '{raw_dataset_name}.csv' cannot be safely converted to numeric; keeping as original type.")

                if cleaned_dataset_name not in dataset_map:
                    dataset_map[cleaned_dataset_name] = df
                    print(f"  📂 Loaded '{cleaned_dataset_name}' (from '{raw_dataset_name}')")
                else:
                    print(f"  ⚠ Duplicate cleaned dataset name '{cleaned_dataset_name}' found. Skipping '{raw_dataset_name}'.")

            except Exception as e:
                print(f"❌ Error reading {file}: {e}")

    return dataset_map

# === Main Generation Function ===
def generate_bullet_visuals():
    """Generates Power BI visual configurations for Bullet Charts."""
    dataset_map = load_all_datasets()
    if not dataset_map:
        print("\n⚠ No datasets loaded. Exiting.")
        return

    positions_file = os.path.join(OUTPUT_DIR, "powerbi_chart_positions.json")
    final_json_file = os.path.join(OUTPUT_DIR, "final.json")
    
    try:
        with open(positions_file, 'r', encoding="utf-8") as f:
            chart_positions_raw = json.load(f)
            # FIX: Ensure chart_positions is always a list of dictionaries
            if not isinstance(chart_positions_raw, list):
                chart_positions = [chart_positions_raw]
            else:
                chart_positions = chart_positions_raw
    except FileNotFoundError as e:
        print(f"❌ Required file not found: {e}. Exiting.")
        return
    except json.JSONDecodeError as e:
        print(f"❌ Error decoding JSON from '{e.doc}': {e}. Exiting.")
        return

    try:
        with open(final_json_file, 'r', encoding="utf-8") as f:
            final_json_data = json.load(f) 
    except FileNotFoundError as e:
        print(f"❌ Required file not found: {e}. Exiting.")
        return
    except json.JSONDecodeError as e:
        print(f"❌ Error decoding JSON from '{e.doc}': {e}. Exiting.")
        return

    if isinstance(final_json_data, dict):
        final_json_data = [final_json_data] 

    bullet_charts = [e for e in final_json_data if "bullet" in normalize_name(e.get("chart_type", ""))]
    if not bullet_charts:
        print("\n⚠ No Bullet charts found in final.json.")
        return
    
    print(f"\n✅ Found {len(bullet_charts)} bullet chart(s) to process.")
    new_visuals = []
    z_coordinate = 0

    extracted_objects = {}
    extracted_vc_objects = {}
    
    bullet_style_sample = find_bullet_chart_style_sample(REFERENCE_TEXT_PATH)

    if bullet_style_sample:
        extracted_objects, extracted_vc_objects = extract_style_objects_from_sample(bullet_style_sample)
        if extracted_objects or extracted_vc_objects:
            print(f"\n✅ Successfully identified and loaded style properties for Bullet Charts from '{REFERENCE_TEXT_PATH}'.")
        else:
            print(f"\nℹ️ Although a matching visual config was found in '{REFERENCE_TEXT_PATH}', it did not contain expected styling properties ('objects'/'vcObjects'). Styling will not be applied from a sample.")

    else:
        print(f"\n❌ Warning: No specific Bar Chart (Bullet Style) visual config with detailed styling found in '{REFERENCE_TEXT_PATH}'. Visual styling (colors, etc.) will not be applied from a sample.")

    all_available_columns_flat_list = get_all_available_columns_with_datasets(dataset_map)

    # --- Process each bullet chart entry ---
    for chart_entry in bullet_charts:
        worksheet_id = f"bullet_chart_{z_coordinate}" 
        
        extracted_fields = extract_bullet_fields(chart_entry)
        
        conceptual_category_field = extracted_fields["conceptual_category_field"]
        category_entity_table_name_raw = extracted_fields["category_entity_table"] 
        measures_aggregation_strings = extracted_fields["measures_aggregation_strings"]
        chart_display_title = extracted_fields["chart_display_title"]
        legend_field_details = extracted_fields["legend_field_details"] 
        
        worksheet_for_pos_lookup = chart_entry.get("worksheet", chart_display_title)


        if not all([conceptual_category_field, category_entity_table_name_raw, measures_aggregation_strings, len(measures_aggregation_strings) >= 2]):
            print(f"⚠ Skipping '{chart_display_title}' ({worksheet_id}): Missing conceptual category, its table, or two measure fields.")
            continue

        dataset_name = None
        current_df_for_chart = None 
        for ds_name, df in dataset_map.items(): 
            if normalize_name(ds_name) == normalize_name(category_entity_table_name_raw):
                dataset_name = ds_name
                current_df_for_chart = df
                break
        
        if not dataset_name:
            print(f"❌ Dataset '{category_entity_table_name_raw}' not loaded/found for chart '{chart_display_title}'.")
            ai_dataset_suggestion = get_gemini_field_suggestion(
                requested_category_conceptual_name=f"conceptual: {conceptual_category_field} from {category_entity_table_name_raw}",
                requested_measures_raw_strings=measures_aggregation_strings,
                available_columns_full_names=all_available_columns_flat_list,
                chart_worksheet_title=chart_display_title
            )
            print(f"    AI Dataset Search Result: {ai_dataset_suggestion}")
            continue
            
        print(f"✅ Matched '{chart_display_title}' to cleaned dataset '{dataset_name}' (from raw '{category_entity_table_name_raw}')")

        resolved_category_pbi_column = None
        
        if conceptual_category_field and conceptual_category_field in current_df_for_chart.columns:
            resolved_category_pbi_column = conceptual_category_field
            print(f"DEBUG: Found direct conceptual category column: '{resolved_category_pbi_column}'.")
        
        if not resolved_category_pbi_column:
            common_category_alternatives = [
                conceptual_category_field, 
                "State", "State/Province", "Segment", "Region", "Country", "City", 
                "Category Name", "Product Category", "Customer Segment", "Product Name", "Type"
            ]
            
            combined_alternatives_set = set(common_category_alternatives)
            combined_alternatives_list = [c for c in common_category_alternatives if c] 
            
            for col in current_df_for_chart.columns:
                if col not in combined_alternatives_set: 
                    combined_alternatives_list.append(col)
            
            for alt_col in combined_alternatives_list:
                if alt_col and alt_col in current_df_for_chart.columns: 
                    resolved_category_pbi_column = alt_col
                    print(f"DEBUG: Auto-detected category column: '{resolved_category_pbi_column}' in '{dataset_name}'.")
                    break
        
        if not resolved_category_pbi_column:
            print(f"⚠ Could not deterministically find a specific category column for conceptual '{conceptual_category_field}' in '{dataset_name}'.")
            available_cols_from_this_dataset = [f"{dataset_name}.{c}" for c in current_df_for_chart.columns]
            ai_category_suggestion_raw = get_gemini_field_suggestion(
                requested_category_conceptual_name=conceptual_category_field,
                requested_measures_raw_strings=[], 
                available_columns_full_names=available_cols_from_this_dataset,
                chart_worksheet_title=f"category field for '{chart_display_title}'",
                table_context_for_category_search=dataset_name
            )
            
            if "No clear category column found" in ai_category_suggestion_raw or "Gemini AI failed" in ai_category_suggestion_raw:
                print(f"❌ AI also could not suggest a category column for '{chart_display_title}'. Skipping chart.")
                continue
            
            inferred_col_from_ai_base_name = extract_base_field_name(ai_category_suggestion_raw)
            if inferred_col_from_ai_base_name and inferred_col_from_ai_base_name in current_df_for_chart.columns:
                resolved_category_pbi_column = inferred_col_from_ai_base_name
                print(f"✅ AI suggested and validated category column: '{resolved_category_pbi_column}'.")
            else:
                print(f"❌ AI suggested category '{ai_category_suggestion_raw}', but it could not be resolved in '{dataset_name}'. Skipping.")
                continue

        measure_details_for_pql_gen = [] 
        resolved_measure_pql_strings = [] 

        skip_chart_due_to_measure = False
        for i, measure_agg_str_raw in enumerate(measures_aggregation_strings):
            base_f = extract_base_field_name(measure_agg_str_raw)
            
            if base_f not in current_df_for_chart.columns:
                print(f"❌ Measure column '{base_f}' (from '{measure_agg_str_raw}') NOT found in dataset '{dataset_name}'. Skipping chart '{chart_display_title}'.")
                skip_chart_due_to_measure = True
                break

            agg_func_id = get_pbi_agg_function_id(measure_agg_str_raw)
            
            agg_type_from_str = re.match(r'(\w+)\(.*\)', measure_agg_str_raw, re.IGNORECASE)
            agg_prefix = agg_type_from_str.group(1).capitalize() if agg_type_from_str else 'Sum' 
            pql_query_ref_string = f"{agg_prefix}({dataset_name}.{base_f})"
            
            clean_agg_for_native = get_clean_agg_name(measure_agg_str_raw)
            pql_native_ref_name = f"{clean_agg_for_native} of {base_f}"

            measure_details_for_pql_gen.append({
                'resolved_base_field': base_f,
                'agg_func_id': agg_func_id,
                'pql_query_ref_string': pql_query_ref_string,
                'pql_native_ref_name': pql_native_ref_name
            })
            resolved_measure_pql_strings.append(pql_query_ref_string)

        if skip_chart_due_to_measure:
            continue
        
        # Determine position
        def flexible_position_match(pos):
            # Normalize all possible keys and values for robust matching
            def norm(val):
                return normalize_name(str(val)) if val is not None else ""
            chart_keys = ["chart", "worksheet", "chart_type"]
            chart_values = [worksheet_for_pos_lookup, chart_display_title, chart_entry.get("chart_type", "")] 
            for key in chart_keys:
                pos_val = pos.get(key, "")
                for val in chart_values:
                    if norm(pos_val) == norm(val):
                        return True
            return False

        chart_position = next((pos for pos in chart_positions if flexible_position_match(pos)), None)

        if not chart_position:
            chart_position = {"x": 32.0, "y": 305.1, "width": 1240.32, "height": 406.8}
            print(f"⚠ No specific position found in powerbi_chart_positions.json for chart '{worksheet_for_pos_lookup}' (chart 'bullet'). Using default position.")
        
        position_x = chart_position.get("x", 32.0)
        position_y = chart_position.get("y", 305.1)
        position_width = chart_position.get("width", 1240.32)
        position_height = chart_position.get("height", 406.8)

        # --- PROGRAMMATICALLY BUILD THE PROTOTYPE QUERY ---
        alias = dataset_name[0].lower() 
        
        select_clauses = []
        select_clauses.append({
            "Column": {"Expression": {"SourceRef": {"Source": alias}}, "Property": resolved_category_pbi_column},
            "Name": f"{dataset_name}.{resolved_category_pbi_column}", 
            "NativeReferenceName": extract_base_field_name(resolved_category_pbi_column) 
        })

        for md in measure_details_for_pql_gen:
            select_clauses.append({
                "Aggregation": {"Expression": {"Column": {"Expression": {"SourceRef": {"Source": alias}}, "Property": md['resolved_base_field']}}, "Function": md['agg_func_id']},
                "Name": md['pql_query_ref_string'], 
                "NativeReferenceName": md['pql_native_ref_name']
            })
        
        if len(measure_details_for_pql_gen) > 1:
            second_measure_for_order_pql = measure_details_for_pql_gen[1]
            order_by_expression = {
                "Aggregation": {
                    "Expression": {
                        "Column": {
                            "Expression": {"SourceRef": {"Source": alias}},
                            "Property": second_measure_for_order_pql['resolved_base_field']
                        }
                    },
                    "Function": second_measure_for_order_pql['agg_func_id']
                }
            }
            prototype_query_orderby = [{"Direction": 2, "Expression": order_by_expression}] 
        else:
            first_measure_for_order_pql = measure_details_for_pql_gen[0]
            order_by_expression = {
                "Aggregation": {
                    "Expression": {
                        "Column": {
                            "Expression": {"SourceRef": {"Source": alias}},
                            "Property": first_measure_for_order_pql['resolved_base_field']
                        }
                    },
                    "Function": first_measure_for_order_pql['agg_func_id']
                }
            }
            prototype_query_orderby = [{"Direction": 2, "Expression": order_by_expression}]
            print("⚠ Warning: Less than two measures for OrderBy. Defaulting to sort by the first measure.")


        prototype_query = {
            "Version": 2,
            "From": [{"Name": alias, "Entity": dataset_name, "Type": 0}], 
            "Select": select_clauses,
            "OrderBy": prototype_query_orderby 
        }
        
        suggested_title = get_gemini_title_suggestion(
            category_field=resolved_category_pbi_column, 
            measures=resolved_measure_pql_strings, 
            dataset_name=dataset_name 
        )
        final_chart_title = suggested_title if suggested_title else chart_display_title 

        # --- Build base_config_object ---
        base_config_object = {
            "name": generate_unique_id(),
            "layouts": [{"id": 0, "position": {"x": position_x, "y": position_y, "z": z_coordinate, "width": position_width, "height": position_height, "tabOrder": z_coordinate*100}}],
            "singleVisual": {
                "visualType": "barChart", 
                "projections": {
                    "Category": [{"queryRef": f"{dataset_name}.{resolved_category_pbi_column}", "active": True}], 
                    "Y": [{"queryRef": m} for m in resolved_measure_pql_strings], 
                },
                "prototypeQuery": prototype_query,
                "drillFilterOtherVisuals": True,
                "vcObjects": {
                    "title": [{"properties": {"text": {"expr": {"Literal": {"Value": f"'{final_chart_title}'"}}}}}] 
                }
            }
        }
        
        # --- MODIFIED: Conditional Series Projection ---
        # Add Series Projection only if the visual is *not* a barChart (bullet-simulated)
        # or if it's a bar chart specifically configured to handle a Series (unlikely for a simple bullet)
        # Your Power BI sample (the bullet chart reference) does NOT have a Series field.
        # So, to avoid "too many columns in Legend" error for bullet charts, we omit Series if it's a barChart.
        if base_config_object["singleVisual"]["visualType"] != "barChart" and legend_field_details and legend_field_details['field'] and legend_field_details['table']:
             legend_table_cleaned = clean_dataset_filename_for_reference(legend_field_details['table'])
             legend_col_base_name = extract_base_field_name(legend_field_details['field'])
             
             legend_df = dataset_map.get(legend_table_cleaned)
             if legend_df is not None and legend_col_base_name in legend_df.columns:
                 series_query_ref = f"{legend_table_cleaned}.{legend_col_base_name}"
                 if "Series" not in base_config_object["singleVisual"]["projections"]: 
                      base_config_object["singleVisual"]["projections"]["Series"] = []
                 base_config_object["singleVisual"]["projections"]["Series"].append({"queryRef": series_query_ref})
                 print(f"✅ Added Series projection from Legend: '{series_query_ref}'")
             else:
                 print(f"⚠ Legend field '{legend_col_base_name}' from table '{legend_field_details['table']}' could not be resolved. Skipping Series projection for {chart_display_title}.")
        elif base_config_object["singleVisual"]["visualType"] == "barChart" and (legend_field_details and legend_field_details['field'] and legend_field_details['table']):
            # This branch catches when a bullet chart *would* have a legend but we're intentionally skipping it.
            print(f"ℹ️ Skipping Series projection for '{chart_display_title}' (barChart/bullet) as it typically conflicts with category/Y roles in a simple bar visual, aligning with sample. Check Power BI field well limits if desired otherwise.")
        # --- END MODIFIED Series Projection ---

        print(f"✅ Generated base config for {chart_display_title}.")

        temp_visual_output_format = {
            "config": json.dumps(base_config_object),
            "filters": "[]",
            "height": position_height, "width": position_width,
            "x": position_x, "y": position_y, "z": z_coordinate
        }

        if extracted_objects or extracted_vc_objects:
            final_visual_config_for_entry = integrate_style_objects(
                temp_visual_output_format, 
                extracted_objects, 
                extracted_vc_objects
            )
            print(f"✅ Applied styles to config for {chart_display_title}.")
        else:
            final_visual_config_for_entry = temp_visual_output_format
            print(f"ℹ️ No styling applied for {chart_display_title} (no valid style properties found in sample data from '{REFERENCE_TEXT_PATH}').")

        new_visuals.append(final_visual_config_for_entry)
        z_coordinate += 1

    if new_visuals:
        save_to_json(new_visuals, OUTPUT_DIR, "bullet_visuals_output.json")
        print(f"\n🎉 Successfully generated {len(new_visuals)} bullet chart configurations.")
    else:
        print("\n⚠ No valid bullet chart configurations were generated.")

# Main function for overall script execution
def run_generate_report():
    """Runs the entire script and generates the report visuals."""
    generate_bullet_visuals() 

if __name__ == "__main__":
    run_generate_report()