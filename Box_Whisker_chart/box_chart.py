import os
import json
import pandas as pd
import google.generativeai as genai
import re
import uuid

# === Configuration Paths ===
BASE_DIR = r"C:\Users\ksaik\OneDrive\Desktop\Box_Whisker_chart"
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
CSV_OUTPUT_DIR = os.path.join(OUTPUT_DIR, "csv_output")
REFERENCE_TEXT_PATH = os.path.join(BASE_DIR, "Reference Chart Configurations.txt")

# Validate directories
if not os.path.exists(CSV_OUTPUT_DIR):
    print(f"❌ CSV folder not found at: {CSV_OUTPUT_DIR}")
    os.makedirs(CSV_OUTPUT_DIR, exist_ok=True)
    print(f"✅ Created CSV folder: {CSV_OUTPUT_DIR}")

# === API Key ===
api_key = "your api key"  # Replace with your actual API key
genai.configure(api_key=api_key)
model = genai.GenerativeModel("gemini-2.5-flash")

# === Load Reference Configuration ===
if os.path.exists(REFERENCE_TEXT_PATH):
    with open(REFERENCE_TEXT_PATH, "r", encoding="utf-8") as f:
        reference_text = f.read()
else:
    print("❌ Reference Chart Configurations.txt not found.")
    reference_text = ""

# === Dynamic Dataset Name Cleaning Function ===
def clean_dataset_name(raw_name):
    """
    Cleans the dataset name by removing suffixes like '_3759F66AE19340B5A44DC7B40426AAA0_data'.
    For example, 'Orders_3759F66AE19340B5A44DC7B40426AAA0_data' becomes 'Orders'.
    """
    match = re.match(r'^(.*?)_[\dA-F]{32}_data$', raw_name)
    if match:
        return match.group(1)
    return raw_name

# === Load and Clean Datasets ===
def load_all_datasets():
    """Load and clean CSV datasets from CSV_OUTPUT_DIR."""
    if not os.path.exists(CSV_OUTPUT_DIR):
        print("❌ Dataset folder missing.")
        return {}, {}

    csv_files = [f for f in os.listdir(CSV_OUTPUT_DIR) if f.endswith(".csv")]
    if not csv_files:
        print("❌ No CSVs found.")
        return {}, {}

    dataset_map = {}
    dataset_column_map = {}

    for file in csv_files:
        dataset_name_raw = file.replace(".csv", "").strip()
        dataset_name = clean_dataset_name(dataset_name_raw)
        file_path = os.path.join(CSV_OUTPUT_DIR, file)
        try:
            df = pd.read_csv(file_path, encoding='utf-8-sig', on_bad_lines='skip')
        except Exception:
            df = pd.read_csv(file_path, encoding='latin1', on_bad_lines='skip')
        dataset_map[dataset_name] = df
        for col in df.columns:
            dataset_column_map[col.strip().lower()] = dataset_name

    print("\n📌 Datasets Loaded:")
    for name, df in dataset_map.items():
        print(f"📂 {name} ➜ {list(df.columns)}")

    return dataset_map, dataset_column_map

# === Helper Functions ===
def generate_unique_id():
    """Generate a unique identifier."""
    return uuid.uuid4().hex

def save_to_json(data, output_path, filename):
    """Save data to a JSON file."""
    os.makedirs(output_path, exist_ok=True)
    path = os.path.join(output_path, filename)
    try:
        with open(path, 'w', encoding="utf-8") as file:
            json.dump(data, file, indent=2)
        print(f"✅ Saved: {path}")
    except Exception as e:
        print(f"❌ Save Error: {e}")

def validate_output_with_gemini(generated_output):
    """Validate the generated JSON using Gemini API."""
    prompt = f"""
    Please check the following JSON for a Power BI Box-and-Whisker chart visual:

    {generated_output}

    ✅ Say "Valid" if it's a single object with "name", "layouts", and "singleVisual" keys at the top level, and no nested "config" key.
    ❌ Otherwise, provide a corrected version (keep table names as-is) removing any nested "config" keys and ensuring the structure matches a valid Power BI visual configuration.
    """
    try:
        response = model.generate_content(prompt)
        return response.text.strip()
    except Exception as e:
        print(f"❌ Gemini Validation Error: {e}")
        return ""

def generate_box_whisker_visuals():
    """Generate Box-and-Whisker chart configurations dynamically."""
    dataset_map, dataset_column_map = load_all_datasets()
    if not dataset_map:
        print("⚠ No datasets loaded. Exiting.")
        return

    try:
        positions_file = os.path.join(OUTPUT_DIR, "powerbi_chart_positions.json")
        with open(positions_file, 'r', encoding="utf-8") as f:
            chart_positions = json.load(f)
    except FileNotFoundError:
        print(f"❌ Error: powerbi_chart_positions.json not found.")
        return
    except json.JSONDecodeError:
        print(f"❌ Error: powerbi_chart_positions.json contains invalid JSON.")
        return
    except Exception as e:
        print(f"❌ An unexpected error occurred while reading powerbi_chart_positions.json: {e}")
        return

    try:
        final_json_file = os.path.join(OUTPUT_DIR, "final_json.json")
        with open(final_json_file, 'r', encoding="utf-8") as f:
            final_json = json.load(f)
    except FileNotFoundError:
        print(f"❌ Error: final_json.json not found.")
        return
    except json.JSONDecodeError:
        print(f"❌ Error: final_json.json contains invalid JSON.")
        return
    except Exception as e:
        print(f"❌ An unexpected error occurred while reading final_json.json: {e}")
        return

    new_visuals = []
    z_coordinate = 0

    box_charts = [entry for entry in final_json if entry.get("chart_type") == "boxWhiskerPlot"]
    if not box_charts:
        print("⚠ No Box-and-Whisker charts found in final_json.json.")
        return

    for i, chart_entry in enumerate(box_charts):
        worksheet = chart_entry.get("Source", "Unknown")
        
        # Extract keys from Columns and Rows objects
        columns_data = chart_entry.get("Columns", {})
        rows_data = chart_entry.get("Rows", {})
        legend_data = chart_entry.get("Legend", {})

        axis = list(columns_data.keys())[0] if columns_data else None
        value = list(rows_data.keys())[0] if rows_data else None
        
        # The legend can be used for color or detail
        legend_field = list(legend_data.keys())[0] if legend_data else None

        if not all([axis, value]):
            print(f"⚠ Missing 'Columns' or 'Rows' data for {worksheet}. Skipping.")
            continue

        category_field = legend_field # Use the legend field for the main category
        
        fields = [axis, value]
        if category_field:
            fields.append(category_field)

        dataset_name = None
        # Determine the dataset from the values of the Rows, Columns, Legend objects
        all_sources = list(columns_data.values()) + list(rows_data.values()) + list(legend_data.values())
        if all_sources:
            # Assume all fields come from the same dataset, take the first one
            dataset_name = all_sources[0]

        if not dataset_name or dataset_name not in dataset_map:
            print(f"❌ Could not determine a valid dataset for {worksheet}. Skipping.")
            continue

        # Check if all fields exist in the determined dataset
        if not all(field in dataset_map[dataset_name].columns for field in fields):
             print(f"❌ Not all fields ({fields}) found in dataset '{dataset_name}' for worksheet {worksheet}. Skipping.")
             continue
        
        if i >= len(chart_positions):
            print(f"⚠ No position available for {worksheet} (index {i}). Skipping.")
            continue
        # In the original file, 'w' and 'h' are used for width/height in powerbi_chart_positions.json
        # but the script was looking for 'width' and 'height'. Let's check for both.
        chart_pos = chart_positions[i]
        chart_w = chart_pos.get('width', chart_pos.get('w'))
        chart_h = chart_pos.get('height', chart_pos.get('h'))
        chart_x = chart_pos.get('x')
        chart_y = chart_pos.get('y')

        if not all([chart_x, chart_y, chart_w, chart_h]):
             print(f"⚠ Position data is incomplete for {worksheet} (index {i}). Skipping.")
             continue

        unique_name = generate_unique_id()
        title = worksheet

        max_attempts = 3
        for attempt in range(max_attempts):
            print(f"\n🌀 Generating config for {worksheet} (Attempt {attempt + 1})...")
            prompt = f"""
            Create a Power BI Box-and-Whisker chart visual configuration in valid JSON format.

            Inputs:
            - Dataset name: '{dataset_name}'
            - xCategoryParent: '{dataset_name}.{axis}'  # X Axis
            {f"- category: '{dataset_name}.{category_field}'" if category_field else ""}
            - measure: 'Sum({dataset_name}.{value})'  # Y Axis
            - Position: x={chart_x}, y={chart_y}, z={z_coordinate}
            - Size: width={chart_w}, height={chart_h}
            - Visual Name: {unique_name}
            - Chart Type: BoxandWhiskerByMAQ1823AD39DT234AB532063E128AX

            **Important**: Use '{dataset_name}' as the entity name in the JSON configuration.

            Here is an example of the expected JSON structure:
            ```json
            {{
                "name": "{unique_name}",
                "layouts": [
                    {{
                        "id": 0,
                        "position": {{
                            "x": {chart_x},
                            "y": {chart_y},
                            "z": {z_coordinate},
                            "width": {chart_w},
                            "height": {chart_h},
                            "tabOrder": 0
                        }}
                    }}
                ],
                "singleVisual": {{
                    "visualType": "BoxandWhiskerByMAQ1823AD39DT234AB532063E128AX",
                    "projections": {{
                        "category": [
                            {{
                                "queryRef": "{dataset_name}.{category_field if category_field else 'default_field'}"
                            }}
                        ],
                        "xCategoryParent": [
                            {{
                                "queryRef": "{dataset_name}.{axis}"
                            }}
                        ],
                        "measure": [
                            {{
                                "queryRef": "Sum({dataset_name}.{value})"
                            }}
                        ]
                    }},
                    "prototypeQuery": {{
                        "Version": 2,
                        "From": [
                            {{
                                "Name": "o",
                                "Entity": "{dataset_name}",
                                "Type": 0
                            }}
                        ],
                        "Select": [
                            {{
                                "Column": {{
                                    "Expression": {{
                                        "SourceRef": {{
                                            "Source": "o"
                                        }}
                                    }},
                                    "Property": "{axis}"
                                }},
                                "Name": "{dataset_name}.{axis}",
                                "NativeReferenceName": "{axis}"
                            }},
                            {{
                                "Column": {{
                                    "Expression": {{
                                        "SourceRef": {{
                                            "Source": "o"
                                        }}
                                    }},
                                    "Property": "{category_field if category_field else 'default_field'}"
                                }},
                                "Name": "{dataset_name}.{category_field if category_field else 'default_field'}",
                                "NativeReferenceName": "{category_field if category_field else 'default_field'}"
                            }},
                            {{
                                "Aggregation": {{
                                    "Expression": {{
                                        "Column": {{
                                            "Expression": {{
                                                "SourceRef": {{
                                                    "Source": "o"
                                                }}
                                            }},
                                            "Property": "{value}"
                                        }}
                                    }},
                                    "Function": 0
                                }},
                                "Name": "Sum({dataset_name}.{value})",
                                "NativeReferenceName": "Sum of {value}"
                            }}
                        ]
                    }},
                    "drillFilterOtherVisuals": true,
                    "vcObjects": {{
                        "title": [
                            {{
                                "properties": {{
                                    "text": {{
                                        "expr": {{
                                            "Literal": {{
                                                "Value": "'{title}'"
                                            }}
                                        }}
                                    }}
                                }}
                            }}
                        ]
                    }}
                }}
            }}
            ```

            Please generate a similar JSON object based on the provided inputs.

            Ensure all property names are enclosed in double quotes and the JSON is valid.

            Return the JSON configuration inside a code block:
            ```json
            {{ ... }}
            ```
            """
            response_text = ""
            try:
                response = model.generate_content(prompt)
                response_text = response.text.strip()
            except Exception as e:
                print(f"❌ Gemini API Error on attempt {attempt + 1} for {worksheet}: {e}")
                continue  # Try again

            print(f"🔍 Model Response:\n{response_text}")

            json_match = re.search(r'```json\n(.*?)\n```', response_text, re.DOTALL)
            extracted_json_str = json_match.group(1) if json_match else response_text

            print(f"🔍 Extracted JSON:\n{extracted_json_str}")

            feedback = validate_output_with_gemini(extracted_json_str)
            print(f"🔍 Validation Feedback: {feedback}")

            if "valid" in feedback.lower():
                try:
                    final_config = json.loads(extracted_json_str)
                    
                    # Create the final wrapped structure
                    wrapped_config = {
                        "config": json.dumps(final_config), # Stringify the visual config
                        "filters": "[]", # Default to empty filters
                        "height": chart_h,
                        "width": chart_w,
                        "x": chart_x,
                        "y": chart_y,
                        "z": z_coordinate
                    }
                    new_visuals.append(wrapped_config)
                    z_coordinate += 1
                    break  # Exit loop on success
                except json.JSONDecodeError:
                    print(f"⚠ JSON parsing failed for {worksheet} despite positive validation. Retrying...")
            else:
                # If validation fails, try to use the corrected version if provided
                corrected_match = re.search(r'```json\n(.*?)\n```', feedback, re.DOTALL)
                if corrected_match:
                    corrected_json_str = corrected_match.group(1)
                    try:
                        final_config = json.loads(corrected_json_str)
                        # Create the final wrapped structure for the corrected version
                        wrapped_config = {
                            "config": json.dumps(final_config),
                            "filters": "[]",
                            "height": chart_h,
                            "width": chart_w,
                            "x": chart_x,
                            "y": chart_y,
                            "z": z_coordinate
                        }
                        new_visuals.append(wrapped_config)
                        z_coordinate += 1
                        break # Exit loop on success
                    except json.JSONDecodeError:
                        print(f"⚠ JSON parsing failed for corrected output on {worksheet}. Retrying...")
                else:
                    print(f"⚠ Validation failed for {worksheet}. Retrying...")
        else:
            print(f"❌ Failed to generate a valid config for {worksheet} after {max_attempts} attempts.")

    if new_visuals:
        save_to_json(new_visuals, OUTPUT_DIR, "box_whisker_visuals_output.json")
    else:
        print("\n⚠ No visuals were generated.")

if __name__ == "__main__":
    generate_box_whisker_visuals()            


    

