# # import xml.etree.ElementTree as ET
# # import json

# # # def is_bullet_chart_from_panes(worksheet, ns):
# # #     panes = worksheet.findall(".//t:pane", ns)
# # #     bar_panes = 0
# # #     for pane in panes:
# # #         marks = pane.findall(".//t:mark", ns)
# # #         encodings = pane.find("t:encodings", ns)

# # #         has_bar = any(m.get("class", "").lower() == "bar" for m in marks)
# # #         has_text = encodings is not None and encodings.find("t:text", ns) is not None
# # #         has_color = encodings is not None and encodings.find("t:color", ns) is not None

# # #         if has_bar and (has_text or has_color):
# # #             bar_panes += 1

# # #     return bar_panes >= 2

# # # def extract_bullet_charts_from_twb(twb_path):
# # #     tree = ET.parse(twb_path)
# # #     root = tree.getroot()

# # #     # Extract namespace
# # #     ns_uri = root.tag.split("}")[0].strip("{")
# # #     ns = {'t': ns_uri}

# # #     # Map worksheet name to bullet chart status
# # #     bullet_worksheets = {}
# # #     for worksheet in root.findall(".//t:worksheet", ns):
# # #         name = worksheet.get("name")
# # #         if is_bullet_chart_from_panes(worksheet, ns):
# # #             bullet_worksheets[name] = True

# # #     # Now scan dashboards to find those that use bullet charts
# # #     dashboards = []
# # #     for dashboard in root.findall(".//t:dashboard", ns):
# # #         dashboard_name = dashboard.get("name")
# # #         zones = dashboard.findall(".//t:zone", ns)
# # #         for zone in zones:
# # #             sheet_name = zone.get("name")
# # #             if sheet_name in bullet_worksheets:
# # #                 dashboards.append({
# # #                     "dashboard": dashboard_name,
# # #                     "worksheet": sheet_name,
# # #                     "x": int(zone.get("x", 0)),
# # #                     "y": int(zone.get("y", 0)),
# # #                     "w": int(zone.get("w", 0)),
# # #                     "h": int(zone.get("h", 0))
# # #                 })

# # #     return dashboards

# # # # Run the function and save output
# # # if __name__ == "__main__":
# # #     input_path = r"D:\tableau_to_powerbi_migration\tableau_to_powerbi_migration\output\input_tableau_output\Bullat Chart.twb" # <-- Change this to your TWB file
# # #     result = extract_bullet_charts_from_twb(input_path)

# # #     # Save to JSON
# # #     with open("bullet_charts.json", "w") as f:
# # #         json.dump(result, f, indent=2)

# # #     print(f"✅ Extracted {len(result)} bullet charts.")

# import xml.etree.ElementTree as ET
# import json

# # def is_bullet_chart_from_panes(worksheet, ns, prefix):
# #     panes = worksheet.findall(f".//{prefix}pane", ns)
# #     bar_panes = 0
# #     for pane in panes:
# #         marks = pane.findall(f".//{prefix}mark", ns)
# #         encodings = pane.find(f"{prefix}encodings", ns)

# #         has_bar = any(m.get("class", "").lower() == "bar" for m in marks)
# #         has_text = encodings is not None and encodings.find(f"{prefix}text", ns) is not None
# #         has_color = encodings is not None and encodings.find(f"{prefix}color", ns) is not None

# #         if has_bar and (has_text or has_color):
# #             bar_panes += 1

# #     return bar_panes >= 2

# # def extract_bullet_charts_from_twb(twb_path):
# #     tree = ET.parse(twb_path)
# #     root = tree.getroot()

# #     # Detect namespace
# #     if "}" in root.tag:
# #         ns_uri = root.tag.split("}")[0].strip("{")
# #         ns = {'t': ns_uri}
# #         prefix = "t:"
# #     else:
# #         ns = {}
# #         prefix = ""

# #     # Step 1: Identify bullet worksheets
# #     bullet_worksheets = {}
# #     for worksheet in root.findall(f".//{prefix}worksheet", ns):
# #         name = worksheet.get("name")
# #         if is_bullet_chart_from_panes(worksheet, ns, prefix):
# #             bullet_worksheets[name] = True
# #     print("✅ Bullet chart worksheets found:", list(bullet_worksheets.keys()))

# #     # Step 2: Find matching zones in dashboards
# #     dashboards = []
# #     for dashboard in root.findall(f".//{prefix}dashboard", ns):
# #         dashboard_name = dashboard.get("name")
# #         zones = dashboard.findall(f".//{prefix}zone", ns)
# #         for zone in zones:
# #             sheet_name = zone.get("name")
# #             if sheet_name in bullet_worksheets:
# #                 dashboards.append({
# #                     "dashboard": dashboard_name,
# #                     "worksheet": sheet_name,
# #                     "x": int(zone.get("x", 0)),
# #                     "y": int(zone.get("y", 0)),
# #                     "w": int(zone.get("w", 0)),
# #                     "h": int(zone.get("h", 0))
# #                 })

# #     return dashboards

# # import xml.etree.ElementTree as ET
# # import json
# # import re

# # def clean_column_name(col_str):
# #     """Extract clean column name from federated syntax."""
# #     if not col_str:
# #         return None
# #     # Extract the readable part of the field name (e.g., 'Profit' from 'sum:Profit:qk')
# #     match = re.search(r':([^:\]]+):[^:\]]*\]?', col_str)
# #     if match:
# #         return match.group(1)
# #     return col_str
# #     # if not col_str:
# #     #     return None
# #     # match = re.search(r":(.*?):", col_str)
# #     # return match.group(1) if match else col_str
# #     # # match = re.search(r"\.\[(.*?)\]$", col_str)
# #     # # return match.group(1) if match else col_str

# # def clean_column_name(full_col):
# #     if not full_col:
# #         return None
# #     match = re.search(r":(.*?):", full_col)
# #     return match.group(1) if match else full_col
# def clean_column_name(full_col):
#     if not full_col:
#         return None

#     # Remove leading [ and trailing ] if present
#     full_col = full_col.strip("[]")

#     # Handle federated names like [federated.something].[Column]
#     if "].[" in full_col:
#         return full_col.split("].[")[-1].strip("[]")

#     # Handle [:Measure Names] and similar
#     match = re.search(r":(.*?)\]?$", full_col)
#     if match:
#         return match.group(1)

#     # Fallback - return cleaned string
#     return full_col

# def extract_bullet_charts_metadata(twb_path):
#     tree = ET.parse(twb_path)
#     root = tree.getroot()

#     # Handle namespace
#     if "}" in root.tag:
#         ns_uri = root.tag.split("}")[0].strip("{")
#         ns = {"t": ns_uri}
#         prefix = "t:"
#     else:
#         ns = {}
#         prefix = ""

#     bullet_charts = []

#     for worksheet in root.findall(f".//{prefix}worksheet", ns):
#         sheet_name = worksheet.get("name")
#         panes = worksheet.findall(f".//{prefix}pane", ns)

#         for pane in panes:
#             mark_elem = pane.find(f"{prefix}mark", ns)
#             if mark_elem is None or mark_elem.get("class") != "Bar":
#                 continue

#             encodings = pane.find(f"{prefix}encodings", ns)
#             if encodings is None:
#                 continue

#             color = encodings.find(f"{prefix}color", ns)
#             text = encodings.find(f"{prefix}text", ns)
#             size = encodings.find(f"{prefix}size", ns)

#             # Require at least text or color for bullet chart
#             if all(e is None for e in [color, text]):
#                 continue

#             x_axis = pane.get("x-axis-name")
#             y_axis = pane.get("y-axis-name")

#             bullet_charts.append({
#                 "worksheet": sheet_name,
#                 "chart_type": "Bullet Chart",
#                 "mark_type": "Bar",
#                 "X Axis": clean_column_name(x_axis) if x_axis else None,
#                 "Y Axis": clean_column_name(y_axis) if y_axis else None,
#                 "Color": clean_column_name(color.get("column")) if color is not None else None,
#                 "Text": clean_column_name(text.get("column")) if text is not None else None,
#                 "Size": clean_column_name(size.get("column")) if size is not None else None
#             })

#     return bullet_charts

# import xml.etree.ElementTree as ET
# import re
# import json
# if __name__ == "__main__":
#     input_path = r"D:\tableau_to_powerbi_migration\tableau_to_powerbi_migration\output\input_tableau_output\Bullat Chart.twb"
#     result = extract_bullet_charts_metadata(input_path)

#     with open("bullet_charts_metadata.json", "w") as f:
#         json.dump(result, f, indent=2)

#     print(f"✅ Extracted {len(result)} bullet charts.")

import xml.etree.ElementTree as ET
import re
import json

# def clean_column_name(full_col):
#     if not full_col:
#         return None

#     full_col = full_col.strip("[]")

#     if "].[" in full_col:
#         return full_col.split("].[")[-1].strip("[]")

#     match = re.search(r":(.*?)\]?$", full_col)
#     if match:
#         return match.group(1)

#     return full_col

import re

def clean_column_name(full_col):
    if not full_col:
        return None

    # Handle fields like 'sum:Shipping Cost:qk' or 'avg:Sales:qk'
    if ':' in full_col:
        parts = full_col.split(':')
        if len(parts) == 3:
            return parts[1]  # e.g., 'Shipping Cost' from 'sum:Shipping Cost:qk'
        elif len(parts) == 2:
            return parts[1]  # fallback: 'Measure Names' from ':Measure Names'

    # Handle federated format like [federated.xxx].[Column]
    match = re.search(r"\.\[(.*?)\]$", full_col)
    if match:
        return match.group(1)

    # Handle Measure Names or similar inside [:...]
    match = re.search(r":([^:\]]+)\]?$", full_col)
    if match:
        return match.group(1)

    # Default fallback
    return full_col.strip("[]")

def extract_bullet_charts_metadata(twb_path):
    tree = ET.parse(twb_path)
    root = tree.getroot()

    # Handle namespace
    if "}" in root.tag:
        ns_uri = root.tag.split("}")[0].strip("{")
        ns = {"t": ns_uri}
        prefix = "t:"
    else:
        ns = {}
        prefix = ""

    bullet_charts = []

    for worksheet in root.findall(f".//{prefix}worksheet", ns):
        sheet_name = worksheet.get("name")
        panes = worksheet.findall(f".//{prefix}pane", ns)

        for pane in panes:
            mark_elem = pane.find(f"{prefix}mark", ns)
            if mark_elem is None or mark_elem.get("class") != "Bar":
                continue

            encodings = pane.find(f"{prefix}encodings", ns)
            if encodings is None:
                continue

            color = encodings.find(f"{prefix}color", ns)
            text = encodings.find(f"{prefix}text", ns)
            size = encodings.find(f"{prefix}size", ns)

            # Require at least text or color for bullet chart
            if all(e is None for e in [color, text]):
                continue

            x_axis = pane.get("x-axis-name")
            y_axis = pane.get("y-axis-name")

            bullet_charts.append({
                "worksheet": sheet_name,
                "chart_type": "Bullet Chart",
                "mark_type": "Bar",
                "X Axis": clean_column_name(x_axis) if x_axis else None,
                "Y Axis": clean_column_name(y_axis) if y_axis else None,
                "Color": clean_column_name(color.get("column")) if color is not None else None,
                "Text": clean_column_name(text.get("column")) if text is not None else None,
                "Size": clean_column_name(size.get("column")) if size is not None else None
            })

    return bullet_charts

def merge_duplicate_charts(charts):
    merged = {}
    for chart in charts:
        key = (chart["worksheet"], chart["chart_type"])
        if key not in merged:
            merged[key] = {
                **chart,
                "mark_type": [chart["mark_type"]] if chart["mark_type"] else [],
                "X Axis": set([chart["X Axis"]]) if chart["X Axis"] else set(),
                "Y Axis": set([chart["Y Axis"]]) if chart["Y Axis"] else set(),
                "Color": set([chart["Color"]]) if chart["Color"] else set(),
                "Text": set([chart["Text"]]) if chart["Text"] else set(),
                "Size": set([chart["Size"]]) if chart["Size"] else set(),
            }
        else:
            if chart["mark_type"] and chart["mark_type"] not in merged[key]["mark_type"]:
                merged[key]["mark_type"].append(chart["mark_type"])
            for field in ["X Axis", "Y Axis", "Color", "Text", "Size"]:
                value = chart[field]
                if value:
                    merged[key][field].add(value)

    # Convert all sets to list/single value
    for chart in merged.values():
        for field in ["X Axis", "Y Axis", "Color", "Text", "Size"]:
            values = list(chart[field])  # sets -> list
            if not values:
                chart[field] = None
            elif len(values) == 1:
                chart[field] = values[0]
            else:
                chart[field] = values

    return list(merged.values())


# def merge_duplicate_charts(charts):
#     merged = {}
#     for chart in charts:
#         key = (chart["worksheet"], chart["chart_type"])
#         if key not in merged:
#             merged[key] = {
#                 **chart,
#                 "mark_type": [chart["mark_type"]] if chart["mark_type"] else [],
#                 "X Axis": set([chart["X Axis"]] if chart["X Axis"] else set()),
#                 "Y Axis": set([chart["Y Axis"]] if chart["Y Axis"] else set()),
#                 "Color": set([chart["Color"]] if chart["Color"] else []),
#                 "Text": set([chart["Text"]] if chart["Text"] else []),
#                 "Size": set([chart["Size"]] if chart["Size"] else [])
#             }
#         else:
#             if chart["mark_type"] and chart["mark_type"] not in merged[key]["mark_type"]:
#                 merged[key]["mark_type"].append(chart["mark_type"])
#             if chart["Color"]:
#                 merged[key]["Color"].add(chart["Color"])
#             if chart["Text"]:
#                 merged[key]["Text"].add(chart["Text"])
#             if chart["Size"]:
#                 merged[key]["Size"].add(chart["Size"])

#     for chart in merged.values():
#         for field in ["Color", "Text", "Size"]:
#             values = list(chart[field])
#             chart[field] = values[0] if len(values) == 1 else (values if values else None)

#     return list(merged.values())


def deduplicate_positions(positions):
    seen = set()
    unique = []
    for pos in positions:
        key = (pos["worksheet"], pos["dashboard"], pos["x"], pos["y"], pos["w"], pos["h"])
        if key not in seen:
            seen.add(key)
            unique.append(pos)
    return unique

# Power BI canvas size
POWER_BI_WIDTH = 1280
POWER_BI_HEIGHT = 720

# Tableau virtual coordinate system size
TABLEAU_MAX_WIDTH = 100000
TABLEAU_MAX_HEIGHT = 100000

def scale_to_powerbi(x, y, width, height):
    """Scale Tableau dimensions to Power BI dimensions."""
    new_x = (x / TABLEAU_MAX_WIDTH) * POWER_BI_WIDTH
    new_y = (y / TABLEAU_MAX_HEIGHT) * POWER_BI_HEIGHT
    new_width = (width / TABLEAU_MAX_WIDTH) * POWER_BI_WIDTH
    new_height = (height / TABLEAU_MAX_HEIGHT) * POWER_BI_HEIGHT
    return round(new_x, 2), round(new_y, 2), round(new_width, 2), round(new_height, 2)

def is_bullet_chart_from_panes(worksheet, ns, prefix):
    panes = worksheet.findall(f".//{prefix}pane", ns)
    bar_panes = 0
    for pane in panes:
        marks = pane.findall(f".//{prefix}mark", ns)
        encodings = pane.find(f"{prefix}encodings", ns)

        has_bar = any(m.get("class", "").lower() == "bar" for m in marks)
        has_text = encodings is not None and encodings.find(f"{prefix}text", ns) is not None
        has_color = encodings is not None and encodings.find(f"{prefix}color", ns) is not None

        if has_bar and (has_text or has_color):
            bar_panes += 1

    return bar_panes >= 2

def extract_bullet_charts_from_twb(twb_path):
    tree = ET.parse(twb_path)
    root = tree.getroot()

    # Detect namespace
    if "}" in root.tag:
        ns_uri = root.tag.split("}")[0].strip("{")
        ns = {'t': ns_uri}
        prefix = "t:"
    else:
        ns = {}
        prefix = ""

    # Step 1: Identify bullet worksheets
    bullet_worksheets = {}
    for worksheet in root.findall(f".//{prefix}worksheet", ns):
        name = worksheet.get("name")
        if is_bullet_chart_from_panes(worksheet, ns, prefix):
            bullet_worksheets[name] = True
    print("✅ Bullet chart worksheets found:", list(bullet_worksheets.keys()))

    # Step 2: Match dashboard zones and scale dimensions
    results = []
    for dashboard in root.findall(f".//{prefix}dashboard", ns):
        dashboard_name = dashboard.get("name")
        zones = dashboard.findall(f".//{prefix}zone", ns)
        for zone in zones:
            sheet_name = zone.get("name")
            if sheet_name in bullet_worksheets:
                x = int(zone.get("x", 0))
                y = int(zone.get("y", 0))
                w = int(zone.get("w", 0))
                h = int(zone.get("h", 0))
                scaled_x, scaled_y, scaled_w, scaled_h = scale_to_powerbi(x, y, w, h)
                results.append({
                    "dashboard": dashboard_name,
                    "worksheet": sheet_name,
                    "x": scaled_x,
                    "y": scaled_y,
                    "w": scaled_w,
                    "h": scaled_h
                })
                return results  # Return only the first match

    return results


# def extract_bullet_charts_metadata(twb_path):
#     tree = ET.parse(twb_path)
#     root = tree.getroot()

#     # Handle namespace
#     if "}" in root.tag:
#         ns_uri = root.tag.split("}")[0].strip("{")
#         ns = {"t": ns_uri}
#         prefix = "t:"
#     else:
#         ns = {}
#         prefix = ""

#     bullet_charts = []

#     for worksheet in root.findall(f".//{prefix}worksheet", ns):
#         sheet_name = worksheet.get("name")
#         panes = worksheet.findall(f".//{prefix}pane", ns)

#         for pane in panes:
#             mark_elem = pane.find(f"{prefix}mark", ns)
#             if mark_elem is None or mark_elem.get("class") != "Bar":
#                 continue

#             encodings = pane.find(f"{prefix}encodings", ns)
#             if encodings is None:
#                 continue

#             color = encodings.find(f"{prefix}color", ns)
#             text = encodings.find(f"{prefix}text", ns)
#             size = encodings.find(f"{prefix}size", ns)

#             if all(e is None for e in [color, text]):
#                 continue

#             x_axis = pane.get("x-axis-name")
#             y_axis = pane.get("y-axis-name")

#             bullet_charts.append({
#                 "worksheet": sheet_name,
#                 "chart_type": "Bullet Chart",
#                 "mark_type": "Bar",
#                 "X Axis": clean_column_name(x_axis) if x_axis else None,
#                 "Y Axis": clean_column_name(y_axis) if y_axis else None,
#                 "Color": clean_column_name(color.get("column")) if color is not None else None,
#                 "Text": clean_column_name(text.get("column")) if text is not None else None,
#                 "Size": clean_column_name(size.get("column")) if size is not None else None
#             })

#     return bullet_charts
import argparse

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract chart details and positions from a Tableau TWB file.")
    parser.add_argument("twb_path", help="Path to the Tableau TWB file")
    # parser.add_argument("--chart_types", nargs="+", default=["bump"], help="Chart types to extract (default: bump)")

    args = parser.parse_args()
    twb_path = args.twb_path
    # input_path = r"D:\tableau_to_powerbi_migration\tableau_to_powerbi_migration\output\input_tableau_output\Bullet Chart.twb"
    result = extract_bullet_charts_metadata(twb_path)
    result = merge_duplicate_charts(result)
    # result = deduplicate_positions(result)
    with open("output/final_json.json", "w") as f:
        json.dump(result, f, indent=2)

    print(f"✅ Extracted {len(result)} bullet charts.")

    with open("output/powerbi_chart_positions.json", "w") as f:
        pos_result = extract_bullet_charts_from_twb(twb_path)
        json.dump(pos_result, f, indent=2)
    print(f"✅ Extracted {len(pos_result)} bullet chart positions (with scaled dimensions).")