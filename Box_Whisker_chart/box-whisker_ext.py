import xml.etree.ElementTree as ET
import json
import re
import argparse
import os
# Constants for scaling (adjust as needed)
POWER_BI_WIDTH = 1280
POWER_BI_HEIGHT = 720
TABLEAU_MAX_WIDTH = 100000
TABLEAU_MAX_HEIGHT = 100000

def scale_to_powerbi(x, y, width, height):
    new_x = (x / TABLEAU_MAX_WIDTH) * POWER_BI_WIDTH
    new_y = (y / TABLEAU_MAX_HEIGHT) * POWER_BI_HEIGHT
    new_w = (width / TABLEAU_MAX_WIDTH) * POWER_BI_WIDTH
    new_h = (height / TABLEAU_MAX_HEIGHT) * POWER_BI_HEIGHT
    return new_x, new_y, new_w, new_h

def normalize_column(col):
    if not col:
        return None
    match = re.search(r':([^:\]]+):[^:\]]*\]?', col)
    return match.group(1) if match else col

def parse_column(raw):
    if not raw:
        return None
    if '].[' in raw:
        return raw.split('].[')[-1].replace(']', '')
    return raw.replace('[', '').replace(']', '')

def extract_box_whisker_charts(twb_file):
    tree = ET.parse(twb_file)
    root = tree.getroot()
    chart_metadata = []
    chart_positions = []

    # Determine XML namespace
    if "}" in root.tag:
        ns_uri = root.tag.split("}")[0].strip("{")
        ns = {"t": ns_uri}
        prefix = "t:"
    else:
        ns = {}
        prefix = ""

    worksheet_info = {}
    for worksheet in root.findall(f".//{prefix}worksheet", ns):
        sheet_name = worksheet.get('name', 'Unknown')
        panes = worksheet.findall(f".//{prefix}pane", ns)

        for pane in panes:
            ref_line = pane.find(f"{prefix}reference-line", ns)
            if ref_line is not None and ref_line.get('boxplot-whisker-type') is not None:
                mark = pane.find(f"{prefix}mark", ns)
                mark_type = mark.get('class') if mark is not None else None

                encodings = pane.find(f"{prefix}encodings", ns)
                color = text = size = detail = None

                if encodings is not None:
                    color_tag = encodings.find(f"{prefix}color", ns)
                    text_tag = encodings.find(f"{prefix}text", ns)
                    size_tag = encodings.find(f"{prefix}size", ns)
                    detail_tag = encodings.find(f"{prefix}lod", ns)

                    color = parse_column(color_tag.get('column')) if color_tag is not None else None
                    text = parse_column(text_tag.get('column')) if text_tag is not None else None
                    size = parse_column(size_tag.get('column')) if size_tag is not None else None
                    detail = parse_column(detail_tag.get('column')) if detail_tag is not None else None

                table = worksheet.find(f".//{prefix}table", ns)
                rows = table.findtext(f"{prefix}rows", default="", namespaces=ns)
                cols = table.findtext(f"{prefix}cols", default="", namespaces=ns)

                worksheet_info[sheet_name] = {
                    "worksheet": sheet_name,
                    "chart_type": "Box-and-Whisker Chart",
                    "mark_type": mark_type,
                    "X Axis": normalize_column(parse_column(cols)),
                    "Y Axis": normalize_column(parse_column(rows)),
                    "Color": normalize_column(color),
                    "Text": normalize_column(text),
                    "Size": normalize_column(size),
                    "Detail": normalize_column(detail)
                }

    for dashboard in root.findall(f".//{prefix}dashboard", ns):
        dashboard_name = dashboard.get("name")
        for zone in dashboard.findall(f".//{prefix}zone", ns):
            sheet_name = zone.get("name")
            if sheet_name in worksheet_info:
                x = int(zone.get("x", 0))
                y = int(zone.get("y", 0))
                w = int(zone.get("w", 0))
                h = int(zone.get("h", 0))
                scaled_x, scaled_y, scaled_w, scaled_h = scale_to_powerbi(x, y, w, h)

                chart_positions.append({
                    "dashboard": dashboard_name,
                    "worksheet": sheet_name,
                    "x": scaled_x,
                    "y": scaled_y,
                    "w": scaled_w,
                    "h": scaled_h
                })

                chart_metadata.append(worksheet_info[sheet_name])

    return chart_metadata, chart_positions

def merge_duplicate_charts(charts):
    merged = {}
    for chart in charts:
        key = (chart["worksheet"], chart["chart_type"])
        if key not in merged:
            merged[key] = chart
        else:
            if not merged[key]["mark_type"] and chart["mark_type"]:
                merged[key]["mark_type"] = chart["mark_type"]
    return list(merged.values())

def deduplicate_positions(positions):
    seen = set()
    deduped = []
    for pos in positions:
        key = (pos["worksheet"], pos["dashboard"], pos["x"], pos["y"], pos["w"], pos["h"])
        if key not in seen:
            seen.add(key)
            deduped.append(pos)
    return deduped



def main():
    parser = argparse.ArgumentParser(description="Extract box-and-whisker chart details and positions from a TWB file.")
    parser.add_argument("twb_file", help="Path to the Tableau TWB file")
    parser.add_argument("--metadata_output", default="final_json.json", help="Path to save chart metadata")
    parser.add_argument("--positions_output", default="powerbi_chart_positions.json", help="Path to save chart positions")

    args = parser.parse_args()

    metadata, positions = extract_box_whisker_charts(args.twb_file)
    metadata = merge_duplicate_charts(metadata)
    positions = deduplicate_positions(positions)

    os.makedirs("output", exist_ok=True)

    with open(f"output/{args.metadata_output}", "w") as f:
        json.dump(metadata, f, indent=2)
    print(f"✅ Metadata saved: {args.metadata_output} ({len(metadata)} items)")

    with open(f"output/{args.positions_output}", "w") as f:
        json.dump(positions, f, indent=2)
    print(f"✅ Positions saved: {args.positions_output} ({len(positions)} items)")



if __name__ == "__main__":
    main()
