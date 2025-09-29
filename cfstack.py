"""
# cfstack.py
The following script is used to create a LightBurn file `lbrn2` for a stack of custom flyers provided in a CSV file. 
"""

import argparse
import pandas as pd
import os
import json
from pathlib import Path
from datetime import datetime
import xml.etree.ElementTree as ET
from copy import deepcopy

def parse_args():
    parser = argparse.ArgumentParser(description="LightBurn Flyer Stack Creator")
    parser.add_argument(
        "excel", type=str,
        help="Path to input Excel file cols {maxPower,QPulseWidth,speed,frequency,numPasses}."
    )
    parser.add_argument(
        "json", type=str,
        help="Path to input JSON for material properties."
    )
    return parser.parse_args()

def load_params(json_file):
    path = Path(json_file)
    if not path.exists():
        raise FileNotFoundError(f"JSON file not found: {path}")
    return json.loads(path.read_text())


args = parse_args()
params = load_params(args.json)
INPUT_XLSX = args.excel
df = pd.read_excel(INPUT_XLSX, engine='openpyxl', usecols="A:E") #First 5 columns
df.columns = ['maxPower', 'QPulseWidth', 'speed', 'frequency', 'numPasses']

TEMPLATE_FILE = params.get("template", "6x2stack.lbrn2")
template_tree = ET.parse(TEMPLATE_FILE) 
template_root = template_tree.getroot()
print("Template project loaded.")

def get_cut(root,name):
    search_norm = name.strip().lower()
    for cut in root.findall(".//CutSetting"):
            name_elem = cut.find("./name")
            if name_elem is not None:
                name_val = (name_elem.get("Value") or (name_elem.text or "")).strip()
                if name_val.lower() == search_norm:
                    return cut
    return None

def count_flyers(root,start=1):
    count = 0
    i = start

    while True:
        cut = get_cut(root, f"F{i}")
        if cut is None:
            break
        count += 1
        i += 1

    print(f"Found {count} sequential flyers starting at F{start}")
    return count

flyer_count = count_flyers(template_root)

if (flyer_count == 0):
    print("ERROR: No flyers found in template...")
else:
    print(f"Template has {flyer_count} flyers.")


# UNIT CONVERSION FOR SPEED
def xml_speed_conversion(ui_mm_per_min):
    # UI → XML: mm/min to mm/s
    try:
        return float(ui_mm_per_min) / 60.0
    except (TypeError, ValueError):
        return ui_mm_per_min


def apply_row_to_cut(root, df, row_idx, flyer_prefix="F"):
    cut_name = f"{flyer_prefix}{row_idx+1}"  # row 0 → F1, row 1 → F2, ...
    cut = get_cut(root, cut_name)
    if cut is None:
        print(f"CutSetting '{cut_name}' not found.")
        return False

    row = df.iloc[row_idx]
    for col in df.columns:
        new_value = row[col]

        # Ensure numPasses is written as an integer string
        if col == "numPasses":
            try:
                new_value = int(float(new_value))
            except Exception:
                pass  # fallback, leave as-is if not numeric
        if col == "frequency":
            try:
                new_value = int(float(new_value))
            except Exception:
                pass  # fallback, leave as-is if not numeric

        elem = cut.find(f"./{col}")
        if elem is not None:
            # Update existing
            elem.set("Value", str(new_value))
        else:
            # Create new element with Value attribute
            new_elem = ET.SubElement(cut, col)
            new_elem.set("Value", str(new_value))
            print(f"Created new attribute '{col}' in CutSetting '{cut_name}' with Value={new_value}")

    return True

for i in range(flyer_count):
    apply_row_to_cut(template_root, df, i)

def rename_text_exact(root, old_text, new_text, case_insensitive=True):
    changed = 0
    for shape in root.findall(".//Shape[@Type='Text']"):
        s = shape.get("Str", "")
        if (s.lower() == old_text.lower()) if case_insensitive else (s == old_text):
            shape.set("Str", new_text)
            changed += 1
    return changed

validation_count = rename_text_exact(template_root, params.get("template_placeholder_ID", "FLYIDK"), params.get("StackID", "001"))
print(f"Replaced {validation_count} instances of flyer ID placeholder.")

OUTPUT_FILE = params.get("output_dir", "updated.lbrn2")
template_tree.write(OUTPUT_FILE, encoding="utf-8", xml_declaration=True)

print(f"Updated project saved to {OUTPUT_FILE}.")