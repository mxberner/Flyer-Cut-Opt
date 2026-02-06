"""
# cfstack.py
The following script is used to create a LightBurn file `lbrn2` for a stack of custom flyers.  
"""

# region IMPORTS
import argparse
from email.mime import base
import logging
import pandas as pd
import os
import sys
import json
from pathlib import Path
from datetime import datetime
from copy import deepcopy
import xml.etree.ElementTree as ET
#endregion

# region SETUP
def parse_args():
    """
    Parse command-line arguments.
    """
    parser = argparse.ArgumentParser(description="LightBurn Flyer Stack Creator")
    parser.add_argument(
        "excel", type=str,
        help="Excel input file path. COLS:{maxPower,QPulseWidth,speed,frequency,numPasses}."
    )
    parser.add_argument(
        "json", type=str,
        help="Path to config JSON."
    )
    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Mute console logging."
    )
    parser.add_argument("--verbose", "-v", action="store_true")
    return parser.parse_args()

args = parse_args()

# region LOGGING
LOG_FILE = "cfstack.log"

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
logger.handlers.clear() 
logger.propagate = False

file_handler = logging.FileHandler(LOG_FILE, mode="a")
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(logging.Formatter(
    "%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
))
logger.addHandler(file_handler)

if not args.quiet:
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO if not args.verbose else logging.DEBUG)
    console_handler.setFormatter(logging.Formatter(
        "[%(levelname)s] %(message)s"
    ))
    logger.addHandler(console_handler)

logging.info("-------------------Starting cfstack script-------------------")

#endregion

def load_params(json_file):
    """
    Load parameters from JSON config file.
    """
    path = Path(json_file)
    if not path.exists():
        logging.error(f"JSON file not found: {path}")
        raise FileNotFoundError(f"JSON file not found: {path}")
    return json.loads(path.read_text())

# Load input parameterse
params = load_params(args.json)
igsn_params = load_params(params.get("igsn_config", "igsn-config/default.json"))
logging.debug("Loaded config JSON and IGSN config.")

# Load Excel data
INPUT_XLSX = args.excel
df = pd.read_excel(INPUT_XLSX, engine='openpyxl', usecols="A:E") #First 5 columns
df.columns = ['maxPower', 'QPulseWidth', 'speed', 'frequency', 'numPasses']


# Load LightBurn Template
TEMPLATE_FILE = params.get("tmp_file", None)
if TEMPLATE_FILE is None:
    logging.error("No template specified in config (missing 'tmp_file' field).")
    raise ValueError("No template specified in config (missing 'tmp_file' field).")
if not os.path.exists(TEMPLATE_FILE):
    logging.error(f"Template file not found: {TEMPLATE_FILE}")
    raise FileNotFoundError(f"Template file '{TEMPLATE_FILE}' was not found.")


template_tree = ET.parse(TEMPLATE_FILE) 
template_root = template_tree.getroot()
logging.info(f"Loaded config '{args.json}' and excel data '{INPUT_XLSX}'.")
logging.info(f"Loaded template '{TEMPLATE_FILE}'.")
logging.info(f"Loaded IGSN config '{params.get('igsn_config', '')}' with IGSN '{igsn_params.get('material', {}).get('igsn', 'N/A')}'.")

def _safe_token(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    return "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in s)

def resolve_output_path(params, igsn_params, template_file: str) -> Path:
    logging.debug(igsn_params.get("material", {}).get("igsn", ""))
    out_dir = Path(params.get("output_dir", ".")).expanduser().resolve()
    if params.get("output_dir_append_IGSN", False):
        igsn = _safe_token(igsn_params.get("material", {}).get("igsn", ""))
        if igsn:
            out_dir = out_dir / igsn
    out_dir.mkdir(parents=True, exist_ok=True)

    overwrite = bool(params.get("output_overwrite", False))

    # Base filename
    base_file = params.get("output_base", "STACK")
    if base_file == "":
        base_file = "STACK"
    base_path = Path(base_file)

    stem = _safe_token(base_path.stem)
    suffix = base_path.suffix or (Path(template_file).suffix or ".lbrn2")

    parts = [stem]

    if params.get("output_append_ID", False):
        parts.append(_safe_token(str(params.get("ID", "0000"))))

    if params.get("output_append_template", False):
        parts.append(_safe_token(Path(template_file).stem))

    if params.get("output_append_timestamp", False):
        parts.append(datetime.now().strftime("%Y%m%d-%H%M%S"))

    filename = "-".join(p for p in parts if p) + suffix
    out_path = out_dir / filename

    if overwrite or not out_path.exists():
        return out_path

    i = 1
    while True:
        candidate = out_dir / f"{out_path.stem}-{i}{out_path.suffix}"
        if not candidate.exists():
            return candidate
        i += 1

#endregion

# region MAIN
"""
Return CutSetting by name.
"""
def get_cut(root,name):
    search_norm = name.strip().lower()
    for cut in root.findall(".//CutSetting"):
            name_elem = cut.find("./name")
            if name_elem is not None:
                name_val = (name_elem.get("Value") or (name_elem.text or "")).strip()
                if name_val.lower() == search_norm:
                    return cut
    logging.debug(f"CutSetting '{name}' not found in template.")
    return None


"""
Count sequential flyers in template starting at F{start}.
"""
"""
List flyers found in template within an inclusive numeric range.

Selection format (from config):
  "flyer_selection": {"start": 1, "end": 25, "step": 1}

Returns a list of flyer numbers that actually exist in the template, e.g. [1,2,3,...].
"""
def list_flyers_in_range(root, start: int = 1, end: int | None = None, step: int = 1, flyer_prefix: str = "F"):
    if end is None:
        end = start
    try:
        start_i = int(start)
        end_i = int(end)
        step_i = int(step) if int(step) != 0 else 1
    except Exception:
        raise ValueError(f"Invalid flyer_selection: start={start}, end={end}, step={step}")

    if step_i < 0:
        step_i = abs(step_i)

    lo, hi = (start_i, end_i) if start_i <= end_i else (end_i, start_i)

    found: list[int] = []
    missing: list[int] = []

    for n in range(lo, hi + 1, step_i):
        cut = get_cut(root, f"{flyer_prefix}{n}")
        if cut is None:
            missing.append(n)
        else:
            found.append(n)

    logging.info(f"Flyer selection requested: {lo}..{hi} step {step_i}. Found {len(found)} flyer(s).")
    if missing:
        logging.debug(f"Missing flyers in template (within selection): {missing}")
    logging.debug(f"Flyers found (within selection): {found}")
    return found


# Resolve flyer selection from config (defaults to a single flyer F1 if not provided)
flyer_sel = params.get("flyer_selection", {}) or {}
sel_start = flyer_sel.get("start", 1)
sel_end = flyer_sel.get("end", sel_start)
sel_step = flyer_sel.get("step", 1)

selected_flyers = list_flyers_in_range(template_root, sel_start, sel_end, sel_step, flyer_prefix="F")

if len(selected_flyers) == 0:
    logging.error(
        f"No flyers found in template '{TEMPLATE_FILE}' within selection "
        f"(start={sel_start}, end={sel_end}, step={sel_step})."
    )
    sys.exit(1)

# UNIT CONVERSION FOR SPEED
def xml_speed_conversion(ui_mm_per_min):
    # UI → XML: mm/min to mm/s
    try:
        return float(ui_mm_per_min) / 60.0
    except (TypeError, ValueError):
        return ui_mm_per_min


"""
Apply DataFrame row to CutSetting.
"""
"""
Apply a given DataFrame row to a specific flyer CutSetting (e.g., F17).
"""
def apply_row_to_cut(root, df, flyer_number: int, excel_row_idx: int, flyer_prefix: str = "F") -> bool:
    cut_name = f"{flyer_prefix}{int(flyer_number)}"
    cut = get_cut(root, cut_name)
    if cut is None:
        logging.debug(f"CutSetting '{cut_name}' not found.")
        return False

    if excel_row_idx < 0 or excel_row_idx >= len(df):
        logging.warning(
            f"Excel row index {excel_row_idx} is out of bounds for dataframe with {len(df)} row(s). "
            f"Skipping flyer '{cut_name}'."
        )
        return False

    row = df.iloc[excel_row_idx]
    for col in df.columns:
        new_value = row[col]

        # Ensure some fields are written as integer strings
        if col in ("numPasses", "frequency"):
            try:
                new_value = int(float(new_value))
            except Exception:
                pass  # leave as-is if not numeric

        elem = cut.find(f"./{col}")
        if elem is not None:
            elem.set("Value", str(new_value))
        else:
            new_elem = ET.SubElement(cut, col)
            new_elem.set("Value", str(new_value))
            logging.debug(f"Created new attribute '{col}' in CutSetting '{cut_name}' with Value={new_value}")

    row_str = ", ".join(f"{k[:6]}={row[k]}" for k in df.columns)
    logging.debug(f"Applied excel row [{excel_row_idx}] to CutSetting '{cut_name}': {row_str}")
    return True


def _excel_row_for_flyer_position(pos: int, style: str, x: int) -> int:
    """
    Map the position (0-based) of a flyer within the *selected_flyers* list to an excel row index.

    Styles (from config):
      - exact:     excel_row = pos
      - repeat_x:  excel_row = pos // x      (repeat each excel row x times)
      - modulus_x: excel_row = pos % x       (cycle through first x excel rows)
    """
    style_norm = (style or "exact").strip().lower()
    if style_norm == "exact":
        return pos

    # sanitize x
    try:
        xi = int(x)
    except Exception:
        xi = 1
    if xi <= 0:
        xi = 1

    if style_norm == "repeat_x":
        return pos // xi
    if style_norm == "modulus_x":
        return pos % xi

    logging.warning(f"Unknown flyer_assignment.style '{style}'. Falling back to 'exact'.")
    return pos


# Resolve flyer assignment mode from config
assign_cfg = params.get("flyer_assignment", {}) or {}
assign_style = assign_cfg.get("style", "exact")
assign_x = assign_cfg.get("x", 1)

logging.info(f"Applying excel rows to {len(selected_flyers)} selected flyer(s) using style='{assign_style}', x={assign_x}")

applied = 0
for pos, flyer_num in enumerate(selected_flyers):
    excel_row_idx = _excel_row_for_flyer_position(pos, assign_style, assign_x)
    if apply_row_to_cut(template_root, df, flyer_num, excel_row_idx, flyer_prefix="F"):
        applied += 1

logging.info(f"Applied settings to {applied}/{len(selected_flyers)} selected flyer(s).")


"""
Rename exact text matches in Shape elements.
"""
def rename_text_exact(root, old_text, new_text, case_insensitive=True):
    changed = 0
    for shape in root.findall(".//Shape[@Type='Text']"):
        s = shape.get("Str", "")
        if (s.lower() == old_text.lower()) if case_insensitive else (s == old_text):
            shape.set("Str", new_text)
            changed += 1
    return changed

validation_count = rename_text_exact(template_root, params.get("tmp_ID_placeholder", "FLYID"), params.get("ID", "0000"))
if validation_count == 0:
    logging.warning(f"No instances of tmp_ID_placeholder found in template to replace with {params.get('ID', '0000')}.")
elif params.get("ID") is None:
    logging.warning(f"Stack ID parameter not provided; replaced placeholder with '0000' by default.")
elif validation_count > 1:
    logging.warning(f"Replaced {validation_count} instances of stack ID placeholder with '{params.get('ID', '0000')}'.")
else:
    logging.info(f"Set stack ID to '{params.get('ID', '0000')}'.")

OPERATOR = params.get("operator", "Unknown operator")
if (OPERATOR is None) or (OPERATOR.strip() == ""):
    logging.warning("Operator name not provided in config.")
    OPERATOR = "Unknown operator"

OUTPUT_PATH = resolve_output_path(params, igsn_params, TEMPLATE_FILE)
action = "overwrote" if OUTPUT_PATH.exists() else "generated"
template_tree.write(str(OUTPUT_PATH), encoding="utf-8", xml_declaration=True)

logging.info(
    f"{OPERATOR} "
    f"{action} LightBurn file '{OUTPUT_PATH}' with configs '{args.json}' and '{params.get('igsn_config', '')}' "
    f"from LB-template '{TEMPLATE_FILE}'."
)
logging.info("-------------------Completed cfstack script-------------------")
#endregion