"""
cfstack_new.py
Create one or more LightBurn .lbrn2 outputs for a stack of custom flyers.

New config shape only. No backward compatibility.
"""

from __future__ import annotations

import argparse
import json
import logging
import math
import sys
from datetime import datetime
from pathlib import Path
import xml.etree.ElementTree as ET

import pandas as pd


# ----------------------------
# COMMAND LINE ARGS
# ----------------------------
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="LightBurn Flyer Stack Creator")
    parser.add_argument("json", type=str, help="Path to config JSON.")
    parser.add_argument("--quiet", "-q", action="store_true", help="Mute console logging.")
    parser.add_argument("--verbose", "-v", action="store_true", help="Enable verbose logging.")
    return parser.parse_args()


args = parse_args()


# ----------------------------
# LOGGING
# ----------------------------
LOG_FILE = "cfstack.log"

logger = logging.getLogger()
logger.setLevel(logging.DEBUG)
logger.handlers.clear()
logger.propagate = False

file_handler = logging.FileHandler(LOG_FILE, mode="a")
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(
    logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
)
logger.addHandler(file_handler)

if not args.quiet:
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO if not args.verbose else logging.DEBUG)
    console_handler.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
    logger.addHandler(console_handler)

logging.info("-------------------Starting cfstack script-------------------")


# ----------------------------
# HELPERS
# ----------------------------
def load_json(path_str: str) -> dict:
    path = Path(path_str)
    if not path.exists():
        raise FileNotFoundError(f"JSON file not found: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def require(cond: bool, msg: str) -> None:
    if not cond:
        raise ValueError(msg)


def _safe_token(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    return "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in s)


def _int(value, default: int) -> int:
    try:
        return int(value)
    except Exception:
        return default


def _float(value, default: float) -> float:
    try:
        return float(value)
    except Exception:
        return default


def _is_n_token(value) -> bool:
    return str(value).strip().lower() == "n"


def validate_top_level_config(cfg: dict) -> None:
    require(isinstance(cfg, dict), "Config root must be a JSON object.")

    for key in ("ID", "operator", "igsn_config", "template", "output", "flyer", "laser_params", "thickness"):
        require(key in cfg, f"Missing required config field: {key}")

    require(str(cfg["ID"]).strip(), "ID is required.")
    require(str(cfg["operator"]).strip(), "operator is required.")

    template = cfg["template"]
    output = cfg["output"]
    flyer = cfg["flyer"]
    laser_params = cfg["laser_params"]
    thickness = cfg["thickness"]

    require(isinstance(template, dict), "template must be an object.")
    require(isinstance(output, dict), "output must be an object.")
    require(isinstance(flyer, dict), "flyer must be an object.")
    require(isinstance(laser_params, dict), "laser_params must be an object.")
    require(isinstance(thickness, dict), "thickness must be an object.")

    require(str(template.get("file", "")).strip(), "template.file is required.")
    require(str(template.get("id_placeholder", "")).strip(), "template.id_placeholder is required.")
    require(str(output.get("dir", "")).strip(), "output.dir is required.")
    require(str(output.get("base", "")).strip(), "output.base is required.")

    flyer_selection = flyer.get("selection", {})
    flyer_assignment = flyer.get("assignment", {})
    require(isinstance(flyer_selection, dict), "flyer.selection must be an object.")
    require(isinstance(flyer_assignment, dict), "flyer.assignment must be an object.")

    start = _int(flyer_selection.get("start", 1), 1)
    end_raw = flyer_selection.get("end", start)
    step = _int(flyer_selection.get("step", 1), 1)
    require(step != 0, "flyer.selection.step must not be 0.")
    require(start >= 0, "flyer.selection.start must be >= 0.")
    if not _is_n_token(end_raw):
        end = _int(end_raw, start)
        require(end >= 0, "flyer.selection.end must be >= 0 or 'n'.")

    style = str(flyer_assignment.get("style", "exact")).strip().lower()
    require(style in {"exact", "repeat_x", "modulus_x"}, "flyer.assignment.style must be exact, repeat_x, or modulus_x.")
    require(_int(flyer_assignment.get("x", 1), 1) > 0, "flyer.assignment.x must be > 0.")

    require(_int(laser_params.get("excel_row_start", 1), 1) >= 1, "laser_params.excel_row_start must be >= 1.")

    glass = thickness.get("glass", {})
    foil = thickness.get("foil", {})
    require(isinstance(glass, dict), "thickness.glass must be an object.")
    require(isinstance(foil, dict), "thickness.foil must be an object.")

    for k in ("tl_mm", "tr_mm", "bl_mm", "br_mm"):
        require(_float(glass.get(k, 6.25), 6.25) > 0, f"thickness.glass.{k} must be > 0.")
    for k in ("tl_um", "tr_um", "bl_um", "br_um"):
        v = foil.get(k, None)
        if v is not None:
            require(_float(v, 0.0) > 0, f"thickness.foil.{k} must be > 0 when provided.")


def normalize_thickness(cfg: dict, igsn_cfg: dict) -> dict:
    foil_default = igsn_cfg.get("material", {}).get("thickness_um", None)
    foil_default = _float(foil_default, 0.0) if foil_default is not None else None

    glass_cfg = (cfg.get("thickness", {}) or {}).get("glass", {}) or {}
    foil_cfg = (cfg.get("thickness", {}) or {}).get("foil", {}) or {}

    resolved = {
        "glass": {
            "tl_mm": _float(glass_cfg.get("tl_mm", 6.25), 6.25),
            "tr_mm": _float(glass_cfg.get("tr_mm", 6.25), 6.25),
            "bl_mm": _float(glass_cfg.get("bl_mm", 6.25), 6.25),
            "br_mm": _float(glass_cfg.get("br_mm", 6.25), 6.25),
        },
        "foil": {
            "tl_um": _float(foil_cfg.get("tl_um", foil_default), foil_default or 0.0),
            "tr_um": _float(foil_cfg.get("tr_um", foil_default), foil_default or 0.0),
            "bl_um": _float(foil_cfg.get("bl_um", foil_default), foil_default or 0.0),
            "br_um": _float(foil_cfg.get("br_um", foil_default), foil_default or 0.0),
        },
    }

    for bucket, suffix in ((resolved["glass"], "mm"), (resolved["foil"], "um")):
        for k, v in bucket.items():
            require(v > 0, f"Resolved thickness value must be > 0: {bucket}.{k}")

    return resolved


# ----------------------------
# OUTPUT PATHS
# ----------------------------
def resolve_output_path(cfg: dict, igsn_cfg: dict, template_file: str, extra_name_parts: list[str] | None = None) -> Path:
    output_cfg = cfg["output"]

    out_dir = Path(output_cfg["dir"]).expanduser().resolve()
    if bool(output_cfg.get("dir_append_igsn", False)):
        igsn = _safe_token(igsn_cfg.get("material", {}).get("igsn", ""))
        if igsn:
            out_dir = out_dir / igsn
    out_dir.mkdir(parents=True, exist_ok=True)

    overwrite = bool(output_cfg.get("overwrite", False))
    base_file = output_cfg.get("base", "STACK") or "STACK"
    base_path = Path(base_file)

    stem = _safe_token(base_path.stem)
    suffix = base_path.suffix or (Path(template_file).suffix or ".lbrn2")

    parts: list[str] = [stem]

    if bool(output_cfg.get("append_id", False)):
        parts.append(_safe_token(str(cfg.get("ID", "0000"))))
    if bool(output_cfg.get("append_template", False)):
        parts.append(_safe_token(Path(template_file).stem))
    if extra_name_parts:
        parts.extend(_safe_token(p) for p in extra_name_parts if p)
    if bool(output_cfg.get("append_timestamp", False)):
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


# ----------------------------
# XML HELPERS
# ----------------------------
def get_cut(root: ET.Element, name: str):
    search_norm = name.strip().lower()
    for cut in root.findall(".//CutSetting"):
        name_elem = cut.find("./name")
        if name_elem is None:
            continue
        name_val = (name_elem.get("Value") or (name_elem.text or "")).strip()
        if name_val.lower() == search_norm:
            return cut
    return None


def list_flyers_in_range(root: ET.Element, start: int = 1, end: int | str | None = None, step: int = 1, flyer_prefix: str = "F") -> list[int]:
    if end is None:
        end = start

    start_i = int(start)
    step_i = int(step)
    if step_i == 0:
        raise ValueError("flyer.selection.step must not be 0.")
    if step_i < 0:
        step_i = abs(step_i)

    if _is_n_token(end):
        found: list[int] = []
        n = start_i
        while True:
            cut = get_cut(root, f"{flyer_prefix}{n}")
            if cut is None:
                break
            found.append(n)
            n += step_i

        if found:
            logging.info(
                "Flyer selection requested: %s..n step %s. Resolved through last sequential flyer %s. Found %s flyer(s).",
                start_i,
                step_i,
                found[-1],
                len(found),
            )
        else:
            logging.info(
                "Flyer selection requested: %s..n step %s. No sequential flyers found starting at %s.",
                start_i,
                step_i,
                start_i,
            )
        return found

    end_i = int(end)
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
        logging.debug(f"Missing flyers in template within selection: {missing}")
    return found


def rename_text_exact(root: ET.Element, old_text: str, new_text: str, case_insensitive: bool = True) -> int:
    changed = 0
    for shape in root.findall(".//Shape[@Type='Text']"):
        s = shape.get("Str", "")
        if (s.lower() == old_text.lower()) if case_insensitive else (s == old_text):
            shape.set("Str", new_text)
            changed += 1
    return changed


# ----------------------------
# LASER PARAMS APPLICATION
# ----------------------------
def _excel_row_for_flyer_position(pos: int, style: str, x: int) -> int:
    style_norm = (style or "exact").strip().lower()
    if style_norm == "exact":
        return pos

    xi = _int(x, 1)
    if xi <= 0:
        xi = 1

    if style_norm == "repeat_x":
        return pos // xi
    if style_norm == "modulus_x":
        return pos % xi

    logging.warning(f"Unknown flyer.assignment.style '{style}'. Falling back to 'exact'.")
    return pos


def apply_row_to_cut(root: ET.Element, df: pd.DataFrame, flyer_number: int, excel_row_idx: int, flyer_prefix: str = "F") -> bool:
    cut_name = f"{flyer_prefix}{int(flyer_number)}"
    cut = get_cut(root, cut_name)
    if cut is None:
        logging.debug(f"CutSetting '{cut_name}' not found.")
        return False

    if excel_row_idx < 0 or excel_row_idx >= len(df):
        logging.warning(f"Excel row index {excel_row_idx} is out of bounds for dataframe with {len(df)} rows.")
        return False

    row = df.iloc[excel_row_idx]
    for col in df.columns:
        new_value = row[col]
        if col in ("numPasses", "frequency"):
            try:
                new_value = int(float(new_value))
            except Exception:
                pass

        elem = cut.find(f"./{col}")
        if elem is None:
            elem = ET.SubElement(cut, col)
        elem.set("Value", str(new_value))

    logging.debug(
        "Applied excel row [%s] to CutSetting '%s': %s",
        excel_row_idx,
        cut_name,
        ", ".join(f"{k}={row[k]}" for k in df.columns),
    )
    return True


def apply_cut_defaults_to_flyer(root: ET.Element, flyer_number: int, cut_defaults: dict, flyer_prefix: str = "F") -> bool:
    cut_name = f"{flyer_prefix}{int(flyer_number)}"
    cut = get_cut(root, cut_name)
    if cut is None:
        return False

    applied_any = False
    for key, val in (cut_defaults or {}).items():
        if val is None:
            continue
        elem = cut.find(f"./{key}")
        if elem is None:
            elem = ET.SubElement(cut, key)
        elem.set("Value", str(val))
        applied_any = True
    return applied_any


def igsn_cut_to_lightburn_fields(igsn_cfg: dict) -> dict:
    cut = igsn_cfg.get("cut", {}) or {}
    out: dict = {}

    if "max_power" in cut:
        out["maxPower"] = cut.get("max_power")
    if "min_power" in cut:
        out["minPower"] = cut.get("min_power")
    if "passes" in cut:
        out["numPasses"] = cut.get("passes")
    if "speed_mm_s" in cut:
        out["speed"] = cut.get("speed_mm_s")

    for k in ("frequency", "QPulseWidth", "q_pulse_width", "qPulseWidth"):
        if k in cut:
            out["frequency" if k == "frequency" else "QPulseWidth"] = cut.get(k)

    return out


# ----------------------------
# EXCEL HELPERS
# ----------------------------
def load_excel_df(excel_path: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path, engine="openpyxl", usecols="A:E")
    df.columns = ["maxPower", "QPulseWidth", "speed", "frequency", "numPasses"]
    return df


def compute_rows_per_template(num_flyers: int, assign_style: str, assign_x: int) -> int:
    style = (assign_style or "exact").strip().lower()
    x = _int(assign_x, 1)
    if x <= 0:
        x = 1

    if style == "exact":
        return num_flyers
    if style == "repeat_x":
        return math.ceil(num_flyers / x)
    if style == "modulus_x":
        return x
    return num_flyers


def build_batch_row_indices(df_len: int, start_idx: int, chunk: int, exhaust: bool) -> list[list[int]]:
    if start_idx >= df_len:
        return []

    if chunk <= 0:
        raise ValueError("Batch chunk size must be > 0.")

    if not exhaust:
        return [list(range(start_idx, min(start_idx + chunk, df_len)))]

    batches: list[list[int]] = []
    row_cursor = start_idx

    while row_cursor < df_len:
        rows = list(range(row_cursor, min(row_cursor + chunk, df_len)))

        if len(rows) < chunk:
            wrap_cursor = start_idx
            while len(rows) < chunk and wrap_cursor < df_len:
                rows.append(wrap_cursor)
                wrap_cursor += 1

        batches.append(rows)
        row_cursor += chunk

    return batches


def describe_batch_rows(rows: list[int], start_idx: int) -> str:
    if not rows:
        return "rows-none"

    chunks: list[tuple[int, int]] = []
    run_start = rows[0]
    prev = rows[0]

    for idx in rows[1:]:
        if idx == prev + 1:
            prev = idx
            continue
        chunks.append((run_start, prev))
        run_start = idx
        prev = idx
    chunks.append((run_start, prev))

    rendered = []
    for a, b in chunks:
        if a == b:
            rendered.append(str(a + 1))
        else:
            rendered.append(f"{a + 1}-{b + 1}")

    return "rows" + "_wrap_".join(rendered)


# ----------------------------
# MAIN
# ----------------------------
cfg = load_json(args.json)
validate_top_level_config(cfg)
igsn_cfg = load_json(cfg["igsn_config"])
resolved_thickness = normalize_thickness(cfg, igsn_cfg)

logging.info(f"Loaded config '{args.json}'.")
logging.info(f"Loaded IGSN config '{cfg['igsn_config']}'.")
logging.info(f"Resolved thickness: {json.dumps(resolved_thickness, sort_keys=True)}")

template_cfg = cfg["template"]
flyer_cfg = cfg["flyer"]
laser_cfg = cfg["laser_params"]

TEMPLATE_FILE = template_cfg["file"]
TMP_ID_PLACEHOLDER = template_cfg["id_placeholder"]

if not Path(TEMPLATE_FILE).exists():
    raise FileNotFoundError(f"Template file not found: {TEMPLATE_FILE}")

excel_as_input = bool(laser_cfg.get("excel_as_input", True))
excel_path = str(laser_cfg.get("excel_path", "") or "")
excel_row_start_1based = _int(laser_cfg.get("excel_row_start", 1), 1)
excel_start_idx = max(0, excel_row_start_1based - 1)
excel_exhaust = bool(laser_cfg.get("excel_exhaust", False))

template_tree_base = ET.parse(TEMPLATE_FILE)
template_root_base = template_tree_base.getroot()

selection_cfg = flyer_cfg["selection"]
assignment_cfg = flyer_cfg["assignment"]

selected_flyers = list_flyers_in_range(
    template_root_base,
    selection_cfg.get("start", 1),
    selection_cfg.get("end", 1),
    selection_cfg.get("step", 1),
    flyer_prefix="F",
)
if not selected_flyers:
    raise ValueError("No flyers found in template for the requested flyer.selection range.")

assign_style = assignment_cfg.get("style", "exact")
assign_x = _int(assignment_cfg.get("x", 1), 1)

outputs_written: list[Path] = []

if excel_as_input:
    require(bool(excel_path.strip()), "laser_params.excel_path is required when excel_as_input is true.")
    require(Path(excel_path).exists(), f"Excel file not found: {excel_path}")

    df = load_excel_df(excel_path)
    require(excel_start_idx < len(df), f"laser_params.excel_row_start={excel_row_start_1based} starts beyond available excel rows ({len(df)}).")

    rows_per_template = compute_rows_per_template(len(selected_flyers), assign_style, assign_x)
    batch_row_sets = build_batch_row_indices(
        df_len=len(df),
        start_idx=excel_start_idx,
        chunk=rows_per_template,
        exhaust=excel_exhaust,
    )

    logging.info(
        "Excel-as-input enabled. excel_path='%s', excel_row_start=%s, excel_exhaust=%s, rows_per_template=%s, batches=%s.",
        excel_path,
        excel_row_start_1based,
        excel_exhaust,
        rows_per_template,
        len(batch_row_sets),
    )

    for batch_num, batch_rows in enumerate(batch_row_sets, start=1):
        template_tree = ET.parse(TEMPLATE_FILE)
        template_root = template_tree.getroot()

        applied = 0
        for pos, flyer_num in enumerate(selected_flyers):
            base_row_idx = _excel_row_for_flyer_position(pos, assign_style, assign_x)
            excel_row_idx = batch_rows[base_row_idx]
            if apply_row_to_cut(template_root, df, flyer_num, excel_row_idx, flyer_prefix="F"):
                applied += 1

        validation_count = rename_text_exact(
            template_root,
            TMP_ID_PLACEHOLDER,
            str(cfg.get("ID", "0000")),
        )
        if validation_count == 0:
            logging.warning("No instances of template.id_placeholder were found in template.")

        extra_parts = None
        if excel_exhaust:
            extra_parts = [describe_batch_rows(batch_rows, excel_start_idx), f"batch{batch_num}"]

        out_path = resolve_output_path(cfg, igsn_cfg, TEMPLATE_FILE, extra_name_parts=extra_parts)
        action = "overwrote" if out_path.exists() else "generated"
        template_tree.write(str(out_path), encoding="utf-8", xml_declaration=True)

        logging.info(
            "%s %s LightBurn file '%s' with config '%s' and excel '%s' (batch %s, %s, applied=%s/%s).",
            (cfg.get("operator", "") or "Unknown operator").strip() or "Unknown operator",
            action,
            out_path,
            args.json,
            excel_path,
            batch_num,
            describe_batch_rows(batch_rows, excel_start_idx),
            applied,
            len(selected_flyers),
        )
        outputs_written.append(out_path)
else:
    cut_defaults = igsn_cut_to_lightburn_fields(igsn_cfg)
    if not cut_defaults:
        logging.warning("excel_as_input is false, but no usable defaults were found in the IGSN config.")

    template_tree = ET.parse(TEMPLATE_FILE)
    template_root = template_tree.getroot()

    applied = 0
    for flyer_num in selected_flyers:
        if apply_cut_defaults_to_flyer(template_root, flyer_num, cut_defaults, flyer_prefix="F"):
            applied += 1

    validation_count = rename_text_exact(
        template_root,
        TMP_ID_PLACEHOLDER,
        str(cfg.get("ID", "0000")),
    )
    if validation_count == 0:
        logging.warning("No instances of template.id_placeholder were found in template.")

    out_path = resolve_output_path(cfg, igsn_cfg, TEMPLATE_FILE, extra_name_parts=["defaults"])
    action = "overwrote" if out_path.exists() else "generated"
    template_tree.write(str(out_path), encoding="utf-8", xml_declaration=True)

    logging.info(
        "%s %s LightBurn file '%s' with config '%s' (excel_as_input=false; applied defaults to %s/%s selected flyers).",
        (cfg.get("operator", "") or "Unknown operator").strip() or "Unknown operator",
        action,
        out_path,
        args.json,
        applied,
        len(selected_flyers),
    )
    outputs_written.append(out_path)

logging.info(f"Total LightBurn outputs written: {len(outputs_written)}")
for p in outputs_written:
    logging.info(f"  - {p}")
logging.info("-------------------Completed cfstack script-------------------")
