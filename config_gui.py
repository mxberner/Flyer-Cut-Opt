#!/usr/bin/env python3
from __future__ import annotations

import json
from pathlib import Path
from gooey import Gooey, GooeyParser

ASSIGNMENT_STYLES = ["exact", "repeat_x", "modulus_x"]

def _relpath_str(p: Path, root: Path) -> str:
    try:
        return str(p.relative_to(root))
    except Exception:
        return str(p)

def find_files(root: Path, subdir: str, exts: tuple[str, ...]) -> list[str]:
    base = root / subdir
    if not base.exists():
        return []
    files: list[Path] = []
    for ext in exts:
        files.extend(sorted(base.glob(f"*{ext}")))
    return [_relpath_str(p, root) for p in files]

def require(cond: bool, msg: str) -> None:
    if not cond:
        raise ValueError(msg)

def require_int_range(name: str, value: int, lo: int, hi: int) -> None:
    require(isinstance(value, int), f"{name} must be an integer.")
    require(lo <= value <= hi, f"{name} must be between {lo} and {hi} (got {value}).")

def validate(args, igsn_choices: list[str], tmp_choices: list[str]) -> None:
    require(bool(args.ID.strip()), "ID is required.")
    require(len(args.ID.strip()) == 4, "ID must be exactly 4 digits (e.g., 0001).")
    require(args.ID.strip().isdigit(), "ID must be 4 digits (e.g., 0001).")

    require(bool(args.operator.strip()), "operator is required.")

    require(args.igsn_config in igsn_choices and args.igsn_config != "<none found>",
            "No IGSN config files found (root/igsn-config/*.json) or invalid selection.")
    require(args.tmp_file in tmp_choices and args.tmp_file != "<none found>",
            "No template files found (root/LB-templates/*.lbrn2) or invalid selection.")

    require_int_range("flyer_selection.start", args.start, 0, 30)
    require_int_range("flyer_selection.end", args.end, 0, 30)
    require_int_range("flyer_selection.step", args.step, 0, 30)

    require(args.style in ASSIGNMENT_STYLES, f"flyer_assignment.style must be one of {ASSIGNMENT_STYLES}.")
    require_int_range("flyer_assignment.x", args.x, 0, 30)

def safe_write_no_overwrite(path: Path, text: str) -> None:
    """
    Write a file WITHOUT overwriting if it already exists.
    If it exists, create name like foo_1.json, foo_2.json, ...
    """
    path = path.resolve()
    if not path.exists():
        path.write_text(text)
        print(f"Wrote config: {path}")
        return

    stem = path.stem
    suffix = path.suffix or ".json"
    parent = path.parent

    i = 1
    while True:
        candidate = parent / f"{stem}_{i}{suffix}"
        if not candidate.exists():
            candidate.write_text(text)
            print(f"Config already existed; wrote instead: {candidate}")
            return
        i += 1

@Gooey(
    program_name="Flyer Config Builder",
    required_cols=2,
    default_size=(820, 780),
    navigation="TABBED",
)
def main():
    root = Path(".").resolve()

    igsn_choices = find_files(root, "igsn-config", (".json",))
    tmp_choices = find_files(root, "LB-templates", (".lbrn2",))

    if not igsn_choices:
        igsn_choices = ["<none found>"]
    if not tmp_choices:
        tmp_choices = ["<none found>"]

    parser = GooeyParser(description="Generate a constrained config.json with dropdowns + validation.")

    # EXPERIMENT
    exp = parser.add_argument_group("EXPERIMENT")
    exp.add_argument("--ID", help="Required. Exactly 4 digits (e.g., 0001).", required=True)
    exp.add_argument("--operator", help="Required. (Whatever cfstack.py expects)", required=True)
    exp.add_argument(
        "--igsn_config",
        choices=igsn_choices,
        default=igsn_choices[0],
        help="Dropdown of root/igsn-config/*.json",
        required=True,
    )
    exp.add_argument(
        "--tmp_file",
        choices=tmp_choices,
        default=tmp_choices[0],
        help="Dropdown of root/LB-templates/*.lbrn2",
        required=True,
    )
    exp.add_argument(
        "--tmp_ID_placeholder",
        default="FLYID",
        help="Default FLYID. Usually do NOT change unless your template uses a different placeholder.",
    )

    # OUTPUT (LightBurn output behavior; NOT config file overwrite)
    outg = parser.add_argument_group("OUTPUT")
    outg.add_argument("--output_dir", default="output", help="Output directory name.")
    outg.add_argument("--output_dir_append_IGSN", action="store_true", default=True,
                      help="(default: checked) Append IGSN to output directory.")
    outg.add_argument("--output_base", default="STACK", help='Base output name (default: "STACK").')
    outg.add_argument("--output_append_ID", action="store_true", default=True,
                      help="(default: checked) Append ID to output name.")
    outg.add_argument("--output_append_template", action="store_true", default=False,
                      help="(default: unchecked) Append template name to output name.")
    outg.add_argument("--output_append_timestamp", action="store_true", default=True,
                      help="(default: checked) Append timestamp to output name.")
    outg.add_argument("--output_overwrite", action="store_true", default=False,
                      help="(default: unchecked) Overwrite existing *LightBurn outputs*.")

    # FLYER SELECTION
    fs = parser.add_argument_group("FLYER SELECTION")
    fs.add_argument("--start", type=int, default=1, help="Start index (0-30).")
    fs.add_argument("--end", type=int, default=25, help="End index (0-30).")
    fs.add_argument("--step", type=int, default=1, help="Step (0-30).")

    # FLYER ASSIGNMENT
    fa = parser.add_argument_group("FLYER ASSIGNMENT")
    fa.add_argument("--style", choices=ASSIGNMENT_STYLES, default="exact",
                    help="Assignment style: exact | repeat_x | modulus_x")
    fa.add_argument("--x", type=int, default=1, help="Style parameter x (0-30).")

    # CONFIG FILE (saved config.json) — separate & never overwritten
    cfg_group = parser.add_argument_group("CONFIG FILE")
    cfg_group.add_argument(
        "--out_dir",
        default=".",
        help="Folder to save the config file into (default: current folder).",
    )
    cfg_group.add_argument(
        "--out_name",
        default="",
        help="Optional override for config filename. Leave blank to use IDconfig.json.",
    )

    args = parser.parse_args()
    validate(args, igsn_choices, tmp_choices)

    cfg = {
        "ID": args.ID.strip(),
        "operator": args.operator.strip(),
        "igsn_config": args.igsn_config,
        "tmp_file": args.tmp_file,
        "tmp_ID_placeholder": (args.tmp_ID_placeholder.strip() if args.tmp_ID_placeholder else "FLYID"),
        "output_dir": (args.output_dir.strip() if args.output_dir else "output"),
        "output_dir_append_IGSN": bool(args.output_dir_append_IGSN),
        "output_base": (args.output_base.strip() if args.output_base else "STACK"),
        "output_append_ID": bool(args.output_append_ID),
        "output_append_template": bool(args.output_append_template),
        "output_append_timestamp": bool(args.output_append_timestamp),
        "output_overwrite": bool(args.output_overwrite),
        "flyer_selection": {"start": int(args.start), "end": int(args.end), "step": int(args.step)},
        "flyer_assignment": {
            "_comment": "flyer_assignment.style: exact | repeat_x | modulus_x",
            "style": args.style,
            "x": int(args.x),
        },
    }

    # Default config filename: IDconfig.json
    out_dir = Path(args.out_dir).expanduser().resolve()
    out_dir.mkdir(parents=True, exist_ok=True)

    filename = (args.out_name.strip() if args.out_name and args.out_name.strip()
                else f"{cfg['ID']}config.json")

    out_path = out_dir / filename

    # Never overwrite the CONFIG file; if exists, write suffix _1, _2, ...
    safe_write_no_overwrite(out_path, json.dumps(cfg, indent=2))

if __name__ == "__main__":
    main()
