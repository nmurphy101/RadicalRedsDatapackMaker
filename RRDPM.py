import math
import pandas as pd
import json
import os
import argparse
import sys
import re
from typing import Any


# Default sheets (pokemon types). Used when `--sheet all` is passed.
DEFAULT_SHEETS = [
    'Normal', 'Fire', 'Water', 'Grass', 'Electric', 'Ice', 'Fighting', 'Poison', 'Ground',
    'Flying', 'Psychic', 'Bug', 'Rock', 'Ghost', 'Dragon', 'Dark', 'Steel', 'Fairy'
]

def _parse_cell_value(val: Any):
    """Try to convert string cell values into native Python types where sensible.

    - If val is already not a str, return it.
    - If val looks like JSON (starts with "[" or "{" or is true/false/null), try json.loads.
    - If val contains semicolons, split into a list.
    - If val looks like an int or float, convert.
    - Otherwise return the original string.
    """
    if val is None:
        return None
    if not isinstance(val, str):
        return val
    s = val.strip()
    if s == "":
        return ""
    # Try JSON
    if s[0] in "[{" or s in ("true", "false", "null"):
        try:
            return json.loads(s)
        except Exception:
            pass
    # Semicolon separated list
    if ";" in s:
        parts = [p.strip() for p in s.split(";") if p.strip() != ""]
        return parts
    # Numeric
    if re.fullmatch(r"-?\d+", s):
        try:
            return int(s)
        except Exception:
            pass
    if re.fullmatch(r"-?\d+\.\d+", s):
        try:
            return float(s)
        except Exception:
            pass
    # fallback
    return s


def _render_node(node: Any, row: pd.Series, pokemon_list: list | None = None):
    """Recursively render a JSON-like structure where strings may be placeholders like {{KEY}}.

    If a string exactly equals a placeholder and the placeholder maps to a complex value
    (e.g. `pokemon_list`), the function will return that Python value instead of a string.
    """
    if isinstance(node, dict):
        return {k: _render_node(v, row, pokemon_list) for k, v in node.items()}
    if isinstance(node, list):
        return [_render_node(v, row, pokemon_list) for v in node]
    if not isinstance(node, str):
        return node

    # Find placeholders in the string
    placeholders = re.findall(r"\{\{([A-Za-z0-9_]+)\}\}", node)
    # If the string is exactly a single placeholder, return the appropriate typed value
    m = re.fullmatch(r"\{\{([A-Za-z0-9_]+)\}\}", node)
    if m:
        key = m.group(1)
        if key == "POKEMON_TEAM_LIST" and pokemon_list is not None:
            return pokemon_list
        # otherwise get from row
        return _parse_cell_value(row.get(key, ""))

    # Otherwise replace any placeholders inside the string with their string values
    def repl(match):
        k = match.group(1)
        if k == "POKEMON_TEAM_LIST" and pokemon_list is not None:
            # embed JSON string representation
            try:
                return json.dumps(pokemon_list, ensure_ascii=False)
            except Exception:
                return str(pokemon_list)
        v = row.get(k, "")
        return "" if v is None else str(v)

    rendered = re.sub(r"\{\{([A-Za-z0-9_]+)\}\}", repl, node)
    return rendered


def _build_pokemon_list(row: pd.Series, pokemon_template_obj: Any):
    """Detect POKEMON_N_* columns, build list of pokemon dicts using the pokemon template object."""
    # Find all columns in the row that match POKEMON_<index>_<FIELD>
    indices = set()
    for col in row.index:
        m = re.match(r"^POKEMON_(\d+)_", str(col))
        if m:
            indices.add(int(m.group(1)))
    if not indices:
        # Maybe there"s a single column "POKEMON_TEAM" with JSON or semicolon list
        team_cell = row.get("POKEMON_TEAM")
        if team_cell is None:
            return []
        parsed = _parse_cell_value(team_cell)
        if isinstance(parsed, list):
            # If items are dict-like strings, try to parse them
            out = []
            for item in parsed:
                if isinstance(item, str):
                    try:
                        out.append(json.loads(item))
                        continue
                    except Exception:
                        pass
                out.append(item)
            return out
        return [parsed]

    pokemon_list = []
    for i in sorted(indices):
        # For each pokemon, render the template using columns like POKEMON_{i}_FIELD
        def _get_val_for_placeholder(placeholder_name: str):
            # placeholder_name like POKEMON_NAME or POKEMON_LEVEL
            if placeholder_name.startswith("POKEMON_"):
                suffix = placeholder_name[len("POKEMON_"):]
                colname = f"POKEMON_{i}_{suffix}"
                if colname in row.index:
                    return _parse_cell_value(row.get(colname))
            # fallback to column without index
            if placeholder_name in row.index:
                return _parse_cell_value(row.get(placeholder_name))
            return ""

        # Render a copy of the pokemon template object, replacing placeholders exactly
        def _render_pokemon_node(node):
            if isinstance(node, dict):
                return {k: _render_pokemon_node(v) for k, v in node.items()}
            if isinstance(node, list):
                return [_render_pokemon_node(v) for v in node]
            if not isinstance(node, str):
                return node
            m = re.fullmatch(r"\{\{([A-Za-z0-9_]+)\}\}", node)
            if m:
                key = m.group(1)
                return _get_val_for_placeholder(key)
            # replace any placeholders inside the string
            def repl(match):
                k = match.group(1)
                v = _get_val_for_placeholder(k)
                return "" if v is None else str(v)
            return re.sub(r"\{\{([A-Za-z0-9_]+)\}\}", repl, node)

        rendered = _render_pokemon_node(pokemon_template_obj)
        pokemon_list.append(rendered)

    return pokemon_list


def _pokemon_from_row(row: pd.Series, pokemon_template_obj: Any):
    """Build a single pokemon dict from a row using the pokemon template and expected column names.

    Expects columns: "Pokemon", "Gender", "Aspect", "Level", "Slot 1".."Slot 4", "Ability", "Item", "IVs", "EVs", "Nature", "Format", "Badge Level"
    """
    # print("Building pokemon for row:", row.to_dict())
    # Build a placeholder mapping
    def cell(name):
        v = None
        # support both exact names and common lowercase variants
        if name.strip() in row.index:
            v = row.get(name)
        else:
            # try case-insensitive
            for col in row.index:
                if str(col).lower().strip() == name.lower().strip():
                    v = row.get(col)
                    break
        return v

    mapping = {}
    mapping["POKEMON_NAME"] = cell("Pokemon").lower() or ""
    mapping["POKEMON_LEVEL"] = _parse_cell_value(cell("Level") or cell("level") or "")

    # Aspect is a list of comma separated values
    aspect_cell = cell("Aspect")
    if type(aspect_cell) in (str, int, float) and aspect_cell != "" and not (isinstance(aspect_cell, float) and math.isnan(aspect_cell)):
        # ensure string, split by comma and ignore empty items
        mapping["POKEMON_ASPECT_LIST"] = [x.strip() for x in str(aspect_cell).split(",") if x.strip() != ""]

    # moveset from Slot 1..Slot 4
    moves = []
    for i in range(0, 5):
        move = cell(f"Slot {i}")
        if move is not None and move is not False and str(move).strip() != "":
            # Sanitize move name to have no spaces at all and all lowercase and no special characters
            move = re.sub(r"\s+", "", str(move).strip()).lower()
            move = re.sub(r"[^a-z0-9_]", "", move)
            moves.append(move)

    mapping["POKEMON_MOVESET_LIST"] = moves
    # Sanitize ability to have no spaces at all and all lowercase and no special characters
    mapping["POKEMON_ABILITY"] = re.sub(r"\s+", "", cell("Ability") or "").lower()
    mapping["POKEMON_ABILITY"] = re.sub(r"[^a-z0-9_]", "", mapping["POKEMON_ABILITY"])
    # Sanitize held item to have no spaces at all replaced with underscores and all lowercase and no special characters
    held_item = cell("Item") or ""
    held_item = re.sub(r"\s+", "_", str(held_item).strip()).lower()
    held_item = re.sub(r"[^a-z0-9_]", "", held_item)
    mapping["POKEMON_HELD_ITEM"] = held_item
    mapping["POKEMON_NATURE"] = cell("Nature").lower() or "docile"
    mapping["POKEMON_GENDER"] = cell("Gender").upper() or "MALE"

    # IVs and EVs: single integer used for all stats
    iv_val = _parse_cell_value(cell("IVs"))
    ev_val = _parse_cell_value(cell("EVs"))
    for stat in ("HP", "ATK", "DEF", "SPA", "SPD", "SPE"):
        mapping[f"POKEMON_IV_{stat}"] = int(iv_val) if isinstance(iv_val, (int, float)) else iv_val
        mapping[f"POKEMON_EV_{stat}"] = int(ev_val) if isinstance(ev_val, (int, float)) else ev_val

    # Render pokemon template object by replacing placeholders from mapping
    def _render_pokemon_node_from_map(node):
        if isinstance(node, dict):
            return {k: _render_pokemon_node_from_map(v) for k, v in node.items()}
        if isinstance(node, list):
            return [_render_pokemon_node_from_map(v) for v in node]
        if not isinstance(node, str):
            return node
        m = re.fullmatch(r"\{\{([A-Za-z0-9_]+)\}\}", node)
        if m:
            key = m.group(1)
            return mapping.get(key, "")
        def repl(match):
            k = match.group(1)
            return "" if mapping.get(k) is None else str(mapping.get(k))
        return re.sub(r"\{\{([A-Za-z0-9_]+)\}\}", repl, node)

    return _render_pokemon_node_from_map(pokemon_template_obj)


def main(
        excel_file,
        sheet_name,
        output_dir="output_jsons",
        leader_template_path="templates/gym_leader_template.json",
        pokemon_template_path="templates/pokemon_template.json",
        mobs_template_path="templates/mob_trainer_group_template.json",
):
    # Validate input file
    if not os.path.isfile(excel_file):
        print(f"Error: Excel file '{excel_file}' not found.", file=sys.stderr)
        sys.exit(1)

    trainers_dir = f"{output_dir}/trainers"
    mobs_dir = f"{output_dir}/mobs/trainers/groups"

    # Ensure an output directory exists
    os.makedirs(trainers_dir, exist_ok=True)
    os.makedirs(mobs_dir, exist_ok=True)

    # Load templates
    if not os.path.isfile(leader_template_path):
        print(f"Error: template file '{leader_template_path}' not found.", file=sys.stderr)
        sys.exit(1)
    with open(leader_template_path, "r", encoding="utf-8") as tf:
        gym_template_obj = json.load(tf)

    pokemon_template_obj = None
    if os.path.isfile(pokemon_template_path):
        with open(pokemon_template_path, "r", encoding="utf-8") as pf:
            pokemon_template_obj = json.load(pf)

    mobs_template_obj = None
    if os.path.isfile(mobs_template_path):
        with open(mobs_template_path, "r", encoding="utf-8") as mf:
            mobs_template_obj = json.load(mf)

    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel(excel_file, sheet_name=sheet_name)

    # Forward-fill merged header cells (Excel often uses merged cells for leader name/badge level)
    ffill_cols = []
    for c in ("Badge Level", "Leader Name"):
        if c in df.columns:
            ffill_cols.append(c)
    if ffill_cols:
        df[ffill_cols] = df[ffill_cols].ffill()

    # If the spreadsheet uses rows for individual pokemon (grouped by Badge Level),
    # group them and create one gym JSON per badge level.
    if "Badge Level" in df.columns:
        grouped = df.groupby("Badge Level")
        created = 0
        for badge_level, group in grouped:
            # Use the first row as representative for gym-level fields
            first = group.iloc[0]
            group_map = dict(first.to_dict())
            leader_name = first.get('Leader Name') or first.get('Leader', '')
            group_map['LEADER_NAME'] = f"{leader_name}_{badge_level}"

            # filename: prefer provided LEADER_NAME
            base_name = group_map.get("LEADER_NAME")

            # ignore groups without a leader name
            if "Leader Name" in base_name:
                continue

            safe = re.sub(r"[^A-Za-z0-9._-]+", "_", str(base_name)).strip("_")

            # Battle format from "Format" column if present
            if "BATTLE_FORMAT" not in group_map or not group_map.get("BATTLE_FORMAT"):
                if "Format" in group.columns:
                    group_map["BATTLE_FORMAT"] = first.get("Format")

            # Build pokemon list from all rows in this group
            pokemon_list = []
            if pokemon_template_obj is not None:
                for _, prow in group.iterrows():
                    p = _pokemon_from_row(prow, pokemon_template_obj)
                    pokemon_list.append(p)

            rendered_obj = _render_node(gym_template_obj, group_map, pokemon_list)
            filename = os.path.join(trainers_dir, f"{safe}.json")
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(rendered_obj, f, indent=2, ensure_ascii=False)
            created += 1

            if mobs_template_obj is not None:
                # Also create mob trainer group JSON
                mob_rendered = _render_node(mobs_template_obj, group_map, pokemon_list)
                mob_filename = os.path.join(mobs_dir, f"chickencoopleader_{safe}.json")
                with open(mob_filename, "w", encoding="utf-8") as mf:
                    json.dump(mob_rendered, mf, indent=2, ensure_ascii=False)

        print(f"Successfully created {created} JSON files in the '{trainers_dir}' directory.")
        return


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert Excel rows to individual JSON files.")
    parser.add_argument("--excel_file", "-e", default="trainers.xlsx", help="Path to the Excel file to convert")
    parser.add_argument("--sheet", "-s", default="all", help="Sheet name to read (default: all)")
    parser.add_argument("--outdir", "-o", default="output_jsons", help="Output directory (default: output_jsons)")
    parser.add_argument("--leader-template", "-l", default="templates/gym_leader_template.json", help="Gym Leader template JSON path")
    parser.add_argument("--pokemon-template", "-p", default="templates/pokemon_template.json", help="Pokemon template JSON path")
    parser.add_argument("--mobs-template", "-m", default="templates/mob_trainer_group_template.json", help="Mob trainer group template JSON path")
    parser.add_argument("--trainer-type-template", "-t", default="templates/trainer_type_template.json", help="Trainer type template JSON path")

    args = parser.parse_args()

    if args.sheet == "all":
        for sheet in DEFAULT_SHEETS:
            print(f"Processing sheet: {sheet}")
            main(
                args.excel_file,
                sheet_name=sheet,
                output_dir=args.outdir,
                leader_template_path=args.leader_template,
                pokemon_template_path=args.pokemon_template,
                mobs_template_path=args.mobs_template,

            )
    else:
        main(
            args.excel_file,
            sheet_name=args.sheet,
            output_dir=args.outdir,
            leader_template_path=args.leader_template,
            pokemon_template_path=args.pokemon_template,
            mobs_template_path=args.mobs_template,
        )

    trainer_types_dir = f"{args.outdir}/trainer_types"
    os.makedirs(trainer_types_dir, exist_ok=True)

    trainer_type_template_obj = None
    if os.path.isfile(args.trainer_type_template):
        with open(args.trainer_type_template, "r", encoding="utf-8") as tf:
            trainer_type_template_obj = json.load(tf)

    if trainer_type_template_obj is not None:
        # Create trainer types JSON
        trainer_types_filename = os.path.join(trainer_types_dir, f"gymleader_chickencoop.json")
        with open(trainer_types_filename, "w", encoding="utf-8") as tf:
            json.dump(trainer_type_template_obj, tf, indent=2, ensure_ascii=False)

    print(f"Successfully created trainer types JSON file in the '{trainer_types_dir}' directory.")
