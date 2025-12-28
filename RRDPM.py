import math
import shutil
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
    'Flying', 'Psychic', 'Bug', 'Rock', 'Ghost', 'Dragon', 'Dark', 'Steel', 'Fairy',
    "E1A", "E1B", "E2A", "E2B", "E3A", "E3B", "E4A", "E4B", "C1", "C2", "C3", "C4",
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
    # Sanitize move name to have no spaces at all and all lowercase and no special characters
    mapping["POKEMON_NAME"] = re.sub(r"\s+", "", cell("Pokemon") or "").lower()
    mapping["POKEMON_NAME"] = re.sub(r"[^a-z0-9_]", "", mapping["POKEMON_NAME"])

    mapping["POKEMON_LEVEL"] = _parse_cell_value(cell("Level") or cell("level") or "")

    # Aspect is a list of comma separated values
    aspect_cell = cell("Aspect")
    if type(aspect_cell) in (str, int, float) and aspect_cell != "" and not (isinstance(aspect_cell, float) and math.isnan(aspect_cell)):
        # ensure list, split by comma and ignore empty items
        mapping["POKEMON_ASPECT_LIST"] = [x.strip() for x in str(aspect_cell).split(",") if x.strip() != ""]
    else:
        mapping["POKEMON_ASPECT_LIST"] = []

    # moveset from Slot 1..Slot 4
    moves = []
    for i in range(0, 5):
        move = cell(f"Slot {i}")
        if move is not None and move is not False and str(move).strip() != "" and not (isinstance(move, float) and math.isnan(move)):
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
        leader_names=[],
        elite_4_names=[],
):
    # Validate input file
    if not os.path.isfile(excel_file):
        print(f"Error: Excel file '{excel_file}' not found.", file=sys.stderr)
        sys.exit(1)

    trainers_dir = f"{output_dir}/trainers"
    # Ensure an output directory exists
    os.makedirs(trainers_dir, exist_ok=True)

    # Load templates
    if not os.path.isfile(leader_template_path):
        print(f"Error: template file '{leader_template_path}' not found.", file=sys.stderr)
        sys.exit(1)
    with open(leader_template_path, "r", encoding="utf-8") as tf:
        gym_template_obj = json.load(tf)

    # Pokemon template JSON creation
    pokemon_template_obj = None
    if os.path.isfile(pokemon_template_path):
        with open(pokemon_template_path, "r", encoding="utf-8") as pf:
            pokemon_template_obj = json.load(pf)

    # Advancement trainer JSON creation
    advancement_template_obj = None
    advancement_template_filename = "templates/advancement_trainer_template.json"
    if os.path.isfile(advancement_template_filename):
        with open(advancement_template_filename, "r", encoding="utf-8") as af:
            advancement_template_obj = json.load(af)

    # Mob trainer group JSON creation
    mobs_dir = f"{output_dir}/mobs/trainers/groups"
    os.makedirs(mobs_dir, exist_ok=True)

    mobs_template_obj = None
    if os.path.isfile(args.mobs_template):
        with open(args.mobs_template, "r", encoding="utf-8") as mf:
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
            if "Badge Level" == badge_level: continue  # skip header row if present
            badge_level = int(badge_level)
            # Use the first row as representative for gym-level fields
            first = group.iloc[0]
            group_map = dict(first.to_dict())
            leader_name = (first.get('Leader Name') or first.get('Leader', '')).lower().strip()
            group_map['LEADER_NAME'] = f"{leader_name}" if leader_name else f"leader_{badge_level}"
            group_map['LEADER_DISPLAY_NAME'] = f"{leader_name[0].upper()}{leader_name[1:]}" if leader_name else f"leader_{badge_level}"
            group_map['BADGE_LEVEL'] = badge_level

            # ignore groups without a leader name
            if "leader name" in leader_name:
                continue

            safe = re.sub(r"[^A-Za-z0-9._-]+", "_", str(leader_name)).strip("_")

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
            filename = os.path.join(trainers_dir, f"chickencoopleader_{safe}_{badge_level}.json")
            with open(filename, "w", encoding="utf-8") as f:
                json.dump(rendered_obj, f, indent=2, ensure_ascii=False)

            # Also create advancement JSON if template provided
            if advancement_template_obj is not None:
                advancement_obj = _render_node(advancement_template_obj, group_map)
                advancement_dir = f"{output_dir}/advancement/trainers"
                os.makedirs(advancement_dir, exist_ok=True)
                advancement_filename = os.path.join(advancement_dir, f"defeat_chickencoopleader_{safe}_{badge_level}.json")
                with open(advancement_filename, "w", encoding="utf-8") as af:
                    json.dump(advancement_obj, af, indent=2, ensure_ascii=False)

            # Also update mob trainer group template if present elite_4_names
            if mobs_template_obj is not None:
                required_defeats = []
                if badge_level == 9:
                    required_defeats = [f"chickencoopleader_{known_leader_name}_{badge_level-1}" for index, known_leader_name in enumerate(elite_4_names) if known_leader_name != leader_name and index < 8]
                elif badge_level == 8:
                    required_defeats = [f"chickencoopleader_{known_leader_name}_{badge_level-1}" for known_leader_name in leader_names if known_leader_name != leader_name]
                elif badge_level >= 1:
                    required_defeats = [f"chickencoopleader_{known_leader_name}_{badge_level-1}" for known_leader_name in leader_names if known_leader_name != leader_name]
                group_map["REQUIRED_DEFEATS"] = [required_defeats] if required_defeats else []
                mobs_obj = _render_node(mobs_template_obj, group_map)
                mob_filename = os.path.join(mobs_dir, f"chickencoopleader_{safe}_{badge_level}.json")
                with open(mob_filename, "w", encoding="utf-8") as mf:
                    json.dump(mobs_obj, mf, indent=2, ensure_ascii=False)

            created += 1

        return


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert Excel rows to individual JSON files.")
    parser.add_argument("--excel_file", "-e", default="trainers.xlsx", help="Path to the Excel file to convert")
    parser.add_argument("--sheet", "-s", default="all", help="Sheet name to read (default: all)")
    parser.add_argument("--outdir", "-o", default="MoreGymLeaders", help="Output directory (default: MoreGymLeaders)")
    parser.add_argument("--leader-template", "-l", default="templates/gym_leader_template.json", help="Gym Leader template JSON path")
    parser.add_argument("--pokemon-template", "-p", default="templates/pokemon_template.json", help="Pokemon template JSON path")
    parser.add_argument("--mobs-template", "-m", default="templates/mob_trainer_group_template.json", help="Mob trainer group template JSON path")
    parser.add_argument("--trainer-type-template", "-t", default="templates/trainer_type_template.json", help="Trainer type template JSON path")

    args = parser.parse_args()

    # remove the output directory if it exists
    if os.path.isdir(args.outdir):
        shutil.rmtree(args.outdir)

    # Load optional command/settings config json
    settings_json = None
    settings_filename = "settings.json"
    if os.path.isfile(settings_filename):
        with open(settings_filename, "r", encoding="utf-8") as tf:
            settings_json = json.load(tf)

    if settings_json is not None:
        # Override args with settings.json values if present
        for key in vars(args).keys():
            if key in settings_json:
                setattr(args, key, settings_json[key])

    output_dir = f"{args.outdir}/data/rctmod"

    if args.sheet == "all":
        leader_names = []
        elite_4_names = []
        for sheet in DEFAULT_SHEETS:
            # get all the leader names from the excel sheets and save them in a list
            # Load the Excel file into a pandas DataFrame
            print(f"Reading leader names from sheet: {sheet}", end=f"{' ' * 30}\r")
            df = pd.read_excel(args.excel_file, sheet_name=sheet)
            badge_level_grouped = df.groupby("Badge Level")
            for badge_level, group in badge_level_grouped:
                if "Badge Level" == badge_level: continue  # skip header row if present
                badge_level = int(badge_level)
                # Use the first row as representative for gym-level fields
                first = group.iloc[0]
                group_map = dict(first.to_dict())
                leader_name = (first.get('Leader Name') or first.get('Leader', '')).lower().strip()
                if "Leader Name" in leader_name:
                    continue
                if badge_level >= 8 and leader_name not in elite_4_names:
                    # sanitize including replacing spaces with underscores
                    elite_4_names.append(df["Leader Name"].dropna().unique().tolist()[0].lower().strip().replace(" ", "_"))
                elif badge_level < 8 and leader_name not in leader_names:
                    leader_names.append(df["Leader Name"].dropna().unique().tolist()[0].lower().strip().replace(" ", "_"))

        print(f"Found leader names: {leader_names}\nFound elite 4 names: {elite_4_names}")

        for sheet in DEFAULT_SHEETS:
            print(f"Processing sheet: {sheet}", end=f"{' ' * 30}\r")
            main(
                args.excel_file,
                sheet_name=sheet,
                output_dir=output_dir,
                leader_template_path=args.leader_template,
                pokemon_template_path=args.pokemon_template,
                leader_names=leader_names,
                elite_4_names=elite_4_names,

            )
    else:
        leader_names = []
        # get all the leader names from the excel sheets and save them in a list
        # Load the Excel file into a pandas DataFrame
        df = pd.read_excel(args.excel_file, sheet_name=args.sheet)
        if "Leader Name" in df.columns:
            leader_names.append(df["Leader Name"].dropna().unique().tolist()[0].lower().strip())

        print(f"Found leader names: {leader_names}")

        main(
            args.excel_file,
            sheet_name=args.sheet,
            output_dir=output_dir,
            leader_template_path=args.leader_template,
            pokemon_template_path=args.pokemon_template,
            leader_names=leader_names,
        )

    # move pack.mcmeta file in templates to output dir
    pack_dir = f"{args.outdir}"
    os.makedirs(pack_dir, exist_ok=True)

    pack_mcmeta_src = "templates/pack.mcmeta"
    pack_mcmeta_dst = os.path.join(pack_dir, "pack.mcmeta")
    if os.path.isfile(pack_mcmeta_src):
        shutil.copyfile(pack_mcmeta_src, pack_mcmeta_dst)
        print(f"Successfully copied pack.mcmeta file to the '{pack_dir}' directory.", end=f"{' ' * 30}\r")

    # move pack.png file in templates to output dir
    pack_image_src = "templates/pack.png"
    pack_image_dst = os.path.join(pack_dir, "pack.png")
    if os.path.isfile(pack_image_src):
        shutil.copyfile(pack_image_src, pack_image_dst)
        print(f"Successfully copied pack.png file to the '{pack_dir}' directory.", end=f"{' ' * 30}\r")

    # Trainer types JSON creation
    trainer_types_dir = f"{output_dir}/trainer_types"
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

        print(f"Successfully created trainer types JSON file in the '{trainer_types_dir}' directory.", end=f"{' ' * 30}\r")

    # Loot table JSON creation
    loot_tables_dir = f"{output_dir}/loot_tables/mobs/trainers"
    os.makedirs(loot_tables_dir, exist_ok=True)
    loot_table_template_filename = "templates/loot_table_template.json"
    loot_table_template_obj = None
    if os.path.isfile(loot_table_template_filename):
        with open(loot_table_template_filename, "r", encoding="utf-8") as lf:
            loot_table_template_obj = json.load(lf)

    if loot_table_template_obj is not None:
        # Also create loot table JSON
        loot_table_filename = os.path.join(loot_tables_dir, f"chickencoopleader.json")
        with open(loot_table_filename, "w", encoding="utf-8") as lf:
            json.dump(loot_table_template_obj, lf, indent=2, ensure_ascii=False)

        print(f"Successfully created loot table JSON file in the '{loot_tables_dir}' directory.", end=f"{' ' * 30}\r")

    # Series JSON creation
    series_dir = f"{output_dir}/series"
    os.makedirs(series_dir, exist_ok=True)
    series_template_filename = "templates/series_template.json"
    series_template_obj = None
    if os.path.isfile(series_template_filename):
        with open(series_template_filename, "r", encoding="utf-8") as sf:
            series_template_obj = json.load(sf)

    if series_template_obj is not None:
        # Also create series JSON
        series_filename = os.path.join(series_dir, f"chickencoopgymchallenge.json")
        with open(series_filename, "w", encoding="utf-8") as sf:
            json.dump(series_template_obj, sf, indent=2, ensure_ascii=False)

        print(f"Successfully created series JSON file in the '{series_dir}' directory.", end=f"{' ' * 30}\r")

    # zip the output dir
    shutil.make_archive(args.outdir, 'zip', args.outdir)
    print(f"Successfully created zip archive '{args.outdir}.zip'.")

    # remove the output dir after zipping if settings.json says to
    if settings_json is None or settings_json.get("delete_output_dir") is True:
        shutil.rmtree(args.outdir)
        print(f"Successfully removed temporary output directory '{args.outdir}'.", end=f"{' ' * 30}\r")
