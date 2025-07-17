import json
import openpyxl
import re
from difflib import SequenceMatcher

name = "CALENTONES"
# EXCEL_FILE_PATH = "/home/miral/temp/geopandas/codificaciones/normalizer/result.xlsx"
EXCEL_FILE_PATH = "/home/miral/problems/python_geocode/split_col_dir/CALENTONES.xlsx"
JSON_FILE_PATH = "/home/miral/temp/geopandas/codificaciones/calles_col/MapaColoniasCallesGPD.json"
OUTPUT_EXCEL_FILE = f"{name}.xlsx"
SEQUENCE_MATCHER_THRESHOLD = 0.4
try:
    wb = openpyxl.load_workbook(EXCEL_FILE_PATH)
    sheet = wb[name]
except FileNotFoundError:
    print(f"Error: Excel file not found at {EXCEL_FILE_PATH}")
    exit()

try:
    with open(JSON_FILE_PATH, "r") as file:
        json_data = json.load(file)
except FileNotFoundError:
    print(f"Error: JSON file not found at {JSON_FILE_PATH}")
    exit()
except json.JSONDecodeError:
    print(f"Error: Could not decode JSON from {JSON_FILE_PATH}")
    exit()

# Create a set of all unique, normalized directions for quick 'in' checks
# And a dictionary to map normalized directions to their parent colonias (can be multiple)
normalized_directions_set = set()
# Maps a normalized direction to a list of its colonias
direction_to_colonia_map = {}

# Create a map in which each direction has a colonia
for colonia_name, directions_list in json_data.items():
    normalized_colonia = colonia_name.strip().upper()
    for direction_name in directions_list:
        normalized_direction = direction_name.strip().upper()
        normalized_directions_set.add(normalized_direction)
        if normalized_direction not in direction_to_colonia_map:
            direction_to_colonia_map[normalized_direction] = []
        direction_to_colonia_map[normalized_direction].append(
            normalized_colonia)

# Create a set of normalized colonias for quick 'in' checks
normalized_colonias_set = {
    col.strip().upper() for col in json_data.keys()}


def normalize_string(text):
    if not isinstance(text, str):
        return None
    # text = re.sub(r"\d", "", text)
    text = text.strip()
    text = text.upper()
    return text


processed_count = 0
matched_directions_count = 0
matched_colonias_count = 0

DIRECCION_COL_INDEX = 10
COLONIA_COL_INDEX = 11

for row_idx, row in enumerate(sheet.iter_rows(min_row=2, values_only=False), start=8):
    direccion_cell = row[DIRECCION_COL_INDEX - 1]
    colonia_cell = row[COLONIA_COL_INDEX - 1]

    original_direccion = direccion_cell.value
    original_colonia = colonia_cell.value

    normalized_direccion = normalize_string(original_direccion)
    normalized_colonia = normalize_string(original_colonia)

    if normalized_colonia:
        if normalized_colonia in normalized_colonias_set:
            colonia_cell.value = normalized_colonia
            matched_colonias_count += 1
            if normalized_direccion:
                colonia_directions = json_data.get(original_colonia, [])
                normalized_colonia_directions = {
                    normalize_string(d) for d in colonia_directions}

                if normalized_direccion in normalized_colonia_directions:
                    direccion_cell.value = normalized_direccion
                    matched_directions_count += 1
                else:
                    best_match_dir = None
                    max_ratio_dir = SEQUENCE_MATCHER_THRESHOLD
                    for json_dir in normalized_colonia_directions:
                        if json_dir is None:
                            continue
                        s_ratio = SequenceMatcher(
                            None, normalized_direccion, json_dir).ratio()
                        if s_ratio > max_ratio_dir:
                            max_ratio_dir = s_ratio
                            best_match_dir = json_dir
                            if max_ratio_dir == 1:
                                break
                    if best_match_dir:
                        direccion_cell.value = best_match_dir
                        matched_directions_count += 1

        else:
            best_match_col = None
            max_ratio_col = SEQUENCE_MATCHER_THRESHOLD
            for json_col in normalized_colonias_set:
                s_ratio = SequenceMatcher(
                    None, normalized_colonia, json_col).ratio()
                if s_ratio > max_ratio_col:
                    max_ratio_col = s_ratio
                    best_match_col = json_col
                    if max_ratio_col == 1:
                        break
            if best_match_col:
                colonia_cell.value = best_match_col
                matched_colonias_count += 1
                if normalized_direccion:
                    original_best_match_col_key = next(
                        (k for k, v in json_data.items() if normalize_string(k) == best_match_col), None)
                    if original_best_match_col_key:
                        colonia_directions = json_data.get(
                            original_best_match_col_key, [])
                        normalized_colonia_directions = {
                            normalize_string(d) for d in colonia_directions}

                        if normalized_direccion in normalized_colonia_directions:
                            direccion_cell.value = normalized_direccion
                            matched_directions_count += 1
                        else:
                            best_match_dir = None
                            max_ratio_dir = SEQUENCE_MATCHER_THRESHOLD
                            for json_dir in normalized_colonia_directions:
                                if json_dir is None:
                                    continue
                                s_ratio = SequenceMatcher(
                                    None, normalized_direccion, json_dir).ratio()
                                if s_ratio > max_ratio_dir:
                                    max_ratio_dir = s_ratio
                                    best_match_dir = json_dir
                                    if max_ratio_dir == 1:
                                        break
                            if best_match_dir:
                                direccion_cell.value = best_match_dir
                                matched_directions_count += 1

    elif normalized_direccion:
        if normalized_direccion in normalized_directions_set:
            direccion_cell.value = normalized_direccion
            matched_directions_count += 1
        else:
            best_match_dir = None
            max_ratio_dir = SEQUENCE_MATCHER_THRESHOLD
            for json_dir in normalized_directions_set:
                s_ratio = SequenceMatcher(
                    None, normalized_direccion, json_dir).ratio()
                if s_ratio > max_ratio_dir:
                    max_ratio_dir = s_ratio
                    best_match_dir = json_dir
                    if max_ratio_dir == 1:
                        break
            if best_match_dir:
                direccion_cell.value = best_match_dir
                matched_directions_count += 1

    processed_count += 1

print(f"\nProcessing complete!")
print(f"Total rows processed: {processed_count}")
print(f"Total Direcciones matched/updated: {matched_directions_count}")
print(f"Total Colonias matched/updated: {matched_colonias_count}")

try:
    wb.save(OUTPUT_EXCEL_FILE)
    print(f"Updated data saved to {OUTPUT_EXCEL_FILE}")
except Exception as e:
    print(f"Error saving workbook: {e}")
finally:
    wb.close()
    print("Workbook closed.")
