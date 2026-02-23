import os
import sys
import time
from collections import defaultdict

# Added error status for user feedback while creating final output.
# This script is designed to be run on CX-Supervisor PC to provide final wagon assignments.

if len(sys.argv) > 1:
    FileName_PLC = sys.argv[1]
else:
    FileName_PLC = "default_value"
    with open(r"C:\FBtemp\data\Sorting\ConfigWagon.txt") as f:
        for line in f:
            if line.startswith("FileName_PLC="):
                FileName_PLC = line.split("=", 1)[1].strip()
                break

File_Path = r"C:\FBtemp\data\Sorting"
INPUT_FILE = f"{File_Path}\\{FileName_PLC}_OPTwagon.txt"
SHELF_FILE = r"C:\FBtemp\data\Sorting\settings.txt"
OUTPUT_FILE = f"{File_Path}\\{FileName_PLC}_final.txt"

def read_shelves(file_path):
    wagons_config = {}
    with open(file_path, "r") as f:
        lines = [line for line in f if line.strip() and not line.strip().startswith("#")]
        if not lines:
            return wagons_config
        header_fields = [s.lower() for s in lines[0].strip().split(",")]
        possible_headers = {"wagon_number", "rows", "slots_per_row", "total_width", "total_height", "allowed_vip_keys", "allow_vip_key_mixing", "stacking_per_slot"}
        is_header = any(h in possible_headers for h in [s.replace(" ", "_") for s in header_fields])
        start_idx = 1 if is_header else 0
        for lineno, line in enumerate(lines[start_idx:], start=start_idx+1):
            parts = [s.strip() for s in line.strip().split(",")]
            if len(parts) < 6:
                raise ValueError("Each line: wagon_number,rows,slots_per_row,total_width,total_height,allowed_vip_keys,[allow_vip_key_mixing],[stacking_per_slot]")
            wagon_raw = parts[0]
            rows_raw = parts[1]
            slots_raw = parts[2]
            width_raw = parts[3]
            height_raw = parts[4]

            # allowed_vip_keys are everything from index 5 up to the control fields at the end
            allowed_vip_keys = [m.strip().lower() for m in parts[5:-2] if m.strip()]
            # allow_vip_key_mixing is second-to-last
            allow_vip_key_mixing = (parts[-2].strip().lower() == "yes") if len(parts) > 6 else False
            # stacking per slot is last
            stacking_per_slot = int(parts[-1]) if len(parts) > 7 and parts[-1].isdigit() else 1

            try:
                wagon = int(wagon_raw)
            except Exception as e:
                raise ValueError(f"Invalid wagon number on line {lineno}: '{wagon_raw}'") from e
            try:
                rows = int(rows_raw)
            except Exception as e:
                raise ValueError(f"Invalid rows for wagon {wagon} on line {lineno}: '{rows_raw}'") from e
            try:
                slots_per_row = int(slots_raw)
            except Exception as e:
                raise ValueError(f"Invalid slots_per_row for wagon {wagon} on line {lineno}: '{slots_raw}'") from e
            try:
                width = float(width_raw)
                height = float(height_raw)
            except Exception as e:
                raise ValueError(f"Invalid width/height for wagon {wagon} on line {lineno}: '{width_raw}', '{height_raw}'") from e

            # Additional sanity checks
            if rows <= 0 or slots_per_row <= 0:
                raise ValueError(f"Wagon {wagon} has non-positive rows/slots_per_row on line {lineno}: rows={rows}, slots={slots_per_row}")
            if width <= 0 or height <= 0:
                raise ValueError(f"Wagon {wagon} has non-positive width/height on line {lineno}: width={width}, height={height}")
            if stacking_per_slot <= 0:
                raise ValueError(f"Wagon {wagon} has invalid stacking_per_slot on line {lineno}: {stacking_per_slot}")

            wagons_config[wagon] = {
                "rows": rows,
                "slots_per_row": slots_per_row,
                "width": width,
                "height": height,
                "allowed_vip_keys": allowed_vip_keys,
                "allow_vip_key_mixing": allow_vip_key_mixing,
                "stacking_per_slot": stacking_per_slot
            }
    return wagons_config

class Shape:
    def __init__(self, fields, header):
        self.fields = {k.lower(): v for k, v in fields.items()}
        self.header = header
        self.pol = int(self.fields.get("pol", 0))
        self.id = self.fields.get("serieid", "")
        self.width = float(self.fields.get("mouldedwidth", 0))
        self.height = float(self.fields.get("mouldedheight", 0))
        self.length = float(self.fields.get("finishedlength", 0))
        self.material = self.fields.get("material", "").strip().lower()
        self.vip_key_original = self.fields.get("vipkey", "").strip()
        self.vip_key = self.vip_key_original.lower()
        self.placed = 0  # Always 0 initially in sorter output
        self.slot = None
        self.row = None
        self.wagon = None
        self.rotated = False
        self._final_placement = None
        self.error_unplaced = False

        self.error_code = None
        self.error_msg = None

    def output_line(self, filtered_header):
        values_filtered = [self.fields.get(h.lower(), "") for h in filtered_header]
        if self.error_unplaced:
            wagon_val = f"Err{self.error_code}" if self.error_code else "Err"
            row_val = "0"
            slot_val = "0"
        else:
            wagon_val = str(self.wagon) if self.wagon is not None else (str(self._final_placement.get("wagon", "")) if self._final_placement else "")
            row_val = str(self.row) if self.row is not None else (str(self._final_placement.get("row", "")) if self._final_placement else "")
            slot_val = str(self.slot) if self.slot is not None else (str(self._final_placement.get("slot", "")) if self._final_placement else "")
        return ",".join([str(self.placed), wagon_val, row_val, slot_val] + values_filtered)

def read_shapes(file_path):
    shapes_input_order = []
    header = None
    header_map = {}
    with open(file_path, "r") as f:
        for line in f:
            if not line.strip():
                continue
            if header is None:
                header = [h.strip() for h in line.strip().split(",")]
                header_map = {name.lower(): idx for idx, name in enumerate(header)}
                continue
            if line.startswith("#"):
                continue
            parts = [p.strip() for p in line.strip().split(",")]
            fields = {name.lower(): parts[idx] if idx < len(parts) else "" for name, idx in header_map.items()}
            shapes_input_order.append(Shape(fields, header))
    shapes_sorted = sorted(shapes_input_order, key=lambda s: (s.vip_key, s.pol, s.id))
    return shapes_input_order, shapes_sorted, header

def assign_parts_to_slots(shapes_sorted, wagons_config):
    placement_dict = {}
    slot_occupancy = defaultdict(lambda: defaultdict(lambda: defaultdict(set)))
    row_vip_map = defaultdict(dict)
    vip_to_shapes = defaultdict(list)
    for shape in shapes_sorted:
        vip_to_shapes[shape.vip_key].append(shape)

    def _shape_fits_in_slot(shape, slot_width, row_height_divided):
        if shape.width <= slot_width and shape.height <= row_height_divided:
            return True, False
        if shape.height <= slot_width and shape.width <= row_height_divided:
            return True, True
        return False, False

    vip_key_counts = {vip: len(vip_to_shapes[vip]) for vip in vip_to_shapes.keys()}
    sorted_vip_keys = sorted(vip_key_counts.keys(), key=lambda k: -vip_key_counts[k])

    max_ops = max(200000, len(shapes_sorted) * 200)
    ops = 0
    start_time = time.time()
    time_limit_seconds = 8.0

    for vip_key in sorted_vip_keys:
        shapes = [s for s in vip_to_shapes[vip_key]]
        part_index = 0

        allowed_nonmixing = [w for w, cfg in sorted(wagons_config.items()) if (vip_key in cfg.get("allowed_vip_keys")) and not cfg.get("allow_vip_key_mixing")]
        allowed_mixing = [w for w, cfg in sorted(wagons_config.items()) if (vip_key in cfg.get("allowed_vip_keys")) and cfg.get("allow_vip_key_mixing")]

        for wagon in allowed_nonmixing:
            if part_index >= len(shapes):
                break
            cfg = wagons_config[wagon]
            stacking_per_slot = cfg.get('stacking_per_slot', 1)
            rows = cfg["rows"]
            slots_per_row = cfg["slots_per_row"]
            row_height = cfg["height"] / rows
            slot_width = cfg["width"] / slots_per_row

            for row in range(1, rows+1):
                current_row_vip = row_vip_map[wagon].get(row)
                if current_row_vip is not None and current_row_vip != vip_key:
                    continue

                for slot in range(1, slots_per_row+1):
                    for stack_pos in reversed(range(stacking_per_slot)):
                        ops += 1
                        if ops > max_ops or (time.time() - start_time) > time_limit_seconds:
                            for i in range(part_index, len(shapes)):
                                s = shapes[i]
                                placement_dict[(s.pol, s.id)] = None
                            break
                        if part_index >= len(shapes):
                            break
                        shape = shapes[part_index]
                        if (shape.pol, shape.id) in placement_dict and placement_dict[(shape.pol, shape.id)] is not None:
                            part_index += 1
                            continue
                        fits, rot = _shape_fits_in_slot(shape, slot_width, row_height / stacking_per_slot)
                        if fits and stack_pos not in slot_occupancy[wagon][row][slot]:
                            if row not in row_vip_map[wagon]:
                                row_vip_map[wagon][row] = vip_key
                            placement_dict[(shape.pol, shape.id)] = {
                                "wagon": wagon,
                                "row": row,
                                "slot": slot,
                                "slot_width": slot_width,
                                "row_height": row_height / stacking_per_slot,
                                "rotated": rot,
                                "stack_pos": stack_pos,
                                "stacking_per_slot": stacking_per_slot
                            }
                            slot_occupancy[wagon][row][slot].add(stack_pos)
                            part_index += 1
                    if part_index >= len(shapes):
                        break
                if part_index >= len(shapes):
                    break

        if part_index < len(shapes):
            for wagon in allowed_mixing:
                if part_index >= len(shapes):
                    break
                cfg = wagons_config[wagon]
                stacking_per_slot = cfg.get('stacking_per_slot', 1)
                rows = cfg["rows"]
                slots_per_row = cfg["slots_per_row"]
                row_height = cfg["height"] / rows
                slot_width = cfg["width"] / slots_per_row

                for row in range(1, rows+1):
                    for slot in range(1, slots_per_row+1):
                        for stack_pos in reversed(range(stacking_per_slot)):
                            ops += 1
                            if ops > max_ops or (time.time() - start_time) > time_limit_seconds:
                                for i in range(part_index, len(shapes)):
                                    s = shapes[i]
                                    placement_dict[(s.pol, s.id)] = None
                                break
                            if part_index >= len(shapes):
                                break
                            shape = shapes[part_index]
                            if (shape.pol, shape.id) in placement_dict and placement_dict[(shape.pol, shape.id)] is not None:
                                part_index += 1
                                continue
                            fits, rot = _shape_fits_in_slot(shape, slot_width, row_height / stacking_per_slot)
                            if fits and stack_pos not in slot_occupancy[wagon][row][slot]:
                                placement_dict[(shape.pol, shape.id)] = {
                                    "wagon": wagon,
                                    "row": row,
                                    "slot": slot,
                                    "slot_width": slot_width,
                                    "row_height": row_height / stacking_per_slot,
                                    "rotated": rot,
                                    "stack_pos": stack_pos,
                                    "stacking_per_slot": stacking_per_slot
                                }
                                slot_occupancy[wagon][row][slot].add(stack_pos)
                                part_index += 1
                        if part_index >= len(shapes):
                            break
                    if part_index >= len(shapes):
                        break

        for i in range(part_index, len(shapes)):
            shape = shapes[i]
            if (shape.pol, shape.id) not in placement_dict:
                placement_dict[(shape.pol, shape.id)] = None

    vip_allowed_wagons = defaultdict(list)
    for w, cfg in wagons_config.items():
        for vip in cfg.get("allowed_vip_keys", []):
            vip_allowed_wagons[vip].append((w, cfg))

    vip_total_counts = {vip: len(vip_to_shapes[vip]) for vip in vip_to_shapes.keys()}

    seen_ids = set()
    for s in shapes_sorted:
        key = (s.pol, s.id)
        if key in seen_ids:
            s.error_unplaced = True
            s.error_code = 7
            s.error_msg = "Duplicate part ID found (conflicts in input)."
        else:
            seen_ids.add(key)

    def inspect_shape_vs_wagon(shape, wagon, cfg, current_slot_occupancy):
        stacking_per_slot = cfg.get('stacking_per_slot', 1)
        rows = cfg["rows"]
        slots_per_row = cfg["slots_per_row"]
        row_height = cfg["height"] / rows
        slot_width = cfg["width"] / slots_per_row

        fits_any_slot = False
        fits_and_has_capacity = False
        total_available_positions = 0

        for row in range(1, rows+1):
            for slot in range(1, slots_per_row+1):
                fits, rot = _shape_fits_in_slot(shape, slot_width, row_height / stacking_per_slot)
                if not fits:
                    continue
                fits_any_slot = True
                occupied = len(current_slot_occupancy[wagon][row][slot])
                free_positions = max(0, stacking_per_slot - occupied)
                total_available_positions += free_positions
                if free_positions > 0:
                    fits_and_has_capacity = True
        return fits_any_slot, fits_and_has_capacity, total_available_positions

    ERRORS = {
        1: "VIP key not accepted by any wagon (no wagon allows this VIP key).",
        2: "Part dimensions do not fit any slot in allowed wagons (even with rotation).",
        3: "No available slot capacity in allowed wagons (slots that fit are all occupied or stack positions used).",
        4: "Not enough rows to allocate this VIP key to a dedicated row in any non-mixing wagon.",
        5: "Total stacking capacity across allowed wagons insufficient for this VIP key.",
        6: "Invalid input data or missing headers for this part.",
        7: "Duplicate part ID found (conflicts in input).",
        8: "Placement algorithm skipped this part (unknown conflict).",
        9: "Other / unspecified error."
    }

    for s in shapes_sorted:
        key = (s.pol, s.id)
        placement = placement_dict.get(key)
        if placement:
            s.error_unplaced = False
            s.error_code = None
            s.error_msg = None
            continue

        if s.error_unplaced and s.error_code in (6,7):
            continue

        allowed = vip_allowed_wagons.get(s.vip_key, [])
        if not allowed:
            s.error_unplaced = True
            s.error_code = 1
            s.error_msg = ERRORS[1]
            continue

        fits_any = False
        fits_capacity = False
        total_capacity_across_allowed = 0
        nonmixing_wagons_checked = 0
        nonmixing_violations = 0

        for (wagon, cfg) in allowed:
            fits, has_capacity, avail_positions = inspect_shape_vs_wagon(s, wagon, cfg, slot_occupancy)
            if fits:
                fits_any = True
            if has_capacity:
                fits_capacity = True
            total_capacity_across_allowed += avail_positions

            if not cfg.get("allow_vip_key_mixing"):
                nonmixing_wagons_checked += 1
                assigned_vips = set()
                for t in shapes_sorted:
                    pl = placement_dict.get((t.pol, t.id))
                    if pl and pl.get("wagon") == wagon:
                        assigned_vips.add(t.vip_key)
                projected_vips = set(assigned_vips)
                projected_vips.add(s.vip_key)
                if len(projected_vips) > cfg["rows"]:
                    nonmixing_violations += 1

        if not fits_any:
            s.error_unplaced = True
            s.error_code = 2
            s.error_msg = ERRORS[2]
            continue

        if not fits_capacity and total_capacity_across_allowed == 0:
            s.error_unplaced = True
            s.error_code = 3
            s.error_msg = ERRORS[3]
            continue

        if nonmixing_wagons_checked > 0 and nonmixing_violations >= nonmixing_wagons_checked:
            s.error_unplaced = True
            s.error_code = 4
            s.error_msg = ERRORS[4]
            continue

        needed_for_vip = vip_total_counts.get(s.vip_key, 1)
        if total_capacity_across_allowed < needed_for_vip:
            s.error_unplaced = True
            s.error_code = 5
            s.error_msg = ERRORS[5] + f" (needed {needed_for_vip}, available {total_capacity_across_allowed})"
            continue

        s.error_unplaced = True
        s.error_code = 8
        s.error_msg = ERRORS[8]

    return placement_dict, slot_occupancy

def write_final_shapes_output(shapes_input_order, output_file, header):
    filtered_header = [h for h in header if h.lower() not in ("placed", "wagon", "row", "slot")]
    new_header = ["Placed", "Wagon", "Row", "Slot"] + filtered_header
    with open(output_file, "w", encoding="utf-8") as fout:
        fout.write(",".join(new_header) + "\n")
        for shape in shapes_input_order:
            fout.write(shape.output_line(filtered_header) + "\n")

if __name__ == "__main__":
    wagons_config = read_shelves(SHELF_FILE)
    shapes_input_order, shapes_sorted, header = read_shapes(INPUT_FILE)
    placement_dict, slot_occupancy = assign_parts_to_slots(shapes_sorted, wagons_config)
    for s in shapes_sorted:
        place = placement_dict.get((s.pol, s.id))
        if place:
            s._final_placement = place
            s.wagon = place.get("wagon", "")
            s.row = place.get("row", "")
            s.slot = place.get("slot", "")
            s.rotated = place.get("rotated", False)
            s.error_unplaced = False
            s.error_code = None
            s.error_msg = None
        else:
            s._final_placement = None
            s.wagon = None
            s.row = None
            s.slot = None
            s.rotated = False
            s.error_unplaced = True
            if not getattr(s, "error_code", None):
                s.error_code = 9
                s.error_msg = "Unplaced: unspecified reason."

    write_final_shapes_output(shapes_input_order, OUTPUT_FILE, header)