import sys
import os
from collections import defaultdict
from tempfile import NamedTemporaryFile
from pylogix import PLC

from PyQt5 import QtWidgets, QtCore, QtGui
import pyqtgraph as pg

import time

PLC_IP = "192.168.100.100"   # Replace with actual PLC IP

# This version polls plc for filename changes - added PLC tag - removed Filename assignment from config file
# Added POL instead of ID on Main display, Added ID,POL,VIP-Key on overview
# Text for slot labels is now horizontal
# font adjustments for better visuals
# removed should_poll_plc variable, QT uses timer
# removed _out file creation and now writes directly to _final file.
# displays current series name. Error handling for missing/invalid file with serie name mismatch
# 10 second buffer for empty filename tag to avoid constant error messages
# Invalid Series name handled properly.
# PLC fix for ID 1, changed to -2.
# Added button to save current wagon layout as images.
# Several adjustments to wagon drawing area according to customer requirements.
# Overhaul to status text message for better user experience. Keeps track of last 4 IDs previously placed.
# Added Force Refresh button to main window.
# Fixed Vip-Key sorting issues with configuration file. Better sorting logic for parts placement.

File_Path = r"C:\FBData\Stacker"

SHELF_FILE = "C:/FBData/Stacker/settings.txt"

SERIE_ID = "Wacon_INT[1]"
SERIE_ID2 = "Wacon_INT[2]"
OPERATING_TRIGGER = "Wacon_INT[20]"

def write_plc_value(tag, value):
    try:
        with PLC() as comm:
            comm.IPAddress = PLC_IP
            comm.Write(tag, value)
    except Exception as e:
        print(f"PLC write to {tag} failed: {e}")

def get_pair_ids_from_plc():
    with PLC() as comm:
        comm.IPAddress = PLC_IP
        response1 = comm.Read(SERIE_ID)
        response2 = comm.Read(SERIE_ID2)
        id1 = response1.Value if response1.Status == 'Success' else None
        id2 = response2.Value if response2.Status == 'Success' else None

        # Only treat -2 and empty string as "no part". 0 is a valid part.
        def normalize_id(val):
            if val is None or val == "" or str(val).strip() == "-2":
                return None
            return str(val)

        id1 = normalize_id(id1)
        id2 = normalize_id(id2)
        return id1, id2

def get_filename_from_plc():
    with PLC() as comm:
        comm.IPAddress = PLC_IP
        response = comm.Read("Wacon_SerieName")  # Get filename from PLC tag now implemented
        if response.Status == 'Success' and response.Value:
            return str(response.Value)
        else:
            return None

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
        for line in lines[start_idx:]:
            parts = [s.strip() for s in line.strip().split(",")]
            if len(parts) < 6:
                raise ValueError("Each line: wagon_number,rows,slots_per_row,total_width,total_height,allowed_vip_keys,[allow_vip_key_mixing],[stacking_per_slot]")
            wagon = int(parts[0])
            rows = int(parts[1])
            slots_per_row = int(parts[2])
            width = float(parts[3])
            height = float(parts[4])
            allowed_vip_keys = [m.strip().lower() for m in parts[5:-2] if m.strip()]
            allow_vip_key_mixing = (parts[-2].strip().lower() == "yes") if len(parts) > 6 else False
            stacking_per_slot = int(parts[-1]) if len(parts) > 7 and parts[-1].isdigit() else 1
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
        self.placed = int(self.fields.get("placed", "0"))
        self.slot = None
        self.row = None
        self.wagon = None
        self.rotated = False
        self._final_placement = None

    def output_line(self, filtered_header):
        values_filtered = [self.fields.get(h.lower(), "") for h in filtered_header]
        wagon_val = str(self.wagon) if self.wagon is not None else (str(self._final_placement.get("wagon", "")) if self._final_placement else "")
        row_val = str(self.row) if self.row is not None else (str(self._final_placement.get("row", "")) if self._final_placement else "")
        slot_val = str(self.slot) if self.slot is not None else (str(self._final_placement.get("slot", "")) if self._final_placement else "")
        return ",".join([str(self.placed), wagon_val, row_val, slot_val] + values_filtered)

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

    return placement_dict, slot_occupancy

def _atomic_write_text(file_path: str, text: str):
    directory = os.path.dirname(file_path) or "."
    with NamedTemporaryFile("w", delete=False, dir=directory, encoding="utf-8", newline="") as tmp:
        tmp.write(text)
        tmp.flush()
        os.fsync(tmp.fileno())
        temp_name = tmp.name
    # os.replace is atomic on Windows when replacing within the same volume
    os.replace(temp_name, file_path)

def write_shapes_output(shapes_input_order, output_file, header):
    filtered_header = [h for h in header if h.lower() not in ("placed", "wagon", "row", "slot")]
    new_header = ["Placed", "Wagon", "Row", "Slot"] + filtered_header

    lines = []
    lines.append(",".join(new_header))
    for shape in shapes_input_order:
        lines.append(shape.output_line(filtered_header))

    csv_text = "\n".join(lines) + "\n"
    _atomic_write_text(output_file, csv_text)

class WagonWidget(QtWidgets.QWidget):
    def __init__(self, wagon_number, cfg, shapes_by_id, slot_occupancy, highlight_ids=None, short_row_labels=False, label_mode="pol", series_name=None):
        super().__init__()
        self.label_mode = label_mode
        self.wagon_number = wagon_number
        self.cfg = cfg
        self.shapes_by_id = shapes_by_id
        self.slot_occupancy = slot_occupancy
        self.highlight_ids = set(highlight_ids) if highlight_ids is not None else set()
        self.short_row_labels = short_row_labels
        self.series_name = series_name

        title_text = f"Wagon {wagon_number}"
        if series_name:
            try:
                display = os.path.basename(series_name)
            except Exception:
                display = series_name
            if display.endswith("_final.txt"):
                display = display[:-10]
            elif display.endswith(".txt"):
                display = display[:-4]
            title_text = f"{title_text} - {display}"

        self.setWindowTitle(title_text)
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(0,0,0,0)
        self.wagon_label = QtWidgets.QLabel(title_text)
        self.wagon_label.setAlignment(QtCore.Qt.AlignCenter)
        font = QtGui.QFont("Arial", 18, QtGui.QFont.Bold)
        self.wagon_label.setFont(font)
        self.wagon_label.setStyleSheet(
            "background-color: #2ecc40; color: white; border-radius: 6px; padding: 10px;"
        )
        layout.addWidget(self.wagon_label)

        self.plot_widget = pg.GraphicsLayoutWidget()
        self.plot_widget.setBackground('w')
        self.plot_widget.setMinimumWidth(770)
        self.plot_widget.setMinimumHeight(880)
        self.plot_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        layout.addWidget(self.plot_widget)

        self.view_box = self.plot_widget.addViewBox()
        self.view_box.setAspectLocked(False)
        self.view_box.setContentsMargins(0,0,0,0)
        self.view_box.invertY(True)
        self.view_box.setBackgroundColor(QtGui.QColor('white'))
        self.view_box.setMouseEnabled(x=False, y=False)
        self.slot_items = {}
        self.label_items = []
        self.initialized = False
        self.last_highlight_ids = set()
        if self.cfg is not None:
            self.setup_graphics_items()
            self.update_graphics(self.shapes_by_id, self.highlight_ids)

        self.setAutoFillBackground(True)
        pal = self.palette()
        pal.setColor(self.backgroundRole(), QtCore.Qt.white)
        self.setPalette(pal)

    def set_series_name(self, series_name):
        self.series_name = series_name
        title_text = f"Wagon {self.wagon_number}"
        if series_name:
            try:
                display = os.path.basename(series_name)
            except Exception:
                display = series_name
            if display.endswith("_final.txt"):
                display = display[:-10]
            elif display.endswith(".txt"):
                display = display[:-4]
            title_text = f"{title_text} - {display}"
        try:
            self.setWindowTitle(title_text)
        except Exception:
            pass
        try:
            self.wagon_label.setText(title_text)
        except Exception:
            pass

    def setup_graphics_items(self):
        self.view_box.clear()
        self.slot_items = {}
        self.label_items = []
        rows = self.cfg["rows"]
        slots_per_row = self.cfg["slots_per_row"]
        wagon_width = self.cfg["width"] * 1.3       # Scale factor for better visibility
        wagon_height = self.cfg["height"]
        stacking_per_slot = self.cfg.get("stacking_per_slot", 1)
        row_height = wagon_height / rows
        slot_width = wagon_width / slots_per_row

        outer_rect = QtCore.QRectF(0, 0, wagon_width, wagon_height)
        border = QtGui.QPen(QtCore.Qt.black, 3)
        outer_item = QtWidgets.QGraphicsRectItem(outer_rect)
        outer_item.setPen(border)
        outer_item.setBrush(pg.mkBrush(None))
        self.view_box.addItem(outer_item)
        self.label_items.append(outer_item)

        margin = 5
        self.view_box.setRange(
            xRange=[-margin, self.cfg["width"] * 1.25 + margin],
            yRange=[0, wagon_height]
        )

        label_font = QtGui.QFont()
        label_font.setPointSize(12)
        label_font.setBold(True)

        for slot in range(1, slots_per_row + 1):
            x = (slot - 1) * slot_width + slot_width / 2
            y = -5
            slot_label = pg.TextItem(str(slot), anchor=(0.5, 1), color='black', fill=None)
            slot_label.setFont(label_font)
            slot_label.setPos(x, y)
            self.view_box.addItem(slot_label)
            self.label_items.append(slot_label)

        for row in range(1, rows+1):
            y_base = (row - 1) * row_height
            row_label_y = y_base + row_height / 2
            row_label_x = -6
            row_text = f"{row}" if self.short_row_labels else f"R{row}"
            row_label = pg.TextItem(row_text, anchor=(1, 0.5), color='black', fill=None)
            row_label.setFont(label_font)
            row_label.setPos(row_label_x, row_label_y)
            self.view_box.addItem(row_label)
            self.label_items.append(row_label)

            for slot in range(1, slots_per_row+1):
                x = (slot - 1) * slot_width
                for stack_pos in reversed(range(stacking_per_slot)):
                    slot_height = row_height / stacking_per_slot
                    y = y_base + slot_height * stack_pos
                    rect_item = QtWidgets.QGraphicsRectItem(x, y, slot_width, slot_height)
                    rect_item.setBrush(pg.mkBrush((255, 255, 255, 180)))
                    rect_item.setPen(QtGui.QPen(QtGui.QColor('gray')))
                    self.view_box.addItem(rect_item)

                    base_font_size = int(slot_width * 0.18)
                    if stacking_per_slot == 2:
                        font_size = max(int(base_font_size*0.8), 9)
                    else:
                        font_size = max(int(base_font_size*1.2), 8)
                    slot_font = QtGui.QFont()
                    slot_font.setPointSize(font_size)
                    slot_font.setBold(True)

                    text_item = QtWidgets.QGraphicsTextItem("")
                    text_item.setTextWidth(slot_width - 6)
                    text_item.setDefaultTextColor(QtGui.QColor('black'))
                    text_item.setFont(slot_font)
                    text_rect = text_item.boundingRect()
                    text_item.setPos(x + (slot_width - text_item.textWidth()) / 2, y + (slot_height - text_rect.height()) / 2 -15)
                    self.view_box.addItem(text_item)
                    self.slot_items[(row, slot, stack_pos)] = (rect_item, text_item)

        for row in range(1, rows):
            y = round(row * row_height, 2)
            line = QtWidgets.QGraphicsLineItem(0, y, wagon_width, y)
            line.setPen(QtGui.QPen(QtCore.Qt.black, 2))
            line.setZValue(2)
            self.view_box.addItem(line)
            self.label_items.append(line)
        for slot in range(1, slots_per_row):
            x = round(slot * slot_width, 2)
            line = QtWidgets.QGraphicsLineItem(x, 0, x, wagon_height)
            line.setPen(QtGui.QPen(QtGui.QColor('black'), 2))
            line.setZValue(2)
            self.view_box.addItem(line)
            self.label_items.append(line)
        if stacking_per_slot > 1:
            for row in range(1, rows+1):
                y_base = (row - 1) * row_height
                for stack_div in range(1, stacking_per_slot):
                    y = y_base + row_height * stack_div / stacking_per_slot
                    line = QtWidgets.QGraphicsLineItem(0, y, wagon_width, y)
                    pen = QtGui.QPen(QtGui.QColor('lightgrey'))
                    pen.setStyle(QtCore.Qt.DashLine)
                    line.setPen(pen)
                    self.view_box.addItem(line)
                    self.label_items.append(line)
        self.initialized = True

    def update_graphics(self, shapes_by_id, highlight_ids):
        planned_slots = {}
        placed_slots = set()
        for s in shapes_by_id.values():
            if s._final_placement and s._final_placement.get("wagon") == self.wagon_number:
                row = s._final_placement["row"]
                slot = s._final_placement["slot"]
                stack_pos = s._final_placement.get("stack_pos", 0)
                planned_slots[(row, slot, stack_pos)] = s.id
                if s.placed == 1:
                    placed_slots.add((row, slot, stack_pos))
        for key, (rect, text) in self.slot_items.items():
            row, slot, stack_pos = key
            part_id = planned_slots.get(key, None)
            is_highlight = False
            for sid in highlight_ids:
                shape = shapes_by_id.get(sid)
                if shape and shape._final_placement and shape._final_placement.get("wagon") == self.wagon_number \
                   and shape._final_placement.get("row") == row and shape._final_placement.get("slot") == slot \
                   and shape._final_placement.get("stack_pos", 0) == stack_pos:
                    is_highlight = True
                    break
            if is_highlight:
                color = "yellow"
            elif key in placed_slots:
                color = "limegreen"
            else:
                color = "white"
            rect.setBrush(pg.mkBrush(color))
            label_str = ""
            if part_id and part_id in shapes_by_id:
                shape = shapes_by_id[part_id]
                if self.label_mode == "id":
                    id_or_pol_str = "  " + str(shape.pol) + "\n"
                else:
                    id_or_pol_str = "  " + str(shape.pol) + "\n"
                vip_str = shape.vip_key_original
                stacking_per_slot = self.cfg.get("stacking_per_slot", 1) if self.cfg else 1
                if stacking_per_slot == 2:
                    if self.label_mode == "id":
                        id_or_pol_str = " P " + str(shape.pol)
                    else:
                        id_or_pol_str = "  " + str(shape.pol)
                    label_str = f"{id_or_pol_str}\n{vip_str}"
                else:
                    label_str = f"{id_or_pol_str}{vip_str}"
            text.setPlainText(label_str)

    def draw_wagon(self, highlight_ids=None):
        if self.cfg is None:
            return
        if not self.initialized:
            self.setup_graphics_items()
        if highlight_ids is not None:
            self.highlight_ids = set(highlight_ids)
        self.update_graphics(self.shapes_by_id, self.highlight_ids)

class PartParameterTable(QtWidgets.QTableWidget):
    def __init__(self, fields, parent=None):
        super().__init__(len(fields), 2, parent)
        self.setHorizontalHeaderLabels(["Part 1", "Part 2"])
        self.setVerticalHeaderLabels(fields)
        header_font = QtGui.QFont("Arial", 10, QtGui.QFont.Bold)
        self.horizontalHeader().setFont(header_font)
        self.verticalHeader().setFont(header_font)
        self.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.verticalHeader().setSectionResizeMode(QtWidgets.QHeaderView.Stretch)
        self.setShowGrid(True)
        self.setStyleSheet("""
            QTableWidget {
                gridline-color: #444;
                font-size: 8pt;
                border: none;
            }
            QTableView::item {
                border: 2px solid #555;
                padding: 4px;
            }
            QHeaderView::section {
                background-color: #f7f7f7;
                font-weight: bold;
                border: 2px solid #555;
                padding: 6px;
            }
        """)
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.setFocusPolicy(QtCore.Qt.NoFocus)

    def set_part_data(self, part1: dict, part2: dict, fields: list):
        for row, field in enumerate(fields):
            val1 = part1.get(field, "")
            val2 = part2.get(field, "")
            item1 = QtWidgets.QTableWidgetItem(str(val1))
            item2 = QtWidgets.QTableWidgetItem(str(val2))
            item1.setTextAlignment(QtCore.Qt.AlignCenter)
            item2.setTextAlignment(QtCore.Qt.AlignCenter)
            bold_font = QtGui.QFont()
            bold_font.setBold(True)
            item1.setFont(bold_font)
            item2.setFont(bold_font)
            self.setItem(row, 0, item1)
            self.setItem(row, 1, item2)

class MainPackingWindow(QtWidgets.QWidget):
    def __init__(self, shapes_by_id, wagons_config, slot_occupancy, overview_window=None):
        super().__init__()
        self.setMinimumSize(1920, 1080)  # For 1920x1080 screen
        self.resize(1920, 1080)
        self.shapes_by_id = shapes_by_id
        self.wagons_config = wagons_config
        self.slot_occupancy = slot_occupancy
        self.overview_window = overview_window
        self.setWindowTitle("Wagon Packing")
        self.layout = QtWidgets.QHBoxLayout(self)
        self.layout.setContentsMargins(10,10,10,10)

        # 88% for wagon area, 12% for info panel
        self.figure_widget = QtWidgets.QWidget()
        self.figure_layout = QtWidgets.QVBoxLayout(self.figure_widget)
        self.figure_layout.setContentsMargins(0,0,0,0)
        self.layout.addWidget(self.figure_widget, stretch=88)

        self.info_panel = QtWidgets.QWidget()
        self.info_panel.setAutoFillBackground(True)
        pal = self.info_panel.palette()
        pal.setColor(self.info_panel.backgroundRole(), QtCore.Qt.white)
        self.info_panel.setPalette(pal)
        self.info_layout = QtWidgets.QVBoxLayout(self.info_panel)
        self.info_layout.setContentsMargins(0,0,0,0)
        self.layout.addWidget(self.info_panel, stretch=12)

        self.figure_widget.setMinimumWidth(1700)
        self.info_panel.setMinimumWidth(200)

        self.current_wagons_label = QtWidgets.QLabel("")
        big_font = QtGui.QFont("Arial", 18, QtGui.QFont.Bold)
        self.current_wagons_label.setFont(big_font)
        self.current_wagons_label.setAlignment(QtCore.Qt.AlignCenter)
        self.current_wagons_label.setStyleSheet(
            "background-color: #2ecc40; color: white; border-radius: 6px; padding: 10px;")
        self.info_layout.addWidget(self.current_wagons_label)
        self.info_layout.addStretch(1)
        self.table_fields = ["ID", "POL", "VIP-Key", "Wagon #", "Row #", "Place #", "Length", "Rotated"]
        self.param_table = PartParameterTable(self.table_fields)
        self.param_table.setFixedHeight(480)
        self.info_layout.addWidget(self.param_table)
        self.info_layout.addStretch(1)
        self.status_text = QtWidgets.QLabel("Status")
        self.status_text.setWordWrap(True)
        self.status_text.setFont(QtGui.QFont("Arial", 14))
        self.status_text.setStyleSheet("""
            background-color: yellow;
            border: 2px solid gray;
            border-radius: 5px;
            padding: 6px;
        """)
        self.info_layout.addWidget(self.status_text)
        self.info_layout.addStretch(1)
        btn_container = QtWidgets.QWidget()
        btn_layout = QtWidgets.QVBoxLayout(btn_container)
        btn_layout.setContentsMargins(0,0,0,0)
        btn_layout.addStretch(1)

        button_font = QtGui.QFont("Arial", 12, QtGui.QFont.Bold)
        button_width = 200

        self.save_all_img_btn = QtWidgets.QPushButton("Save Layout as Images")
        self.save_all_img_btn.setFont(button_font)
        self.save_all_img_btn.setMinimumWidth(button_width)
        self.save_all_img_btn.setMaximumWidth(button_width)
        self.save_all_img_btn.setMinimumHeight(50)
        self.save_all_img_btn.clicked.connect(self.save_all_overview_pages_as_images)
        btn_layout.addWidget(self.save_all_img_btn, alignment=QtCore.Qt.AlignHCenter)
        btn_layout.addSpacing(15)

        self.force_refresh_btn = QtWidgets.QPushButton("Force Refresh")
        self.force_refresh_btn.setFont(button_font)
        self.force_refresh_btn.setMinimumWidth(button_width)
        self.force_refresh_btn.setMaximumWidth(button_width)
        self.force_refresh_btn.setMinimumHeight(40)
        btn_layout.addWidget(self.force_refresh_btn, alignment=QtCore.Qt.AlignHCenter)
        btn_layout.addSpacing(8)

        self.close_btn = QtWidgets.QPushButton("Close")
        self.close_btn.setFont(button_font)
        self.close_btn.setMinimumWidth(button_width)
        self.close_btn.setMaximumWidth(button_width)
        self.close_btn.setMinimumHeight(40)
        btn_layout.addWidget(self.close_btn, alignment=QtCore.Qt.AlignHCenter)
        btn_layout.addStretch(1)

        self.info_layout.addWidget(btn_container)
        self.info_layout.addStretch(1)

        self.wagon_area = QtWidgets.QWidget()
        self.wagon_layout = QtWidgets.QHBoxLayout(self.wagon_area)
        self.wagon_layout.setContentsMargins(0,0,0,0)
        self.figure_layout.addWidget(self.wagon_area)
        self.left_wagon_widget = None
        self.right_wagon_widget = None
        self.last_wagon_ids = None
        self.last_highlight_ids = None
        self.last_series_name = None
        self.setAutoFillBackground(True)
        pal_main = self.palette()
        pal_main.setColor(self.backgroundRole(), QtCore.Qt.white)
        self.setPalette(pal_main)
        self.figure_widget.setAutoFillBackground(True)
        pal2 = self.figure_widget.palette()
        pal2.setColor(self.figure_widget.backgroundRole(), QtCore.Qt.white)
        self.figure_widget.setPalette(pal2)
        self.wagon_area.setAutoFillBackground(True)
        pal3 = self.wagon_area.palette()
        pal3.setColor(self.wagon_area.backgroundRole(), QtCore.Qt.white)
        self.wagon_area.setPalette(pal3)

    def set_wagons(self, wagon_ids, shapes_by_id, slot_occupancy, highlight_ids, series_name=None):
        if self.last_wagon_ids == wagon_ids and self.last_highlight_ids == highlight_ids and self.last_series_name == series_name:
            self.update_current_wagons_label(wagon_ids)
            return
        self.last_wagon_ids = list(wagon_ids) if wagon_ids else []
        self.last_highlight_ids = list(highlight_ids) if highlight_ids else []
        self.last_series_name = series_name

        for i in reversed(range(self.wagon_layout.count())):
            widget = self.wagon_layout.itemAt(i).widget()
            if widget is not None:
                widget.setParent(None)

        wagon_area_width = self.figure_widget.width()
        if wagon_area_width < 100:
            wagon_area_width = 800

        wagon_width = int(wagon_area_width // 2)

        if len(wagon_ids) == 2 and wagon_ids[0] != wagon_ids[1]:
            self.left_wagon_widget = WagonWidget(
                wagon_ids[0], self.wagons_config[wagon_ids[0]], shapes_by_id, slot_occupancy, highlight_ids, label_mode="pol", series_name=series_name
            )
            self.right_wagon_widget = WagonWidget(
                wagon_ids[1], self.wagons_config[wagon_ids[1]], shapes_by_id, slot_occupancy, highlight_ids, label_mode="pol", series_name=series_name
            )
            self.left_wagon_widget.setFixedWidth(wagon_width)
            self.right_wagon_widget.setFixedWidth(wagon_width)
            self.left_wagon_widget.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Expanding)
            self.right_wagon_widget.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Expanding)
            self.wagon_layout.addWidget(self.left_wagon_widget)
            self.wagon_layout.addWidget(self.right_wagon_widget)
        elif len(wagon_ids) == 1 or (len(wagon_ids) == 2 and wagon_ids[0] == wagon_ids[1]):
            left_spacer = QtWidgets.QWidget()
            left_spacer.setFixedWidth(wagon_width)
            right_spacer = QtWidgets.QWidget()
            right_spacer.setFixedWidth(wagon_width)
            wagon_widget = WagonWidget(
                wagon_ids[0], self.wagons_config[wagon_ids[0]], shapes_by_id, slot_occupancy, highlight_ids, label_mode="pol", series_name=series_name
            )
            wagon_widget.setFixedWidth(wagon_width)
            wagon_widget.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Expanding)
            self.wagon_layout.addWidget(left_spacer)
            self.wagon_layout.addWidget(wagon_widget)
            self.wagon_layout.addWidget(right_spacer)
            self.left_wagon_widget = wagon_widget
            self.right_wagon_widget = None
        else:
            first_wagon = next(iter(self.wagons_config), None)
            if first_wagon is not None:
                left_spacer = QtWidgets.QWidget()
                left_spacer.setFixedWidth(wagon_width)
                right_spacer = QtWidgets.QWidget()
                right_spacer.setFixedWidth(wagon_width)
                wagon_widget = WagonWidget(
                    first_wagon, self.wagons_config[first_wagon], shapes_by_id, slot_occupancy, [], label_mode="pol", series_name=series_name
                )
                wagon_widget.setFixedWidth(wagon_width)
                wagon_widget.setSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Expanding)
                self.wagon_layout.addWidget(left_spacer)
                self.wagon_layout.addWidget(wagon_widget)
                self.wagon_layout.addWidget(right_spacer)
                self.left_wagon_widget = wagon_widget
            self.right_wagon_widget = None

        self.update_current_wagons_label(wagon_ids)

    def set_info(self, shape1: 'Shape', shape2: 'Shape'):
        def shape_to_dict(shape):
            if not shape:
                return {}
            return {
                "ID": shape.id,
                "POL": shape.pol,
                "VIP-Key": getattr(shape, "vip_key_original", ""),
                "Wagon #": getattr(shape, "wagon", ""),
                "Row #": getattr(shape, "row", ""),
                "Place #": getattr(shape, "slot", ""),
                "Length": getattr(shape, "length", ""),
                "Rotated": "Yes" if getattr(shape, "rotated", False) else "No"
            }
        part1 = shape_to_dict(shape1)
        part2 = shape_to_dict(shape2)
        self.param_table.set_part_data(part1, part2, self.table_fields)

    def set_status(self, status_msg):
        self.status_text.setText(status_msg)

    def update_current_wagons_label(self, wagon_ids):
        if wagon_ids:
            self.current_wagons_label.setText(
                "Wagons: " + ", ".join(str(w) for w in wagon_ids)
            )
        else:
            self.current_wagons_label.setText("No wagons")

    def save_all_overview_pages_as_images(self):

        BASE_IMAGE_PATH = r"C:\FBData\DoneSeries_Images"

        overview = self.overview_window
        if overview is None:
            return

        series_name = None
        parent = self.parent()
        if hasattr(parent, "current_filename"):
            series_name = parent.current_filename
        if not series_name:
            series_name = get_filename_from_plc()
        if not series_name:
            series_name = "UnknownSeries"

        if series_name.endswith("_final.txt"):
            series_name = series_name[:-10]
        elif series_name.endswith(".txt"):
            series_name = series_name[:-4]

        series_folder = os.path.join(BASE_IMAGE_PATH, series_name)
        os.makedirs(series_folder, exist_ok=True)

        for filename in os.listdir(series_folder):
            if filename.endswith(".png"):
                try:
                    os.remove(os.path.join(series_folder, filename))
                except Exception as e:
                    print(f"Could not delete {filename}: {e}")

        used_wagon_numbers = set()
        for shape in self.shapes_by_id.values():
            if shape._final_placement and shape._final_placement.get("wagon", None) is not None:
                used_wagon_numbers.add(shape._final_placement["wagon"])

        for wagon_number in used_wagon_numbers:
            try:
                wagon_idx = overview.all_wagon_numbers.index(wagon_number)
            except ValueError:
                continue

            page_size = overview.page_size
            page = wagon_idx // page_size
            overview.current_page = page
            overview.update_page()
            QtWidgets.QApplication.processEvents()

            page_wagons = overview.all_wagon_numbers[page * page_size : (page + 1) * page_size]
            try:
                pos_in_page = page_wagons.index(wagon_number)
            except ValueError:
                continue

            wagon_widget = overview.wagon_widgets[pos_in_page]
            pixmap = wagon_widget.grab()

            filename = f"{series_name}_wagon_{wagon_number}.png"
            file_path = os.path.join(series_folder, filename)
            pixmap.save(file_path, "PNG")

        overview.current_page = 0
        overview.update_page()

class OverviewWindow(QtWidgets.QWidget):
    def __init__(self, wagons_config, shapes_by_id, slot_occupancy, rows=1, columns=4, series_name=None):
        super().__init__()
        self.setWindowTitle("Wagons Overview")
        self.wagons_config = wagons_config
        self.shapes_by_id = shapes_by_id
        self.slot_occupancy = slot_occupancy
        self.rows = rows
        self.columns = columns
        self.page_size = self.rows * self.columns
        self.all_wagon_numbers = sorted(list(self.wagons_config.keys()))
        self.current_page = 0
        self.series_name = series_name

        # Main vertical layout
        self.main_layout = QtWidgets.QVBoxLayout(self)
        self.main_layout.setContentsMargins(0, 0, 0, 0)
        self.main_layout.setSpacing(0)

        # Grid widget and layout
        self.grid_widget = QtWidgets.QWidget()
        self.grid_layout = QtWidgets.QGridLayout(self.grid_widget)
        self.grid_layout.setContentsMargins(8, 8, 8, 8)
        self.grid_layout.setSpacing(8)

        self.main_layout.addWidget(self.grid_widget, stretch=1)

        # Navigation buttons
        btn_layout = QtWidgets.QHBoxLayout()
        self.prev_btn = QtWidgets.QPushButton("Previous")
        self.next_btn = QtWidgets.QPushButton("Next")
        btn_layout.addWidget(self.prev_btn)
        btn_layout.addStretch(1)
        btn_layout.addWidget(self.next_btn)
        self.main_layout.addLayout(btn_layout, stretch=0)
        self.prev_btn.clicked.connect(self.prev_page)
        self.next_btn.clicked.connect(self.next_page)

        self.wagon_widgets = []
        start_idx = self.current_page * self.page_size
        end_idx = start_idx + self.page_size
        page_wagons = self.all_wagon_numbers[start_idx:end_idx]
        for i in range(self.page_size):
            if i < len(page_wagons):
                wagon_number = page_wagons[i]
                cfg = self.wagons_config[wagon_number]
            else:
                wagon_number = None
                cfg = None
            w_widget = WagonWidget(
                wagon_number,
                cfg,
                self.shapes_by_id,
                self.slot_occupancy,
                highlight_ids=[],
                short_row_labels=True,
                label_mode="id",
                series_name=self.series_name
            )
            w_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
            row = i // self.columns if self.rows > 1 else 0
            col = i % self.columns
            self.grid_layout.addWidget(w_widget, row, col)
            self.wagon_widgets.append(w_widget)

        for r in range(self.rows):
            self.grid_layout.setRowStretch(r, 1)
        for c in range(self.columns):
            self.grid_layout.setColumnStretch(c, 1)

        self.last_page_wagons = []
        self.last_highlight_ids = None

    def update_page(self, highlight_ids=None):
        start_idx = self.current_page * self.page_size
        end_idx = start_idx + self.page_size
        page_wagons = self.all_wagon_numbers[start_idx:end_idx]
        # Only update if page or highlight_ids have changed
        if self.last_page_wagons == page_wagons and self.last_highlight_ids == highlight_ids:
            return
        self.last_page_wagons = list(page_wagons)
        self.last_highlight_ids = list(highlight_ids) if highlight_ids else []
        for i, w_widget in enumerate(self.wagon_widgets):
            if i < len(page_wagons):
                wagon_number = page_wagons[i]
                cfg = self.wagons_config.get(wagon_number)
                if cfg is None:
                    w_widget.hide()
                    continue
                if w_widget.wagon_number != wagon_number or w_widget.cfg != cfg:
                    w_widget.setParent(None)
                    new_widget = WagonWidget(
                        wagon_number,
                        cfg,
                        self.shapes_by_id,
                        self.slot_occupancy,
                        highlight_ids=highlight_ids if highlight_ids else [],
                        short_row_labels=True,
                        label_mode="id",
                        series_name=self.series_name
                    )
                    new_widget.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
                    row = i // self.columns if self.rows > 1 else 0
                    col = i % self.columns
                    self.grid_layout.addWidget(new_widget, row, col)
                    self.wagon_widgets[i] = new_widget
                    w_widget = new_widget
                else:
                    w_widget.set_series_name(self.series_name)
                    w_widget.shapes_by_id = self.shapes_by_id
                    w_widget.slot_occupancy = self.slot_occupancy
                    w_widget.draw_wagon(highlight_ids=highlight_ids)
                w_widget.show()
            else:
                w_widget.hide()
        self.prev_btn.setEnabled(self.current_page > 0)
        self.next_btn.setEnabled(end_idx < len(self.all_wagon_numbers))

    def next_page(self):
        if (self.current_page + 1) * self.page_size < len(self.all_wagon_numbers):
            self.current_page += 1
            self.update_page()

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.update_page()

class StackerApp(QtWidgets.QApplication):
    def __init__(
        self, shapes_input_order, shapes_sorted, header, shapes_by_id,
        wagons_config, slot_occupancy, output_file, plc_filename
    ):
        super().__init__(sys.argv)
        self.shapes_input_order = shapes_input_order
        self.shapes_sorted = shapes_sorted
        self.header = header
        self.shapes_by_id = shapes_by_id
        self.wagons_config = wagons_config
        self.slot_occupancy = slot_occupancy
        self.output_file = output_file
        self.current_filename = plc_filename
        self.empty_filename_start_time = None
        self.empty_filename_timeout = 10

        def _normalize_series_name(filename):
            if not filename:
                return ""
            base = os.path.basename(str(filename))
            if base.endswith("_final.txt"):
                base = base[:-10]
            elif base.endswith(".txt"):
                base = base[:-4]
            return base
        self._normalize_series_name = _normalize_series_name

        self.overview_window = OverviewWindow(
            self.wagons_config, self.shapes_by_id, self.slot_occupancy, rows=1, columns=2, series_name=self._normalize_series_name(self.current_filename)
        )
        self.overview_window.resize(1300, 900)
        self.overview_window.showMaximized()

        self.main_window = MainPackingWindow(
            self.shapes_by_id, self.wagons_config, self.slot_occupancy, overview_window=self.overview_window
        )
        self.main_window.close_btn.clicked.connect(self.do_close)
        
        try:
            self.main_window.force_refresh_btn.clicked.connect(self.do_force_refresh)
        except Exception:
            pass
        self.main_window.showMaximized()

        self.already_placed_ids = set(s.id for s in shapes_sorted if s.placed == 1)

        self.placed_history = [s.id for s in shapes_sorted if s.placed == 1]
        self.last_placed_parts = []

        self.last_polled_ids = None
        self.last_highlight_ids = (None, None)
        self.last_drawn_ids = (None, None)

        # Poll PLC at interval
        self.timer = QtCore.QTimer()
        self.timer.timeout.connect(self.poll_plc_and_save)
        self.timer.start(500)
        # For testing without PLC polling. (No longer in use, Demo_script is used for testing)
        # PLC_POLLING = False
        # self.timer = QtCore.QTimer()
        # self.timer.timeout.connect(self.poll_plc_and_save)
        # if PLC_POLLING:
        #     self.timer.start(500)
        # else:
        #     # Ensure timer is stopped in case it was running
        #     try:
        #         self.timer.stop()
        #     except Exception:
        #         pass
        self.update_main_window(self.last_highlight_ids)

        self.last_saved_state = set((s.id, s.placed) for s in self.shapes_input_order)

    def get_file_paths(self, plc_filename):
        input_file = f"{File_Path}\\{plc_filename}_final.txt"
        output_file = input_file
        settings_file = SHELF_FILE  # Always use latest settings file updated by CX-Supervisor
        return input_file, output_file, settings_file

    def force_refresh_views(self, highlight_ids=(None, None)):
        self.last_drawn_ids = (None, None)
        self.previous_drawn_ids = (None, None)
        self.last_highlight_ids = highlight_ids
        self.update_main_window(highlight_ids)

        if hasattr(self.overview_window, 'last_page_wagons'):
            self.overview_window.last_page_wagons = []
        self.overview_window.last_highlight_ids = None
        self.overview_window.current_page = 0
        self.overview_window.all_wagon_numbers = sorted(list(self.wagons_config.keys()))
        self.overview_window.wagons_config = self.wagons_config
        self.overview_window.shapes_by_id = self.shapes_by_id
        self.overview_window.slot_occupancy = self.slot_occupancy
        self.overview_window.series_name = self._normalize_series_name(self.current_filename)
        self.overview_window.update_page(highlight_ids=highlight_ids)
        self.main_window.set_status(f"Loaded new series: {self.current_filename}")

    def reload_data_if_filename_changed(self, new_filename):
        input_file, output_file, settings_file = self.get_file_paths(new_filename)
        if not os.path.exists(input_file):
            self.current_filename = new_filename
            self.shapes_input_order = []
            self.shapes_sorted = []
            self.header = []
            self.shapes_by_id = {}
            self.wagons_config = {}
            self.slot_occupancy = {}
            self.output_file = output_file

            status_message = (
                f"Series '{new_filename}' not found. Waiting for valid series.\n"
                f"Current Series: {self.current_filename}"
            )
            self.main_window.set_status(status_message)
            if hasattr(self.main_window, "set_wagons"):
                self.main_window.set_wagons([], {}, {}, [], series_name=self._normalize_series_name(self.current_filename))
            if hasattr(self.overview_window, "update_page"):
                self.overview_window.wagons_config = {}
                self.overview_window.shapes_by_id = {}
                self.overview_window.slot_occupancy = {}
                self.overview_window.series_name = self._normalize_series_name(self.current_filename)
                self.overview_window.update_page(highlight_ids=[])
            self.already_placed_ids = set()
            return

        shapes_input_order, shapes_sorted, header = read_shapes(input_file)
        shapes_by_id = {s.id: s for s in shapes_sorted}
        wagons_config = read_shelves(settings_file)
        placement_dict, slot_occupancy = assign_parts_to_slots(shapes_sorted, wagons_config)
        for s in shapes_sorted:
            place = placement_dict.get((s.pol, s.id))
            if place:
                s._final_placement = place
                s.wagon = place.get("wagon", "")
                s.row = place.get("row", "")
                s.slot = place.get("slot", "")
                s.rotated = place.get("rotated", False)
            else:
                s._final_placement = None
                s.wagon = None
                s.row = None
                s.slot = None
                s.rotated = False
        self.shapes_input_order = shapes_input_order
        self.shapes_sorted = shapes_sorted
        self.header = header
        self.shapes_by_id = shapes_by_id
        self.wagons_config = wagons_config
        self.slot_occupancy = slot_occupancy
        self.output_file = output_file
        self.current_filename = new_filename
        self.main_window.shapes_by_id = shapes_by_id
        self.main_window.wagons_config = wagons_config
        self.main_window.slot_occupancy = slot_occupancy
        self.overview_window.shapes_by_id = shapes_by_id
        self.overview_window.wagons_config = wagons_config
        self.overview_window.slot_occupancy = slot_occupancy
        self.overview_window.series_name = self._normalize_series_name(new_filename)
        self.already_placed_ids = set(s.id for s in shapes_sorted if s.placed == 1)
        self.force_refresh_views(highlight_ids=(None, None))

    def poll_plc_and_save(self):

        plc_filename = get_filename_from_plc()
        now = time.time()

        if not plc_filename:
            if self.empty_filename_start_time is None:
                self.empty_filename_start_time = now
            elapsed = now - self.empty_filename_start_time

            if elapsed < self.empty_filename_timeout:
                # Less than 10 seconds: wait and do nothing
                # Status box will continue displaying latest normal status
                return
            else:
                # 10 seconds or more: blank out all data and show error status
                if self.current_filename != "":
                    self.current_filename = ""
                    self.shapes_input_order = []
                    self.shapes_sorted = []
                    self.header = []
                    self.shapes_by_id = {}
                    self.wagons_config = {}
                    self.slot_occupancy = {}
                    self.output_file = ""
                    status_message = (
                        "Series name not found.\n"
                        "Waiting for valid series.\n"
                        "Current Series: "
                    )
                    self.main_window.set_status(status_message)
                    if hasattr(self.main_window, "set_wagons"):
                        self.main_window.set_wagons([], {}, {}, [], series_name=self._normalize_series_name(self.current_filename))
                    if hasattr(self.overview_window, "update_page"):
                        self.overview_window.wagons_config = {}
                        self.overview_window.shapes_by_id = {}
                        self.overview_window.slot_occupancy = {}
                        self.overview_window.series_name = self._normalize_series_name(self.current_filename)
                        self.overview_window.update_page(highlight_ids=[])
                    self.already_placed_ids = set()
                else:
                    status_message = (
                        "Series name not found.\n"
                        "Waiting for valid series.\n"
                        "Current Series: "
                    )
                    self.main_window.set_status(status_message)
                return
        else:
            # PLC filename is not empty, reset timer
            self.empty_filename_start_time = None

        input_file, output_file, settings_file = self.get_file_paths(plc_filename)

        # If PLC filename is not blank and file does not exist, clear everything and show not found status
        if not os.path.exists(input_file):
            if self.current_filename != plc_filename:
                self.current_filename = plc_filename
            self.shapes_input_order = []
            self.shapes_sorted = []
            self.header = []
            self.shapes_by_id = {}
            self.wagons_config = {}
            self.slot_occupancy = {}
            self.output_file = ""
            status_message = (
                f"Series '{plc_filename}' not found.\n"
                "Waiting for valid series.\n"
                f"Current Series: {plc_filename}"
            )
            self.main_window.set_status(status_message)
            if hasattr(self.main_window, "set_wagons"):
                self.main_window.set_wagons([], {}, {}, [], series_name=self._normalize_series_name(self.current_filename))
            if hasattr(self.overview_window, "update_page"):
                self.overview_window.wagons_config = {}
                self.overview_window.shapes_by_id = {}
                self.overview_window.slot_occupancy = {}
                self.overview_window.series_name = self._normalize_series_name(self.current_filename)
                self.overview_window.update_page(highlight_ids=[])
            self.already_placed_ids = set()
            return

        # If PLC filename changed and file exists, always reload data
        if plc_filename != self.current_filename or not self.shapes_input_order:
            self.reload_data_if_filename_changed(plc_filename)
            self.current_filename = plc_filename

        # If no valid shapes loaded, show "not found" for the current file
        if not self.shapes_input_order:
            status_message = (
                f"Series '{self.current_filename}' not found. Waiting for valid series.\n"
                f"Current Series: {self.current_filename}"
            )
            self.main_window.set_status(status_message)
            return

        id1, id2 = get_pair_ids_from_plc()

        ids = (id1, id2)
        self.update_parameter_table(ids)

        series_display = self._normalize_series_name(self.current_filename)
        self.overview_window.series_name = series_display
        self.overview_window.update_page(highlight_ids=ids)

        def is_real_id(id_val):
            return id_val not in (None, "", "-2")
        
        
        current_ids = [i for i in ids if is_real_id(i)]

        def _pol_and_coords(part_id):
            if not part_id:
                return "-", "-"
            s = self.shapes_by_id.get(part_id)
            if not s:
                return "-", "-"
            pol = str(getattr(s, "pol", "-") or "-")

            wagon = getattr(s, "wagon", None)
            if wagon in (None, "", []):
                wagon = (s._final_placement.get("wagon") if s._final_placement else None)
            row = getattr(s, "row", None)
            if row in (None, "", []):
                row = (s._final_placement.get("row") if s._final_placement else None)
            place = getattr(s, "slot", None)
            if place in (None, "", []):
                place = (s._final_placement.get("slot") if s._final_placement else None)
            wagon_str = str(wagon) if wagon is not None else "?"
            row_str = str(row) if row is not None else "?"
            place_str = str(place) if place is not None else "?"
            coords = f"{wagon_str},{row_str},{place_str}"
            return pol, coords

        # Build previous-placed list from chronological history but exclude current PLC IDs
        placed_history = getattr(self, "placed_history", []) or []
        placed_history_filtered = [pid for pid in placed_history if pid not in current_ids]

        # pick the four most recent placed parts (after filtering out current IDs)
        prev1 = placed_history_filtered[-1] if len(placed_history_filtered) >= 1 else None
        prev2 = placed_history_filtered[-2] if len(placed_history_filtered) >= 2 else None
        prev3 = placed_history_filtered[-3] if len(placed_history_filtered) >= 3 else None
        prev4 = placed_history_filtered[-4] if len(placed_history_filtered) >= 4 else None

        status_lines = []

        # Current IDs: show "ID (POL)" for each PLC ID
        if current_ids:
            cur_display = []
            for cid in current_ids:
                s = self.shapes_by_id.get(cid)
                pol_str = str(getattr(s, "pol", "-")) if s is not None else "-"
                cur_display.append(f"{cid} (P{pol_str})")
            status_lines.append("Current IDs: " + ", ".join(cur_display))
        else:
            status_lines.append("Current IDs: -")

        if prev1:
            pol1, coords1 = _pol_and_coords(prev1)
            status_lines.append(f"1:  P{pol1} - {coords1}")
        else:
            status_lines.append("1: -")

        if prev2:
            pol2, coords2 = _pol_and_coords(prev2)
            status_lines.append(f"2: P{pol2} - {coords2}")
        else:
            status_lines.append("2: -")

        if prev3:
            pol3, coords3 = _pol_and_coords(prev3)
            status_lines.append(f"3: P{pol3} - {coords3}")
        else:
            status_lines.append("3: -")

        if prev4:
            pol4, coords4 = _pol_and_coords(prev4)
            status_lines.append(f"4: P{pol4} - {coords4}")
        else:
            status_lines.append("4: -")

        status_lines.append(f"Current Series: {self.current_filename or '-'}")

        if self.last_placed_parts:
            status_lines.append("Parts Placed: " + ", ".join(str(x) for x in self.last_placed_parts))
        else:
            status_lines.append("Parts Placed: -")

        status_msg = "\n".join(status_lines)

        if not (is_real_id(id1) or is_real_id(id2)):
            self.main_window.set_status(status_msg)
            self.last_polled_ids = ids
            self.last_highlight_ids = ids
            return
        error_lines = []
        new_ids = [id for id in ids if is_real_id(id) and id not in self.already_placed_ids]
        valid_ids = []
        for part_id in new_ids:
            s = self.shapes_by_id.get(part_id)
            if s and s.placed == 0:
                if s._final_placement:
                    fits = (s.width <= s._final_placement["slot_width"] and s.height <= s._final_placement["row_height"])
                    rotated = False
                    if not fits and (s.height <= s._final_placement["slot_width"] and s.width <= s._final_placement["row_height"]):
                        fits = True
                        rotated = True
                    if not fits:
                        error_lines.append(f"Part {s.id} does not fit in assigned slot (Wagon {s._final_placement['wagon']} Row {s._final_placement['row']} Slot {s._final_placement['slot']})!")
                        continue
                    s.placed = 1
                    s.wagon = s._final_placement["wagon"]
                    s.row = s._final_placement["row"]
                    s.slot = s._final_placement["slot"]
                    s.rotated = rotated or s._final_placement.get("rotated", False)
                    self.already_placed_ids.add(part_id)
                    valid_ids.append(s.id)
                else:
                    error_lines.append(f"Part {part_id} could not be assigned to any slot!")
            elif not s:
                error_lines.append(f"Invalid part ID: {part_id}")
        if valid_ids:
            error_lines.append(f"Parts placed: {', '.join(valid_ids)}")
            self.last_placed_parts = list(valid_ids)
            for pid in valid_ids:
                self.placed_history.append(pid)

        if any(is_real_id(i) for i in ids) and ids != getattr(self, 'last_drawn_ids', None):
            self.previous_drawn_ids = getattr(self, 'last_drawn_ids', (None, None))
            series_display = self._normalize_series_name(self.current_filename)
            self.update_main_window(ids, series_name=series_display)
            self.main_window.activateWindow()
            self.last_drawn_ids = ids

        if error_lines:
            status_msg += "\n" + "\n".join(error_lines)

        self.main_window.set_status(status_msg)
        self.last_polled_ids = ids
        self.last_highlight_ids = ids

        current_state = set((s.id, s.placed) for s in self.shapes_input_order)
        if current_state != self.last_saved_state:
            write_shapes_output(self.shapes_input_order, self.output_file, self.header)
            self.last_saved_state = set(current_state)

    def update_main_window(self, highlight_ids, series_name=None):
        wagon_ids = self.get_current_wagons(highlight_ids)
        self.main_window.set_wagons(wagon_ids, self.shapes_by_id, self.slot_occupancy, highlight_ids, series_name=series_name)
        shape1 = self.shapes_by_id.get(highlight_ids[0]) if len(highlight_ids) > 0 and highlight_ids[0] else None
        shape2 = self.shapes_by_id.get(highlight_ids[1]) if len(highlight_ids) > 1 and highlight_ids[1] else None
        self.main_window.set_info(shape1, shape2)

    def get_current_wagons(self, ids):
        wagons = []
        for id in ids:
            if id in self.shapes_by_id and self.shapes_by_id[id].wagon is not None:
                wagons.append(self.shapes_by_id[id].wagon)
        wagons = [w for i, w in enumerate(wagons) if w is not None and w not in wagons[:i]]
        if not wagons and any(ids):
            for id in ids:
                shape = self.shapes_by_id.get(id)
                if shape and shape._final_placement:
                    w = shape._final_placement.get("wagon")
                    if w is not None and w not in wagons:
                        wagons.append(w)
        return wagons

    def update_parameter_table(self, ids):
        shape1 = self.shapes_by_id.get(ids[0]) if ids[0] else None
        shape2 = self.shapes_by_id.get(ids[1]) if ids[1] else None
        self.main_window.set_info(shape1, shape2)

    def do_force_refresh(self):
        if not getattr(self, "current_filename", ""):
            self.main_window.set_status("Force refresh: no current series loaded.")
            return

        try:
            self.reload_data_if_filename_changed(self.current_filename)
            self.main_window.set_status(f"Force refreshed series: {self.current_filename}")
        except Exception as e:
            self.main_window.set_status(f"Force refresh failed: {e}")

    def do_save(self):
        write_shapes_output(self.shapes_input_order, self.output_file, self.header)
        self.main_window.set_status(f"Packed parts saved to {self.output_file}\nCurrent Series: {self.current_filename}")

    def do_close(self):
        self.timer.stop()
        self.quit()

if __name__ == "__main__":
    plc_filename = get_filename_from_plc()
    if not plc_filename:
        plc_filename = "DoNotDelete"  # fallback if PLC not available yet (only first startup)
    input_file = f"{File_Path}\\{plc_filename}_final.txt"

    output_file = input_file
    settings_file = SHELF_FILE

    wagons_config = read_shelves(settings_file)
    shapes_input_order, shapes_sorted, header = read_shapes(input_file)
    shapes_by_id = {s.id: s for s in shapes_sorted}
    placement_dict, slot_occupancy = assign_parts_to_slots(shapes_sorted, wagons_config)
    for s in shapes_sorted:
        place = placement_dict.get((s.pol, s.id))
        if place:
            s._final_placement = place
            s.wagon = place.get("wagon", "")
            s.row = place.get("row", "")
            s.slot = place.get("slot", "")
            s.rotated = place.get("rotated", False)
        else:
            s._final_placement = None
            s.wagon = None
            s.row = None
            s.slot = None
            s.rotated = False

        if not s._final_placement:
            compatible_wagons = [w for w, cfg in wagons_config.items()
                                 if s.vip_key in cfg["allowed_vip_keys"]]
            if not compatible_wagons:
                reason = f"No wagon accepts VIP-Key '{s.vip_key}'"
            else:
                reason = "No available slot or part does not fit"
            print(f"Part {s.id} could not be placed: {reason}")

    write_shapes_output(shapes_input_order, output_file, header)

    app = StackerApp(
        shapes_input_order, shapes_sorted, header,
        shapes_by_id, wagons_config, slot_occupancy, output_file, plc_filename
    )
    sys.exit(app.exec_())