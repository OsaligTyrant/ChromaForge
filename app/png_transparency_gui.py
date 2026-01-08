import os
import re
import json
import csv
import time
import threading
import queue
import shutil
import traceback
import sys
import math
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
except Exception:
    TkinterDnD = None
    DND_FILES = None

try:
    from PIL import Image, ImageTk
except Exception as exc:
    Image = None
    ImageTk = None
    PIL_IMPORT_ERROR = exc
else:
    PIL_IMPORT_ERROR = None

APP_DIR = os.path.dirname(os.path.abspath(__file__))
APP_NAME = "ChromaForge"
APP_VERSION = "Beta 3.1.1"
APP_TITLE = f"{APP_NAME} - {APP_VERSION}"
THEMES = ["Light", "Dark"]
def get_app_home():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.abspath(os.path.join(APP_DIR, os.pardir))


APP_HOME = get_app_home()
SETTINGS_DIR = os.path.join(APP_HOME, "settings")
LAST_SETTINGS_PATH = os.path.join(SETTINGS_DIR, "last_settings.json")
ERROR_LOG_PATH = os.path.join(SETTINGS_DIR, "startup_error.log")
LOGO_PNG = "ChromaForge_logo.png"
ICON_ICO = "ChromaForge_logo.ico"


def resource_path(rel_path):
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, rel_path)
    return os.path.join(APP_DIR, rel_path)
RECENT_PRESETS_MAX = 8

DEFAULT_INPUT = r"C:\Users\TA-Ko\Downloads\DirectorCastRipper_D12\Exports\TWO\Sprite Materials\Characters\Old"
DEFAULT_OUTPUT = r"C:\Users\TA-Ko\Downloads\DirectorCastRipper_D12\Exports\TWO\Sprite Materials\Characters\New"
DEFAULT_COLOR = "#00FF00"
DEFAULT_PREFIXES = ["Characters", "Inventory", "FX", "Chars", "MapGFX"]
FOLDER_MODES = ["Process", "Skip", "Rename only"]


def normalize_hex(hex_str):
    value = hex_str.strip()
    if value.startswith("#"):
        value = value[1:]
    if not re.fullmatch(r"[0-9A-Fa-f]{6}", value):
        raise ValueError("Color must be a 6-digit hex value like #00FF00")
    return "#" + value.upper()


def hex_to_rgb(hex_str):
    value = normalize_hex(hex_str)[1:]
    r = int(value[0:2], 16)
    g = int(value[2:4], 16)
    b = int(value[4:6], 16)
    return (r, g, b)


def parse_hex_list(hex_list):
    if not hex_list.strip():
        return []
    parts = re.split(r"[\s,]+", hex_list.strip())
    colors = []
    for part in parts:
        if not part:
            continue
        colors.append(normalize_hex(part))
    return colors


def first_hex_from_list(hex_list):
    colors = parse_hex_list(hex_list)
    return colors[0] if colors else ""


def build_prefix_regex(prefixes):
    if not prefixes:
        return None
    escaped = [re.escape(p) for p in prefixes]
    return re.compile(r"^(?P<prefix>" + "|".join(escaped) + r")_\d+_")


def process_image(in_path, out_path, mode, target_rgbs, replace_pair, fill_color, fill_shadows, dry_run):
    with Image.open(in_path) as img:
        img = img.convert("RGBA")
        pixels = img.load()
        width, height = img.size

        converted = 0
        if mode == "transparent":
            target_set = set(target_rgbs)
            for y in range(height):
                for x in range(width):
                    r, g, b, a = pixels[x, y]
                    if (r, g, b) in target_set:
                        converted += 1
                        if not dry_run:
                            pixels[x, y] = (r, g, b, 0)
        elif mode == "fill":
            if fill_color is None:
                return 0
            fr, fg, fb = fill_color
            for y in range(height):
                for x in range(width):
                    r, g, b, a = pixels[x, y]
                    if a == 0:
                        converted += 1
                        if not dry_run:
                            pixels[x, y] = (fr, fg, fb, 255)
                    elif fill_shadows and a < 255:
                        converted += 1
                        if not dry_run:
                            pixels[x, y] = (fr, fg, fb, 255)
        else:
            src_rgb, dst_rgb = replace_pair
            for y in range(height):
                for x in range(width):
                    r, g, b, a = pixels[x, y]
                    if (r, g, b) == src_rgb:
                        converted += 1
                        if not dry_run:
                            pixels[x, y] = (dst_rgb[0], dst_rgb[1], dst_rgb[2], a)

        if not dry_run:
            os.makedirs(os.path.dirname(out_path), exist_ok=True)
            img.save(out_path, format="PNG")

    return converted


def output_filename(name, prefix_re, keep_prefixes):
    base, ext = os.path.splitext(name)
    if prefix_re is not None and not keep_prefixes:
        base = prefix_re.sub("", base)
    return base + ext


def output_prefix_folder(name, prefix_re):
    if prefix_re is None:
        return "needs_sorting"
    base = os.path.splitext(name)[0]
    match = prefix_re.match(base)
    if not match:
        return "needs_sorting"
    return match.group("prefix").lower()


def extract_group_prefix(name):
    base = os.path.splitext(name)[0]
    parts = re.split(r"[-_]", base)
    if not parts:
        return base
    first = parts[0].strip()
    if not first:
        return base
    if first.isdigit() and len(parts) > 1:
        second = parts[1].strip()
        if second:
            second_upper = second.upper()
            if re.fullmatch(r"F\d+", second_upper) or second_upper in {"N", "E", "S", "W"} or second.isdigit():
                return first
            if re.search(r"[A-Z]", second_upper):
                return second
    return first


def sprite_sort_key(name):
    base = os.path.splitext(name)[0]
    base_upper = base.upper()
    direction_match = re.search(r"(?:^|[-_])(N|E|S|W)(?:$|[-_])", base_upper)
    direction = direction_match.group(1) if direction_match else ""
    direction_order = {"N": 0, "E": 1, "S": 2, "W": 3}
    direction_index = direction_order.get(direction, 99)

    frame_match = re.search(r"(?:^|[-_])F(\d+)(?:$|[-_])", base_upper)
    if frame_match:
        return (0, direction_index, int(frame_match.group(1)), base_upper)

    number_match = re.search(r"(?:^|[-_])(\d+)(?:$)", base_upper)
    if number_match:
        return (1, int(number_match.group(1)), base_upper)

    return (2, base_upper)


def ensure_csv_header(path):
    if os.path.exists(path):
        return
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(["input", "output", "action", "converted"])


class Tooltip:
    def __init__(self, widget, text_func, enabled_var):
        self.widget = widget
        self.text_func = text_func
        self.enabled_var = enabled_var
        self.tip = None
        self.after_id = None
        widget.bind("<Enter>", self.schedule)
        widget.bind("<Leave>", self.hide)
        widget.bind("<ButtonPress>", self.hide)

    def schedule(self, _event=None):
        if not self.enabled_var.get():
            return
        self.after_id = self.widget.after(400, self.show)

    def show(self):
        if not self.enabled_var.get():
            return
        text = self.text_func()
        if not text:
            return
        x = self.widget.winfo_rootx() + 16
        y = self.widget.winfo_rooty() + 20
        self.tip = tk.Toplevel(self.widget)
        self.tip.wm_overrideredirect(True)
        self.tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            self.tip,
            text=text,
            background="#ffffe0",
            foreground="#000000",
            relief="solid",
            borderwidth=1,
            padx=6,
            pady=3,
        )
        label.pack()

    def hide(self, _event=None):
        if self.after_id:
            self.widget.after_cancel(self.after_id)
            self.after_id = None
        if self.tip is not None:
            self.tip.destroy()
            self.tip = None


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_TITLE)
        self.queue = queue.Queue()
        self.running = False
        self.previewing = False
        self.window_size = None
        self._resize_save_after_id = None
        icon_path = resource_path(ICON_ICO)
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)

        self.input_var = tk.StringVar(value=DEFAULT_INPUT)
        self.output_var = tk.StringVar(value=DEFAULT_OUTPUT)
        self.mode_var = tk.StringVar(value="transparent")
        self.color_list_var = tk.StringVar(value=DEFAULT_COLOR)
        self.fill_color_var = tk.StringVar(value="")
        self.fill_shadows_var = tk.BooleanVar(value=False)
        self.replace_from_var = tk.StringVar(value="")
        self.replace_to_var = tk.StringVar(value="")
        self.extra_prefix_var = tk.StringVar(value="")
        self.keep_prefixes_var = tk.BooleanVar(value=False)
        self.process_all_var = tk.BooleanVar(value=True)
        self.skip_existing_var = tk.BooleanVar(value=True)
        self.skip_existing_files_var = tk.BooleanVar(value=False)
        self.dry_run_var = tk.BooleanVar(value=False)
        self.csv_log_var = tk.BooleanVar(value=False)
        self.csv_path_var = tk.StringVar(value="")
        self.exclude_folders_var = tk.StringVar(value="")
        self.recent_preset_var = tk.StringVar(value="")
        self.recent_presets = []
        self.tooltips_enabled_var = tk.BooleanVar(value=True)
        self.theme_var = tk.StringVar(value="Light")
        self.theme_colors = {}
        self.sprite_layout_var = tk.StringVar(value="Grid")
        self.sprite_columns_var = tk.IntVar(value=4)
        self.sprite_padding_mode_var = tk.StringVar(value="Fixed")
        self.sprite_padding_var = tk.IntVar(value=1)
        self.sheet_running = False
        self.tile_layout_var = tk.StringVar(value="Grid")
        self.tile_columns_var = tk.IntVar(value=15)
        self.tile_size_var = tk.IntVar(value=32)
        self.tile_export_meta_var = tk.BooleanVar(value=True)
        self.tile_running = False
        self.layout_type_var = tk.StringVar(value="Spritesheet")
        self.layout_folder_var = tk.StringVar(value="")
        self.layout_output_var = tk.StringVar(value="")
        self.layout_mode_var = tk.StringVar(value="Grid")
        self.layout_grid_columns_var = tk.IntVar(value=4)
        self.layout_grid_rows_var = tk.IntVar(value=0)
        self.layout_snap_var = tk.BooleanVar(value=True)
        self.layout_show_grid_var = tk.BooleanVar(value=True)
        self.layout_show_guides_var = tk.BooleanVar(value=False)
        self.tile_size_mode_var = tk.StringVar(value="Force")
        self.layout_recent_json_var = tk.StringVar(value="")
        self.layout_recent_json_paths = []
        self.layout_dnd_ready = False
        self.layout_items = {}
        self.layout_selected_ids = set()
        self.layout_drag_start = None
        self.layout_drag_positions = {}
        self.layout_drag_moved = False
        self.layout_drag_anchor_id = None
        self.layout_pending_select_id = None
        self.layout_pos_x_var = tk.StringVar(value="")
        self.layout_pos_y_var = tk.StringVar(value="")
        self.layout_zoom = 1.0
        self.layout_layers = []
        self.layout_active_layer_id = None
        self.layout_reference_layer_id = None
        self.layout_layer_counter = 0
        self.layout_item_counter = 0
        self.layout_layer_drag_index = None
        self.layout_select_box_start = None
        self.layout_select_box_id = None
        self.layout_select_box_add = False
        self.layout_select_box_moved = False
        self.layout_undo_stack = []
        self.layout_guides_enabled = False
        self.layout_guides = {"h": [], "v": []}
        self.layout_guide_drag = None
        self.layout_guide_threshold = 6
        self.layout_anchor_mode_var = tk.StringVar(value="Off")
        self.layout_anchor_source_var = tk.StringVar(value="Auto")
        self.layout_anchor_type_var = tk.StringVar(value="Bottom")
        self.layout_anchor_x_var = tk.StringVar(value="")
        self.layout_anchor_y_var = tk.StringVar(value="")
        self.layout_anchor_padding_var = tk.IntVar(value=2)
        self.layout_anchor_inherit_var = tk.BooleanVar(value=True)
        self.layout_anchor_ui_lock = False
        self.split_sheet_path_var = tk.StringVar(value="")
        self.split_output_dir_var = tk.StringVar(value="")
        self.split_base_name_var = tk.StringVar(value="")
        self.split_cell_w_var = tk.IntVar(value=32)
        self.split_cell_h_var = tk.IntVar(value=32)
        self.split_columns_var = tk.IntVar(value=0)
        self.split_rows_var = tk.IntVar(value=0)
        self.split_pad_x_var = tk.IntVar(value=0)
        self.split_pad_y_var = tk.IntVar(value=0)
        self.split_offset_x_var = tk.IntVar(value=0)
        self.split_offset_y_var = tk.IntVar(value=0)
        self.split_show_grid_var = tk.BooleanVar(value=True)
        self.split_zoom = 1.0
        self.split_sheet_pil = None
        self.split_sheet_photo = None
        self.split_sheet_canvas_id = None
        self.split_selected_cells = set()
        self.split_preview_window = None
        self.split_preview_label = None
        self.split_preview_photo = None
        self.split_resize_axis = None
        self.split_resize_anchor = 0
        self.split_resize_start = None
        self.split_resize_moved = False
        self.split_undo_stack = []

        self.prefix_vars = {}
        self.conflict_event = threading.Event()
        self.conflict_choice = None

        self.folder_rows = {}

        self._build_ui()
        self.layout_init_layers()
        self.refresh_folders()
        self.load_last_settings()
        self.refresh_layout_folders()
        self.poll_queue()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self):
        root = self.root

        self.style = ttk.Style()
        self.apply_theme(self.theme_var.get())

        notebook = ttk.Notebook(root)
        notebook.grid(row=0, column=0, sticky="nsew")
        self.main_notebook = notebook

        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        root.rowconfigure(1, weight=0)

        color_tab = ttk.Frame(notebook)
        sheet_tab = ttk.Frame(notebook)
        tile_tab = ttk.Frame(notebook)
        layout_assemble_tab = ttk.Frame(notebook)
        layout_split_tab = ttk.Frame(notebook)
        notebook.add(color_tab, text="Color Mode")
        notebook.add(sheet_tab, text="Spritesheet Mode")
        notebook.add(tile_tab, text="Tilemap Mode")
        notebook.add(layout_assemble_tab, text="Layout Editor (Assemble)")
        notebook.add(layout_split_tab, text="Layout Editor (Split)")

        footer = ttk.Frame(root)
        footer.grid(row=1, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        ttk.Label(footer, text=f"Version: {APP_VERSION}").grid(row=0, column=0, sticky="e", padx=10, pady=(0, 6))

        main = ttk.Frame(color_tab, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        color_tab.columnconfigure(0, weight=1)
        color_tab.rowconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)

        self.main_sheet_tab = sheet_tab
        self.main_tile_tab = tile_tab
        self.main_layout_tab = layout_assemble_tab
        self.main_split_tab = layout_split_tab

        preset_frame = ttk.Frame(main)
        preset_frame.grid(row=0, column=0, columnspan=3, sticky="ew")
        ttk.Button(preset_frame, text="Load Preset", command=self.load_preset).grid(row=0, column=0, padx=4, pady=2)
        ttk.Button(preset_frame, text="Save Preset", command=self.save_preset).grid(row=0, column=1, padx=4, pady=2)
        ttk.Label(preset_frame, text="Recent").grid(row=0, column=2, padx=(10, 4), pady=2)
        self.recent_combo = ttk.Combobox(preset_frame, textvariable=self.recent_preset_var, width=40, state="readonly")
        self.recent_combo.grid(row=0, column=3, padx=4, pady=2, sticky="ew")
        ttk.Button(preset_frame, text="Load Selected", command=self.load_recent_preset).grid(row=0, column=4, padx=4, pady=2)
        ttk.Label(preset_frame, text="Theme").grid(row=0, column=5, padx=(14, 4), pady=2)
        self.theme_combo = ttk.Combobox(preset_frame, textvariable=self.theme_var, values=THEMES, state="readonly", width=10, style="Theme.TCombobox", takefocus=False)
        self.theme_combo.grid(row=0, column=6, padx=4, pady=2)
        self.theme_combo.bind("<<ComboboxSelected>>", lambda _e: self.on_theme_change())
        preset_frame.columnconfigure(3, weight=1)

        ttk.Label(main, text="Input (Old) Folder").grid(row=1, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.input_var).grid(row=1, column=1, sticky="ew", padx=5)
        ttk.Button(main, text="Browse", command=self.pick_input).grid(row=1, column=2)

        ttk.Label(main, text="Output (New) Folder").grid(row=2, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.output_var).grid(row=2, column=1, sticky="ew", padx=5)
        ttk.Button(main, text="Browse", command=self.pick_output).grid(row=2, column=2)

        color_frame = ttk.LabelFrame(main, text="Color Mode")
        color_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=6)
        color_frame.columnconfigure(2, weight=1)

        transparent_radio = ttk.Radiobutton(color_frame, text="Make transparent", variable=self.mode_var, value="transparent", command=self.update_color_mode)
        transparent_radio.grid(row=0, column=0, sticky="w", padx=6, pady=2)
        ttk.Label(color_frame, text="Hex list (comma or space separated)").grid(row=0, column=1, sticky="w")
        self.color_list_entry = ttk.Entry(color_frame, textvariable=self.color_list_var)
        self.color_list_entry.grid(row=0, column=2, sticky="ew", padx=6)
        self.color_list_swatch = tk.Label(color_frame, width=3, relief="sunken")
        self.color_list_swatch.grid(row=0, column=3, padx=4)

        fill_radio = ttk.Radiobutton(color_frame, text="Fill with color", variable=self.mode_var, value="fill", command=self.update_color_mode)
        fill_radio.grid(row=1, column=0, sticky="w", padx=6, pady=2)
        ttk.Label(color_frame, text="Fill hex").grid(row=1, column=1, sticky="w")
        self.fill_color_entry = ttk.Entry(color_frame, textvariable=self.fill_color_var, width=12)
        self.fill_color_entry.grid(row=1, column=2, sticky="w", padx=6)
        self.fill_color_swatch = tk.Label(color_frame, width=3, relief="sunken")
        self.fill_color_swatch.grid(row=1, column=3, padx=4)
        self.fill_shadows_check = ttk.Checkbutton(color_frame, text="Also fill shadows (semi-transparent) [experimental]", variable=self.fill_shadows_var)
        self.fill_shadows_check.grid(row=1, column=4, sticky="w", padx=6)

        replace_radio = ttk.Radiobutton(color_frame, text="Replace color", variable=self.mode_var, value="replace", command=self.update_color_mode)
        replace_radio.grid(row=2, column=0, sticky="w", padx=6, pady=2)
        ttk.Label(color_frame, text="From").grid(row=2, column=1, sticky="e")
        self.replace_from_entry = ttk.Entry(color_frame, textvariable=self.replace_from_var, width=12)
        self.replace_from_entry.grid(row=2, column=2, sticky="w", padx=6)
        self.replace_from_swatch = tk.Label(color_frame, width=3, relief="sunken")
        self.replace_from_swatch.grid(row=2, column=3, padx=4)
        ttk.Label(color_frame, text="To").grid(row=2, column=4, sticky="e")
        self.replace_to_entry = ttk.Entry(color_frame, textvariable=self.replace_to_var, width=12)
        self.replace_to_entry.grid(row=2, column=5, sticky="w", padx=6)
        self.replace_to_swatch = tk.Label(color_frame, width=3, relief="sunken")
        self.replace_to_swatch.grid(row=2, column=6, padx=4)

        prefix_frame = ttk.LabelFrame(main, text="Prefixes")
        prefix_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=6)
        prefix_controls = ttk.Frame(prefix_frame)
        prefix_controls.grid(row=0, column=0, sticky="ew", padx=6, pady=(4, 0))
        self.scan_prefixes_button = ttk.Button(prefix_controls, text="Scan for prefixes", command=self.scan_prefixes)
        self.scan_prefixes_button.grid(row=0, column=0, sticky="w")
        self.keep_prefixes_check = ttk.Checkbutton(prefix_controls, text="Keep prefixes", variable=self.keep_prefixes_var)
        self.keep_prefixes_check.grid(row=0, column=1, sticky="w", padx=(10, 0))
        prefix_controls.columnconfigure(2, weight=1)

        self.prefix_list_frame = ttk.Frame(prefix_frame)
        self.prefix_list_frame.grid(row=1, column=0, sticky="ew", padx=6, pady=4)
        self.prefix_empty_label = ttk.Label(self.prefix_list_frame, text="No prefixes detected. Click Scan for prefixes.")
        self.prefix_empty_label.grid(row=0, column=0, sticky="w")

        ttk.Label(prefix_frame, text="Extra prefixes (comma-separated)").grid(row=2, column=0, sticky="w", padx=6)
        self.extra_prefix_entry = ttk.Entry(prefix_frame, textvariable=self.extra_prefix_var)
        self.extra_prefix_entry.grid(row=3, column=0, sticky="ew", padx=6, pady=(0, 6))
        prefix_frame.columnconfigure(0, weight=1)

        folders_frame = ttk.Frame(main)
        folders_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=6)
        folders_frame.columnconfigure(0, weight=1)
        folders_frame.rowconfigure(1, weight=1)
        ttk.Label(folders_frame, text="Folders").grid(row=0, column=0, sticky="w", padx=2, pady=(0, 4))

        self.folders_box = tk.Frame(folders_frame, bd=1, relief="solid")
        self.folders_box.grid(row=1, column=0, sticky="nsew")
        self.folders_box.columnconfigure(0, weight=1)
        self.folders_box.columnconfigure(1, weight=0)
        self.folders_box.rowconfigure(1, weight=1)

        top_controls = ttk.Frame(self.folders_box)
        top_controls.grid(row=0, column=0, sticky="ew", padx=6, pady=4)
        process_all_btn = ttk.Checkbutton(top_controls, text="Process all folders", variable=self.process_all_var)
        process_all_btn.grid(row=0, column=0, sticky="w")
        ttk.Button(top_controls, text="Refresh", command=self.refresh_folders).grid(row=0, column=1, padx=6)
        ttk.Button(top_controls, text="Select All", command=self.select_all_folders).grid(row=0, column=2, padx=6)
        ttk.Button(top_controls, text="Select None", command=self.select_no_folders).grid(row=0, column=3, padx=6)

        self.folders_canvas = tk.Canvas(self.folders_box, height=160, highlightthickness=0, bd=0)
        scroll = ttk.Scrollbar(self.folders_box, orient="vertical", command=self.folders_canvas.yview)
        self.folders_container = ttk.Frame(self.folders_canvas)

        self.folders_container.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")),
        )

        self.folders_canvas.create_window((0, 0), window=self.folders_container, anchor="nw")
        self.folders_canvas.configure(yscrollcommand=scroll.set)

        self.folders_canvas.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        scroll.grid(row=1, column=1, sticky="ns", pady=(0, 6))

        options_frame = ttk.LabelFrame(main, text="Options")
        options_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=4)
        skip_existing_btn = ttk.Checkbutton(options_frame, text="Skip prefix folders that already exist in New (default)", variable=self.skip_existing_var)
        skip_existing_btn.grid(row=0, column=0, sticky="w")
        skip_files_btn = ttk.Checkbutton(options_frame, text="Skip files that already exist in New", variable=self.skip_existing_files_var)
        skip_files_btn.grid(row=1, column=0, sticky="w")
        dry_run_btn = ttk.Checkbutton(options_frame, text="Dry run (no files written)", variable=self.dry_run_var)
        dry_run_btn.grid(row=2, column=0, sticky="w")
        ttk.Label(options_frame, text="Exclude folders (comma-separated)").grid(row=3, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(options_frame, textvariable=self.exclude_folders_var).grid(row=4, column=0, sticky="ew", pady=(0, 4))
        tooltips_btn = ttk.Checkbutton(options_frame, text="Enable tooltips", variable=self.tooltips_enabled_var)
        tooltips_btn.grid(row=5, column=0, sticky="w", pady=(2, 0))
        options_frame.columnconfigure(0, weight=1)

        csv_frame = ttk.Frame(options_frame)
        csv_frame.grid(row=3, column=0, sticky="ew", pady=2)
        ttk.Checkbutton(csv_frame, text="Export CSV log", variable=self.csv_log_var).grid(row=0, column=0, sticky="w")
        ttk.Entry(csv_frame, textvariable=self.csv_path_var, width=50).grid(row=0, column=1, padx=6, sticky="ew")
        ttk.Button(csv_frame, text="Browse", command=self.pick_csv).grid(row=0, column=2)
        csv_frame.columnconfigure(1, weight=1)

        run_frame = ttk.Frame(main)
        run_frame.grid(row=7, column=0, columnspan=3, sticky="ew", pady=6)
        self.run_button = ttk.Button(run_frame, text="Run", command=self.run)
        self.run_button.grid(row=0, column=0, sticky="w")
        self.preview_button = ttk.Button(run_frame, text="Preview Routing", command=self.preview_routing)
        self.preview_button.grid(row=0, column=1, sticky="w", padx=(8, 0))
        self.progress = ttk.Progressbar(run_frame, length=240)
        self.progress.grid(row=0, column=2, padx=10)
        self.status_label = ttk.Label(run_frame, text="")
        self.status_label.grid(row=0, column=3, sticky="w")
        run_frame.columnconfigure(2, weight=1)

        log_frame = ttk.LabelFrame(main, text="Log")
        log_frame.grid(row=8, column=0, columnspan=3, sticky="nsew", pady=6)
        log_frame.rowconfigure(0, weight=1)
        log_frame.columnconfigure(0, weight=1)
        self.log = tk.Text(log_frame, height=12, wrap="none")
        self.log.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log.yview)
        log_scroll.grid(row=0, column=1, sticky="ns")
        self.log.configure(yscrollcommand=log_scroll.set)

        main.rowconfigure(8, weight=1)

        sheet_main = ttk.Frame(self.main_sheet_tab, padding=10)
        sheet_main.grid(row=0, column=0, sticky="nsew")
        self.main_sheet_tab.columnconfigure(0, weight=1)
        self.main_sheet_tab.rowconfigure(0, weight=1)
        sheet_main.columnconfigure(1, weight=1)

        ttk.Label(sheet_main, text="Input Folder").grid(row=0, column=0, sticky="w")
        ttk.Entry(sheet_main, textvariable=self.input_var).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(sheet_main, text="Browse", command=self.pick_input).grid(row=0, column=2)

        sheet_layout = ttk.LabelFrame(sheet_main, text="Layout")
        sheet_layout.grid(row=1, column=0, columnspan=3, sticky="ew", pady=6)
        ttk.Label(sheet_layout, text="Mode").grid(row=0, column=0, sticky="w", padx=6, pady=2)
        self.sprite_layout_combo = ttk.Combobox(
            sheet_layout,
            textvariable=self.sprite_layout_var,
            values=["Grid", "Horizontal", "Vertical"],
            state="readonly",
            width=12,
        )
        self.sprite_layout_combo.grid(row=0, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(sheet_layout, text="Columns").grid(row=0, column=2, sticky="w", padx=6)
        self.sprite_columns_spin = ttk.Spinbox(sheet_layout, from_=1, to=20, textvariable=self.sprite_columns_var, width=6)
        self.sprite_columns_spin.grid(row=0, column=3, sticky="w", padx=6, pady=2)
        ttk.Label(sheet_layout, text="Padding").grid(row=0, column=4, sticky="w", padx=6)
        self.sprite_padding_mode_combo = ttk.Combobox(
            sheet_layout,
            textvariable=self.sprite_padding_mode_var,
            values=["Fixed", "Frame width", "Frame height"],
            state="readonly",
            width=12,
        )
        self.sprite_padding_mode_combo.grid(row=0, column=5, sticky="w", padx=6, pady=2)
        self.sprite_padding_spin = ttk.Spinbox(sheet_layout, from_=0, to=128, textvariable=self.sprite_padding_var, width=6)
        self.sprite_padding_spin.grid(row=0, column=6, sticky="w", padx=6, pady=2)
        sheet_layout.columnconfigure(7, weight=1)

        ttk.Label(sheet_main, text="Exclude folders (comma-separated)").grid(row=2, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(sheet_main, textvariable=self.exclude_folders_var).grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 4))

        sheet_run_frame = ttk.Frame(sheet_main)
        sheet_run_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=6)
        self.sheet_run_button = ttk.Button(sheet_run_frame, text="Generate Sprite Sheets", command=self.run_sprite_sheets)
        self.sheet_run_button.grid(row=0, column=0, sticky="w")
        self.sheet_progress = ttk.Progressbar(sheet_run_frame, length=240, mode="indeterminate")
        self.sheet_progress.grid(row=0, column=1, padx=10)
        self.sheet_status_label = ttk.Label(sheet_run_frame, text="")
        self.sheet_status_label.grid(row=0, column=2, sticky="w")
        sheet_run_frame.columnconfigure(1, weight=1)

        sheet_log_frame = ttk.LabelFrame(sheet_main, text="Log")
        sheet_log_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=6)
        sheet_log_frame.rowconfigure(0, weight=1)
        sheet_log_frame.columnconfigure(0, weight=1)
        self.sheet_log = tk.Text(sheet_log_frame, height=12, wrap="none")
        self.sheet_log.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        sheet_log_scroll = ttk.Scrollbar(sheet_log_frame, orient="vertical", command=self.sheet_log.yview)
        sheet_log_scroll.grid(row=0, column=1, sticky="ns")
        self.sheet_log.configure(yscrollcommand=sheet_log_scroll.set)
        sheet_main.rowconfigure(5, weight=1)

        tile_main = ttk.Frame(self.main_tile_tab, padding=10)
        tile_main.grid(row=0, column=0, sticky="nsew")
        self.main_tile_tab.columnconfigure(0, weight=1)
        self.main_tile_tab.rowconfigure(0, weight=1)
        tile_main.columnconfigure(1, weight=1)

        ttk.Label(tile_main, text="Input Folder").grid(row=0, column=0, sticky="w")
        ttk.Entry(tile_main, textvariable=self.input_var).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(tile_main, text="Browse", command=self.pick_input).grid(row=0, column=2)

        tile_layout = ttk.LabelFrame(tile_main, text="Layout")
        tile_layout.grid(row=1, column=0, columnspan=3, sticky="ew", pady=6)
        ttk.Label(tile_layout, text="Mode").grid(row=0, column=0, sticky="w", padx=6, pady=2)
        self.tile_layout_combo = ttk.Combobox(
            tile_layout,
            textvariable=self.tile_layout_var,
            values=["Grid", "Horizontal", "Vertical"],
            state="readonly",
            width=12,
        )
        self.tile_layout_combo.grid(row=0, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(tile_layout, text="Columns").grid(row=0, column=2, sticky="w", padx=6)
        self.tile_columns_combo = ttk.Combobox(
            tile_layout,
            textvariable=self.tile_columns_var,
            values=[10, 15, 20, 25],
            state="readonly",
            width=6,
        )
        self.tile_columns_combo.grid(row=0, column=3, sticky="w", padx=6, pady=2)
        ttk.Label(tile_layout, text="Tile size").grid(row=0, column=4, sticky="w", padx=6)
        self.tile_size_combo = ttk.Combobox(
            tile_layout,
            textvariable=self.tile_size_var,
            values=[8, 16, 32, 64],
            state="readonly",
            width=6,
        )
        self.tile_size_combo.grid(row=0, column=5, sticky="w", padx=6, pady=2)
        tile_layout.columnconfigure(6, weight=1)

        ttk.Checkbutton(tile_main, text="Export metadata (CSV)", variable=self.tile_export_meta_var).grid(row=2, column=0, sticky="w", pady=(4, 0))
        ttk.Label(tile_main, text="Exclude folders (comma-separated)").grid(row=3, column=0, sticky="w", pady=(4, 0))
        ttk.Entry(tile_main, textvariable=self.exclude_folders_var).grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 4))

        tile_run_frame = ttk.Frame(tile_main)
        tile_run_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=6)
        self.tile_run_button = ttk.Button(tile_run_frame, text="Generate Tilemaps", command=self.run_tilemaps)
        self.tile_run_button.grid(row=0, column=0, sticky="w")
        self.tile_progress = ttk.Progressbar(tile_run_frame, length=240, mode="indeterminate")
        self.tile_progress.grid(row=0, column=1, padx=10)
        self.tile_status_label = ttk.Label(tile_run_frame, text="")
        self.tile_status_label.grid(row=0, column=2, sticky="w")
        tile_run_frame.columnconfigure(1, weight=1)

        tile_log_frame = ttk.LabelFrame(tile_main, text="Log")
        tile_log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=6)
        tile_log_frame.rowconfigure(0, weight=1)
        tile_log_frame.columnconfigure(0, weight=1)
        self.tile_log = tk.Text(tile_log_frame, height=12, wrap="none")
        self.tile_log.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        tile_log_scroll = ttk.Scrollbar(tile_log_frame, orient="vertical", command=self.tile_log.yview)
        tile_log_scroll.grid(row=0, column=1, sticky="ns")
        self.tile_log.configure(yscrollcommand=tile_log_scroll.set)
        tile_main.rowconfigure(6, weight=1)

        layout_main = ttk.Frame(self.main_layout_tab, padding=10)
        layout_main.grid(row=0, column=0, sticky="nsew")
        self.main_layout_tab.columnconfigure(0, weight=1)
        self.main_layout_tab.rowconfigure(0, weight=1)
        layout_main.columnconfigure(1, weight=1)

        ttk.Label(layout_main, text="Input Folder").grid(row=0, column=0, sticky="w")
        ttk.Entry(layout_main, textvariable=self.input_var).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(layout_main, text="Browse", command=self.pick_input).grid(row=0, column=2)

        ttk.Label(layout_main, text="Source Folder").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.layout_folder_combo = ttk.Combobox(layout_main, textvariable=self.layout_folder_var, state="readonly")
        self.layout_folder_combo.grid(row=1, column=1, sticky="ew", padx=5, pady=(4, 0))
        ttk.Button(layout_main, text="Refresh", command=self.refresh_layout_folders).grid(row=1, column=2, pady=(4, 0))

        ttk.Label(layout_main, text="Type").grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.layout_type_combo = ttk.Combobox(
            layout_main,
            textvariable=self.layout_type_var,
            values=["Spritesheet", "Tilemap"],
            state="readonly",
            width=14,
        )
        self.layout_type_combo.grid(row=2, column=1, sticky="w", padx=5, pady=(4, 0))

        layout_body = ttk.Frame(layout_main)
        layout_body.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=8)
        layout_body.columnconfigure(1, weight=1)
        layout_body.rowconfigure(0, weight=1)
        layout_body.rowconfigure(1, weight=0)
        layout_main.rowconfigure(3, weight=1)

        layout_list_frame = ttk.LabelFrame(layout_body, text="Frames")
        layout_list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        layout_list_frame.rowconfigure(1, weight=1)
        layout_list_frame.columnconfigure(0, weight=1)
        self.layout_listbox = tk.Listbox(layout_list_frame, selectmode="extended", height=18)
        self.layout_listbox.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        layout_list_scroll = ttk.Scrollbar(layout_list_frame, orient="vertical", command=self.layout_listbox.yview)
        layout_list_scroll.grid(row=1, column=1, sticky="ns", pady=6)
        self.layout_listbox.configure(yscrollcommand=layout_list_scroll.set)
        layout_list_buttons = ttk.Frame(layout_list_frame)
        layout_list_buttons.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=(0, 6))
        ttk.Button(layout_list_buttons, text="Add Selected", command=self.layout_add_selected).grid(row=0, column=0, padx=3)
        ttk.Button(layout_list_buttons, text="Remove Selected", command=self.layout_remove_selected).grid(row=0, column=1, padx=3)
        ttk.Button(layout_list_buttons, text="Clear Canvas", command=self.layout_clear_canvas).grid(row=0, column=2, padx=3)

        layout_layers_frame = ttk.LabelFrame(layout_body, text="Layers")
        layout_layers_frame.grid(row=1, column=0, sticky="nsew", padx=(0, 8), pady=(6, 0))
        layout_layers_frame.columnconfigure(0, weight=1)
        layout_layers_frame.rowconfigure(1, weight=1)
        self.layout_layers_listbox = tk.Listbox(layout_layers_frame, height=8, exportselection=False)
        self.layout_layers_listbox.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        layout_layers_scroll = ttk.Scrollbar(layout_layers_frame, orient="vertical", command=self.layout_layers_listbox.yview)
        layout_layers_scroll.grid(row=1, column=1, sticky="ns", pady=6)
        self.layout_layers_listbox.configure(yscrollcommand=layout_layers_scroll.set)
        layout_layers_buttons = ttk.Frame(layout_layers_frame)
        layout_layers_buttons.grid(row=2, column=0, columnspan=2, sticky="ew", padx=6, pady=(0, 6))
        ttk.Button(layout_layers_buttons, text="New", command=self.layout_layer_create).grid(row=0, column=0, padx=3)
        ttk.Button(layout_layers_buttons, text="Rename", command=self.layout_layer_rename).grid(row=0, column=1, padx=3)
        ttk.Button(layout_layers_buttons, text="Delete", command=self.layout_layer_delete).grid(row=0, column=2, padx=3)
        layout_layers_buttons2 = ttk.Frame(layout_layers_frame)
        layout_layers_buttons2.grid(row=3, column=0, columnspan=2, sticky="ew", padx=6, pady=(0, 6))
        ttk.Button(layout_layers_buttons2, text="Assign Frames", command=self.layout_layer_assign_dialog).grid(row=0, column=0, padx=3)
        ttk.Button(layout_layers_buttons2, text="Toggle Layer", command=self.layout_layer_toggle_visibility).grid(row=0, column=1, padx=3)
        ttk.Button(layout_layers_buttons2, text="Set Reference", command=self.layout_layer_set_reference).grid(row=0, column=2, padx=3)

        layout_canvas_frame = ttk.LabelFrame(layout_body, text="Canvas")
        layout_canvas_frame.grid(row=0, column=1, sticky="nsew")
        layout_canvas_frame.rowconfigure(1, weight=1)
        layout_canvas_frame.columnconfigure(0, weight=1)
        layout_canvas_header = ttk.Frame(layout_canvas_frame)
        layout_canvas_header.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 0))
        layout_canvas_header.columnconfigure(0, weight=1)
        ttk.Button(layout_canvas_header, text="Undo", command=self.layout_undo).grid(row=0, column=1, padx=3)
        ttk.Button(layout_canvas_header, text="+", command=lambda: self.layout_set_zoom(self.layout_zoom * 1.1)).grid(
            row=0, column=2, padx=3
        )
        ttk.Button(layout_canvas_header, text="-", command=lambda: self.layout_set_zoom(self.layout_zoom * 0.9)).grid(
            row=0, column=3, padx=3
        )
        self.layout_canvas = tk.Canvas(layout_canvas_frame, width=520, height=360, highlightthickness=1)
        self.layout_canvas.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        layout_canvas_scroll_y = ttk.Scrollbar(layout_canvas_frame, orient="vertical", command=self.layout_canvas.yview)
        layout_canvas_scroll_y.grid(row=1, column=1, sticky="ns", pady=6)
        layout_canvas_scroll_x = ttk.Scrollbar(layout_canvas_frame, orient="horizontal", command=self.layout_canvas.xview)
        layout_canvas_scroll_x.grid(row=2, column=0, sticky="ew", padx=6)
        self.layout_canvas.configure(yscrollcommand=layout_canvas_scroll_y.set, xscrollcommand=layout_canvas_scroll_x.set)

        layout_options_outer = ttk.LabelFrame(layout_body, text="Layout Options")
        layout_options_outer.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        layout_options_outer.rowconfigure(0, weight=1)
        layout_options_outer.columnconfigure(0, weight=1)
        layout_options_canvas = tk.Canvas(layout_options_outer, highlightthickness=0)
        layout_options_canvas.grid(row=0, column=0, sticky="nsew")
        layout_options_scroll = ttk.Scrollbar(layout_options_outer, orient="vertical", command=layout_options_canvas.yview)
        layout_options_scroll.grid(row=0, column=1, sticky="ns")
        layout_options_canvas.configure(yscrollcommand=layout_options_scroll.set)
        layout_options = ttk.Frame(layout_options_canvas)
        layout_options_window = layout_options_canvas.create_window((0, 0), window=layout_options, anchor="nw")
        layout_options.columnconfigure(1, weight=1)

        def _layout_options_resize(_event):
            layout_options_canvas.configure(scrollregion=layout_options_canvas.bbox("all"))

        def _layout_options_canvas_resize(event):
            layout_options_canvas.itemconfigure(layout_options_window, width=event.width)

        layout_options.bind("<Configure>", _layout_options_resize)
        layout_options_canvas.bind("<Configure>", _layout_options_canvas_resize)
        self.layout_name_label = ttk.Label(layout_options, text="Spritesheet name")
        self.layout_name_label.grid(row=0, column=0, sticky="w", padx=6, pady=(6, 2))
        ttk.Entry(layout_options, textvariable=self.layout_output_var).grid(row=0, column=1, columnspan=3, sticky="ew", padx=6, pady=(6, 2))
        ttk.Label(layout_options, text="Mode").grid(row=1, column=0, sticky="w", padx=6, pady=2)
        self.layout_mode_combo = ttk.Combobox(
            layout_options,
            textvariable=self.layout_mode_var,
            values=["Grid", "Free-form"],
            state="readonly",
            width=12,
        )
        self.layout_mode_combo.grid(row=1, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(layout_options, text="Columns").grid(row=2, column=0, sticky="w", padx=6, pady=2)
        self.layout_columns_spin = ttk.Spinbox(layout_options, from_=1, to=50, textvariable=self.layout_grid_columns_var, width=6)
        self.layout_columns_spin.grid(row=2, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(layout_options, text="Rows (0=auto)").grid(row=3, column=0, sticky="w", padx=6, pady=2)
        self.layout_rows_spin = ttk.Spinbox(layout_options, from_=0, to=200, textvariable=self.layout_grid_rows_var, width=6)
        self.layout_rows_spin.grid(row=3, column=1, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(layout_options, text="Snap to grid", variable=self.layout_snap_var, command=self.layout_snap_selected).grid(row=4, column=0, columnspan=2, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(layout_options, text="Show grid lines", variable=self.layout_show_grid_var, command=self.layout_redraw).grid(row=5, column=0, columnspan=2, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(layout_options, text="Show guides", variable=self.layout_show_guides_var, command=self.layout_redraw).grid(row=6, column=0, columnspan=2, sticky="w", padx=6, pady=2)

        layout_pos_frame = ttk.LabelFrame(layout_options, text="Position")
        layout_pos_frame.grid(row=7, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
        ttk.Label(layout_pos_frame, text="X").grid(row=0, column=0, sticky="w", padx=6, pady=2)
        self.layout_pos_x_entry = ttk.Entry(layout_pos_frame, textvariable=self.layout_pos_x_var, width=8)
        self.layout_pos_x_entry.grid(row=0, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(layout_pos_frame, text="Y").grid(row=0, column=2, sticky="w", padx=6, pady=2)
        self.layout_pos_y_entry = ttk.Entry(layout_pos_frame, textvariable=self.layout_pos_y_var, width=8)
        self.layout_pos_y_entry.grid(row=0, column=3, sticky="w", padx=6, pady=2)
        ttk.Button(layout_pos_frame, text="Apply", command=self.layout_apply_position).grid(row=0, column=4, sticky="w", padx=6, pady=2)

        layout_align_frame = ttk.Frame(layout_options)
        layout_align_frame.grid(row=8, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
        ttk.Button(layout_align_frame, text="Align Horizontal", command=lambda: self.layout_align("horizontal")).grid(row=0, column=0, padx=3)
        ttk.Button(layout_align_frame, text="Align Vertical", command=lambda: self.layout_align("vertical")).grid(row=0, column=1, padx=3)
        ttk.Button(layout_align_frame, text="Center in Cell", command=self.layout_center_in_cell).grid(row=0, column=2, padx=3)
        ttk.Button(layout_align_frame, text="Copy Selected", command=self.layout_copy_selected).grid(row=0, column=3, padx=3)
        ttk.Button(layout_align_frame, text="Remove Selected", command=self.layout_remove_selected).grid(row=0, column=4, padx=3)
        ttk.Button(layout_align_frame, text="Toggle Visible", command=self.layout_toggle_selected_visibility).grid(row=1, column=0, padx=3, pady=(4, 0))
        ttk.Button(layout_align_frame, text="Snap Guides", command=self.layout_toggle_guides).grid(row=1, column=1, padx=3, pady=(4, 0))

        layout_anchor_frame = ttk.LabelFrame(layout_options, text="Anchors")
        layout_anchor_frame.grid(row=9, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
        layout_anchor_frame.columnconfigure(1, weight=1)
        ttk.Label(layout_anchor_frame, text="Mode").grid(row=0, column=0, sticky="w", padx=6, pady=2)
        self.layout_anchor_mode_combo = ttk.Combobox(
            layout_anchor_frame,
            textvariable=self.layout_anchor_mode_var,
            values=["Off", "Global", "Per-frame"],
            state="readonly",
            width=10,
        )
        self.layout_anchor_mode_combo.grid(row=0, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(layout_anchor_frame, text="Source").grid(row=0, column=2, sticky="w", padx=6, pady=2)
        self.layout_anchor_source_combo = ttk.Combobox(
            layout_anchor_frame,
            textvariable=self.layout_anchor_source_var,
            values=["Auto", "Manual"],
            state="readonly",
            width=10,
        )
        self.layout_anchor_source_combo.grid(row=0, column=3, sticky="w", padx=6, pady=2)
        ttk.Label(layout_anchor_frame, text="Auto type").grid(row=1, column=0, sticky="w", padx=6, pady=2)
        self.layout_anchor_type_combo = ttk.Combobox(
            layout_anchor_frame,
            textvariable=self.layout_anchor_type_var,
            values=["Bottom", "Top", "Center"],
            state="readonly",
            width=10,
        )
        self.layout_anchor_type_combo.grid(row=1, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(layout_anchor_frame, text="Padding").grid(row=1, column=2, sticky="w", padx=6, pady=2)
        self.layout_anchor_padding_spin = ttk.Spinbox(
            layout_anchor_frame,
            from_=0,
            to=64,
            textvariable=self.layout_anchor_padding_var,
            width=6,
        )
        self.layout_anchor_padding_spin.grid(row=1, column=3, sticky="w", padx=6, pady=2)
        ttk.Label(layout_anchor_frame, text="Anchor X").grid(row=2, column=0, sticky="w", padx=6, pady=2)
        self.layout_anchor_x_entry = ttk.Entry(layout_anchor_frame, textvariable=self.layout_anchor_x_var, width=8)
        self.layout_anchor_x_entry.grid(row=2, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(layout_anchor_frame, text="Anchor Y").grid(row=2, column=2, sticky="w", padx=6, pady=2)
        self.layout_anchor_y_entry = ttk.Entry(layout_anchor_frame, textvariable=self.layout_anchor_y_var, width=8)
        self.layout_anchor_y_entry.grid(row=2, column=3, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(
            layout_anchor_frame,
            text="Use layer anchors (selected)",
            variable=self.layout_anchor_inherit_var,
            command=self.layout_toggle_anchor_inherit,
        ).grid(row=3, column=0, columnspan=4, sticky="w", padx=6, pady=2)
        layout_anchor_buttons = ttk.Frame(layout_anchor_frame)
        layout_anchor_buttons.grid(row=4, column=0, columnspan=4, sticky="ew", padx=6, pady=(4, 2))
        ttk.Button(layout_anchor_buttons, text="Auto Detect", command=self.layout_anchor_auto_detect).grid(row=0, column=0, padx=3)
        ttk.Button(layout_anchor_buttons, text="Set Manual", command=self.layout_anchor_set_manual).grid(row=0, column=1, padx=3)
        ttk.Button(layout_anchor_buttons, text="Align to Reference", command=self.layout_anchor_align_to_reference).grid(row=0, column=2, padx=3)

        self.layout_sprite_options = ttk.LabelFrame(layout_options, text="Spritesheet")
        self.layout_sprite_options.grid(row=10, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
        ttk.Label(self.layout_sprite_options, text="Padding").grid(row=0, column=0, sticky="w", padx=6, pady=2)
        self.layout_sprite_padding_combo = ttk.Combobox(
            self.layout_sprite_options,
            textvariable=self.sprite_padding_mode_var,
            values=["Fixed", "Frame width", "Frame height"],
            state="readonly",
            width=12,
        )
        self.layout_sprite_padding_combo.grid(row=0, column=1, sticky="w", padx=6, pady=2)
        self.layout_sprite_padding_spin = ttk.Spinbox(self.layout_sprite_options, from_=0, to=128, textvariable=self.sprite_padding_var, width=6)
        self.layout_sprite_padding_spin.grid(row=0, column=2, sticky="w", padx=6, pady=2)

        self.layout_tile_options = ttk.LabelFrame(layout_options, text="Tilemap")
        self.layout_tile_options.grid(row=11, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
        ttk.Label(self.layout_tile_options, text="Tile size mode").grid(row=0, column=0, sticky="w", padx=6, pady=2)
        self.layout_tile_size_mode_combo = ttk.Combobox(
            self.layout_tile_options,
            textvariable=self.tile_size_mode_var,
            values=["Force", "Per-tile"],
            state="readonly",
            width=10,
        )
        self.layout_tile_size_mode_combo.grid(row=0, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(self.layout_tile_options, text="Tile size").grid(row=0, column=2, sticky="w", padx=6, pady=2)
        self.layout_tile_size_combo = ttk.Combobox(
            self.layout_tile_options,
            textvariable=self.tile_size_var,
            values=[8, 16, 32, 64],
            state="readonly",
            width=6,
        )
        self.layout_tile_size_combo.grid(row=0, column=3, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(self.layout_tile_options, text="Export metadata (CSV)", variable=self.tile_export_meta_var).grid(row=1, column=0, columnspan=4, sticky="w", padx=6, pady=2)

        layout_recent_frame = ttk.LabelFrame(layout_options, text="Recent JSON")
        layout_recent_frame.grid(row=12, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
        layout_recent_frame.columnconfigure(1, weight=1)
        ttk.Label(layout_recent_frame, text="File").grid(row=0, column=0, sticky="w", padx=6, pady=2)
        self.layout_recent_combo = ttk.Combobox(
            layout_recent_frame,
            textvariable=self.layout_recent_json_var,
            state="readonly",
            width=26,
        )
        self.layout_recent_combo.grid(row=0, column=1, sticky="ew", padx=6, pady=2)
        ttk.Button(layout_recent_frame, text="Load", command=self.layout_load_recent_json).grid(row=0, column=2, padx=6, pady=2)

        layout_action_frame = ttk.Frame(layout_options)
        layout_action_frame.grid(row=13, column=0, columnspan=2, sticky="ew", padx=6, pady=(8, 6))
        ttk.Button(layout_action_frame, text="Load Layout", command=self.layout_load).grid(row=0, column=0, padx=3)
        ttk.Button(layout_action_frame, text="Save Layout", command=self.layout_save).grid(row=0, column=1, padx=3)
        self.layout_export_button = ttk.Button(layout_action_frame, text="Export", command=self.layout_export)
        self.layout_export_button.grid(row=0, column=2, padx=3)
        self.layout_status_label = ttk.Label(layout_action_frame, text="")
        self.layout_status_label.grid(row=0, column=3, sticky="w", padx=(8, 0))

        self.sprite_layout_combo.bind("<<ComboboxSelected>>", lambda _e: self.update_sprite_layout_controls())
        self.sprite_padding_mode_combo.bind("<<ComboboxSelected>>", lambda _e: self.update_sprite_padding_controls())
        self.tile_layout_combo.bind("<<ComboboxSelected>>", lambda _e: self.update_tile_layout_controls())
        self.layout_type_combo.bind("<<ComboboxSelected>>", lambda _e: self.update_layout_type_controls())
        self.layout_folder_combo.bind("<<ComboboxSelected>>", lambda _e: self.refresh_layout_files())
        self.layout_mode_combo.bind("<<ComboboxSelected>>", lambda _e: self.layout_redraw())
        self.layout_columns_spin.bind("<FocusOut>", lambda _e: self.layout_redraw())
        self.layout_rows_spin.bind("<FocusOut>", lambda _e: self.layout_redraw())
        self.layout_tile_size_mode_combo.bind("<<ComboboxSelected>>", lambda _e: self.update_layout_type_controls())
        self.layout_tile_size_combo.bind("<<ComboboxSelected>>", lambda _e: self.layout_redraw())
        self.layout_sprite_padding_combo.bind(
            "<<ComboboxSelected>>",
            lambda _e: (self.update_sprite_padding_controls(), self.layout_redraw()),
        )
        self.layout_sprite_padding_spin.bind("<FocusOut>", lambda _e: self.layout_redraw())
        self.layout_anchor_mode_combo.bind("<<ComboboxSelected>>", lambda _e: self.layout_update_anchor_settings())
        self.layout_anchor_source_combo.bind("<<ComboboxSelected>>", lambda _e: self.layout_update_anchor_settings())
        self.layout_anchor_type_combo.bind("<<ComboboxSelected>>", lambda _e: self.layout_update_anchor_settings())
        self.layout_anchor_padding_spin.bind("<FocusOut>", lambda _e: self.layout_update_anchor_settings())
        self.layout_anchor_x_entry.bind("<Return>", lambda _e: self.layout_anchor_set_manual())
        self.layout_anchor_y_entry.bind("<Return>", lambda _e: self.layout_anchor_set_manual())
        self.layout_listbox.bind("<<ListboxSelect>>", lambda _e: self.layout_maybe_set_output_name())
        self.layout_pos_x_entry.bind("<Return>", lambda _e: self.layout_apply_position())
        self.layout_pos_y_entry.bind("<Return>", lambda _e: self.layout_apply_position())
        self.layout_canvas.bind("<ButtonPress-1>", self.layout_on_press)
        self.layout_canvas.bind("<B1-Motion>", self.layout_on_drag)
        self.layout_canvas.bind("<ButtonRelease-1>", self.layout_on_release)
        self.layout_canvas.bind("<MouseWheel>", self.layout_on_mousewheel)
        self.layout_canvas.bind("<Up>", lambda e: self.layout_nudge(0, -1, e))
        self.layout_canvas.bind("<Down>", lambda e: self.layout_nudge(0, 1, e))
        self.layout_canvas.bind("<Left>", lambda e: self.layout_nudge(-1, 0, e))
        self.layout_canvas.bind("<Right>", lambda e: self.layout_nudge(1, 0, e))
        self.layout_layers_listbox.bind("<<ListboxSelect>>", self.layout_on_layer_select)
        self.layout_layers_listbox.bind("<ButtonPress-1>", self.layout_layers_on_press)
        self.layout_layers_listbox.bind("<B1-Motion>", self.layout_layers_on_drag)
        self.layout_layers_listbox.bind("<ButtonRelease-1>", self.layout_layers_on_release)
        self.setup_layout_dnd()
        self.root.bind_all("<Control-z>", self.on_global_undo)

        split_main = ttk.Frame(self.main_split_tab, padding=10)
        split_main.grid(row=0, column=0, sticky="nsew")
        self.main_split_tab.columnconfigure(0, weight=1)
        self.main_split_tab.rowconfigure(0, weight=1)
        split_main.columnconfigure(1, weight=1)

        ttk.Label(split_main, text="Spritesheet File").grid(row=0, column=0, sticky="w")
        ttk.Entry(split_main, textvariable=self.split_sheet_path_var).grid(row=0, column=1, sticky="ew", padx=5)
        ttk.Button(split_main, text="Browse", command=self.pick_split_sheet).grid(row=0, column=2)

        ttk.Label(split_main, text="Output Folder (auto)").grid(row=1, column=0, sticky="w", pady=(4, 0))
        self.split_output_entry = ttk.Entry(split_main, textvariable=self.split_output_dir_var, state="readonly")
        self.split_output_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=(4, 0))

        ttk.Label(split_main, text="Base Name").grid(row=2, column=0, sticky="w", pady=(4, 0))
        self.split_base_name_entry = ttk.Entry(split_main, textvariable=self.split_base_name_var)
        self.split_base_name_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=(4, 0))
        ttk.Button(split_main, text="Use Sheet Name", command=self.split_use_sheet_name).grid(row=2, column=2, pady=(4, 0))

        ttk.Label(split_main, text="Tip: drag grid lines on the canvas to resize cell width/height.").grid(row=3, column=0, columnspan=3, sticky="w", pady=(4, 0))

        split_body = ttk.Frame(split_main)
        split_body.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=8)
        split_body.columnconfigure(0, weight=1)
        split_body.rowconfigure(0, weight=1)
        split_main.rowconfigure(4, weight=1)

        split_canvas_frame = ttk.LabelFrame(split_body, text="Spritesheet")
        split_canvas_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        split_canvas_frame.rowconfigure(1, weight=1)
        split_canvas_frame.columnconfigure(0, weight=1)
        split_canvas_header = ttk.Frame(split_canvas_frame)
        split_canvas_header.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 0))
        split_canvas_header.columnconfigure(0, weight=1)
        ttk.Button(split_canvas_header, text="Undo", command=self.split_undo).grid(row=0, column=1, padx=3)
        ttk.Button(split_canvas_header, text="+", command=lambda: self.split_set_zoom(self.split_zoom * 1.1)).grid(
            row=0, column=2, padx=3
        )
        ttk.Button(split_canvas_header, text="-", command=lambda: self.split_set_zoom(self.split_zoom * 0.9)).grid(
            row=0, column=3, padx=3
        )
        self.split_canvas = tk.Canvas(split_canvas_frame, width=520, height=360, highlightthickness=1)
        self.split_canvas.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        split_canvas_scroll_y = ttk.Scrollbar(split_canvas_frame, orient="vertical", command=self.split_canvas.yview)
        split_canvas_scroll_y.grid(row=1, column=1, sticky="ns", pady=6)
        split_canvas_scroll_x = ttk.Scrollbar(split_canvas_frame, orient="horizontal", command=self.split_canvas.xview)
        split_canvas_scroll_x.grid(row=2, column=0, sticky="ew", padx=6)
        self.split_canvas.configure(yscrollcommand=split_canvas_scroll_y.set, xscrollcommand=split_canvas_scroll_x.set)

        split_options = ttk.LabelFrame(split_body, text="Split Options")
        split_options.grid(row=0, column=1, sticky="nsew")
        split_options.columnconfigure(1, weight=1)
        ttk.Button(split_options, text="Auto Detect Cells", command=self.split_auto_detect_grid).grid(row=0, column=0, columnspan=2, sticky="w", padx=6, pady=(6, 2))
        ttk.Label(split_options, text="Cell width").grid(row=1, column=0, sticky="w", padx=6, pady=2)
        self.split_cell_w_spin = ttk.Spinbox(split_options, from_=1, to=4096, textvariable=self.split_cell_w_var, width=6)
        self.split_cell_w_spin.grid(row=1, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(split_options, text="Cell height").grid(row=2, column=0, sticky="w", padx=6, pady=2)
        self.split_cell_h_spin = ttk.Spinbox(split_options, from_=1, to=4096, textvariable=self.split_cell_h_var, width=6)
        self.split_cell_h_spin.grid(row=2, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(split_options, text="Columns (0=auto)").grid(row=3, column=0, sticky="w", padx=6, pady=2)
        self.split_columns_spin = ttk.Spinbox(split_options, from_=0, to=512, textvariable=self.split_columns_var, width=6)
        self.split_columns_spin.grid(row=3, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(split_options, text="Rows (0=auto)").grid(row=4, column=0, sticky="w", padx=6, pady=2)
        self.split_rows_spin = ttk.Spinbox(split_options, from_=0, to=512, textvariable=self.split_rows_var, width=6)
        self.split_rows_spin.grid(row=4, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(split_options, text="Padding X").grid(row=5, column=0, sticky="w", padx=6, pady=2)
        self.split_pad_x_spin = ttk.Spinbox(split_options, from_=0, to=1024, textvariable=self.split_pad_x_var, width=6)
        self.split_pad_x_spin.grid(row=5, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(split_options, text="Padding Y").grid(row=6, column=0, sticky="w", padx=6, pady=2)
        self.split_pad_y_spin = ttk.Spinbox(split_options, from_=0, to=1024, textvariable=self.split_pad_y_var, width=6)
        self.split_pad_y_spin.grid(row=6, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(split_options, text="Offset X").grid(row=7, column=0, sticky="w", padx=6, pady=2)
        self.split_offset_x_spin = ttk.Spinbox(split_options, from_=0, to=4096, textvariable=self.split_offset_x_var, width=6)
        self.split_offset_x_spin.grid(row=7, column=1, sticky="w", padx=6, pady=2)
        ttk.Label(split_options, text="Offset Y").grid(row=8, column=0, sticky="w", padx=6, pady=2)
        self.split_offset_y_spin = ttk.Spinbox(split_options, from_=0, to=4096, textvariable=self.split_offset_y_var, width=6)
        self.split_offset_y_spin.grid(row=8, column=1, sticky="w", padx=6, pady=2)
        ttk.Checkbutton(split_options, text="Show grid lines", variable=self.split_show_grid_var, command=self.split_redraw).grid(row=9, column=0, columnspan=2, sticky="w", padx=6, pady=2)

        split_select_frame = ttk.Frame(split_options)
        split_select_frame.grid(row=10, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
        ttk.Button(split_select_frame, text="Select All", command=self.split_select_all).grid(row=0, column=0, padx=3)
        ttk.Button(split_select_frame, text="Clear Selection", command=self.split_clear_selection).grid(row=0, column=1, padx=3)
        ttk.Button(split_select_frame, text="Save Selection", command=self.split_save_selection).grid(row=0, column=2, padx=3)

        split_export_frame = ttk.Frame(split_options)
        split_export_frame.grid(row=11, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
        ttk.Button(split_export_frame, text="Export Selected", command=self.split_export_selected).grid(row=0, column=0, padx=3)
        ttk.Button(split_export_frame, text="Export All", command=self.split_export_all).grid(row=0, column=1, padx=3)

        self.split_status_label = ttk.Label(split_options, text="")
        self.split_status_label.grid(row=12, column=0, columnspan=2, sticky="w", padx=6, pady=(4, 6))

        self.split_cell_w_spin.bind("<FocusOut>", lambda _e: self.split_redraw())
        self.split_cell_h_spin.bind("<FocusOut>", lambda _e: self.split_redraw())
        self.split_columns_spin.bind("<FocusOut>", lambda _e: self.split_redraw())
        self.split_rows_spin.bind("<FocusOut>", lambda _e: self.split_redraw())
        self.split_pad_x_spin.bind("<FocusOut>", lambda _e: self.split_redraw())
        self.split_pad_y_spin.bind("<FocusOut>", lambda _e: self.split_redraw())
        self.split_offset_x_spin.bind("<FocusOut>", lambda _e: self.split_redraw())
        self.split_offset_y_spin.bind("<FocusOut>", lambda _e: self.split_redraw())
        self.split_base_name_entry.bind("<FocusOut>", lambda _e: self.split_update_output_dir())
        self.split_base_name_entry.bind("<Return>", lambda _e: self.split_update_output_dir())
        self.split_canvas.bind("<ButtonPress-1>", self.split_on_press)
        self.split_canvas.bind("<B1-Motion>", self.split_on_drag)
        self.split_canvas.bind("<ButtonRelease-1>", self.split_on_release)
        self.split_canvas.bind("<MouseWheel>", self.split_on_mousewheel)

        self.update_color_mode()
        self.update_sprite_layout_controls()
        self.update_sprite_padding_controls()
        self.update_tile_layout_controls()
        self.update_layout_type_controls()
        self.layout_redraw()
        self.split_redraw()
        self.bind_color_swatch_updates()
        self.setup_tooltips(
            transparent_radio,
            fill_radio,
            replace_radio,
            process_all_btn,
            skip_existing_btn,
            skip_files_btn,
            dry_run_btn,
            tooltips_btn,
            self.preview_button,
            self.scan_prefixes_button,
            self.keep_prefixes_check,
        )
        self.theme_var.trace_add("write", lambda *_: self.on_theme_change())
        self.root.bind("<Configure>", self.on_root_configure)

    def apply_theme(self, theme):
        if theme not in THEMES:
            theme = "Light"

        if theme == "Dark":
            self.style.theme_use("clam")
            colors = {
                "bg": "#1f1f1f",
                "fg": "#a7a7a7",
                "entry_bg": "#1f1f1f",
                "entry_fg": "#a7a7a7",
                "border": "#5e5e5e",
                "button_bg": "#202020",
                "button_fg": "#a7a7a7",
                "log_bg": "#1b1b1b",
                "log_fg": "#a7a7a7",
            }
        else:
            self.style.theme_use("clam")
            colors = {
                "bg": "#f3f3f3",
                "fg": "#1f1f1f",
                "entry_bg": "#ffffff",
                "entry_fg": "#1f1f1f",
                "border": "#b0b0b0",
                "button_bg": "#f6f6f6",
                "button_fg": "#1f1f1f",
                "log_bg": "#ffffff",
                "log_fg": "#1f1f1f",
            }

        self.theme_colors = colors
        self.root.configure(background=colors["bg"])

        self.style.configure("TFrame", background=colors["bg"])
        self.style.configure("TLabel", background=colors["bg"], foreground=colors["fg"])
        self.style.configure("TLabelframe", background=colors["bg"], foreground=colors["fg"])
        self.style.configure("TLabelframe.Label", background=colors["bg"], foreground=colors["fg"])
        self.style.configure("TCheckbutton", background=colors["bg"], foreground=colors["fg"])
        self.style.configure("TRadiobutton", background=colors["bg"], foreground=colors["fg"])
        self.style.configure("TButton", background=colors["button_bg"], foreground=colors["button_fg"])
        self.style.configure("TEntry", fieldbackground=colors["entry_bg"], foreground=colors["entry_fg"])
        self.style.configure("TCombobox", fieldbackground=colors["entry_bg"], foreground=colors["entry_fg"])
        self.style.map("TCombobox", fieldbackground=[("readonly", colors["entry_bg"])])
        self.style.configure(
            "Theme.TCombobox",
            fieldbackground=colors["entry_bg"],
            foreground=colors["entry_fg"],
            background=colors["entry_bg"],
            selectbackground=colors["entry_bg"],
            selectforeground=colors["entry_fg"],
        )
        self.style.map(
            "Theme.TCombobox",
            fieldbackground=[("readonly", colors["entry_bg"]), ("focus", colors["entry_bg"]), ("active", colors["entry_bg"])],
            selectbackground=[("readonly", colors["entry_bg"]), ("focus", colors["entry_bg"]), ("active", colors["entry_bg"])],
            selectforeground=[("readonly", colors["entry_fg"]), ("focus", colors["entry_fg"]), ("active", colors["entry_fg"])],
        )
        self.style.configure("TProgressbar", background="#3d7bfd", troughcolor=colors["entry_bg"])
        self.style.configure("TScrollbar", background=colors["entry_bg"])

        self.style.configure("NoFocus.TLabelframe", bordercolor=colors["border"], lightcolor=colors["border"], darkcolor=colors["border"])
        self.style.map(
            "NoFocus.TLabelframe",
            bordercolor=[("focus", colors["border"])],
            lightcolor=[("focus", colors["border"])],
            darkcolor=[("focus", colors["border"])],
        )

        self.style.configure(
            "Folder.TCombobox",
            fieldbackground=colors["entry_bg"],
            foreground=colors["entry_fg"],
            background=colors["entry_bg"],
            selectbackground=colors["entry_bg"],
            selectforeground=colors["entry_fg"],
        )
        self.style.map(
            "Folder.TCombobox",
            fieldbackground=[("readonly", colors["entry_bg"]), ("focus", colors["entry_bg"]), ("active", colors["entry_bg"])],
            selectbackground=[("readonly", colors["entry_bg"]), ("focus", colors["entry_bg"]), ("active", colors["entry_bg"])],
            selectforeground=[("readonly", colors["entry_fg"]), ("focus", colors["entry_fg"]), ("active", colors["entry_fg"])],
        )

        if hasattr(self, "log"):
            self.log.configure(background=colors["log_bg"], foreground=colors["log_fg"], insertbackground=colors["log_fg"])
        if hasattr(self, "sheet_log"):
            self.sheet_log.configure(background=colors["log_bg"], foreground=colors["log_fg"], insertbackground=colors["log_fg"])
        if hasattr(self, "tile_log"):
            self.tile_log.configure(background=colors["log_bg"], foreground=colors["log_fg"], insertbackground=colors["log_fg"])
        if hasattr(self, "layout_canvas"):
            self.layout_canvas.configure(background=colors["entry_bg"], highlightbackground=colors["border"], highlightcolor=colors["border"])
        if hasattr(self, "split_canvas"):
            self.split_canvas.configure(background=colors["entry_bg"], highlightbackground=colors["border"], highlightcolor=colors["border"])
        if hasattr(self, "layout_listbox"):
            self.layout_listbox.configure(background=colors["entry_bg"], foreground=colors["entry_fg"])
        if hasattr(self, "layout_layers_listbox"):
            self.layout_layers_listbox.configure(background=colors["entry_bg"], foreground=colors["entry_fg"])
        if hasattr(self, "folders_box"):
            self.folders_box.configure(background=colors["bg"], highlightbackground=colors["border"], highlightcolor=colors["border"], highlightthickness=1)
        if hasattr(self, "folders_canvas"):
            self.folders_canvas.configure(background=colors["bg"], highlightthickness=0)
        self.update_swatches()

    def on_theme_change(self):
        self.apply_theme(self.theme_var.get())
        self.root.after(0, self.root.focus_set)
        self.write_settings_or_error()

    def on_root_configure(self, event):
        if event.widget is not self.root:
            return
        if event.width < 200 or event.height < 200:
            return
        self.window_size = (event.width, event.height)
        if self._resize_save_after_id:
            self.root.after_cancel(self._resize_save_after_id)
        self._resize_save_after_id = self.root.after(400, self.write_settings_or_error)

    def on_global_undo(self, _event=None):
        if not hasattr(self, "main_notebook"):
            return
        try:
            tab_text = self.main_notebook.tab(self.main_notebook.select(), "text")
        except Exception:
            return
        if tab_text == "Layout Editor (Assemble)":
            self.layout_undo()
        elif tab_text == "Layout Editor (Split)":
            self.split_undo()

    def update_color_mode(self):
        mode = self.mode_var.get()
        if mode == "transparent":
            self.color_list_entry.configure(state="normal")
            self.fill_color_entry.configure(state="disabled")
            self.fill_shadows_check.configure(state="disabled")
            self.replace_from_entry.configure(state="disabled")
            self.replace_to_entry.configure(state="disabled")
        elif mode == "fill":
            self.color_list_entry.configure(state="normal")
            self.fill_color_entry.configure(state="normal")
            self.fill_shadows_check.configure(state="normal")
            self.replace_from_entry.configure(state="disabled")
            self.replace_to_entry.configure(state="disabled")
        else:
            self.color_list_entry.configure(state="disabled")
            self.fill_color_entry.configure(state="disabled")
            self.fill_shadows_check.configure(state="disabled")
            self.replace_from_entry.configure(state="normal")
            self.replace_to_entry.configure(state="normal")
        self.update_swatches()

    def update_sprite_layout_controls(self):
        mode = self.sprite_layout_var.get()
        if mode == "Grid":
            self.sprite_columns_spin.configure(state="normal")
        else:
            self.sprite_columns_spin.configure(state="disabled")

    def update_sprite_padding_controls(self):
        mode = self.sprite_padding_mode_var.get()
        state = "normal" if mode == "Fixed" else "disabled"
        self.sprite_padding_spin.configure(state=state)
        if hasattr(self, "layout_sprite_padding_spin"):
            self.layout_sprite_padding_spin.configure(state=state)

    def update_tile_layout_controls(self):
        mode = self.tile_layout_var.get()
        if mode == "Grid":
            self.tile_columns_combo.configure(state="readonly")
        else:
            self.tile_columns_combo.configure(state="disabled")

    def update_layout_mode_controls(self):
        mode = self.layout_mode_var.get()
        if mode == "Grid":
            self.layout_columns_spin.configure(state="normal")
            self.layout_rows_spin.configure(state="normal")
        else:
            self.layout_columns_spin.configure(state="disabled")
            self.layout_rows_spin.configure(state="disabled")

    def update_layout_tile_size_controls(self):
        if self.tile_size_mode_var.get() == "Force":
            self.layout_tile_size_combo.configure(state="readonly")
        else:
            self.layout_tile_size_combo.configure(state="disabled")

    def update_layout_type_controls(self):
        if self.layout_type_var.get() == "Spritesheet":
            if hasattr(self, "layout_name_label"):
                self.layout_name_label.configure(text="Spritesheet name")
            self.layout_sprite_options.grid()
            self.layout_tile_options.grid_remove()
            self.layout_export_button.configure(text="Export Spritesheet")
        else:
            if hasattr(self, "layout_name_label"):
                self.layout_name_label.configure(text="Tilemap name")
            self.layout_tile_options.grid()
            self.layout_sprite_options.grid_remove()
            self.layout_export_button.configure(text="Export Tilemap")
            if not self.layout_output_var.get().strip():
                folder_path = self.layout_get_folder_path()
                if folder_path:
                    self.layout_output_var.set(os.path.basename(folder_path))
        self.update_layout_mode_controls()
        self.update_sprite_padding_controls()
        self.update_layout_tile_size_controls()
        self.layout_refresh_recent_jsons()
        self.layout_redraw()

    def layout_get_folder_path(self):
        input_root = self.input_var.get().strip()
        folder = self.layout_folder_var.get().strip()
        if not input_root or not folder:
            return ""
        if folder in (".", "./"):
            return input_root
        return os.path.join(input_root, folder)

    def layout_refresh_recent_jsons(self):
        if not hasattr(self, "layout_recent_combo"):
            return
        input_root = self.input_var.get().strip()
        if not os.path.isdir(input_root):
            self.layout_recent_combo["values"] = []
            self.layout_recent_json_var.set("")
            self.layout_recent_json_paths = []
            return
        target_dir = "sprite_sheets" if self.layout_type_var.get() == "Spritesheet" else "tilemaps"
        entries = []
        for root, _dirs, files in os.walk(input_root):
            if os.path.basename(root).lower() != target_dir:
                continue
            for file in files:
                if not file.lower().endswith(".json"):
                    continue
                path = os.path.join(root, file)
                try:
                    mtime = os.path.getmtime(path)
                except OSError:
                    mtime = 0
                entries.append((mtime, path))
        entries.sort(key=lambda item: item[0], reverse=True)
        entries = entries[:30]
        labels = []
        paths = []
        for _mtime, path in entries:
            try:
                label = os.path.relpath(path, input_root)
            except ValueError:
                label = os.path.basename(path)
            labels.append(label)
            paths.append(path)
        self.layout_recent_combo["values"] = labels
        self.layout_recent_json_paths = paths
        if labels:
            if self.layout_recent_json_var.get() not in labels:
                self.layout_recent_json_var.set(labels[0])
        else:
            self.layout_recent_json_var.set("")

    def refresh_layout_folders(self):
        input_root = self.input_var.get().strip()
        self.layout_folder_combo["values"] = []
        if not os.path.isdir(input_root):
            self.layout_folder_var.set("")
            self.layout_listbox.delete(0, tk.END)
            return
        folders = []
        for root, dirs, files in os.walk(input_root):
            dirs[:] = [d for d in dirs if d not in ("sprite_sheets", "tilemaps")]
            if any(file.lower().endswith(".png") for file in files):
                rel = os.path.relpath(root, input_root)
                folders.append("." if rel == "." else rel)
        folders.sort(key=str.lower)
        self.layout_folder_combo["values"] = folders
        if self.layout_folder_var.get() not in folders:
            self.layout_folder_var.set(folders[0] if folders else "")
        self.refresh_layout_files()
        self.layout_refresh_recent_jsons()

    def refresh_layout_files(self):
        folder_path = self.layout_get_folder_path()
        self.layout_listbox.delete(0, tk.END)
        if not folder_path or not os.path.isdir(folder_path):
            return
        files = [f for f in os.listdir(folder_path) if f.lower().endswith(".png")]
        files.sort(key=str.lower)
        for file in files:
            self.layout_listbox.insert(tk.END, file)
        self.layout_clear_canvas()
        if self.layout_type_var.get() == "Tilemap" and not self.layout_output_var.get().strip():
            self.layout_output_var.set(os.path.basename(folder_path))
        self.layout_refresh_recent_jsons()

    def layout_maybe_set_output_name(self):
        if self.layout_output_var.get().strip():
            return
        selections = self.layout_listbox.curselection()
        if not selections:
            return
        file = self.layout_listbox.get(selections[0])
        if self.layout_type_var.get() == "Spritesheet":
            self.layout_output_var.set(extract_group_prefix(file))
        else:
            folder_path = self.layout_get_folder_path()
            if folder_path:
                self.layout_output_var.set(os.path.basename(folder_path))

    def layout_init_layers(self):
        if self.layout_layers:
            return
        self.layout_create_layer("Layer 1")

    def layout_create_layer(self, name=None):
        self.layout_layer_counter += 1
        layer_id = f"layer_{self.layout_layer_counter}"
        layer_name = name or f"Layer {self.layout_layer_counter}"
        layer = {
            "id": layer_id,
            "name": layer_name,
            "visible": True,
            "anchor_mode": "Off",
            "anchor_source": "Auto",
            "anchor_type": "Bottom",
            "anchor_padding": 2,
            "anchor_x": None,
            "anchor_y": None,
        }
        self.layout_layers.append(layer)
        self.layout_active_layer_id = layer_id
        if self.layout_reference_layer_id is None:
            self.layout_reference_layer_id = layer_id
        self.layout_refresh_layers_list()
        self.layout_apply_layer_order()
        self.layout_update_anchor_fields()
        return layer_id

    def layout_layer_create(self):
        self.layout_create_layer()

    def layout_layer_rename(self):
        layer = self.layout_get_active_layer()
        if not layer:
            return
        name = simpledialog.askstring("Rename layer", "New layer name:", initialvalue=layer["name"])
        if not name:
            return
        layer["name"] = name.strip()
        self.layout_refresh_layers_list()

    def layout_layer_delete(self):
        if not self.layout_layers:
            return
        if len(self.layout_layers) == 1:
            messagebox.showinfo("Layers", "At least one layer is required.")
            return
        layer = self.layout_get_active_layer()
        if not layer:
            return
        self.layout_layers = [l for l in self.layout_layers if l["id"] != layer["id"]]
        fallback = self.layout_layers[0]
        self.layout_active_layer_id = fallback["id"]
        if self.layout_reference_layer_id == layer["id"]:
            self.layout_reference_layer_id = fallback["id"]
        for item in self.layout_items.values():
            if item.get("layer_id") == layer["id"]:
                item["layer_id"] = fallback["id"]
        self.layout_refresh_layers_list()
        self.layout_apply_layer_order()
        self.layout_redraw()
        self.layout_update_anchor_fields()

    def layout_layer_toggle_visibility(self):
        layer = self.layout_get_active_layer()
        if not layer:
            return
        layer["visible"] = not layer.get("visible", True)
        self.layout_refresh_layers_list()
        self.layout_redraw()

    def layout_layer_set_reference(self):
        layer = self.layout_get_active_layer()
        if not layer:
            return
        self.layout_reference_layer_id = layer["id"]
        self.layout_refresh_layers_list()

    def layout_layer_assign_dialog(self):
        if not self.layout_selected_ids:
            messagebox.showinfo("Assign frames", "Select one or more frames first.")
            return
        if not self.layout_layers:
            self.layout_create_layer("Layer 1")
        dialog = tk.Toplevel(self.root)
        dialog.title("Assign to layer")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.columnconfigure(0, weight=1)

        ttk.Label(dialog, text="Layer").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))
        layer_list = tk.Listbox(dialog, height=6, exportselection=False)
        layer_list.grid(row=1, column=0, sticky="nsew", padx=10)
        for layer in self.layout_layers:
            layer_list.insert("end", layer["name"])
        if self.layout_active_layer_id:
            active_index = self.layout_get_layer_index(self.layout_active_layer_id)
            if active_index is not None:
                layer_list.selection_set(active_index)
                layer_list.activate(active_index)
        inherit_values = {
            self.layout_items[item_id].get("anchor_inherit", True)
            for item_id in self.layout_selected_ids
            if item_id in self.layout_items
        }
        inherit_var = tk.BooleanVar(value=True)
        if inherit_values:
            inherit_var.set(len(inherit_values) == 1 and True in inherit_values)
        ttk.Checkbutton(dialog, text="Inherit layer anchors", variable=inherit_var).grid(
            row=2, column=0, sticky="w", padx=10, pady=(6, 4)
        )

        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=3, column=0, sticky="e", padx=10, pady=(4, 10))

        def assign_to_selected():
            if not layer_list.curselection():
                return
            index = layer_list.curselection()[0]
            target_layer = self.layout_layers[index]
            for item_id in list(self.layout_selected_ids):
                item = self.layout_items.get(item_id)
                if not item:
                    continue
                item["layer_id"] = target_layer["id"]
                item["anchor_inherit"] = bool(inherit_var.get())
            self.layout_active_layer_id = target_layer["id"]
            self.layout_refresh_layers_list()
            self.layout_redraw()
            dialog.destroy()

        def create_layer():
            name = simpledialog.askstring("New layer", "Layer name:")
            if not name:
                return
            new_id = self.layout_create_layer(name.strip())
            new_index = self.layout_get_layer_index(new_id)
            if new_index is not None:
                layer_list.insert("end", name.strip())
                layer_list.selection_clear(0, "end")
                layer_list.selection_set(new_index)
                layer_list.activate(new_index)

        ttk.Button(button_frame, text="New Layer", command=create_layer).grid(row=0, column=0, padx=4)
        ttk.Button(button_frame, text="Assign", command=assign_to_selected).grid(row=0, column=1, padx=4)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).grid(row=0, column=2, padx=4)

    def layout_get_layer_index(self, layer_id):
        for index, layer in enumerate(self.layout_layers):
            if layer["id"] == layer_id:
                return index
        return None

    def layout_get_layer_by_id(self, layer_id):
        for layer in self.layout_layers:
            if layer["id"] == layer_id:
                return layer
        return None

    def layout_get_active_layer(self):
        if not self.layout_layers:
            return None
        if not self.layout_active_layer_id:
            self.layout_active_layer_id = self.layout_layers[0]["id"]
        return self.layout_get_layer_by_id(self.layout_active_layer_id)

    def layout_refresh_layers_list(self):
        if not hasattr(self, "layout_layers_listbox"):
            return
        self.layout_layers_listbox.delete(0, "end")
        for layer in self.layout_layers:
            visible = "[x]" if layer.get("visible", True) else "[ ]"
            ref = "*" if layer["id"] == self.layout_reference_layer_id else " "
            self.layout_layers_listbox.insert("end", f"{visible} {ref} {layer['name']}")
        if self.layout_active_layer_id:
            index = self.layout_get_layer_index(self.layout_active_layer_id)
            if index is not None:
                self.layout_layers_listbox.selection_set(index)
                self.layout_layers_listbox.activate(index)

    def layout_layers_on_press(self, event):
        index = self.layout_layers_listbox.nearest(event.y)
        if index >= 0:
            self.layout_layer_drag_index = index
            self.layout_layers_listbox.selection_clear(0, "end")
            self.layout_layers_listbox.selection_set(index)
            self.layout_layers_listbox.activate(index)
            layer = self.layout_layers[index]
            self.layout_active_layer_id = layer["id"]
            self.layout_update_anchor_fields()

    def layout_layers_on_drag(self, event):
        if self.layout_layer_drag_index is None:
            return
        target = self.layout_layers_listbox.nearest(event.y)
        if target < 0 or target == self.layout_layer_drag_index:
            return
        layer = self.layout_layers.pop(self.layout_layer_drag_index)
        self.layout_layers.insert(target, layer)
        self.layout_layer_drag_index = target
        self.layout_refresh_layers_list()
        self.layout_apply_layer_order()

    def layout_layers_on_release(self, _event):
        self.layout_layer_drag_index = None

    def layout_on_layer_select(self, _event=None):
        selection = self.layout_layers_listbox.curselection()
        if not selection:
            return
        layer = self.layout_layers[selection[0]]
        self.layout_active_layer_id = layer["id"]
        self.layout_update_anchor_fields()

    def layout_apply_layer_order(self):
        if not hasattr(self, "layout_canvas"):
            return
        order_map = {layer["id"]: idx for idx, layer in enumerate(self.layout_layers)}
        sorted_items = sorted(
            self.layout_items.items(),
            key=lambda item: (
                order_map.get(item[1].get("layer_id"), 0),
                item[1].get("order", 0),
            ),
        )
        for item_id, _item in sorted_items:
            self.layout_canvas.tag_raise(item_id)
        for item_id in self.layout_selected_ids:
            item = self.layout_items.get(item_id)
            if item and item.get("rect_id"):
                self.layout_canvas.tag_raise(item["rect_id"])

    def layout_item_is_visible(self, item):
        if not item:
            return False
        layer = self.layout_get_layer_by_id(item.get("layer_id"))
        if layer and not layer.get("visible", True):
            return False
        return item.get("visible", True)

    def layout_get_ordered_items(self, include_hidden=False):
        order_map = {layer["id"]: idx for idx, layer in enumerate(self.layout_layers)}
        ordered = sorted(
            self.layout_items.values(),
            key=lambda item: (
                order_map.get(item.get("layer_id"), 0),
                item.get("order", 0),
            ),
        )
        if include_hidden:
            return ordered
        return [item for item in ordered if self.layout_item_is_visible(item)]

    def layout_push_undo(self):
        if not self.layout_items:
            return
        snapshot = {item_id: (item["x"], item["y"]) for item_id, item in self.layout_items.items()}
        if self.layout_undo_stack and self.layout_undo_stack[-1] == snapshot:
            return
        self.layout_undo_stack.append(snapshot)
        if len(self.layout_undo_stack) > 10:
            self.layout_undo_stack.pop(0)

    def layout_undo(self, _event=None):
        if not self.layout_undo_stack:
            return
        snapshot = self.layout_undo_stack.pop()
        for item_id, pos in snapshot.items():
            item = self.layout_items.get(item_id)
            if not item:
                continue
            item["x"], item["y"] = pos
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_nudge(self, dx, dy, event=None):
        if not self.layout_selected_ids:
            return
        step = 4 if event and (event.state & 0x0001) else 1
        dx *= step
        dy *= step
        self.layout_push_undo()
        for item_id in self.layout_selected_ids:
            item = self.layout_items.get(item_id)
            if not item:
                continue
            item["x"] += dx
            item["y"] += dy
        if self.layout_snap_var.get() and self.layout_mode_var.get() == "Grid":
            self.layout_snap_selected(record_undo=False)
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_toggle_guides(self):
        self.layout_guides_enabled = not self.layout_guides_enabled
        if self.layout_guides_enabled and not (self.layout_guides["h"] or self.layout_guides["v"]):
            cell_w, cell_h, pad = self.layout_get_cell_and_padding()
            cols = max(1, int(self.layout_grid_columns_var.get()))
            rows = int(self.layout_grid_rows_var.get())
            max_x = max((item["x"] + item["width"]) for item in self.layout_items.values()) if self.layout_items else 0
            max_y = max((item["y"] + item["height"]) for item in self.layout_items.values()) if self.layout_items else 0
            if self.layout_mode_var.get() == "Grid":
                step_x = cell_w + pad
                step_y = cell_h + pad
                if rows <= 0 and step_x > 0 and step_y > 0 and self.layout_items:
                    row_indices = [int(item["y"] // step_y) for item in self.layout_items.values()]
                    rows = max(row_indices) + 1 if row_indices else 1
                if rows <= 0:
                    rows = 1
                grid_w = (cols * cell_w) + (pad * max(0, cols - 1))
                grid_h = (rows * cell_h) + (pad * max(0, rows - 1))
            else:
                grid_w = max_x
                grid_h = max_y
            center_x = grid_w / 2 if grid_w else 0
            center_y = grid_h / 2 if grid_h else 0
            self.layout_guides["v"] = [center_x]
            self.layout_guides["h"] = [center_y]
        self.layout_redraw()

    def layout_guide_hit(self, canvas_x, canvas_y):
        if not self.layout_guides_enabled:
            return None
        x = canvas_x / self.layout_zoom
        y = canvas_y / self.layout_zoom
        threshold = self.layout_guide_threshold / self.layout_zoom
        for idx, gx in enumerate(self.layout_guides.get("v", [])):
            if abs(x - gx) <= threshold:
                return ("v", idx)
        for idx, gy in enumerate(self.layout_guides.get("h", [])):
            if abs(y - gy) <= threshold:
                return ("h", idx)
        return None

    def layout_get_anchor_offset(self, item):
        anchor = item.get("anchor")
        if anchor:
            return anchor["x"], anchor["y"]
        return item["width"] / 2, item["height"]

    def layout_apply_guide_snap(self, anchor_x, anchor_y):
        snap_x = anchor_x
        snap_y = anchor_y
        threshold = self.layout_guide_threshold
        for gx in self.layout_guides.get("v", []):
            if abs(anchor_x - gx) <= threshold:
                snap_x = gx
                break
        for gy in self.layout_guides.get("h", []):
            if abs(anchor_y - gy) <= threshold:
                snap_y = gy
                break
        return snap_x, snap_y

    def layout_update_position_fields(self):
        if not self.layout_selected_ids:
            self.layout_pos_x_var.set("")
            self.layout_pos_y_var.set("")
            self.layout_update_anchor_fields()
            return
        xs = {int(self.layout_items[item_id]["x"]) for item_id in self.layout_selected_ids}
        ys = {int(self.layout_items[item_id]["y"]) for item_id in self.layout_selected_ids}
        self.layout_pos_x_var.set(str(xs.pop()) if len(xs) == 1 else "")
        self.layout_pos_y_var.set(str(ys.pop()) if len(ys) == 1 else "")
        self.layout_update_anchor_fields()

    def layout_update_anchor_fields(self):
        layer = self.layout_get_active_layer()
        if not layer or self.layout_anchor_ui_lock:
            return
        self.layout_anchor_ui_lock = True
        try:
            self.layout_anchor_mode_var.set(layer.get("anchor_mode", "Off"))
            self.layout_anchor_source_var.set(layer.get("anchor_source", "Auto"))
            self.layout_anchor_type_var.set(layer.get("anchor_type", "Bottom"))
            self.layout_anchor_padding_var.set(int(layer.get("anchor_padding", 2)))
            anchor_x = ""
            anchor_y = ""
            if layer.get("anchor_mode") == "Global":
                if layer.get("anchor_x") is not None:
                    anchor_x = str(int(layer.get("anchor_x")))
                if layer.get("anchor_y") is not None:
                    anchor_y = str(int(layer.get("anchor_y")))
            elif layer.get("anchor_mode") == "Per-frame" and len(self.layout_selected_ids) == 1:
                item_id = next(iter(self.layout_selected_ids))
                item = self.layout_items.get(item_id)
                if item and item.get("anchor"):
                    anchor_x = str(int(item["anchor"]["x"]))
                    anchor_y = str(int(item["anchor"]["y"]))
            self.layout_anchor_x_var.set(anchor_x)
            self.layout_anchor_y_var.set(anchor_y)

            inherit_values = {
                self.layout_items[item_id].get("anchor_inherit", True)
                for item_id in self.layout_selected_ids
                if item_id in self.layout_items
            }
            if not inherit_values:
                self.layout_anchor_inherit_var.set(True)
            else:
                self.layout_anchor_inherit_var.set(len(inherit_values) == 1 and True in inherit_values)
        finally:
            self.layout_anchor_ui_lock = False

    def layout_update_anchor_settings(self):
        layer = self.layout_get_active_layer()
        if not layer:
            return
        if self.layout_anchor_ui_lock:
            return
        layer["anchor_mode"] = self.layout_anchor_mode_var.get()
        layer["anchor_source"] = self.layout_anchor_source_var.get()
        layer["anchor_type"] = self.layout_anchor_type_var.get()
        layer["anchor_padding"] = int(self.layout_anchor_padding_var.get())
        self.layout_update_anchor_fields()

    def layout_toggle_anchor_inherit(self):
        if not self.layout_selected_ids:
            return
        value = bool(self.layout_anchor_inherit_var.get())
        for item_id in self.layout_selected_ids:
            item = self.layout_items.get(item_id)
            if not item:
                continue
            item["anchor_inherit"] = value
        self.layout_update_anchor_fields()

    def layout_toggle_selected_visibility(self):
        if not self.layout_selected_ids:
            return
        for item_id in list(self.layout_selected_ids):
            item = self.layout_items.get(item_id)
            if not item:
                continue
            item["visible"] = not item.get("visible", True)
            if not item["visible"]:
                self.layout_clear_selection_rect(item_id)
                self.layout_selected_ids.discard(item_id)
        self.layout_redraw()

    def layout_compute_auto_anchor(self, pil, anchor_type, padding):
        if pil is None:
            return None
        alpha = pil.split()[-1]
        threshold = 5
        mask = alpha.point(lambda v: 255 if v > threshold else 0)
        bbox = mask.getbbox()
        if not bbox:
            return None
        left, top, right, bottom = bbox
        right -= 1
        bottom -= 1
        center_x = (left + right) / 2
        if anchor_type == "Top":
            anchor_y = top + padding
        elif anchor_type == "Center":
            anchor_y = (top + bottom) / 2
        else:
            anchor_y = bottom - padding
        anchor_x = center_x
        anchor_x = max(0, min(pil.width - 1, anchor_x))
        anchor_y = max(0, min(pil.height - 1, anchor_y))
        return {
            "x": int(round(anchor_x)),
            "y": int(round(anchor_y)),
        }

    def layout_get_item_anchor(self, item, layer, compute_auto=False):
        if not item.get("anchor_inherit", True):
            return item.get("anchor")
        mode = layer.get("anchor_mode", "Off")
        if mode == "Off":
            return None
        source = layer.get("anchor_source", "Auto")
        if mode == "Global":
            if source == "Manual":
                if layer.get("anchor_x") is None or layer.get("anchor_y") is None:
                    return None
                return {"x": int(layer.get("anchor_x")), "y": int(layer.get("anchor_y"))}
            if not compute_auto:
                return None
            if layer.get("anchor_x") is not None and layer.get("anchor_y") is not None:
                return {"x": int(layer.get("anchor_x")), "y": int(layer.get("anchor_y"))}
            anchor = self.layout_compute_auto_anchor(
                item["pil"],
                layer.get("anchor_type", "Bottom"),
                int(layer.get("anchor_padding", 2)),
            )
            if anchor:
                layer["anchor_x"] = anchor["x"]
                layer["anchor_y"] = anchor["y"]
            return anchor
        if mode == "Per-frame":
            if source == "Manual":
                return item.get("anchor")
            if not compute_auto:
                return item.get("anchor")
            anchor = self.layout_compute_auto_anchor(
                item["pil"],
                layer.get("anchor_type", "Bottom"),
                int(layer.get("anchor_padding", 2)),
            )
            if anchor:
                item["anchor"] = anchor
            return anchor
        return None

    def layout_anchor_auto_detect(self):
        layer = self.layout_get_active_layer()
        if not layer:
            return
        mode = layer.get("anchor_mode", "Off")
        if mode == "Off":
            messagebox.showinfo("Anchors", "Set anchor mode to Global or Per-frame first.")
            return
        target_ids = list(self.layout_selected_ids) if self.layout_selected_ids else [
            item_id for item_id, item in self.layout_items.items() if item.get("layer_id") == layer["id"]
        ]
        if not target_ids:
            messagebox.showinfo("Anchors", "No frames available to detect anchors.")
            return
        if mode == "Global":
            item = self.layout_items[target_ids[0]]
            anchor = self.layout_compute_auto_anchor(
                item["pil"],
                layer.get("anchor_type", "Bottom"),
                int(layer.get("anchor_padding", 2)),
            )
            if not anchor:
                return
            layer["anchor_x"] = anchor["x"]
            layer["anchor_y"] = anchor["y"]
        else:
            for item_id in target_ids:
                item = self.layout_items.get(item_id)
                if not item:
                    continue
                anchor = self.layout_compute_auto_anchor(
                    item["pil"],
                    layer.get("anchor_type", "Bottom"),
                    int(layer.get("anchor_padding", 2)),
                )
                if anchor:
                    item["anchor"] = anchor
        self.layout_update_anchor_fields()

    def layout_anchor_set_manual(self):
        layer = self.layout_get_active_layer()
        if not layer:
            return
        mode = layer.get("anchor_mode", "Off")
        if mode == "Off":
            messagebox.showinfo("Anchors", "Set anchor mode to Global or Per-frame first.")
            return
        try:
            anchor_x = int(self.layout_anchor_x_var.get())
            anchor_y = int(self.layout_anchor_y_var.get())
        except ValueError:
            messagebox.showerror("Anchors", "Anchor X/Y must be whole numbers.")
            return
        if mode == "Global":
            layer["anchor_x"] = anchor_x
            layer["anchor_y"] = anchor_y
        else:
            if not self.layout_selected_ids:
                messagebox.showinfo("Anchors", "Select frames to set per-frame anchors.")
                return
            for item_id in self.layout_selected_ids:
                item = self.layout_items.get(item_id)
                if not item:
                    continue
                item["anchor"] = {"x": anchor_x, "y": anchor_y}
        self.layout_update_anchor_fields()

    def layout_anchor_align_to_reference(self):
        if self.layout_mode_var.get() != "Grid":
            messagebox.showinfo("Anchors", "Anchor alignment requires Grid mode.")
            return
        if not self.layout_reference_layer_id:
            messagebox.showinfo("Anchors", "Select a reference layer first.")
            return
        active_layer = self.layout_get_active_layer()
        if not active_layer:
            return
        if active_layer.get("anchor_mode", "Off") == "Off":
            messagebox.showinfo("Anchors", "Enable anchors for the active layer first.")
            return
        if active_layer["id"] == self.layout_reference_layer_id:
            messagebox.showinfo("Anchors", "Active layer is already the reference layer.")
            return

        ref_layer = self.layout_get_layer_by_id(self.layout_reference_layer_id)
        if not ref_layer:
            messagebox.showinfo("Anchors", "Reference layer not found.")
            return

        cell_w, cell_h, pad = self.layout_get_cell_and_padding()
        step_x = cell_w + pad
        step_y = cell_h + pad
        if step_x <= 0 or step_y <= 0:
            return

        ref_anchors = {}
        for item in self.layout_items.values():
            if item.get("layer_id") != ref_layer["id"]:
                continue
            anchor = self.layout_get_item_anchor(item, ref_layer, compute_auto=True)
            if not anchor:
                continue
            col = int(round(item["x"] / step_x))
            row = int(round(item["y"] / step_y))
            ref_anchors[(row, col)] = (item["x"] + anchor["x"], item["y"] + anchor["y"])

        if not ref_anchors:
            messagebox.showinfo("Anchors", "Reference layer has no anchors to align to.")
            return

        target_ids = list(self.layout_selected_ids) if self.layout_selected_ids else [
            item_id for item_id, item in self.layout_items.items() if item.get("layer_id") == active_layer["id"]
        ]
        self.layout_push_undo()
        moved = False
        for item_id in target_ids:
            item = self.layout_items.get(item_id)
            if not item or item.get("layer_id") != active_layer["id"]:
                continue
            anchor = self.layout_get_item_anchor(item, active_layer, compute_auto=True)
            if not anchor:
                continue
            col = int(round(item["x"] / step_x))
            row = int(round(item["y"] / step_y))
            ref_anchor = ref_anchors.get((row, col))
            if not ref_anchor:
                continue
            current_anchor = (item["x"] + anchor["x"], item["y"] + anchor["y"])
            delta_x = ref_anchor[0] - current_anchor[0]
            delta_y = ref_anchor[1] - current_anchor[1]
            item["x"] += delta_x
            item["y"] += delta_y
            moved = True
            canvas_x, canvas_y = self.layout_to_canvas(item["x"], item["y"])
            self.layout_canvas.coords(item_id, canvas_x, canvas_y)
            if item.get("rect_id"):
                width = item["width"] * self.layout_zoom
                height = item["height"] * self.layout_zoom
                self.layout_canvas.coords(item["rect_id"], canvas_x, canvas_y, canvas_x + width, canvas_y + height)
        if moved:
            self.layout_redraw()
            self.layout_update_position_fields()

    def layout_apply_position(self):
        if not self.layout_selected_ids:
            return
        x_text = self.layout_pos_x_var.get().strip()
        y_text = self.layout_pos_y_var.get().strip()
        try:
            x_val = int(x_text) if x_text else None
            y_val = int(y_text) if y_text else None
        except ValueError:
            messagebox.showerror("Invalid position", "X and Y must be whole numbers.")
            return
        self.layout_push_undo()
        for item_id in self.layout_selected_ids:
            item = self.layout_items[item_id]
            new_x = item["x"] if x_val is None else x_val
            new_y = item["y"] if y_val is None else y_val
            item["x"] = new_x
            item["y"] = new_y
            canvas_x, canvas_y = self.layout_to_canvas(new_x, new_y)
            self.layout_canvas.coords(item_id, canvas_x, canvas_y)
            if item.get("rect_id"):
                width = item["width"] * self.layout_zoom
                height = item["height"] * self.layout_zoom
                self.layout_canvas.coords(item["rect_id"], canvas_x, canvas_y, canvas_x + width, canvas_y + height)
        if self.layout_snap_var.get() and self.layout_mode_var.get() == "Grid":
            self.layout_snap_selected(record_undo=False)
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_align(self, direction):
        if len(self.layout_selected_ids) < 2:
            return
        self.layout_push_undo()
        selected = list(self.layout_selected_ids)
        ref = self.layout_items[selected[0]]
        ref_x = ref["x"]
        ref_y = ref["y"]
        for item_id in self.layout_selected_ids:
            item = self.layout_items[item_id]
            if direction == "horizontal":
                item["y"] = ref_y
            else:
                item["x"] = ref_x
            canvas_x, canvas_y = self.layout_to_canvas(item["x"], item["y"])
            self.layout_canvas.coords(item_id, canvas_x, canvas_y)
            if item.get("rect_id"):
                width = item["width"] * self.layout_zoom
                height = item["height"] * self.layout_zoom
                self.layout_canvas.coords(item["rect_id"], canvas_x, canvas_y, canvas_x + width, canvas_y + height)
        if self.layout_snap_var.get() and self.layout_mode_var.get() == "Grid":
            self.layout_snap_selected(record_undo=False)
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_center_in_cell(self):
        if not self.layout_selected_ids:
            return
        if self.layout_mode_var.get() != "Grid":
            messagebox.showinfo("Center in cell", "Centering requires Grid mode.")
            return
        self.layout_push_undo()
        cell_w, cell_h, pad = self.layout_get_cell_and_padding()
        step_x = cell_w + pad
        step_y = cell_h + pad
        if step_x <= 0 or step_y <= 0:
            return
        for item_id in self.layout_selected_ids:
            item = self.layout_items[item_id]
            col = int(round(item["x"] / step_x))
            row = int(round(item["y"] / step_y))
            new_x = col * step_x + (cell_w - item["width"]) / 2
            new_y = row * step_y + (cell_h - item["height"]) / 2
            item["x"] = new_x
            item["y"] = new_y
            canvas_x, canvas_y = self.layout_to_canvas(new_x, new_y)
            self.layout_canvas.coords(item_id, canvas_x, canvas_y)
            if item.get("rect_id"):
                width = item["width"] * self.layout_zoom
                height = item["height"] * self.layout_zoom
                self.layout_canvas.coords(item["rect_id"], canvas_x, canvas_y, canvas_x + width, canvas_y + height)
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_get_cell_and_padding(self, extra_sizes=None):
        extra_sizes = extra_sizes or []
        widths = [item["width"] for item in self.layout_items.values()] + [w for w, _ in extra_sizes]
        heights = [item["height"] for item in self.layout_items.values()] + [h for _, h in extra_sizes]

        if self.layout_type_var.get() == "Tilemap":
            if self.tile_size_mode_var.get() == "Force":
                cell_w = cell_h = int(self.tile_size_var.get())
            else:
                cell_w = max(widths) if widths else int(self.tile_size_var.get())
                cell_h = max(heights) if heights else int(self.tile_size_var.get())
            pad = 0
        else:
            cell_w = max(widths) if widths else 32
            cell_h = max(heights) if heights else 32
            if self.sprite_padding_mode_var.get() == "Frame width":
                pad = cell_w
            elif self.sprite_padding_mode_var.get() == "Frame height":
                pad = cell_h
            else:
                pad = int(self.sprite_padding_var.get())
        return cell_w, cell_h, pad

    def layout_to_canvas(self, x, y):
        return x * self.layout_zoom, y * self.layout_zoom

    def layout_from_canvas(self, x, y):
        return x / self.layout_zoom, y / self.layout_zoom

    def layout_make_photo(self, pil_image):
        if self.layout_zoom == 1.0:
            resized = pil_image
        else:
            width = max(1, int(round(pil_image.width * self.layout_zoom)))
            height = max(1, int(round(pil_image.height * self.layout_zoom)))
            resample = Image.NEAREST
            if hasattr(Image, "Resampling"):
                resample = Image.Resampling.NEAREST
            resized = pil_image.resize((width, height), resample)
        return ImageTk.PhotoImage(resized)

    def layout_set_zoom(self, zoom):
        zoom = max(0.2, min(zoom, 8.0))
        if abs(zoom - self.layout_zoom) < 0.001:
            return
        self.layout_zoom = zoom
        for item_id, item in self.layout_items.items():
            item["image"] = self.layout_make_photo(item["pil"])
            self.layout_canvas.itemconfigure(item_id, image=item["image"])
        self.layout_redraw()

    def layout_on_mousewheel(self, event):
        delta = event.delta
        if delta == 0:
            return
        factor = 1.1 if delta > 0 else 0.9
        self.layout_set_zoom(self.layout_zoom * factor)

    def setup_layout_dnd(self):
        if TkinterDnD is None or DND_FILES is None:
            self.layout_dnd_ready = False
            return
        try:
            self.layout_canvas.drop_target_register(DND_FILES)
            self.layout_canvas.dnd_bind("<<Drop>>", self.layout_on_dnd_drop)
            self.layout_dnd_ready = True
        except Exception:
            self.layout_dnd_ready = False

    def layout_parse_dnd_files(self, data):
        if not data:
            return []
        try:
            return list(self.root.tk.splitlist(data))
        except Exception:
            return data.split()

    def layout_on_dnd_drop(self, event):
        files = self.layout_parse_dnd_files(getattr(event, "data", ""))
        json_files = [path for path in files if path.lower().endswith(".json")]
        if not json_files:
            return "break"
        self.layout_load_file(json_files[0])
        return "break"

    def layout_next_position(self, index, cell_w, cell_h, pad):
        if self.layout_mode_var.get() == "Grid":
            cols = max(1, int(self.layout_grid_columns_var.get()))
            row = index // cols
            col = index % cols
            x = col * (cell_w + pad)
            y = row * (cell_h + pad)
            return x, y
        offset = 10
        return index * offset, index * offset

    def layout_add_selected(self):
        folder_path = self.layout_get_folder_path()
        if not folder_path or not os.path.isdir(folder_path):
            messagebox.showerror("Invalid folder", "Select a valid source folder first.")
            return
        selections = self.layout_listbox.curselection()
        if not selections:
            return
        existing = {item["file"] for item in self.layout_items.values()}
        new_items = []
        for idx in selections:
            file = self.layout_listbox.get(idx)
            if file in existing:
                continue
            path = os.path.join(folder_path, file)
            try:
                with Image.open(path) as img:
                    pil = img.convert("RGBA")
            except Exception:
                continue
            photo = self.layout_make_photo(pil)
            new_items.append((file, path, pil, photo, pil.width, pil.height))

        if not new_items:
            return
        extra_sizes = [(item[4], item[5]) for item in new_items]
        cell_w, cell_h, pad = self.layout_get_cell_and_padding(extra_sizes=extra_sizes)
        start_index = len(self.layout_items)
        layer = self.layout_get_active_layer()
        if not layer:
            self.layout_create_layer("Layer 1")
            layer = self.layout_get_active_layer()
        for i, (file, path, pil, photo, width, height) in enumerate(new_items):
            x, y = self.layout_next_position(start_index + i, cell_w, cell_h, pad)
            canvas_x, canvas_y = self.layout_to_canvas(x, y)
            item_id = self.layout_canvas.create_image(canvas_x, canvas_y, anchor="nw", image=photo, tags=("layout_item",))
            self.layout_item_counter += 1
            self.layout_items[item_id] = {
                "file": file,
                "path": path,
                "image": photo,
                "pil": pil,
                "x": x,
                "y": y,
                "width": width,
                "height": height,
                "rect_id": None,
                "layer_id": layer["id"],
                "visible": True,
                "anchor_inherit": True,
                "anchor": None,
                "order": self.layout_item_counter,
            }
        self.layout_redraw()
        self.layout_maybe_set_output_name()

    def layout_remove_selected(self):
        for item_id in list(self.layout_selected_ids):
            item = self.layout_items.pop(item_id, None)
            if not item:
                continue
            if item.get("rect_id"):
                self.layout_canvas.delete(item["rect_id"])
            self.layout_canvas.delete(item_id)
        self.layout_selected_ids.clear()
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_copy_selected(self):
        if not self.layout_selected_ids:
            return
        cell_w, cell_h, pad = self.layout_get_cell_and_padding()
        step_x = cell_w + pad
        step_y = cell_h + pad
        use_grid = self.layout_mode_var.get() == "Grid"
        cols = max(1, int(self.layout_grid_columns_var.get())) if use_grid else 1
        occupied = set()
        if use_grid and step_x > 0 and step_y > 0:
            for existing in self.layout_items.values():
                col = int(round(existing["x"] / step_x))
                row = int(round(existing["y"] / step_y))
                occupied.add((row, col))
        new_ids = []
        for item_id in list(self.layout_selected_ids):
            item = self.layout_items.get(item_id)
            if not item:
                continue
            pil = item["pil"]
            photo = self.layout_make_photo(pil)
            new_x = item["x"]
            new_y = item["y"]
            if use_grid and step_x > 0 and step_y > 0:
                base_col = int(round(item["x"] / step_x))
                base_row = int(round(item["y"] / step_y))
                index = base_row * cols + base_col + 1
                while True:
                    col = index % cols
                    row = index // cols
                    if (row, col) not in occupied:
                        break
                    index += 1
                occupied.add((row, col))
                new_x = col * step_x
                new_y = row * step_y
            else:
                new_x += 10
                new_y += 10
            canvas_x, canvas_y = self.layout_to_canvas(new_x, new_y)
            new_id = self.layout_canvas.create_image(
                canvas_x,
                canvas_y,
                anchor="nw",
                image=photo,
                tags=("layout_item",),
            )
            self.layout_item_counter += 1
            self.layout_items[new_id] = {
                "file": item["file"],
                "path": item["path"],
                "image": photo,
                "pil": pil,
                "x": new_x,
                "y": new_y,
                "width": item["width"],
                "height": item["height"],
                "rect_id": None,
                "layer_id": item.get("layer_id"),
                "visible": item.get("visible", True),
                "anchor_inherit": item.get("anchor_inherit", True),
                "anchor": item.get("anchor"),
                "order": self.layout_item_counter,
            }
            new_ids.append(new_id)
        for new_id in new_ids:
            self.layout_selected_ids.add(new_id)
            self.layout_update_selection_rect(new_id)
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_clear_canvas(self):
        for item_id, item in list(self.layout_items.items()):
            if item.get("rect_id"):
                self.layout_canvas.delete(item["rect_id"])
            self.layout_canvas.delete(item_id)
        self.layout_items.clear()
        self.layout_selected_ids.clear()
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_find_item_at(self, x, y):
        items = self.layout_canvas.find_overlapping(x, y, x, y)
        for item_id in reversed(items):
            if item_id in self.layout_items:
                item = self.layout_items.get(item_id)
                if not self.layout_item_is_visible(item):
                    continue
                return item_id
        return None

    def layout_update_selection_rect(self, item_id):
        item = self.layout_items.get(item_id)
        if not item:
            return
        if not self.layout_item_is_visible(item):
            self.layout_clear_selection_rect(item_id)
            self.layout_selected_ids.discard(item_id)
            return
        x, y = self.layout_to_canvas(item["x"], item["y"])
        width = item["width"] * self.layout_zoom
        height = item["height"] * self.layout_zoom
        rect = item.get("rect_id")
        if rect is None:
            rect = self.layout_canvas.create_rectangle(
                x,
                y,
                x + width,
                y + height,
                outline="#3d7bfd",
                width=2,
                tags=("selection",),
            )
            item["rect_id"] = rect
        else:
            self.layout_canvas.coords(rect, x, y, x + width, y + height)
        self.layout_canvas.tag_raise(rect)

    def layout_clear_selection_rect(self, item_id):
        item = self.layout_items.get(item_id)
        if item and item.get("rect_id"):
            self.layout_canvas.delete(item["rect_id"])
            item["rect_id"] = None

    def layout_select_item(self, item_id, add=False, toggle=False):
        if toggle and item_id in self.layout_selected_ids:
            self.layout_selected_ids.remove(item_id)
            self.layout_clear_selection_rect(item_id)
            self.layout_update_position_fields()
            return
        if not add and not toggle:
            for selected in list(self.layout_selected_ids):
                self.layout_clear_selection_rect(selected)
            self.layout_selected_ids.clear()
        self.layout_selected_ids.add(item_id)
        self.layout_update_selection_rect(item_id)
        self.layout_update_position_fields()

    def layout_on_press(self, event):
        self.layout_canvas.focus_set()
        canvas_x = self.layout_canvas.canvasx(event.x)
        canvas_y = self.layout_canvas.canvasy(event.y)
        guide_hit = self.layout_guide_hit(canvas_x, canvas_y)
        if guide_hit:
            self.layout_guide_drag = guide_hit
            self.layout_drag_start = None
            self.layout_select_box_start = None
            self.layout_pending_select_id = None
            return
        item_id = self.layout_find_item_at(canvas_x, canvas_y)
        ctrl = bool(event.state & 0x0004)
        if not item_id:
            if not ctrl:
                for selected in list(self.layout_selected_ids):
                    self.layout_clear_selection_rect(selected)
                self.layout_selected_ids.clear()
                self.layout_update_position_fields()
            self.layout_drag_start = None
            self.layout_drag_moved = False
            self.layout_drag_anchor_id = None
            self.layout_pending_select_id = None
            self.layout_select_box_start = (canvas_x, canvas_y)
            self.layout_select_box_add = ctrl
            self.layout_select_box_moved = False
            self.layout_select_box_id = self.layout_canvas.create_rectangle(
                canvas_x,
                canvas_y,
                canvas_x,
                canvas_y,
                outline="#3d7bfd",
                dash=(3, 2),
                tags=("selection_box",),
            )
            return
        if ctrl:
            self.layout_select_item(item_id, add=False, toggle=True)
            self.layout_pending_select_id = None
        else:
            if item_id in self.layout_selected_ids:
                self.layout_pending_select_id = None
            elif self.layout_selected_ids:
                self.layout_pending_select_id = item_id
            else:
                self.layout_select_item(item_id, add=False, toggle=False)
                self.layout_pending_select_id = None
        x, y = self.layout_from_canvas(canvas_x, canvas_y)
        self.layout_drag_start = (x, y)
        self.layout_drag_positions = {
            selected: (self.layout_items[selected]["x"], self.layout_items[selected]["y"])
            for selected in self.layout_selected_ids
        }
        self.layout_drag_moved = False
        self.layout_drag_anchor_id = item_id

    def layout_on_drag(self, event):
        if self.layout_guide_drag:
            canvas_x = self.layout_canvas.canvasx(event.x)
            canvas_y = self.layout_canvas.canvasy(event.y)
            axis, index = self.layout_guide_drag
            if axis == "v":
                self.layout_guides["v"][index] = canvas_x / self.layout_zoom
            else:
                self.layout_guides["h"][index] = canvas_y / self.layout_zoom
            self.layout_redraw()
            return
        if self.layout_select_box_start:
            canvas_x = self.layout_canvas.canvasx(event.x)
            canvas_y = self.layout_canvas.canvasy(event.y)
            x0, y0 = self.layout_select_box_start
            if self.layout_select_box_id:
                self.layout_canvas.coords(self.layout_select_box_id, x0, y0, canvas_x, canvas_y)
            if abs(canvas_x - x0) > 2 or abs(canvas_y - y0) > 2:
                self.layout_select_box_moved = True
            return
        if not self.layout_drag_start:
            return
        canvas_x = self.layout_canvas.canvasx(event.x)
        canvas_y = self.layout_canvas.canvasy(event.y)
        x, y = self.layout_from_canvas(canvas_x, canvas_y)
        dx = x - self.layout_drag_start[0]
        dy = y - self.layout_drag_start[1]
        cell_w, cell_h, pad = self.layout_get_cell_and_padding()
        step_x = cell_w + pad
        step_y = cell_h + pad
        snap = self.layout_snap_var.get() and self.layout_mode_var.get() == "Grid"

        if self.layout_guides_enabled and not snap and self.layout_selected_ids:
            anchor_id = self.layout_drag_anchor_id
            if anchor_id not in self.layout_drag_positions:
                anchor_id = next(iter(self.layout_selected_ids))
            anchor_start = self.layout_drag_positions.get(anchor_id)
            anchor_item = self.layout_items.get(anchor_id)
            if anchor_start and anchor_item:
                offset_x, offset_y = self.layout_get_anchor_offset(anchor_item)
                anchor_x = anchor_start[0] + dx + offset_x
                anchor_y = anchor_start[1] + dy + offset_y
                snap_x, snap_y = self.layout_apply_guide_snap(anchor_x, anchor_y)
                dx += snap_x - anchor_x
                dy += snap_y - anchor_y

        moved_positions = {}
        moved = False
        for item_id in self.layout_selected_ids:
            start_x, start_y = self.layout_drag_positions.get(item_id, (0, 0))
            new_x = start_x + dx
            new_y = start_y + dy
            if snap and step_x > 0 and step_y > 0:
                col = int(round(new_x / step_x))
                row = int(round(new_y / step_y))
                new_x = col * step_x
                new_y = row * step_y
            if new_x != start_x or new_y != start_y:
                moved = True
            moved_positions[item_id] = (new_x, new_y)

        if moved and not self.layout_drag_moved:
            self.layout_push_undo()
            self.layout_drag_moved = True
            self.layout_pending_select_id = None

        for item_id, (new_x, new_y) in moved_positions.items():
            item = self.layout_items[item_id]
            item["x"] = new_x
            item["y"] = new_y
            canvas_x, canvas_y = self.layout_to_canvas(new_x, new_y)
            self.layout_canvas.coords(item_id, canvas_x, canvas_y)
            if item.get("rect_id"):
                width = item["width"] * self.layout_zoom
                height = item["height"] * self.layout_zoom
                self.layout_canvas.coords(item["rect_id"], canvas_x, canvas_y, canvas_x + width, canvas_y + height)

        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_on_release(self, _event):
        if self.layout_guide_drag:
            self.layout_guide_drag = None
            return
        if self.layout_select_box_start:
            x0, y0 = self.layout_select_box_start
            x1, y1 = self.layout_canvas.coords(self.layout_select_box_id)[2:]
            x_min, x_max = sorted([x0, x1])
            y_min, y_max = sorted([y0, y1])
            if self.layout_select_box_moved:
                selected_items = self.layout_canvas.find_enclosed(x_min, y_min, x_max, y_max)
                if not self.layout_select_box_add:
                    for selected in list(self.layout_selected_ids):
                        self.layout_clear_selection_rect(selected)
                    self.layout_selected_ids.clear()
                for item_id in selected_items:
                    if item_id in self.layout_items:
                        item = self.layout_items.get(item_id)
                        if not self.layout_item_is_visible(item):
                            continue
                        self.layout_selected_ids.add(item_id)
                        self.layout_update_selection_rect(item_id)
            if self.layout_select_box_id:
                self.layout_canvas.delete(self.layout_select_box_id)
            self.layout_select_box_start = None
            self.layout_select_box_id = None
            self.layout_select_box_add = False
            self.layout_select_box_moved = False
            self.layout_update_position_fields()
            return
        if self.layout_pending_select_id:
            for selected in list(self.layout_selected_ids):
                self.layout_clear_selection_rect(selected)
            self.layout_selected_ids.clear()
            self.layout_selected_ids.add(self.layout_pending_select_id)
            self.layout_update_selection_rect(self.layout_pending_select_id)
        self.layout_drag_start = None
        self.layout_drag_positions = {}
        self.layout_drag_moved = False
        self.layout_drag_anchor_id = None
        self.layout_pending_select_id = None
        self.layout_update_position_fields()

    def layout_snap_selected(self, record_undo=True):
        if not self.layout_snap_var.get() or self.layout_mode_var.get() != "Grid":
            return
        if record_undo:
            self.layout_push_undo()
        cell_w, cell_h, pad = self.layout_get_cell_and_padding()
        step_x = cell_w + pad
        step_y = cell_h + pad
        if step_x <= 0 or step_y <= 0:
            return
        for item_id, item in self.layout_items.items():
            col = int(round(item["x"] / step_x))
            row = int(round(item["y"] / step_y))
            new_x = col * step_x
            new_y = row * step_y
            item["x"] = new_x
            item["y"] = new_y
            canvas_x, canvas_y = self.layout_to_canvas(new_x, new_y)
            self.layout_canvas.coords(item_id, canvas_x, canvas_y)
            if item.get("rect_id"):
                width = item["width"] * self.layout_zoom
                height = item["height"] * self.layout_zoom
                self.layout_canvas.coords(item["rect_id"], canvas_x, canvas_y, canvas_x + width, canvas_y + height)
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_redraw(self):
        self.update_layout_mode_controls()
        self.update_layout_tile_size_controls()
        cell_w, cell_h, pad = self.layout_get_cell_and_padding()
        cols = max(1, int(self.layout_grid_columns_var.get()))
        rows = int(self.layout_grid_rows_var.get())
        max_x = max((item["x"] + item["width"]) for item in self.layout_items.values()) if self.layout_items else 0
        max_y = max((item["y"] + item["height"]) for item in self.layout_items.values()) if self.layout_items else 0

        if self.layout_mode_var.get() == "Grid":
            step_x = cell_w + pad
            step_y = cell_h + pad
            if rows <= 0 and step_x > 0 and step_y > 0 and self.layout_items:
                row_indices = [int(item["y"] // step_y) for item in self.layout_items.values()]
                rows = max(row_indices) + 1 if row_indices else 1
            if rows <= 0:
                rows = 1
            grid_w = (cols * cell_w) + (pad * max(0, cols - 1))
            grid_h = (rows * cell_h) + (pad * max(0, rows - 1))
        else:
            grid_w = 0
            grid_h = 0

        canvas_w = max(grid_w, max_x, 1) * self.layout_zoom
        canvas_h = max(grid_h, max_y, 1) * self.layout_zoom
        self.layout_canvas.configure(scrollregion=(0, 0, canvas_w, canvas_h))

        self.layout_canvas.delete("grid_line")
        self.layout_canvas.delete("guide_line")
        self.layout_canvas.delete("snap_guide")

        if self.layout_show_grid_var.get() and self.layout_mode_var.get() == "Grid":
            step_x = cell_w + pad
            step_y = cell_h + pad
            for col in range(cols):
                x0 = col * step_x
                x1 = x0 + cell_w
                x0s = x0 * self.layout_zoom
                x1s = x1 * self.layout_zoom
                self.layout_canvas.create_line(x0s, 0, x0s, grid_h * self.layout_zoom, fill="#808080", tags=("grid_line",))
                self.layout_canvas.create_line(x1s, 0, x1s, grid_h * self.layout_zoom, fill="#808080", tags=("grid_line",))
            for row in range(rows):
                y0 = row * step_y
                y1 = y0 + cell_h
                y0s = y0 * self.layout_zoom
                y1s = y1 * self.layout_zoom
                self.layout_canvas.create_line(0, y0s, grid_w * self.layout_zoom, y0s, fill="#808080", tags=("grid_line",))
                self.layout_canvas.create_line(0, y1s, grid_w * self.layout_zoom, y1s, fill="#808080", tags=("grid_line",))
            self.layout_canvas.tag_lower("grid_line")

        if self.layout_show_guides_var.get():
            cx = canvas_w // 2
            cy = canvas_h // 2
            self.layout_canvas.create_line(cx, 0, cx, canvas_h, fill="#a0a0a0", dash=(4, 2), tags=("guide_line",))
            self.layout_canvas.create_line(0, cy, canvas_w, cy, fill="#a0a0a0", dash=(4, 2), tags=("guide_line",))
            self.layout_canvas.tag_lower("guide_line")

        if self.layout_guides_enabled:
            for gx in self.layout_guides.get("v", []):
                xs = gx * self.layout_zoom
                self.layout_canvas.create_line(xs, 0, xs, canvas_h, fill="#d97904", dash=(2, 2), tags=("snap_guide",))
            for gy in self.layout_guides.get("h", []):
                ys = gy * self.layout_zoom
                self.layout_canvas.create_line(0, ys, canvas_w, ys, fill="#d97904", dash=(2, 2), tags=("snap_guide",))
            self.layout_canvas.tag_lower("snap_guide")

        for item_id, item in self.layout_items.items():
            canvas_x, canvas_y = self.layout_to_canvas(item["x"], item["y"])
            self.layout_canvas.coords(item_id, canvas_x, canvas_y)
            if self.layout_item_is_visible(item):
                self.layout_canvas.itemconfigure(item_id, state="normal")
            else:
                self.layout_canvas.itemconfigure(item_id, state="hidden")
            if item.get("rect_id"):
                width = item["width"] * self.layout_zoom
                height = item["height"] * self.layout_zoom
                self.layout_canvas.coords(item["rect_id"], canvas_x, canvas_y, canvas_x + width, canvas_y + height)

        for item_id in list(self.layout_selected_ids):
            self.layout_update_selection_rect(item_id)
        self.layout_apply_layer_order()

    def layout_export(self):
        if self.running or self.previewing or self.sheet_running or self.tile_running:
            messagebox.showerror("Busy", "A batch run is already in progress.")
            return
        folder_path = self.layout_get_folder_path()
        if not folder_path or not os.path.isdir(folder_path):
            messagebox.showerror("Invalid folder", "Select a valid source folder first.")
            return
        if not self.layout_items:
            messagebox.showerror("No frames", "Add frames to the canvas first.")
            return
        visible_items = self.layout_get_ordered_items()
        if not visible_items:
            messagebox.showerror("No visible frames", "All frames are hidden.")
            return
        output_name = self.layout_output_var.get().strip()
        if not output_name:
            missing_kind = "spritesheet" if self.layout_type_var.get() == "Spritesheet" else "tilemap"
            messagebox.showerror("Missing name", f"Provide a {missing_kind} name first.")
            return
        if output_name.lower().endswith(".png"):
            output_name = output_name[:-4]
        elif output_name.lower().endswith(".json"):
            output_name = output_name[:-5]

        cell_w, cell_h, pad = self.layout_get_cell_and_padding()
        cols = max(1, int(self.layout_grid_columns_var.get()))
        rows = int(self.layout_grid_rows_var.get())
        max_x = max((item["x"] + item["width"]) for item in visible_items)
        max_y = max((item["y"] + item["height"]) for item in visible_items)

        if self.layout_mode_var.get() == "Grid":
            step_x = cell_w + pad
            step_y = cell_h + pad
            if rows <= 0 and step_x > 0 and step_y > 0:
                row_indices = [int(item["y"] // step_y) for item in visible_items]
                rows = max(row_indices) + 1 if row_indices else 1
            if rows <= 0:
                rows = 1
            grid_w = (cols * cell_w) + (pad * max(0, cols - 1))
            grid_h = (rows * cell_h) + (pad * max(0, rows - 1))
        else:
            grid_w = 0
            grid_h = 0

        out_w = max(grid_w, max_x)
        out_h = max(grid_h, max_y)

        if self.layout_type_var.get() == "Spritesheet":
            out_dir = os.path.join(folder_path, "sprite_sheets")
            out_name = f"{output_name}_sprite_sheet.png"
        else:
            out_dir = os.path.join(folder_path, "tilemaps")
            out_name = f"{output_name}.png"

        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, out_name)
        canvas = Image.new("RGBA", (out_w, out_h), (0, 0, 0, 0))
        metadata = []
        for item in visible_items:
            try:
                with Image.open(item["path"]) as img:
                    img = img.convert("RGBA")
                    canvas.paste(img, (int(item["x"]), int(item["y"])))
                    metadata.append((item["file"], int(item["x"]), int(item["y"]), img.width, img.height))
            except Exception:
                continue
        canvas.save(out_path, format="PNG")

        if self.layout_type_var.get() == "Tilemap" and self.tile_export_meta_var.get():
            meta_path = os.path.join(out_dir, f"{output_name}.csv")
            with open(meta_path, "w", newline="", encoding="utf-8") as handle:
                writer = csv.writer(handle)
                writer.writerow(["tilemap", "tile", "x", "y", "width", "height"])
                for file, x, y, w, h in metadata:
                    writer.writerow([out_name, file, x, y, w, h])

        self.layout_save()
        self.layout_status_label.configure(text="Exported")
        messagebox.showinfo("Exported", f"Saved to:\n{out_path}")

    def layout_build_json(self):
        return {
            "type": self.layout_type_var.get(),
            "input_root": self.input_var.get(),
            "source_folder": self.layout_folder_var.get(),
            "output_name": self.layout_output_var.get(),
            "layout_mode": self.layout_mode_var.get(),
            "columns": self.layout_grid_columns_var.get(),
            "rows": self.layout_grid_rows_var.get(),
            "snap": self.layout_snap_var.get(),
            "show_grid": self.layout_show_grid_var.get(),
            "show_guides": self.layout_show_guides_var.get(),
            "sprite_padding_mode": self.sprite_padding_mode_var.get(),
            "sprite_padding": self.sprite_padding_var.get(),
            "tile_size_mode": self.tile_size_mode_var.get(),
            "tile_size": self.tile_size_var.get(),
            "tile_export_meta": self.tile_export_meta_var.get(),
            "layers": [
                {
                    "id": layer["id"],
                    "name": layer["name"],
                    "visible": layer.get("visible", True),
                    "anchor_mode": layer.get("anchor_mode", "Off"),
                    "anchor_source": layer.get("anchor_source", "Auto"),
                    "anchor_type": layer.get("anchor_type", "Bottom"),
                    "anchor_padding": layer.get("anchor_padding", 2),
                    "anchor_x": layer.get("anchor_x"),
                    "anchor_y": layer.get("anchor_y"),
                }
                for layer in self.layout_layers
            ],
            "active_layer_id": self.layout_active_layer_id,
            "reference_layer_id": self.layout_reference_layer_id,
            "items": [
                {
                    "file": item["file"],
                    "x": int(item["x"]),
                    "y": int(item["y"]),
                    "width": int(item["width"]),
                    "height": int(item["height"]),
                    "layer_id": item.get("layer_id"),
                    "visible": item.get("visible", True),
                    "anchor_inherit": item.get("anchor_inherit", True),
                    "anchor": item.get("anchor"),
                    "order": item.get("order", 0),
                }
                for item in self.layout_items.values()
            ],
        }

    def layout_save(self):
        folder_path = self.layout_get_folder_path()
        if not folder_path:
            return
        output_name = self.layout_output_var.get().strip()
        if not output_name:
            return
        if output_name.lower().endswith(".png"):
            output_name = output_name[:-4]
        elif output_name.lower().endswith(".json"):
            output_name = output_name[:-5]
        if self.layout_type_var.get() == "Spritesheet":
            out_dir = os.path.join(folder_path, "sprite_sheets")
            layout_name = f"{output_name}_sprite_sheet.json"
        else:
            out_dir = os.path.join(folder_path, "tilemaps")
            layout_name = f"{output_name}.json"
        os.makedirs(out_dir, exist_ok=True)
        layout_path = os.path.join(out_dir, layout_name)
        with open(layout_path, "w", encoding="utf-8") as handle:
            json.dump(self.layout_build_json(), handle, indent=2)
        self.layout_refresh_recent_jsons()

    def layout_load_file(self, path):
        if not path or not os.path.isfile(path):
            messagebox.showerror("Layout error", "Select a valid JSON file.")
            return
        if not path.lower().endswith(".json"):
            messagebox.showerror("Layout error", "Only JSON layout files can be loaded.")
            return
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception as exc:
            messagebox.showerror("Layout error", f"Failed to load layout:\n{exc}")
            return

        input_root = data.get("input_root")
        if input_root and os.path.isdir(input_root):
            self.input_var.set(input_root)
        self.layout_type_var.set(data.get("type", "Spritesheet"))
        self.layout_folder_var.set(data.get("source_folder", ""))
        self.layout_output_var.set(data.get("output_name", ""))
        self.layout_mode_var.set(data.get("layout_mode", "Grid"))
        self.layout_grid_columns_var.set(int(data.get("columns", 4)))
        self.layout_grid_rows_var.set(int(data.get("rows", 0)))
        self.layout_snap_var.set(bool(data.get("snap", True)))
        self.layout_show_grid_var.set(bool(data.get("show_grid", True)))
        self.layout_show_guides_var.set(bool(data.get("show_guides", False)))
        self.sprite_padding_mode_var.set(data.get("sprite_padding_mode", "Fixed"))
        self.sprite_padding_var.set(int(data.get("sprite_padding", 1)))
        self.tile_size_mode_var.set(data.get("tile_size_mode", "Force"))
        self.tile_size_var.set(int(data.get("tile_size", 32)))
        self.tile_export_meta_var.set(bool(data.get("tile_export_meta", True)))
        self.update_layout_type_controls()
        self.refresh_layout_folders()

        self.layout_layers = []
        self.layout_layer_counter = 0
        self.layout_active_layer_id = None
        self.layout_reference_layer_id = None
        max_layer_index = 0
        for layer_data in data.get("layers", []):
            layer_id = layer_data.get("id")
            if not layer_id:
                self.layout_layer_counter += 1
                layer_id = f"layer_{self.layout_layer_counter}"
            else:
                if layer_id.startswith("layer_"):
                    suffix = layer_id.split("_", 1)[1]
                    if suffix.isdigit():
                        max_layer_index = max(max_layer_index, int(suffix))
            self.layout_layers.append(
                {
                    "id": layer_id,
                    "name": layer_data.get("name") or layer_id,
                    "visible": layer_data.get("visible", True),
                    "anchor_mode": layer_data.get("anchor_mode", "Off"),
                    "anchor_source": layer_data.get("anchor_source", "Auto"),
                    "anchor_type": layer_data.get("anchor_type", "Bottom"),
                    "anchor_padding": int(layer_data.get("anchor_padding", 2)),
                    "anchor_x": layer_data.get("anchor_x"),
                    "anchor_y": layer_data.get("anchor_y"),
                }
            )
        if max_layer_index:
            self.layout_layer_counter = max(self.layout_layer_counter, max_layer_index)
        if not self.layout_layers:
            self.layout_create_layer("Layer 1")
        self.layout_active_layer_id = data.get("active_layer_id") or self.layout_layers[0]["id"]
        if not self.layout_get_layer_by_id(self.layout_active_layer_id):
            self.layout_active_layer_id = self.layout_layers[0]["id"]
        self.layout_reference_layer_id = data.get("reference_layer_id") or self.layout_layers[0]["id"]
        if not self.layout_get_layer_by_id(self.layout_reference_layer_id):
            self.layout_reference_layer_id = self.layout_layers[0]["id"]
        self.layout_refresh_layers_list()

        self.layout_clear_canvas()
        folder_path = self.layout_get_folder_path()
        if not folder_path:
            return
        self.layout_item_counter = 0
        for item in data.get("items", []):
            file = item.get("file")
            if not file:
                continue
            item_path = os.path.join(folder_path, file)
            if not os.path.exists(item_path):
                continue
            try:
                with Image.open(item_path) as img:
                    pil = img.convert("RGBA")
            except Exception:
                continue
            photo = self.layout_make_photo(pil)
            width = pil.width
            height = pil.height
            x = int(item.get("x", 0))
            y = int(item.get("y", 0))
            canvas_x, canvas_y = self.layout_to_canvas(x, y)
            item_id = self.layout_canvas.create_image(canvas_x, canvas_y, anchor="nw", image=photo, tags=("layout_item",))
            order = int(item.get("order", 0))
            if order <= 0:
                self.layout_item_counter += 1
                order = self.layout_item_counter
            else:
                self.layout_item_counter = max(self.layout_item_counter, order)
            layer_id = item.get("layer_id") or self.layout_active_layer_id
            self.layout_items[item_id] = {
                "file": file,
                "path": item_path,
                "image": photo,
                "pil": pil,
                "x": x,
                "y": y,
                "width": width,
                "height": height,
                "rect_id": None,
                "layer_id": layer_id,
                "visible": item.get("visible", True),
                "anchor_inherit": item.get("anchor_inherit", True),
                "anchor": item.get("anchor"),
                "order": order,
            }
        self.layout_redraw()
        self.layout_update_anchor_fields()
        self.layout_refresh_recent_jsons()
        if hasattr(self, "layout_recent_combo"):
            input_root = self.input_var.get().strip()
            try:
                label = os.path.relpath(path, input_root)
            except ValueError:
                label = os.path.basename(path)
            if label in self.layout_recent_combo["values"]:
                self.layout_recent_json_var.set(label)

    def layout_load(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            initialdir=APP_DIR,
        )
        if not path:
            return
        self.layout_load_file(path)

    def layout_load_recent_json(self):
        if not getattr(self, "layout_recent_combo", None):
            return
        selected = self.layout_recent_json_var.get().strip()
        path = None
        if selected and selected in self.layout_recent_combo["values"]:
            index = list(self.layout_recent_combo["values"]).index(selected)
            if 0 <= index < len(self.layout_recent_json_paths):
                path = self.layout_recent_json_paths[index]
        if not path and self.layout_recent_json_paths:
            path = self.layout_recent_json_paths[0]
        if path:
            self.layout_load_file(path)

    def pick_split_sheet(self):
        path = filedialog.askopenfilename(
            filetypes=[("PNG files", "*.png"), ("All files", "*.*")],
            initialdir=APP_DIR,
        )
        if path:
            self.split_load_sheet(path)

    def pick_split_output_dir(self):
        path = filedialog.askdirectory(initialdir=self.split_output_dir_var.get() or os.getcwd())
        if path:
            self.split_output_dir_var.set(path)

    def split_use_sheet_name(self):
        path = self.split_sheet_path_var.get().strip()
        if not path:
            return
        base = os.path.splitext(os.path.basename(path))[0]
        self.split_base_name_var.set(base)
        self.split_update_output_dir()

    def split_update_output_dir(self):
        path = self.split_sheet_path_var.get().strip()
        if not path:
            return
        base = self.split_base_name_var.get().strip()
        if not base:
            base = os.path.splitext(os.path.basename(path))[0]
        sheet_dir = os.path.dirname(path)
        self.split_output_dir_var.set(os.path.join(sheet_dir, f"{base}_sprite_frames"))

    def split_load_sheet(self, path):
        if Image is None:
            messagebox.showerror("Pillow not installed", f"Install Pillow first.\n\n{PIL_IMPORT_ERROR}")
            return
        if not path or not os.path.isfile(path):
            messagebox.showerror("Invalid file", "Select a valid spritesheet file.")
            return
        try:
            with Image.open(path) as img:
                self.split_sheet_pil = img.convert("RGBA")
        except Exception as exc:
            messagebox.showerror("Load error", f"Failed to load spritesheet:\n{exc}")
            return
        self.split_sheet_path_var.set(path)
        if not self.split_base_name_var.get().strip():
            self.split_use_sheet_name()
        self.split_update_output_dir()
        self.split_zoom = 1.0
        self.split_selected_cells.clear()
        if not self.split_auto_detect_grid(notify=False):
            self.split_redraw()

    def split_make_photo(self, pil_image):
        if self.split_zoom == 1.0:
            resized = pil_image
        else:
            width = max(1, int(round(pil_image.width * self.split_zoom)))
            height = max(1, int(round(pil_image.height * self.split_zoom)))
            resample = Image.NEAREST
            if hasattr(Image, "Resampling"):
                resample = Image.Resampling.NEAREST
            resized = pil_image.resize((width, height), resample)
        return ImageTk.PhotoImage(resized)

    def split_set_zoom(self, zoom):
        zoom = max(0.2, min(zoom, 8.0))
        if abs(zoom - self.split_zoom) < 0.001:
            return
        self.split_zoom = zoom
        self.split_redraw()

    def split_on_mousewheel(self, event):
        delta = event.delta
        if delta == 0:
            return
        factor = 1.1 if delta > 0 else 0.9
        self.split_set_zoom(self.split_zoom * factor)

    def split_hide_preview(self):
        if self.split_preview_window and self.split_preview_window.winfo_exists():
            self.split_preview_window.destroy()
        self.split_preview_window = None
        self.split_preview_label = None
        self.split_preview_photo = None

    def split_show_preview(self, row, col):
        grid = self.split_get_grid()
        if not grid or self.split_sheet_pil is None or ImageTk is None:
            return
        x = grid["offset_x"] + col * grid["step_x"]
        y = grid["offset_y"] + row * grid["step_y"]
        crop = self.split_sheet_pil.crop((x, y, x + grid["cell_w"], y + grid["cell_h"]))
        max_size = 220
        scale = min(1.0, max_size / max(crop.width, crop.height)) if max(crop.width, crop.height) > 0 else 1.0
        if scale < 1.0:
            width = max(1, int(round(crop.width * scale)))
            height = max(1, int(round(crop.height * scale)))
            resample = Image.NEAREST
            if hasattr(Image, "Resampling"):
                resample = Image.Resampling.NEAREST
            crop = crop.resize((width, height), resample)
        photo = ImageTk.PhotoImage(crop)
        if not self.split_preview_window or not self.split_preview_window.winfo_exists():
            self.split_preview_window = tk.Toplevel(self.root)
            self.split_preview_window.title("Cell Preview")
            self.split_preview_window.resizable(False, False)
            self.split_preview_label = ttk.Label(self.split_preview_window)
            self.split_preview_label.grid(row=0, column=0, padx=6, pady=6)
        self.split_preview_photo = photo
        self.split_preview_label.configure(image=self.split_preview_photo)
        try:
            x_pos = self.root.winfo_pointerx() + 12
            y_pos = self.root.winfo_pointery() + 12
            self.split_preview_window.geometry(f"+{x_pos}+{y_pos}")
        except Exception:
            pass

    def split_find_runs(self, values):
        runs = []
        start = None
        for idx, flag in enumerate(values):
            if flag and start is None:
                start = idx
            elif not flag and start is not None:
                runs.append((start, idx - 1))
                start = None
        if start is not None:
            runs.append((start, len(values) - 1))
        return runs

    def split_pick_common_length(self, runs):
        lengths = {}
        for start, end in runs:
            length = end - start + 1
            lengths[length] = lengths.get(length, 0) + 1
        if not lengths:
            return 0
        return sorted(lengths.items(), key=lambda item: (-item[1], -item[0]))[0][0]

    def split_pick_common_gap(self, runs):
        gaps = {}
        for idx in range(1, len(runs)):
            gap = runs[idx][0] - runs[idx - 1][1] - 1
            if gap < 0:
                continue
            gaps[gap] = gaps.get(gap, 0) + 1
        if not gaps:
            return 0
        return sorted(gaps.items(), key=lambda item: (-item[1], -item[0]))[0][0]

    def split_auto_detect_grid(self, notify=True):
        if self.split_sheet_pil is None:
            if notify:
                messagebox.showerror("No spritesheet", "Load a spritesheet first.")
            return False
        alpha = self.split_sheet_pil.getchannel("A")
        width, height = alpha.size
        pixels = alpha.load()
        row_has = [False] * height
        col_has = [False] * width
        any_transparent = False
        for y in range(height):
            for x in range(width):
                a = pixels[x, y]
                if a != 0:
                    row_has[y] = True
                    col_has[x] = True
                else:
                    any_transparent = True
        if not any_transparent:
            self.split_status_label.configure(text="Auto-detect failed: no transparency found.")
            if notify:
                messagebox.showerror("Auto-detect failed", "No transparency was detected in the spritesheet.")
            return False

        row_runs = self.split_find_runs(row_has)
        col_runs = self.split_find_runs(col_has)
        if not row_runs or not col_runs:
            self.split_status_label.configure(text="Auto-detect failed: no grid runs found.")
            if notify:
                messagebox.showerror("Auto-detect failed", "Could not detect grid rows/columns.")
            return False

        cell_h = self.split_pick_common_length(row_runs)
        cell_w = self.split_pick_common_length(col_runs)
        rows = len(row_runs)
        cols = len(col_runs)

        if cell_w <= 0 or cell_h <= 0:
            self.split_status_label.configure(text="Auto-detect failed: invalid cell size.")
            if notify:
                messagebox.showerror("Auto-detect failed", "Could not determine a valid cell size.")
            return False

        self.split_cell_w_var.set(cell_w)
        self.split_cell_h_var.set(cell_h)
        self.split_columns_var.set(cols)
        self.split_rows_var.set(rows)
        self.split_selected_cells.clear()
        self.split_status_label.configure(text=f"Auto-detected {cols}x{rows} cells at {cell_w}x{cell_h}")
        self.split_redraw()
        return True

    def split_get_grid(self):
        if self.split_sheet_pil is None:
            return None
        try:
            cell_w = int(self.split_cell_w_var.get())
            cell_h = int(self.split_cell_h_var.get())
            cols = int(self.split_columns_var.get())
            rows = int(self.split_rows_var.get())
            pad_x = max(0, int(self.split_pad_x_var.get()))
            pad_y = max(0, int(self.split_pad_y_var.get()))
            offset_x = max(0, int(self.split_offset_x_var.get()))
            offset_y = max(0, int(self.split_offset_y_var.get()))
        except ValueError:
            return None
        if cell_w <= 0 or cell_h <= 0:
            return None
        step_x = cell_w + pad_x
        step_y = cell_h + pad_y
        sheet_w, sheet_h = self.split_sheet_pil.size
        if cols <= 0:
            avail_w = sheet_w - offset_x
            cols = int((avail_w + pad_x) // step_x) if step_x > 0 and avail_w > 0 else 0
        if rows <= 0:
            avail_h = sheet_h - offset_y
            rows = int((avail_h + pad_y) // step_y) if step_y > 0 and avail_h > 0 else 0
        grid_w = (cols * cell_w) + (pad_x * max(0, cols - 1))
        grid_h = (rows * cell_h) + (pad_y * max(0, rows - 1))
        return {
            "cell_w": cell_w,
            "cell_h": cell_h,
            "pad_x": pad_x,
            "pad_y": pad_y,
            "offset_x": offset_x,
            "offset_y": offset_y,
            "cols": cols,
            "rows": rows,
            "grid_w": grid_w,
            "grid_h": grid_h,
            "step_x": step_x,
            "step_y": step_y,
            "sheet_w": sheet_w,
            "sheet_h": sheet_h,
        }

    def split_push_undo(self, snapshot):
        if not snapshot:
            return
        if self.split_undo_stack and self.split_undo_stack[-1] == snapshot:
            return
        self.split_undo_stack.append(snapshot)
        if len(self.split_undo_stack) > 10:
            self.split_undo_stack.pop(0)

    def split_undo(self, _event=None):
        if not self.split_undo_stack:
            return
        snapshot = self.split_undo_stack.pop()
        self.split_cell_w_var.set(int(snapshot.get("cell_w", self.split_cell_w_var.get())))
        self.split_cell_h_var.set(int(snapshot.get("cell_h", self.split_cell_h_var.get())))
        self.split_redraw()

    def split_clear_selection(self):
        self.split_selected_cells.clear()
        self.split_redraw()

    def split_select_all(self):
        grid = self.split_get_grid()
        if not grid or grid["cols"] <= 0 or grid["rows"] <= 0:
            return
        self.split_selected_cells = {
            (row, col)
            for row in range(grid["rows"])
            for col in range(grid["cols"])
        }
        self.split_redraw()

    def split_cell_from_point(self, canvas_x, canvas_y):
        grid = self.split_get_grid()
        if not grid or grid["cols"] <= 0 or grid["rows"] <= 0:
            return None
        x = canvas_x / self.split_zoom
        y = canvas_y / self.split_zoom
        rel_x = x - grid["offset_x"]
        rel_y = y - grid["offset_y"]
        if rel_x < 0 or rel_y < 0:
            return None
        step_x = grid["step_x"]
        step_y = grid["step_y"]
        if step_x <= 0 or step_y <= 0:
            return None
        col = int(rel_x // step_x)
        row = int(rel_y // step_y)
        if col < 0 or row < 0 or col >= grid["cols"] or row >= grid["rows"]:
            return None
        if rel_x - (col * step_x) >= grid["cell_w"]:
            return None
        if rel_y - (row * step_y) >= grid["cell_h"]:
            return None
        return row, col

    def split_resize_hit(self, canvas_x, canvas_y):
        grid = self.split_get_grid()
        if not grid or grid["cols"] <= 0 or grid["rows"] <= 0:
            return None
        x = canvas_x / self.split_zoom
        y = canvas_y / self.split_zoom
        threshold = max(1.0, 6.0 / self.split_zoom)
        best = None

        for col in range(grid["cols"]):
            left = grid["offset_x"] + col * grid["step_x"]
            right = left + grid["cell_w"]
            dist = abs(x - right)
            if dist <= threshold and (best is None or dist < best[0]):
                best = (dist, "x", left)

        for row in range(grid["rows"]):
            top = grid["offset_y"] + row * grid["step_y"]
            bottom = top + grid["cell_h"]
            dist = abs(y - bottom)
            if dist <= threshold and (best is None or dist < best[0]):
                best = (dist, "y", top)

        if best is None:
            return None
        return best[1], best[2]

    def split_on_press(self, event):
        canvas_x = self.split_canvas.canvasx(event.x)
        canvas_y = self.split_canvas.canvasy(event.y)
        hit = self.split_resize_hit(canvas_x, canvas_y)
        if hit:
            self.split_resize_axis, self.split_resize_anchor = hit
            self.split_resize_start = (
                int(self.split_cell_w_var.get()),
                int(self.split_cell_h_var.get()),
            )
            self.split_resize_moved = False
            return
        cell = self.split_cell_from_point(canvas_x, canvas_y)
        add = bool(event.state & 0x0001)
        toggle = bool(event.state & 0x0004)
        if cell is None:
            if not add and not toggle:
                self.split_selected_cells.clear()
                self.split_hide_preview()
                self.split_redraw()
            return
        if toggle and cell in self.split_selected_cells:
            self.split_selected_cells.remove(cell)
        elif add and cell in self.split_selected_cells:
            pass
        elif add:
            self.split_selected_cells.add(cell)
        else:
            self.split_selected_cells = {cell}
        self.split_redraw()
        self.split_show_preview(cell[0], cell[1])

    def split_on_drag(self, event):
        if not self.split_resize_axis:
            return
        canvas_x = self.split_canvas.canvasx(event.x)
        canvas_y = self.split_canvas.canvasy(event.y)
        x = canvas_x / self.split_zoom
        y = canvas_y / self.split_zoom
        if self.split_resize_axis == "x":
            new_w = int(round(x - self.split_resize_anchor))
            new_w = max(1, new_w)
            if (
                self.split_resize_start
                and not self.split_resize_moved
                and new_w != self.split_resize_start[0]
            ):
                self.split_push_undo(
                    {
                        "cell_w": self.split_resize_start[0],
                        "cell_h": self.split_resize_start[1],
                    }
                )
                self.split_resize_moved = True
            self.split_cell_w_var.set(new_w)
        else:
            new_h = int(round(y - self.split_resize_anchor))
            new_h = max(1, new_h)
            if (
                self.split_resize_start
                and not self.split_resize_moved
                and new_h != self.split_resize_start[1]
            ):
                self.split_push_undo(
                    {
                        "cell_w": self.split_resize_start[0],
                        "cell_h": self.split_resize_start[1],
                    }
                )
                self.split_resize_moved = True
            self.split_cell_h_var.set(new_h)
        self.split_redraw()

    def split_on_release(self, _event):
        self.split_resize_axis = None
        self.split_resize_anchor = 0
        self.split_resize_start = None
        self.split_resize_moved = False
    def split_redraw(self):
        self.split_canvas.delete("split_grid")
        self.split_canvas.delete("split_selection")
        if self.split_sheet_pil is None or ImageTk is None:
            self.split_canvas.configure(scrollregion=(0, 0, 1, 1))
            return
        self.split_sheet_photo = self.split_make_photo(self.split_sheet_pil)
        if self.split_sheet_canvas_id is None:
            self.split_sheet_canvas_id = self.split_canvas.create_image(
                0,
                0,
                anchor="nw",
                image=self.split_sheet_photo,
                tags=("split_image",),
            )
        else:
            self.split_canvas.itemconfigure(self.split_sheet_canvas_id, image=self.split_sheet_photo)
        grid = self.split_get_grid()
        sheet_w = self.split_sheet_pil.width
        sheet_h = self.split_sheet_pil.height
        canvas_w = sheet_w
        canvas_h = sheet_h
        if grid:
            canvas_w = max(canvas_w, grid["offset_x"] + grid["grid_w"])
            canvas_h = max(canvas_h, grid["offset_y"] + grid["grid_h"])
        self.split_canvas.configure(scrollregion=(0, 0, canvas_w * self.split_zoom, canvas_h * self.split_zoom))
        if grid and self.split_show_grid_var.get() and grid["cols"] > 0 and grid["rows"] > 0:
            for col in range(grid["cols"]):
                x0 = grid["offset_x"] + col * grid["step_x"]
                x1 = x0 + grid["cell_w"]
                x0s = x0 * self.split_zoom
                x1s = x1 * self.split_zoom
                self.split_canvas.create_line(x0s, 0, x0s, canvas_h * self.split_zoom, fill="#808080", tags=("split_grid",))
                self.split_canvas.create_line(x1s, 0, x1s, canvas_h * self.split_zoom, fill="#808080", tags=("split_grid",))
            for row in range(grid["rows"]):
                y0 = grid["offset_y"] + row * grid["step_y"]
                y1 = y0 + grid["cell_h"]
                y0s = y0 * self.split_zoom
                y1s = y1 * self.split_zoom
                self.split_canvas.create_line(0, y0s, canvas_w * self.split_zoom, y0s, fill="#808080", tags=("split_grid",))
                self.split_canvas.create_line(0, y1s, canvas_w * self.split_zoom, y1s, fill="#808080", tags=("split_grid",))
            self.split_canvas.tag_raise("split_grid")
        if grid:
            cleaned = set()
            for row, col in self.split_selected_cells:
                if row < 0 or col < 0 or row >= grid["rows"] or col >= grid["cols"]:
                    continue
                cleaned.add((row, col))
            self.split_selected_cells = cleaned
            for row, col in self.split_selected_cells:
                x0 = (grid["offset_x"] + col * grid["step_x"]) * self.split_zoom
                y0 = (grid["offset_y"] + row * grid["step_y"]) * self.split_zoom
                x1 = x0 + grid["cell_w"] * self.split_zoom
                y1 = y0 + grid["cell_h"] * self.split_zoom
                self.split_canvas.create_rectangle(
                    x0,
                    y0,
                    x1,
                    y1,
                    outline="#3d7bfd",
                    width=2,
                    tags=("split_selection",),
                )
        if not self.split_selected_cells:
            self.split_hide_preview()

    def split_export_selected(self):
        if not self.split_selected_cells:
            messagebox.showerror("No selection", "Select one or more cells to export.")
            return
        self.split_export_cells(self.split_selected_cells, skip_blank=False)

    def split_export_all(self):
        grid = self.split_get_grid()
        if not grid or grid["cols"] <= 0 or grid["rows"] <= 0:
            messagebox.showerror("Grid error", "Set a valid grid before exporting.")
            return
        all_cells = {(row, col) for row in range(grid["rows"]) for col in range(grid["cols"])}
        self.split_export_cells(all_cells, skip_blank=True)

    def split_export_cells(self, cells, skip_blank=False):
        if self.split_sheet_pil is None:
            messagebox.showerror("No spritesheet", "Load a spritesheet first.")
            return
        grid = self.split_get_grid()
        if not grid or grid["cols"] <= 0 or grid["rows"] <= 0:
            messagebox.showerror("Grid error", "Set a valid grid before exporting.")
            return
        base_name = self.split_base_name_var.get().strip()
        if not base_name:
            base_name = os.path.splitext(os.path.basename(self.split_sheet_path_var.get().strip()))[0]
        output_dir = os.path.join(os.path.dirname(self.split_sheet_path_var.get().strip()), f"{base_name}_sprite_frames")
        self.split_output_dir_var.set(output_dir)
        os.makedirs(output_dir, exist_ok=True)
        exported = 0
        skipped = 0
        exported_cells = []
        for row, col in sorted(cells):
            if row < 0 or col < 0 or row >= grid["rows"] or col >= grid["cols"]:
                continue
            x = grid["offset_x"] + col * grid["step_x"]
            y = grid["offset_y"] + row * grid["step_y"]
            crop = self.split_sheet_pil.crop((x, y, x + grid["cell_w"], y + grid["cell_h"]))
            if skip_blank:
                alpha = crop.getchannel("A")
                if alpha.getbbox() is None:
                    skipped += 1
                    continue
            out_name = f"{base_name}_r{row + 1}_c{col + 1}.png"
            out_path = os.path.join(output_dir, out_name)
            crop.save(out_path, format="PNG")
            exported += 1
            exported_cells.append({"row": row, "col": col, "file": out_name})
        self.split_write_export_json(output_dir, base_name, grid, exported_cells, skipped, len(cells))
        if skip_blank:
            self.split_status_label.configure(text=f"Exported {exported} frame(s), skipped {skipped} empty")
        else:
            self.split_status_label.configure(text=f"Exported {exported} frame(s)")

    def split_write_export_json(self, output_dir, base_name, grid, exported_cells, skipped, requested):
        payload = {
            "sheet": self.split_sheet_path_var.get().strip(),
            "base_name": base_name,
            "cell_w": grid["cell_w"],
            "cell_h": grid["cell_h"],
            "columns": grid["cols"],
            "rows": grid["rows"],
            "pad_x": grid["pad_x"],
            "pad_y": grid["pad_y"],
            "offset_x": grid["offset_x"],
            "offset_y": grid["offset_y"],
            "exported_count": len(exported_cells),
            "requested_count": requested,
            "skipped_empty": skipped,
            "exported": exported_cells,
        }
        out_path = os.path.join(output_dir, f"{base_name}_sprite_frames.json")
        with open(out_path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)

    def split_save_selection(self):
        if self.split_sheet_pil is None:
            messagebox.showerror("No spritesheet", "Load a spritesheet first.")
            return
        if not self.split_selected_cells:
            messagebox.showerror("No selection", "Select one or more cells to save.")
            return
        grid = self.split_get_grid()
        if not grid:
            messagebox.showerror("Grid error", "Set a valid grid before saving.")
            return
        base_name = self.split_base_name_var.get().strip()
        if not base_name:
            base_name = os.path.splitext(os.path.basename(self.split_sheet_path_var.get().strip()))[0]
        output_dir = os.path.join(os.path.dirname(self.split_sheet_path_var.get().strip()), f"{base_name}_sprite_frames")
        os.makedirs(output_dir, exist_ok=True)
        payload = {
            "sheet": self.split_sheet_path_var.get().strip(),
            "base_name": base_name,
            "cell_w": grid["cell_w"],
            "cell_h": grid["cell_h"],
            "columns": grid["cols"],
            "rows": grid["rows"],
            "pad_x": grid["pad_x"],
            "pad_y": grid["pad_y"],
            "offset_x": grid["offset_x"],
            "offset_y": grid["offset_y"],
            "selected": sorted([{"row": row, "col": col} for row, col in self.split_selected_cells], key=lambda r: (r["row"], r["col"])),
        }
        out_path = os.path.join(output_dir, f"{base_name}_selection.json")
        with open(out_path, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)
        self.split_status_label.configure(text=f"Saved selection to {os.path.basename(out_path)}")

    def bind_color_swatch_updates(self):
        for var in (self.color_list_var, self.fill_color_var, self.replace_from_var, self.replace_to_var):
            var.trace_add("write", lambda *_: self.update_swatches())

    def update_swatch(self, label, hex_value):
        try:
            color = normalize_hex(hex_value)
        except ValueError:
            color = self.theme_colors.get("border", "#777777")
        label.configure(background=color)

    def update_swatches(self):
        if not hasattr(self, "color_list_swatch"):
            return
        first_color = first_hex_from_list(self.color_list_var.get()) or ""
        self.update_swatch(self.color_list_swatch, first_color)
        self.update_swatch(self.fill_color_swatch, self.fill_color_var.get())
        self.update_swatch(self.replace_from_swatch, self.replace_from_var.get())
        self.update_swatch(self.replace_to_swatch, self.replace_to_var.get())

    def setup_tooltips(self, transparent_radio, fill_radio, replace_radio, process_all_btn, skip_existing_btn, skip_files_btn, dry_run_btn, tooltips_btn, preview_btn, scan_prefixes_btn, keep_prefixes_check):
        Tooltip(transparent_radio, lambda: "Turn matching colors fully transparent.", self.tooltips_enabled_var)
        Tooltip(fill_radio, lambda: "Fill transparent pixels with a solid color.", self.tooltips_enabled_var)
        Tooltip(replace_radio, lambda: "Swap one exact color for another.", self.tooltips_enabled_var)
        Tooltip(process_all_btn, lambda: "When enabled, ignore individual folder checkboxes.", self.tooltips_enabled_var)
        Tooltip(self.color_list_entry, lambda: "Exact colors to change (comma or space separated).", self.tooltips_enabled_var)
        Tooltip(self.fill_color_entry, lambda: "Fill color used when Fill mode is selected.", self.tooltips_enabled_var)
        Tooltip(self.fill_shadows_check, lambda: "Experimental: fills semi-transparent pixels with a solid color.", self.tooltips_enabled_var)
        Tooltip(self.replace_from_entry, lambda: "Exact source color to replace.", self.tooltips_enabled_var)
        Tooltip(self.replace_to_entry, lambda: "Exact destination color.", self.tooltips_enabled_var)
        Tooltip(self.extra_prefix_entry, lambda: "Extra prefixes to detect (comma-separated).", self.tooltips_enabled_var)
        Tooltip(skip_existing_btn, lambda: "Skip prefix-named folders already in the output.", self.tooltips_enabled_var)
        Tooltip(skip_files_btn, lambda: "Skip files that already exist in the output.", self.tooltips_enabled_var)
        Tooltip(dry_run_btn, lambda: "Scan and log only. No files are written.", self.tooltips_enabled_var)
        Tooltip(tooltips_btn, lambda: "Toggle tooltip hints on or off.", self.tooltips_enabled_var)
        Tooltip(preview_btn, lambda: "Preview how files will be grouped by prefix.", self.tooltips_enabled_var)
        Tooltip(scan_prefixes_btn, lambda: "Scan input folders and list detected prefixes.", self.tooltips_enabled_var)
        Tooltip(keep_prefixes_check, lambda: "Keep full filenames instead of stripping prefixes.", self.tooltips_enabled_var)

    def get_selected_prefixes(self):
        prefixes = [name for name, var in self.prefix_vars.items() if var.get()]
        extra = [p.strip() for p in self.extra_prefix_var.get().split(",") if p.strip()]
        prefixes.extend(extra)
        return prefixes

    def render_prefix_checkboxes(self):
        for child in self.prefix_list_frame.winfo_children():
            child.destroy()
        if not self.prefix_vars:
            ttk.Label(self.prefix_list_frame, text="No prefixes detected. Click Scan for prefixes.").grid(row=0, column=0, sticky="w")
            return
        col_count = 4
        for idx, name in enumerate(sorted(self.prefix_vars, key=str.lower)):
            row = idx // col_count
            col = idx % col_count
            ttk.Checkbutton(self.prefix_list_frame, text=name, variable=self.prefix_vars[name]).grid(row=row, column=col, padx=4, sticky="w")

    def set_prefixes(self, prefix_states):
        self.prefix_vars = {}
        for name in sorted(prefix_states, key=str.lower):
            self.prefix_vars[name] = tk.BooleanVar(value=bool(prefix_states[name]))
        self.render_prefix_checkboxes()

    def scan_prefixes(self):
        input_root = self.input_var.get().strip()
        if not os.path.isdir(input_root):
            messagebox.showerror("Invalid input", "Input folder does not exist.")
            return

        rules = self.get_folder_rules()
        if self.process_all_var.get():
            allowed_dirs = None
        else:
            allowed_dirs = {name for name, rule in rules.items() if rule.get("include")}
        exclude_folders = {name.strip() for name in self.exclude_folders_var.get().split(",") if name.strip()}

        prefix_re = re.compile(r"^(?P<prefix>[^_]+)_\d+_")
        found = set()
        top_dirs = [d for d in os.listdir(input_root) if os.path.isdir(os.path.join(input_root, d))]
        top_dirs.sort(key=lambda s: s.lower())
        for top in top_dirs:
            if allowed_dirs is not None and top not in allowed_dirs:
                continue
            if top in exclude_folders:
                continue
            for root, _, files in os.walk(os.path.join(input_root, top)):
                for file in files:
                    if not file.lower().endswith(".png"):
                        continue
                    base = os.path.splitext(file)[0]
                    match = prefix_re.match(base)
                    if match:
                        found.add(match.group("prefix"))

        if not found:
            self.prefix_vars = {}
            self.render_prefix_checkboxes()
            messagebox.showinfo("Scan complete", "No prefixes were detected.")
            return

        existing_states = {name: var.get() for name, var in self.prefix_vars.items()}
        prefix_states = {name: existing_states.get(name, True) for name in found}
        self.set_prefixes(prefix_states)
        messagebox.showinfo("Scan complete", f"Detected {len(prefix_states)} prefixes.")

    def detect_conflicts(self, tasks):
        path_map = {}
        for task in tasks:
            if task["action"] == "skip_existing_file":
                continue
            path_map.setdefault(task["out_path"], []).append(task)
        return {path: items for path, items in path_map.items() if len(items) > 1}

    def apply_copy_suffixes(self, tasks, conflicts, output_root):
        used_paths = {task["out_path"] for task in tasks}
        for path, group in conflicts.items():
            base_root, ext = os.path.splitext(path)
            for task in group[1:]:
                copy_index = 1
                while True:
                    candidate = f"{base_root}_copy{copy_index}{ext}"
                    if candidate not in used_paths and not os.path.exists(candidate):
                        break
                    copy_index += 1
                used_paths.add(candidate)
                task["out_path"] = candidate
                task["rel_out"] = os.path.relpath(candidate, output_root)

    def pick_input(self):
        path = filedialog.askdirectory(initialdir=self.input_var.get() or os.getcwd())
        if path:
            self.input_var.set(path)
            self.refresh_folders()
            self.refresh_layout_folders()

    def pick_output(self):
        path = filedialog.askdirectory(initialdir=self.output_var.get() or os.getcwd())
        if path:
            self.output_var.set(path)

    def pick_csv(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialdir=APP_DIR,
        )
        if path:
            self.csv_path_var.set(path)

    def refresh_folders(self):
        existing = self.get_folder_rules()
        for child in self.folders_container.winfo_children():
            child.destroy()
        self.folder_rows.clear()

        root_path = self.input_var.get().strip()
        if not os.path.isdir(root_path):
            return

        dirs = [d for d in os.listdir(root_path) if os.path.isdir(os.path.join(root_path, d))]
        dirs.sort(key=lambda s: s.lower())
        header = ttk.Frame(self.folders_container)
        header.grid(row=0, column=0, sticky="ew", padx=2)
        ttk.Label(header, text="Include", width=8).grid(row=0, column=0, sticky="w")
        ttk.Label(header, text="Folder", width=30).grid(row=0, column=1, sticky="w")
        ttk.Label(header, text="Mode", width=16).grid(row=0, column=2, sticky="w")

        for idx, name in enumerate(dirs, start=1):
            include_var = tk.BooleanVar(value=True)
            mode_var = tk.StringVar(value="Process")
            if name in existing:
                include_var.set(existing[name].get("include", True))
                mode_var.set(existing[name].get("mode", "Process"))

            row = ttk.Frame(self.folders_container)
            row.grid(row=idx, column=0, sticky="ew", padx=2)

            ttk.Checkbutton(row, variable=include_var).grid(row=0, column=0, sticky="w")
            ttk.Label(row, text=name, width=30, anchor="w").grid(row=0, column=1, sticky="w")
            combo = ttk.Combobox(row, values=FOLDER_MODES, textvariable=mode_var, width=14, state="readonly", style="Folder.TCombobox", takefocus=False)
            combo.grid(row=0, column=2, sticky="w", padx=4)
            combo.bind("<<ComboboxSelected>>", lambda _e: self.root.focus_set())

            self.folder_rows[name] = {
                "include_var": include_var,
                "mode_var": mode_var,
            }

    def select_all_folders(self):
        for row in self.folder_rows.values():
            row["include_var"].set(True)

    def select_no_folders(self):
        for row in self.folder_rows.values():
            row["include_var"].set(False)

    def get_folder_rules(self):
        rules = {}
        for name, row in self.folder_rows.items():
            rules[name] = {
                "include": row["include_var"].get(),
            "mode": row["mode_var"].get(),
        }
        return rules

    def build_settings(self):
        if self.window_size is None:
            self.root.update_idletasks()
            width = self.root.winfo_width()
            height = self.root.winfo_height()
            if width >= 200 and height >= 200:
                self.window_size = (width, height)
        return {
            "input": self.input_var.get(),
            "output": self.output_var.get(),
            "mode": self.mode_var.get(),
            "color_list": self.color_list_var.get(),
            "fill_color": self.fill_color_var.get(),
            "fill_shadows": self.fill_shadows_var.get(),
            "replace_from": self.replace_from_var.get(),
            "replace_to": self.replace_to_var.get(),
            "extra_prefixes": self.extra_prefix_var.get(),
            "keep_prefixes": self.keep_prefixes_var.get(),
            "prefixes": {name: var.get() for name, var in self.prefix_vars.items()},
            "process_all": self.process_all_var.get(),
            "skip_existing_folders": self.skip_existing_var.get(),
            "skip_existing_files": self.skip_existing_files_var.get(),
            "dry_run": self.dry_run_var.get(),
            "csv_log": self.csv_log_var.get(),
            "csv_path": self.csv_path_var.get(),
            "exclude_folders": self.exclude_folders_var.get(),
            "sprite_layout": self.sprite_layout_var.get(),
            "sprite_columns": self.sprite_columns_var.get(),
            "sprite_padding_mode": self.sprite_padding_mode_var.get(),
            "sprite_padding": self.sprite_padding_var.get(),
            "tile_layout": self.tile_layout_var.get(),
            "tile_columns": self.tile_columns_var.get(),
            "tile_size": self.tile_size_var.get(),
            "tile_export_meta": self.tile_export_meta_var.get(),
            "layout_type": self.layout_type_var.get(),
            "layout_folder": self.layout_folder_var.get(),
            "layout_output_name": self.layout_output_var.get(),
            "layout_mode": self.layout_mode_var.get(),
            "layout_columns": self.layout_grid_columns_var.get(),
            "layout_rows": self.layout_grid_rows_var.get(),
            "layout_snap": self.layout_snap_var.get(),
            "layout_show_grid": self.layout_show_grid_var.get(),
            "layout_show_guides": self.layout_show_guides_var.get(),
            "tile_size_mode": self.tile_size_mode_var.get(),
            "split_sheet_path": self.split_sheet_path_var.get(),
            "split_output_dir": self.split_output_dir_var.get(),
            "split_base_name": self.split_base_name_var.get(),
            "split_cell_w": self.split_cell_w_var.get(),
            "split_cell_h": self.split_cell_h_var.get(),
            "split_columns": self.split_columns_var.get(),
            "split_rows": self.split_rows_var.get(),
            "split_pad_x": self.split_pad_x_var.get(),
            "split_pad_y": self.split_pad_y_var.get(),
            "split_offset_x": self.split_offset_x_var.get(),
            "split_offset_y": self.split_offset_y_var.get(),
            "split_show_grid": self.split_show_grid_var.get(),
            "recent_presets": self.recent_presets,
            "tooltips_enabled": self.tooltips_enabled_var.get(),
            "theme": self.theme_var.get(),
            "window_size": list(self.window_size) if self.window_size else None,
            "folder_rules": self.get_folder_rules(),
        }

    def apply_settings(self, data):
        self.input_var.set(data.get("input", DEFAULT_INPUT))
        self.output_var.set(data.get("output", DEFAULT_OUTPUT))
        self.mode_var.set(data.get("mode", "transparent"))
        self.color_list_var.set(data.get("color_list", DEFAULT_COLOR))
        self.fill_color_var.set(data.get("fill_color", ""))
        self.fill_shadows_var.set(bool(data.get("fill_shadows", False)))
        self.replace_from_var.set(data.get("replace_from", ""))
        self.replace_to_var.set(data.get("replace_to", ""))
        self.extra_prefix_var.set(data.get("extra_prefixes", ""))
        self.keep_prefixes_var.set(bool(data.get("keep_prefixes", False)))
        self.process_all_var.set(bool(data.get("process_all", True)))
        self.skip_existing_var.set(bool(data.get("skip_existing_folders", True)))
        self.skip_existing_files_var.set(bool(data.get("skip_existing_files", False)))
        self.dry_run_var.set(bool(data.get("dry_run", False)))
        self.csv_log_var.set(bool(data.get("csv_log", False)))
        self.csv_path_var.set(data.get("csv_path", ""))
        self.exclude_folders_var.set(data.get("exclude_folders", ""))
        self.sprite_layout_var.set(data.get("sprite_layout", "Grid"))
        self.sprite_columns_var.set(int(data.get("sprite_columns", 4)))
        self.sprite_padding_mode_var.set(data.get("sprite_padding_mode", "Fixed"))
        self.sprite_padding_var.set(int(data.get("sprite_padding", 1)))
        self.tile_layout_var.set(data.get("tile_layout", "Grid"))
        self.tile_columns_var.set(int(data.get("tile_columns", 15)))
        self.tile_size_var.set(int(data.get("tile_size", 32)))
        self.tile_export_meta_var.set(bool(data.get("tile_export_meta", True)))
        self.layout_type_var.set(data.get("layout_type", "Spritesheet"))
        self.layout_folder_var.set(data.get("layout_folder", ""))
        self.layout_output_var.set(data.get("layout_output_name", ""))
        self.layout_mode_var.set(data.get("layout_mode", "Grid"))
        self.layout_grid_columns_var.set(int(data.get("layout_columns", 4)))
        self.layout_grid_rows_var.set(int(data.get("layout_rows", 0)))
        self.layout_snap_var.set(bool(data.get("layout_snap", True)))
        self.layout_show_grid_var.set(bool(data.get("layout_show_grid", True)))
        self.layout_show_guides_var.set(bool(data.get("layout_show_guides", False)))
        self.tile_size_mode_var.set(data.get("tile_size_mode", "Force"))
        self.split_sheet_path_var.set(data.get("split_sheet_path", ""))
        self.split_output_dir_var.set(data.get("split_output_dir", ""))
        self.split_base_name_var.set(data.get("split_base_name", ""))
        self.split_cell_w_var.set(int(data.get("split_cell_w", 32)))
        self.split_cell_h_var.set(int(data.get("split_cell_h", 32)))
        self.split_columns_var.set(int(data.get("split_columns", 0)))
        self.split_rows_var.set(int(data.get("split_rows", 0)))
        self.split_pad_x_var.set(int(data.get("split_pad_x", 0)))
        self.split_pad_y_var.set(int(data.get("split_pad_y", 0)))
        self.split_offset_x_var.set(int(data.get("split_offset_x", 0)))
        self.split_offset_y_var.set(int(data.get("split_offset_y", 0)))
        self.split_show_grid_var.set(bool(data.get("split_show_grid", True)))
        self.tooltips_enabled_var.set(bool(data.get("tooltips_enabled", True)))
        self.theme_var.set(data.get("theme", "Light"))
        window_size = data.get("window_size")
        if isinstance(window_size, (list, tuple)) and len(window_size) == 2:
            width, height = window_size
            if isinstance(width, int) and isinstance(height, int) and width >= 200 and height >= 200:
                self.root.geometry(f"{width}x{height}")
                self.window_size = (width, height)

        recent = data.get("recent_presets", [])
        if isinstance(recent, list):
            self.recent_presets = recent[:]
            self.update_recent_presets_dropdown()

        prefixes = data.get("prefixes", {})
        if isinstance(prefixes, dict) and prefixes:
            self.set_prefixes(prefixes)
        elif not self.prefix_vars:
            self.render_prefix_checkboxes()

        self.refresh_folders()
        rules = data.get("folder_rules", {})
        for name, rule in rules.items():
            if name in self.folder_rows:
                self.folder_rows[name]["include_var"].set(rule.get("include", True))
                self.folder_rows[name]["mode_var"].set(rule.get("mode", "Process"))

        self.refresh_layout_folders()
        self.update_color_mode()
        self.update_sprite_layout_controls()
        self.update_sprite_padding_controls()
        self.update_tile_layout_controls()
        self.update_layout_type_controls()
        self.apply_theme(self.theme_var.get())
        split_path = self.split_sheet_path_var.get().strip()
        if split_path and os.path.isfile(split_path):
            self.split_load_sheet(split_path)
        else:
            self.split_redraw()

    def save_preset(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".json",
            filetypes=[("JSON files", "*.json")],
            initialdir=APP_DIR,
        )
        if not path:
            return
        data = self.build_settings()
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2)
        self.record_recent_preset(path)
        messagebox.showinfo("Preset saved", f"Saved preset to:\n{path}")

    def load_preset(self):
        path = filedialog.askopenfilename(
            filetypes=[("JSON files", "*.json")],
            initialdir=APP_DIR,
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception as exc:
            messagebox.showerror("Preset error", f"Failed to load preset:\n{exc}")
            return
        self.apply_settings(data)
        self.record_recent_preset(path)

    def load_recent_preset(self):
        path = self.recent_preset_var.get().strip()
        if not path:
            return
        if not os.path.exists(path):
            messagebox.showerror("Preset missing", f"Preset not found:\n{path}")
            return
        try:
            with open(path, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception as exc:
            messagebox.showerror("Preset error", f"Failed to load preset:\n{exc}")
            return
        self.apply_settings(data)

    def record_recent_preset(self, path):
        path = os.path.abspath(path)
        self.recent_presets = [p for p in self.recent_presets if p != path]
        self.recent_presets.insert(0, path)
        self.recent_presets = self.recent_presets[:RECENT_PRESETS_MAX]
        self.update_recent_presets_dropdown()

    def update_recent_presets_dropdown(self):
        self.recent_combo["values"] = self.recent_presets
        if self.recent_presets and not self.recent_preset_var.get():
            self.recent_preset_var.set(self.recent_presets[0])

    def save_last_settings(self):
        data = self.build_settings()
        os.makedirs(SETTINGS_DIR, exist_ok=True)
        with open(LAST_SETTINGS_PATH, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2)

    def write_settings_or_error(self):
        try:
            self.save_last_settings()
            return True
        except Exception:
            os.makedirs(SETTINGS_DIR, exist_ok=True)
            with open(ERROR_LOG_PATH, "w", encoding="utf-8") as handle:
                handle.write(traceback.format_exc())
            return False

    def load_last_settings(self):
        if not os.path.exists(LAST_SETTINGS_PATH):
            return
        try:
            with open(LAST_SETTINGS_PATH, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except Exception:
            return
        self.apply_settings(data)

    def on_close(self):
        self.save_last_settings()
        self.root.destroy()

    def run(self):
        if self.running or self.sheet_running or self.tile_running:
            return
        if Image is None:
            messagebox.showerror("Pillow not installed", f"Install Pillow first.\n\n{PIL_IMPORT_ERROR}")
            return

        input_root = self.input_var.get().strip()
        output_root = self.output_var.get().strip()
        mode = self.mode_var.get()

        if not os.path.isdir(input_root):
            messagebox.showerror("Invalid input", "Input folder does not exist.")
            return
        if not output_root:
            messagebox.showerror("Invalid output", "Output folder is empty.")
            return

        if mode in ("transparent", "fill"):
            try:
                color_list = parse_hex_list(self.color_list_var.get())
            except ValueError as exc:
                messagebox.showerror("Invalid color list", str(exc))
                return
            if not color_list:
                messagebox.showerror("Invalid color list", "Provide at least one hex color.")
                return
            target_rgbs = [hex_to_rgb(c) for c in color_list]
            replace_pair = None
            if mode == "fill":
                try:
                    fill_color = normalize_hex(self.fill_color_var.get())
                except ValueError as exc:
                    messagebox.showerror("Invalid fill color", str(exc))
                    return
                fill_rgb = hex_to_rgb(fill_color)
            else:
                fill_rgb = None
        else:
            try:
                src = normalize_hex(self.replace_from_var.get())
                dst = normalize_hex(self.replace_to_var.get())
            except ValueError as exc:
                messagebox.showerror("Invalid replace colors", str(exc))
                return
            target_rgbs = []
            replace_pair = (hex_to_rgb(src), hex_to_rgb(dst))
            fill_rgb = None

        prefixes = self.get_selected_prefixes()
        prefix_re = build_prefix_regex(prefixes)
        keep_prefixes = self.keep_prefixes_var.get()

        rules = self.get_folder_rules()

        if self.process_all_var.get():
            allowed_dirs = None
        else:
            allowed_dirs = {name for name, rule in rules.items() if rule.get("include")}

        skip_existing = self.skip_existing_var.get()
        skip_existing_files = self.skip_existing_files_var.get()
        dry_run = self.dry_run_var.get()
        csv_enabled = self.csv_log_var.get()
        csv_path = self.csv_path_var.get().strip()
        exclude_folders = {name.strip() for name in self.exclude_folders_var.get().split(",") if name.strip()}
        if csv_enabled and not csv_path:
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            csv_path = os.path.join(APP_DIR, f"log-{timestamp}.csv")
            self.csv_path_var.set(csv_path)

        self.log.delete("1.0", tk.END)
        self.running = True
        self.run_button.configure(state="disabled")
        if hasattr(self, "preview_button"):
            self.preview_button.configure(state="disabled")
        if hasattr(self, "sheet_run_button"):
            self.sheet_run_button.configure(state="disabled")
        if hasattr(self, "tile_run_button"):
            self.tile_run_button.configure(state="disabled")
        self.progress.configure(value=0, maximum=1)
        self.status_label.configure(text="Scanning...")

        self.save_last_settings()

        thread = threading.Thread(
            target=self.worker,
            args=(
                input_root,
                output_root,
                mode,
                target_rgbs,
                replace_pair,
                fill_rgb,
                self.fill_shadows_var.get(),
                prefix_re,
                keep_prefixes,
                rules,
                allowed_dirs,
                skip_existing,
                skip_existing_files,
                exclude_folders,
                dry_run,
                csv_enabled,
                csv_path,
            ),
            daemon=True,
        )
        thread.start()

    def preview_routing(self):
        if self.running or self.previewing or self.sheet_running or self.tile_running:
            return

        input_root = self.input_var.get().strip()
        output_root = self.output_var.get().strip()

        if not os.path.isdir(input_root):
            messagebox.showerror("Invalid input", "Input folder does not exist.")
            return
        if not output_root:
            messagebox.showerror("Invalid output", "Output folder is empty.")
            return

        prefixes = self.get_selected_prefixes()
        prefix_re = build_prefix_regex(prefixes)
        keep_prefixes = self.keep_prefixes_var.get()
        rules = self.get_folder_rules()
        if self.process_all_var.get():
            allowed_dirs = None
        else:
            allowed_dirs = {name for name, rule in rules.items() if rule.get("include")}

        skip_existing = self.skip_existing_var.get()
        skip_existing_files = self.skip_existing_files_var.get()
        exclude_folders = {name.strip() for name in self.exclude_folders_var.get().split(",") if name.strip()}

        self.previewing = True
        self.run_button.configure(state="disabled")
        if hasattr(self, "preview_button"):
            self.preview_button.configure(state="disabled")
        self.status_label.configure(text="Previewing...")

        thread = threading.Thread(
            target=self.preview_worker,
            args=(
                input_root,
                output_root,
                prefix_re,
                keep_prefixes,
                rules,
                allowed_dirs,
                skip_existing,
                skip_existing_files,
                exclude_folders,
            ),
            daemon=True,
        )
        thread.start()

    def preview_worker(self, input_root, output_root, prefix_re, keep_prefixes, rules, allowed_dirs, skip_existing, skip_existing_files, exclude_folders):
        try:
            counts = {}
            total_scanned = 0
            skipped_existing_files = 0
            skipped_prefixes = set()
            existing_prefix_dirs = {}

            top_dirs = [d for d in os.listdir(input_root) if os.path.isdir(os.path.join(input_root, d))]
            top_dirs.sort(key=lambda s: s.lower())
            for top in top_dirs:
                if allowed_dirs is not None and top not in allowed_dirs:
                    continue
                if top in exclude_folders:
                    continue
                rule = rules.get(top, {})
                folder_mode = rule.get("mode", "Process")
                if folder_mode == "Skip":
                    continue
                for root, _, files in os.walk(os.path.join(input_root, top)):
                    for file in files:
                        if not file.lower().endswith(".png"):
                            continue
                        total_scanned += 1
                        prefix_folder = output_prefix_folder(file, prefix_re)
                        out_dir = output_root if not prefix_folder else os.path.join(output_root, prefix_folder)
                        if skip_existing and prefix_folder:
                            exists = existing_prefix_dirs.get(prefix_folder)
                            if exists is None:
                                exists = os.path.isdir(out_dir)
                                existing_prefix_dirs[prefix_folder] = exists
                            if exists:
                                skipped_prefixes.add(prefix_folder)
                                continue
                        out_name = output_filename(file, prefix_re, keep_prefixes)
                        out_path = os.path.join(out_dir, out_name)
                        if skip_existing_files and os.path.exists(out_path):
                            skipped_existing_files += 1
                            continue
                        counts[prefix_folder] = counts.get(prefix_folder, 0) + 1

            lines = [f"Total PNGs scanned: {total_scanned}"]
            if counts:
                lines.append("")
                lines.append("Routed by prefix:")
                sorted_items = sorted(counts.items(), key=lambda item: (item[0] != "needs_sorting", item[0]))
                for name, count in sorted_items:
                    lines.append(f"- {name}: {count}")
            else:
                lines.append("No files matched the current filters.")

            if skipped_prefixes:
                lines.append("")
                skipped_list = ", ".join(sorted(skipped_prefixes))
                lines.append(f"Skipped (existing prefix folders): {skipped_list}")
            if skipped_existing_files:
                lines.append(f"Skipped (existing files): {skipped_existing_files}")

            self.queue.put(("preview", "\n".join(lines)))
        except Exception as exc:
            self.queue.put(("preview_error", str(exc)))

    def run_sprite_sheets(self):
        if self.running or self.previewing or self.sheet_running or self.tile_running:
            return
        if Image is None:
            messagebox.showerror("Pillow not installed", f"Install Pillow first.\n\n{PIL_IMPORT_ERROR}")
            return

        input_root = self.input_var.get().strip()
        if not os.path.isdir(input_root):
            messagebox.showerror("Invalid input", "Input folder does not exist.")
            return

        layout_mode = self.sprite_layout_var.get()
        padding_mode = self.sprite_padding_mode_var.get()
        try:
            columns = max(1, int(self.sprite_columns_var.get()))
            padding_value = max(0, int(self.sprite_padding_var.get()))
        except ValueError:
            messagebox.showerror("Invalid columns", "Columns must be a number.")
            return
        exclude_folders = {name.strip() for name in self.exclude_folders_var.get().split(",") if name.strip()}

        self.sheet_log.delete("1.0", tk.END)
        self.sheet_running = True
        self.sheet_run_button.configure(state="disabled")
        self.sheet_status_label.configure(text="Scanning...")
        self.sheet_progress.start(10)
        self.run_button.configure(state="disabled")
        if hasattr(self, "preview_button"):
            self.preview_button.configure(state="disabled")
        self.tile_run_button.configure(state="disabled")

        thread = threading.Thread(
            target=self.sprite_worker,
            args=(input_root, layout_mode, columns, padding_mode, padding_value, exclude_folders),
            daemon=True,
        )
        thread.start()

    def sprite_worker(self, input_root, layout_mode, columns, padding_mode, padding_value, exclude_folders):
        try:
            sheet_count = 0
            folder_count = 0
            for root, dirs, files in os.walk(input_root):
                dirs[:] = [d for d in dirs if d not in ("sprite_sheets", "tilemaps")]
                folder_name = os.path.basename(root)
                if folder_name in exclude_folders:
                    dirs[:] = []
                    continue

                png_files = [f for f in files if f.lower().endswith(".png")]
                if not png_files:
                    continue
                folder_count += 1
                groups = {}
                for file in png_files:
                    prefix = extract_group_prefix(file)
                    groups.setdefault(prefix, []).append(file)

                for prefix, group_files in groups.items():
                    group_files.sort(key=sprite_sort_key)
                    images = []
                    try:
                        for file in group_files:
                            img = Image.open(os.path.join(root, file)).convert("RGBA")
                            images.append(img)
                        if not images:
                            continue
                        cell_w = max(img.width for img in images)
                        cell_h = max(img.height for img in images)
                        if padding_mode == "Frame width":
                            pad = cell_w
                        elif padding_mode == "Frame height":
                            pad = cell_h
                        else:
                            pad = padding_value

                        if layout_mode == "Horizontal":
                            cols = len(images)
                            rows = 1
                        elif layout_mode == "Vertical":
                            cols = 1
                            rows = len(images)
                        else:
                            cols = max(1, columns)
                            rows = int(math.ceil(len(images) / cols))

                        sheet_w = (cols * cell_w) + (pad * max(0, cols - 1))
                        sheet_h = (rows * cell_h) + (pad * max(0, rows - 1))
                        sheet = Image.new("RGBA", (sheet_w, sheet_h), (0, 0, 0, 0))
                        for idx, img in enumerate(images):
                            row = idx // cols
                            col = idx % cols
                            x = col * (cell_w + pad)
                            y = row * (cell_h + pad)
                            sheet.paste(img, (x, y))

                        out_dir = os.path.join(root, "sprite_sheets")
                        os.makedirs(out_dir, exist_ok=True)
                        out_name = f"{prefix}_Spritesheet.png"
                        out_path = os.path.join(out_dir, out_name)
                        sheet.save(out_path, format="PNG")
                        sheet_count += 1

                        rel_root = os.path.relpath(root, input_root)
                        rel_root = "." if rel_root == "." else rel_root
                        self.queue.put(("sheet_log", f"{rel_root} -> {out_name} ({len(images)} frames)"))
                    finally:
                        for img in images:
                            img.close()

            self.queue.put(("sheet_done", sheet_count, folder_count))
        except Exception as exc:
            self.queue.put(("sheet_error", str(exc)))

    def run_tilemaps(self):
        if self.running or self.previewing or self.sheet_running or self.tile_running:
            return
        if Image is None:
            messagebox.showerror("Pillow not installed", f"Install Pillow first.\n\n{PIL_IMPORT_ERROR}")
            return

        input_root = self.input_var.get().strip()
        if not os.path.isdir(input_root):
            messagebox.showerror("Invalid input", "Input folder does not exist.")
            return

        layout_mode = self.tile_layout_var.get()
        try:
            columns = int(self.tile_columns_var.get())
            tile_size = int(self.tile_size_var.get())
        except ValueError:
            messagebox.showerror("Invalid settings", "Tile size and columns must be numbers.")
            return
        export_meta = self.tile_export_meta_var.get()
        exclude_folders = {name.strip() for name in self.exclude_folders_var.get().split(",") if name.strip()}

        self.tile_log.delete("1.0", tk.END)
        self.tile_running = True
        self.tile_run_button.configure(state="disabled")
        self.tile_status_label.configure(text="Scanning...")
        self.tile_progress.start(10)
        self.run_button.configure(state="disabled")
        if hasattr(self, "preview_button"):
            self.preview_button.configure(state="disabled")
        self.sheet_run_button.configure(state="disabled")

        thread = threading.Thread(
            target=self.tile_worker,
            args=(input_root, layout_mode, columns, tile_size, export_meta, exclude_folders),
            daemon=True,
        )
        thread.start()

    def tile_worker(self, input_root, layout_mode, columns, tile_size, export_meta, exclude_folders):
        try:
            tilemap_count = 0
            folder_count = 0
            for root, dirs, files in os.walk(input_root):
                dirs[:] = [d for d in dirs if d not in ("sprite_sheets", "tilemaps")]
                folder_name = os.path.basename(root)
                if folder_name in exclude_folders:
                    dirs[:] = []
                    continue

                png_files = sorted([f for f in files if f.lower().endswith(".png")], key=lambda s: s.lower())
                if not png_files:
                    continue
                folder_count += 1
                images = []
                try:
                    for file in png_files:
                        img = Image.open(os.path.join(root, file)).convert("RGBA")
                        images.append((file, img))

                    if not images:
                        continue

                    max_w = max(img.width for _, img in images)
                    max_h = max(img.height for _, img in images)
                    cell_w = max(tile_size, max_w)
                    cell_h = max(tile_size, max_h)

                    if layout_mode == "Horizontal":
                        cols = len(images)
                        rows = 1
                    elif layout_mode == "Vertical":
                        cols = 1
                        rows = len(images)
                    else:
                        cols = max(1, columns)
                        rows = int(math.ceil(len(images) / cols))

                    tilemap = Image.new("RGBA", (cols * cell_w, rows * cell_h), (0, 0, 0, 0))
                    metadata = []
                    for idx, (file, img) in enumerate(images):
                        row = idx // cols
                        col = idx % cols
                        x = col * cell_w
                        y = row * cell_h
                        tilemap.paste(img, (x, y))
                        metadata.append((file, x, y, img.width, img.height))

                    out_dir = os.path.join(root, "tilemaps")
                    os.makedirs(out_dir, exist_ok=True)
                    out_name = f"{folder_name}_tilemap.png"
                    out_path = os.path.join(out_dir, out_name)
                    tilemap.save(out_path, format="PNG")

                    if export_meta:
                        meta_path = os.path.join(out_dir, f"{folder_name}_tilemap.csv")
                        with open(meta_path, "w", newline="", encoding="utf-8") as handle:
                            writer = csv.writer(handle)
                            writer.writerow(["tilemap", "tile", "x", "y", "width", "height"])
                            for file, x, y, w, h in metadata:
                                writer.writerow([out_name, file, x, y, w, h])

                    tilemap_count += 1
                    rel_root = os.path.relpath(root, input_root)
                    rel_root = "." if rel_root == "." else rel_root
                    self.queue.put(("tile_log", f"{rel_root} -> {out_name} ({len(images)} tiles)"))
                finally:
                    for _, img in images:
                        img.close()

            self.queue.put(("tile_done", tilemap_count, folder_count))
        except Exception as exc:
            self.queue.put(("tile_error", str(exc)))

    def worker(self, input_root, output_root, mode, target_rgbs, replace_pair, fill_rgb, fill_shadows, prefix_re, keep_prefixes, rules, allowed_dirs, skip_existing, skip_existing_files, exclude_folders, dry_run, csv_enabled, csv_path):
        tasks = []
        existing_prefix_dirs = {}
        top_dirs = [d for d in os.listdir(input_root) if os.path.isdir(os.path.join(input_root, d))]
        top_dirs.sort(key=lambda s: s.lower())
        for top in top_dirs:
            if allowed_dirs is not None and top not in allowed_dirs:
                continue
            if top in exclude_folders:
                continue
            rule = rules.get(top, {})
            folder_mode = rule.get("mode", "Process")
            if folder_mode == "Skip":
                continue
            custom_colors = rule.get("colors", "").strip()
            for root, _, files in os.walk(os.path.join(input_root, top)):
                for file in files:
                    if not file.lower().endswith(".png"):
                        continue
                    in_path = os.path.join(root, file)
                    prefix_folder = output_prefix_folder(file, prefix_re)
                    out_dir = output_root if not prefix_folder else os.path.join(output_root, prefix_folder)
                    if skip_existing and prefix_folder:
                        exists = existing_prefix_dirs.get(prefix_folder)
                        if exists is None:
                            exists = os.path.isdir(out_dir)
                            existing_prefix_dirs[prefix_folder] = exists
                        if exists:
                            continue
                    out_name = output_filename(file, prefix_re, keep_prefixes)
                    out_path = os.path.join(out_dir, out_name)
                    action = "process"
                    if skip_existing_files and os.path.exists(out_path):
                        action = "skip_existing_file"
                    tasks.append({
                        "in_path": in_path,
                        "out_path": out_path,
                        "rel_in": os.path.relpath(in_path, input_root),
                        "rel_out": os.path.relpath(out_path, output_root),
                        "folder_mode": folder_mode,
                        "custom_colors": custom_colors,
                        "action": action,
                    })

        total = len(tasks)
        if total == 0:
            self.queue.put(("log", "No files to process."))
            self.queue.put(("status", "Done"))
            self.queue.put(("done", 0, 0, 0))
            return

        conflicts = self.detect_conflicts(tasks)
        if conflicts:
            self.conflict_choice = None
            self.conflict_event.clear()
            conflict_files = sum(len(items) for items in conflicts.values())
            self.queue.put(("conflicts", len(conflicts), conflict_files))
            self.conflict_event.wait()
            if self.conflict_choice == "cancel" or self.conflict_choice is None:
                self.queue.put(("log", "Canceled due to naming conflicts."))
                self.queue.put(("status", "Canceled"))
                self.queue.put(("done", 0, 0, 0))
                return
            if self.conflict_choice == "copy":
                self.apply_copy_suffixes(tasks, conflicts, output_root)

        if csv_enabled:
            ensure_csv_header(csv_path)
            csv_handle = open(csv_path, "a", newline="", encoding="utf-8")
            csv_writer = csv.writer(csv_handle)
        else:
            csv_handle = None
            csv_writer = None

        start = time.time()
        processed = 0
        total_converted = 0
        logged_files = 0

        for task in tasks:
            action = task["action"]
            converted = 0

            if action == "skip_existing_file":
                pass
            elif task["folder_mode"] == "Rename only":
                if not dry_run:
                    os.makedirs(os.path.dirname(task["out_path"]), exist_ok=True)
                    shutil.copy2(task["in_path"], task["out_path"])
                action = "rename_only"
            else:
                active_colors = target_rgbs
                if mode == "transparent" and task["folder_mode"] == "Custom colors" and task["custom_colors"]:
                    try:
                        custom_list = parse_hex_list(task["custom_colors"])
                        active_colors = [hex_to_rgb(c) for c in custom_list]
                    except ValueError:
                        action = "invalid_custom_colors"
                    else:
                        converted = process_image(task["in_path"], task["out_path"], mode, active_colors, replace_pair, fill_rgb, fill_shadows, dry_run)
                else:
                    converted = process_image(task["in_path"], task["out_path"], mode, active_colors, replace_pair, fill_rgb, fill_shadows, dry_run)

            if converted > 0:
                logged_files += 1
                self.queue.put(("log", f"{task['rel_in']} -> {task['rel_out']} | converted: {converted}"))

            if csv_writer is not None:
                csv_writer.writerow([task["rel_in"], task["rel_out"], action, converted])

            total_converted += converted
            processed += 1

            if processed % 10 == 0 or processed == total:
                elapsed = time.time() - start
                eta = (elapsed / processed) * (total - processed) if processed else 0
                self.queue.put(("progress", processed, total, eta))

        if csv_handle is not None:
            csv_handle.close()

        self.queue.put(("status", "Done"))
        self.queue.put(("done", total, logged_files, total_converted))

    def poll_queue(self):
        try:
            while True:
                msg = self.queue.get_nowait()
                if isinstance(msg, tuple):
                    kind = msg[0]
                    if kind == "log":
                        self.log.insert(tk.END, msg[1] + "\n")
                        self.log.see(tk.END)
                    elif kind == "progress":
                        processed, total, eta = msg[1], msg[2], msg[3]
                        self.progress.configure(maximum=total, value=processed)
                        if total:
                            eta_str = time.strftime("%M:%S", time.gmtime(eta))
                            self.status_label.configure(text=f"{processed}/{total} | ETA {eta_str}")
                    elif kind == "status":
                        self.status_label.configure(text=msg[1])
                    elif kind == "done":
                        total, logged_files, total_converted = msg[1], msg[2], msg[3]
                        self.log.insert(tk.END, "\n")
                        self.log.insert(
                            tk.END,
                            f"Done. Files processed: {total}, files with changes: {logged_files}, total pixels converted: {total_converted}\n",
                        )
                        if self.dry_run_var.get():
                            self.log.insert(tk.END, "Dry run mode: no files were written.\n")
                        self.running = False
                        self.run_button.configure(state="normal")
                        if hasattr(self, "preview_button") and not self.previewing:
                            self.preview_button.configure(state="normal")
                        if hasattr(self, "sheet_run_button") and not self.sheet_running:
                            self.sheet_run_button.configure(state="normal")
                        if hasattr(self, "tile_run_button") and not self.tile_running:
                            self.tile_run_button.configure(state="normal")
                    elif kind == "conflicts":
                        groups, files = msg[1], msg[2]
                        prompt = (
                            f"Naming conflicts detected.\n\n"
                            f"Conflict groups: {groups}\n"
                            f"Files affected: {files}\n\n"
                            "Overwrite duplicates? (Yes = overwrite, No = add copy)"
                        )
                        choice = messagebox.askyesnocancel("Naming conflicts", prompt)
                        if choice is None:
                            self.conflict_choice = "cancel"
                        elif choice:
                            self.conflict_choice = "overwrite"
                        else:
                            self.conflict_choice = "copy"
                        self.conflict_event.set()
                    elif kind == "sheet_log":
                        self.sheet_log.insert(tk.END, msg[1] + "\n")
                        self.sheet_log.see(tk.END)
                    elif kind == "sheet_done":
                        sheet_count, folder_count = msg[1], msg[2]
                        self.sheet_log.insert(tk.END, "\n")
                        self.sheet_log.insert(
                            tk.END,
                            f"Done. Sheets created: {sheet_count}, folders scanned: {folder_count}\n",
                        )
                        self.sheet_running = False
                        self.sheet_progress.stop()
                        self.sheet_status_label.configure(text="Done")
                        self.sheet_run_button.configure(state="normal")
                        if not self.running and not self.tile_running:
                            self.run_button.configure(state="normal")
                            if hasattr(self, "preview_button"):
                                self.preview_button.configure(state="normal")
                            if hasattr(self, "tile_run_button"):
                                self.tile_run_button.configure(state="normal")
                    elif kind == "sheet_error":
                        messagebox.showerror("Sprite sheet error", msg[1])
                        self.sheet_running = False
                        self.sheet_progress.stop()
                        self.sheet_status_label.configure(text="")
                        self.sheet_run_button.configure(state="normal")
                        if not self.running and not self.tile_running:
                            self.run_button.configure(state="normal")
                            if hasattr(self, "preview_button"):
                                self.preview_button.configure(state="normal")
                            if hasattr(self, "tile_run_button"):
                                self.tile_run_button.configure(state="normal")
                    elif kind == "tile_log":
                        self.tile_log.insert(tk.END, msg[1] + "\n")
                        self.tile_log.see(tk.END)
                    elif kind == "tile_done":
                        tile_count, folder_count = msg[1], msg[2]
                        self.tile_log.insert(tk.END, "\n")
                        self.tile_log.insert(
                            tk.END,
                            f"Done. Tilemaps created: {tile_count}, folders scanned: {folder_count}\n",
                        )
                        self.tile_running = False
                        self.tile_progress.stop()
                        self.tile_status_label.configure(text="Done")
                        self.tile_run_button.configure(state="normal")
                        if not self.running and not self.sheet_running:
                            self.run_button.configure(state="normal")
                            if hasattr(self, "preview_button"):
                                self.preview_button.configure(state="normal")
                            if hasattr(self, "sheet_run_button"):
                                self.sheet_run_button.configure(state="normal")
                    elif kind == "tile_error":
                        messagebox.showerror("Tilemap error", msg[1])
                        self.tile_running = False
                        self.tile_progress.stop()
                        self.tile_status_label.configure(text="")
                        self.tile_run_button.configure(state="normal")
                        if not self.running and not self.sheet_running:
                            self.run_button.configure(state="normal")
                            if hasattr(self, "preview_button"):
                                self.preview_button.configure(state="normal")
                            if hasattr(self, "sheet_run_button"):
                                self.sheet_run_button.configure(state="normal")
                    elif kind == "preview":
                        messagebox.showinfo("Preview routing", msg[1])
                        self.previewing = False
                        if hasattr(self, "preview_button"):
                            self.preview_button.configure(state="normal")
                        if not self.running:
                            self.run_button.configure(state="normal")
                        self.status_label.configure(text="")
                    elif kind == "preview_error":
                        messagebox.showerror("Preview error", msg[1])
                        self.previewing = False
                        if hasattr(self, "preview_button"):
                            self.preview_button.configure(state="normal")
                        if not self.running:
                            self.run_button.configure(state="normal")
                        self.status_label.configure(text="")
                else:
                    self.log.insert(tk.END, str(msg) + "\n")
                    self.log.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(100, self.poll_queue)


if __name__ == "__main__":
    if TkinterDnD is not None:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    root.withdraw()

    splash = tk.Toplevel(root)
    splash.overrideredirect(True)
    splash.configure(background="#1f1f1f")

    logo_img = None
    logo_w = 0
    logo_h = 0
    logo_path = resource_path(LOGO_PNG)
    if os.path.exists(logo_path):
        try:
            logo_img = tk.PhotoImage(file=logo_path)
            logo_w = logo_img.width()
            logo_h = logo_img.height()
        except Exception:
            logo_img = None

    credit_font = ("Segoe UI", 11)
    credit_text = "Created by Koma"
    version_font = ("Segoe UI", 10)
    version_text = APP_VERSION
    credit_label = tk.Label(
        splash,
        text=credit_text,
        font=credit_font,
        background="#1f1f1f",
        foreground="#b5b5b5",
    )
    credit_label.update_idletasks()
    credit_w = credit_label.winfo_reqwidth()
    credit_h = credit_label.winfo_reqheight()
    version_label = tk.Label(
        splash,
        text=version_text,
        font=version_font,
        background="#1f1f1f",
        foreground="#b5b5b5",
    )
    version_label.update_idletasks()
    version_w = version_label.winfo_reqwidth()
    version_h = version_label.winfo_reqheight()

    pad_x = 40
    pad_top = 20
    pad_mid = 10
    pad_mid2 = 4
    pad_bottom = 18
    splash_width = max(logo_w, credit_w, version_w) + pad_x * 2
    splash_height = pad_top + logo_h + pad_mid + credit_h + pad_mid2 + version_h + pad_bottom
    splash_width = max(splash_width, 320)
    splash_height = max(splash_height, 140)

    x = (splash.winfo_screenwidth() - splash_width) // 2
    y = (splash.winfo_screenheight() - splash_height) // 2
    splash.geometry(f"{splash_width}x{splash_height}+{x}+{y}")

    if logo_img is not None:
        logo_label = tk.Label(splash, image=logo_img, background="#1f1f1f")
        logo_label.image = logo_img
        logo_label.pack(pady=(pad_top, pad_mid))
    else:
        tk.Label(
            splash,
            text=APP_NAME,
            font=("Segoe UI", 18, "bold"),
            background="#1f1f1f",
            foreground="#ffffff",
        ).pack(pady=(pad_top, pad_mid))

    credit_label.pack(pady=(0, pad_mid2))
    version_label.pack(pady=(0, pad_bottom))

    def start_app():
        try:
            splash.destroy()
            App(root)
            root.deiconify()
        except Exception:
            os.makedirs(APPDATA_DIR, exist_ok=True)
            with open(ERROR_LOG_PATH, "w", encoding="utf-8") as handle:
                handle.write(traceback.format_exc())
            messagebox.showerror(
                APP_NAME,
                f"Failed to start the app. Details were written to:\n{ERROR_LOG_PATH}",
            )

    root.after(5000, start_app)
    root.mainloop()
