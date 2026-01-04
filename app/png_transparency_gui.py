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
from tkinter import ttk, filedialog, messagebox
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
APP_VERSION = "Beta 2.1"
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
        self.layout_pos_x_var = tk.StringVar(value="")
        self.layout_pos_y_var = tk.StringVar(value="")
        self.layout_zoom = 1.0

        self.prefix_vars = {}
        self.conflict_event = threading.Event()
        self.conflict_choice = None

        self.folder_rows = {}

        self._build_ui()
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

        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        root.rowconfigure(1, weight=0)

        color_tab = ttk.Frame(notebook)
        sheet_tab = ttk.Frame(notebook)
        tile_tab = ttk.Frame(notebook)
        layout_tab = ttk.Frame(notebook)
        notebook.add(color_tab, text="Color Mode")
        notebook.add(sheet_tab, text="Spritesheet Mode")
        notebook.add(tile_tab, text="Tilemap Mode")
        notebook.add(layout_tab, text="Layout Editor")

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
        self.main_layout_tab = layout_tab

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

        layout_canvas_frame = ttk.LabelFrame(layout_body, text="Canvas")
        layout_canvas_frame.grid(row=0, column=1, sticky="nsew")
        layout_canvas_frame.rowconfigure(0, weight=1)
        layout_canvas_frame.columnconfigure(0, weight=1)
        self.layout_canvas = tk.Canvas(layout_canvas_frame, width=520, height=360, highlightthickness=1)
        self.layout_canvas.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        layout_canvas_scroll_y = ttk.Scrollbar(layout_canvas_frame, orient="vertical", command=self.layout_canvas.yview)
        layout_canvas_scroll_y.grid(row=0, column=1, sticky="ns", pady=6)
        layout_canvas_scroll_x = ttk.Scrollbar(layout_canvas_frame, orient="horizontal", command=self.layout_canvas.xview)
        layout_canvas_scroll_x.grid(row=1, column=0, sticky="ew", padx=6)
        self.layout_canvas.configure(yscrollcommand=layout_canvas_scroll_y.set, xscrollcommand=layout_canvas_scroll_x.set)

        layout_options = ttk.LabelFrame(layout_body, text="Layout Options")
        layout_options.grid(row=0, column=2, sticky="nsew", padx=(8, 0))
        layout_options.columnconfigure(1, weight=1)
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

        self.layout_sprite_options = ttk.LabelFrame(layout_options, text="Spritesheet")
        self.layout_sprite_options.grid(row=9, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
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
        self.layout_tile_options.grid(row=10, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
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
        layout_recent_frame.grid(row=11, column=0, columnspan=2, sticky="ew", padx=6, pady=(6, 2))
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
        layout_action_frame.grid(row=12, column=0, columnspan=2, sticky="ew", padx=6, pady=(8, 6))
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
        self.layout_listbox.bind("<<ListboxSelect>>", lambda _e: self.layout_maybe_set_output_name())
        self.layout_pos_x_entry.bind("<Return>", lambda _e: self.layout_apply_position())
        self.layout_pos_y_entry.bind("<Return>", lambda _e: self.layout_apply_position())
        self.layout_canvas.bind("<ButtonPress-1>", self.layout_on_press)
        self.layout_canvas.bind("<B1-Motion>", self.layout_on_drag)
        self.layout_canvas.bind("<ButtonRelease-1>", self.layout_on_release)
        self.layout_canvas.bind("<MouseWheel>", self.layout_on_mousewheel)
        self.setup_layout_dnd()

        self.update_color_mode()
        self.update_sprite_layout_controls()
        self.update_sprite_padding_controls()
        self.update_tile_layout_controls()
        self.update_layout_type_controls()
        self.layout_redraw()
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
        if hasattr(self, "layout_listbox"):
            self.layout_listbox.configure(background=colors["entry_bg"], foreground=colors["entry_fg"])
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

    def layout_update_position_fields(self):
        if not self.layout_selected_ids:
            self.layout_pos_x_var.set("")
            self.layout_pos_y_var.set("")
            return
        xs = {int(self.layout_items[item_id]["x"]) for item_id in self.layout_selected_ids}
        ys = {int(self.layout_items[item_id]["y"]) for item_id in self.layout_selected_ids}
        self.layout_pos_x_var.set(str(xs.pop()) if len(xs) == 1 else "")
        self.layout_pos_y_var.set(str(ys.pop()) if len(ys) == 1 else "")

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
            self.layout_snap_selected()
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_align(self, direction):
        if len(self.layout_selected_ids) < 2:
            return
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
            self.layout_snap_selected()
        self.layout_redraw()
        self.layout_update_position_fields()

    def layout_center_in_cell(self):
        if not self.layout_selected_ids:
            return
        if self.layout_mode_var.get() != "Grid":
            messagebox.showinfo("Center in cell", "Centering requires Grid mode.")
            return
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
        zoom = max(0.2, min(zoom, 4.0))
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
        for i, (file, path, pil, photo, width, height) in enumerate(new_items):
            x, y = self.layout_next_position(start_index + i, cell_w, cell_h, pad)
            canvas_x, canvas_y = self.layout_to_canvas(x, y)
            item_id = self.layout_canvas.create_image(canvas_x, canvas_y, anchor="nw", image=photo, tags=("layout_item",))
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
                return item_id
        return None

    def layout_update_selection_rect(self, item_id):
        item = self.layout_items.get(item_id)
        if not item:
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
        canvas_x = self.layout_canvas.canvasx(event.x)
        canvas_y = self.layout_canvas.canvasy(event.y)
        item_id = self.layout_find_item_at(canvas_x, canvas_y)
        if not item_id:
            for selected in list(self.layout_selected_ids):
                self.layout_clear_selection_rect(selected)
            self.layout_selected_ids.clear()
            self.layout_update_position_fields()
            self.layout_drag_start = None
            return
        add = bool(event.state & 0x0001)
        toggle = bool(event.state & 0x0004)
        self.layout_select_item(item_id, add=add, toggle=toggle)
        x, y = self.layout_from_canvas(canvas_x, canvas_y)
        self.layout_drag_start = (x, y)
        self.layout_drag_positions = {
            selected: (self.layout_items[selected]["x"], self.layout_items[selected]["y"])
            for selected in self.layout_selected_ids
        }

    def layout_on_drag(self, event):
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

        for item_id in self.layout_selected_ids:
            start_x, start_y = self.layout_drag_positions.get(item_id, (0, 0))
            new_x = start_x + dx
            new_y = start_y + dy
            if snap and step_x > 0 and step_y > 0:
                col = int(round(new_x / step_x))
                row = int(round(new_y / step_y))
                new_x = col * step_x
                new_y = row * step_y
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
        self.layout_drag_start = None
        self.layout_drag_positions = {}
        self.layout_update_position_fields()

    def layout_snap_selected(self):
        if not self.layout_snap_var.get() or self.layout_mode_var.get() != "Grid":
            return
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

        for item_id, item in self.layout_items.items():
            canvas_x, canvas_y = self.layout_to_canvas(item["x"], item["y"])
            self.layout_canvas.coords(item_id, canvas_x, canvas_y)
            if item.get("rect_id"):
                width = item["width"] * self.layout_zoom
                height = item["height"] * self.layout_zoom
                self.layout_canvas.coords(item["rect_id"], canvas_x, canvas_y, canvas_x + width, canvas_y + height)

        for item_id in self.layout_selected_ids:
            self.layout_update_selection_rect(item_id)

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
        max_x = max((item["x"] + item["width"]) for item in self.layout_items.values())
        max_y = max((item["y"] + item["height"]) for item in self.layout_items.values())

        if self.layout_mode_var.get() == "Grid":
            step_x = cell_w + pad
            step_y = cell_h + pad
            if rows <= 0 and step_x > 0 and step_y > 0:
                row_indices = [int(item["y"] // step_y) for item in self.layout_items.values()]
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
            out_name = f"{output_name}.png"
        else:
            out_dir = os.path.join(folder_path, "tilemaps")
            out_name = f"{output_name}.png"

        os.makedirs(out_dir, exist_ok=True)
        out_path = os.path.join(out_dir, out_name)
        canvas = Image.new("RGBA", (out_w, out_h), (0, 0, 0, 0))
        metadata = []
        for item in self.layout_items.values():
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
            "items": [
                {
                    "file": item["file"],
                    "x": int(item["x"]),
                    "y": int(item["y"]),
                    "width": int(item["width"]),
                    "height": int(item["height"]),
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
            layout_name = f"{output_name}.json"
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

        self.layout_clear_canvas()
        folder_path = self.layout_get_folder_path()
        if not folder_path:
            return
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
            }
        self.layout_redraw()
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
