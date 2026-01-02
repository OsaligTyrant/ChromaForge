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
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    from PIL import Image
except Exception as exc:
    Image = None
    PIL_IMPORT_ERROR = exc
else:
    PIL_IMPORT_ERROR = None

APP_DIR = os.path.dirname(os.path.abspath(__file__))
APP_NAME = "ChromaForge"
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
    return re.compile(r"^(?:" + "|".join(escaped) + r")_\d+_")


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


def output_filename(name, prefix_re):
    base, ext = os.path.splitext(name)
    if prefix_re is not None:
        base = prefix_re.sub("", base)
    return base + ext


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
        self.root.title(APP_NAME)
        self.queue = queue.Queue()
        self.running = False
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

        self.prefix_vars = {}
        for name in DEFAULT_PREFIXES:
            self.prefix_vars[name] = tk.BooleanVar(value=True)

        self.folder_rows = {}

        self._build_ui()
        self.refresh_folders()
        self.load_last_settings()
        self.poll_queue()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _build_ui(self):
        root = self.root

        self.style = ttk.Style()
        self.apply_theme(self.theme_var.get())

        main = ttk.Frame(root, padding=10)
        main.grid(row=0, column=0, sticky="nsew")

        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main.columnconfigure(1, weight=1)

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

        prefix_frame = ttk.LabelFrame(main, text="Prefixes to Remove")
        prefix_frame.grid(row=4, column=0, columnspan=3, sticky="ew", pady=6)
        pf_inner = ttk.Frame(prefix_frame)
        pf_inner.grid(row=0, column=0, sticky="ew", padx=6, pady=4)

        col = 0
        for name in DEFAULT_PREFIXES:
            ttk.Checkbutton(pf_inner, text=name, variable=self.prefix_vars[name]).grid(row=0, column=col, padx=4, sticky="w")
            col += 1

        ttk.Label(prefix_frame, text="Extra prefixes (comma-separated)").grid(row=1, column=0, sticky="w", padx=6)
        self.extra_prefix_entry = ttk.Entry(prefix_frame, textvariable=self.extra_prefix_var)
        self.extra_prefix_entry.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 6))
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
        skip_existing_btn = ttk.Checkbutton(options_frame, text="Skip folders that already exist in New (default)", variable=self.skip_existing_var)
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
        self.progress = ttk.Progressbar(run_frame, length=240)
        self.progress.grid(row=0, column=1, padx=10)
        self.status_label = ttk.Label(run_frame, text="")
        self.status_label.grid(row=0, column=2, sticky="w")
        run_frame.columnconfigure(1, weight=1)

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

        self.update_color_mode()
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

    def setup_tooltips(self, transparent_radio, fill_radio, replace_radio, process_all_btn, skip_existing_btn, skip_files_btn, dry_run_btn, tooltips_btn):
        Tooltip(transparent_radio, lambda: "Turn matching colors fully transparent.", self.tooltips_enabled_var)
        Tooltip(fill_radio, lambda: "Fill transparent pixels with a solid color.", self.tooltips_enabled_var)
        Tooltip(replace_radio, lambda: "Swap one exact color for another.", self.tooltips_enabled_var)
        Tooltip(process_all_btn, lambda: "When enabled, ignore individual folder checkboxes.", self.tooltips_enabled_var)
        Tooltip(self.color_list_entry, lambda: "Exact colors to change (comma or space separated).", self.tooltips_enabled_var)
        Tooltip(self.fill_color_entry, lambda: "Fill color used when Fill mode is selected.", self.tooltips_enabled_var)
        Tooltip(self.fill_shadows_check, lambda: "Experimental: fills semi-transparent pixels with a solid color.", self.tooltips_enabled_var)
        Tooltip(self.replace_from_entry, lambda: "Exact source color to replace.", self.tooltips_enabled_var)
        Tooltip(self.replace_to_entry, lambda: "Exact destination color.", self.tooltips_enabled_var)
        Tooltip(self.extra_prefix_entry, lambda: "Extra prefixes to strip (comma-separated).", self.tooltips_enabled_var)
        Tooltip(skip_existing_btn, lambda: "Skip top-level folders already in the output.", self.tooltips_enabled_var)
        Tooltip(skip_files_btn, lambda: "Skip files that already exist in the output.", self.tooltips_enabled_var)
        Tooltip(dry_run_btn, lambda: "Scan and log only. No files are written.", self.tooltips_enabled_var)
        Tooltip(tooltips_btn, lambda: "Toggle tooltip hints on or off.", self.tooltips_enabled_var)

    def pick_input(self):
        path = filedialog.askdirectory(initialdir=self.input_var.get() or os.getcwd())
        if path:
            self.input_var.set(path)
            self.refresh_folders()

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
            "prefixes": {name: var.get() for name, var in self.prefix_vars.items()},
            "process_all": self.process_all_var.get(),
            "skip_existing_folders": self.skip_existing_var.get(),
            "skip_existing_files": self.skip_existing_files_var.get(),
            "dry_run": self.dry_run_var.get(),
            "csv_log": self.csv_log_var.get(),
            "csv_path": self.csv_path_var.get(),
            "exclude_folders": self.exclude_folders_var.get(),
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
        self.process_all_var.set(bool(data.get("process_all", True)))
        self.skip_existing_var.set(bool(data.get("skip_existing_folders", True)))
        self.skip_existing_files_var.set(bool(data.get("skip_existing_files", False)))
        self.dry_run_var.set(bool(data.get("dry_run", False)))
        self.csv_log_var.set(bool(data.get("csv_log", False)))
        self.csv_path_var.set(data.get("csv_path", ""))
        self.exclude_folders_var.set(data.get("exclude_folders", ""))
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
        for name, var in self.prefix_vars.items():
            var.set(bool(prefixes.get(name, True)))

        self.refresh_folders()
        rules = data.get("folder_rules", {})
        for name, rule in rules.items():
            if name in self.folder_rows:
                self.folder_rows[name]["include_var"].set(rule.get("include", True))
                self.folder_rows[name]["mode_var"].set(rule.get("mode", "Process"))

        self.update_color_mode()
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
        if self.running:
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

        prefixes = [name for name, var in self.prefix_vars.items() if var.get()]
        extra = [p.strip() for p in self.extra_prefix_var.get().split(",") if p.strip()]
        prefixes.extend(extra)
        prefix_re = build_prefix_regex(prefixes)

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

    def worker(self, input_root, output_root, mode, target_rgbs, replace_pair, fill_rgb, fill_shadows, prefix_re, rules, allowed_dirs, skip_existing, skip_existing_files, exclude_folders, dry_run, csv_enabled, csv_path):
        tasks = []
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
            if skip_existing and os.path.isdir(os.path.join(output_root, top)):
                continue
            custom_colors = rule.get("colors", "").strip()
            for root, _, files in os.walk(os.path.join(input_root, top)):
                rel_dir = os.path.relpath(root, input_root)
                out_dir = output_root if rel_dir == "." else os.path.join(output_root, rel_dir)
                for file in files:
                    if not file.lower().endswith(".png"):
                        continue
                    in_path = os.path.join(root, file)
                    out_name = output_filename(file, prefix_re)
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
                else:
                    self.log.insert(tk.END, str(msg) + "\n")
                    self.log.see(tk.END)
        except queue.Empty:
            pass
        self.root.after(100, self.poll_queue)


if __name__ == "__main__":
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

    pad_x = 40
    pad_top = 20
    pad_mid = 10
    pad_bottom = 18
    splash_width = max(logo_w, credit_w) + pad_x * 2
    splash_height = pad_top + logo_h + pad_mid + credit_h + pad_bottom
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

    credit_label.pack(pady=(0, pad_bottom))

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
