#!/usr/bin/env python3
"""Modern GUI wrapper for MSIKLM with explicit 3-zone keyboard mapping."""

from __future__ import annotations

import argparse
import array
import os
import re
import shlex
import shutil
import subprocess
import sys
import threading
import time

try:
    import tkinter as tk
    from tkinter import colorchooser, messagebox, ttk
except ModuleNotFoundError as exc:
    raise SystemExit(
        "tkinter is not installed. Install it first (Debian/Ubuntu/Mint: sudo apt install python3-tk)."
    ) from exc


PRESET_COLORS = {
    "off": "#000000",
    "none": "#000000",
    "red": "#ff0000",
    "orange": "#ff6400",
    "yellow": "#ffff00",
    "green": "#00ff00",
    "sky": "#00ffff",
    "blue": "#0000ff",
    "purple": "#ff00ff",
    "white": "#ffffff",
}

COLOR_CHOICES = [
    "off",
    "red",
    "orange",
    "yellow",
    "green",
    "sky",
    "blue",
    "purple",
    "white",
    "custom",
]

PRIMARY_ZONES = ["left", "middle", "right"]
OPTIONAL_ZONES = ["logo", "front_left", "front_right", "mouse"]
ALL_ZONES = PRIMARY_ZONES + OPTIONAL_ZONES

KEY_LAYOUT = [
    [
        ("Esc", 1.0, "left"),
        ("1", 1.0, "left"),
        ("2", 1.0, "left"),
        ("3", 1.0, "left"),
        ("4", 1.0, "left"),
        ("5", 1.0, "left"),
        ("6", 1.0, "middle"),
        ("7", 1.0, "middle"),
        ("8", 1.0, "middle"),
        ("9", 1.0, "middle"),
        ("0", 1.0, "middle"),
        ("-", 1.0, "right"),
        ("=", 1.0, "right"),
        ("Back", 2.0, "right"),
        ("", 0.8, None),
        ("Num", 1.0, "right"),
        ("/", 1.0, "right"),
        ("*", 1.0, "right"),
        ("-", 1.0, "right"),
    ],
    [
        ("Tab", 1.5, "left"),
        ("Q", 1.0, "left"),
        ("W", 1.0, "left"),
        ("E", 1.0, "left"),
        ("R", 1.0, "left"),
        ("T", 1.0, "left"),
        ("Y", 1.0, "middle"),
        ("U", 1.0, "middle"),
        ("I", 1.0, "middle"),
        ("O", 1.0, "middle"),
        ("P", 1.0, "middle"),
        ("[", 1.0, "right"),
        ("]", 1.0, "right"),
        ("\\", 1.5, "right"),
        ("", 0.8, None),
        ("7", 1.0, "right"),
        ("8", 1.0, "right"),
        ("9", 1.0, "right"),
        ("+", 1.0, "right"),
    ],
    [
        ("Caps", 1.8, "left"),
        ("A", 1.0, "left"),
        ("S", 1.0, "left"),
        ("D", 1.0, "left"),
        ("F", 1.0, "left"),
        ("G", 1.0, "left"),
        ("H", 1.0, "middle"),
        ("J", 1.0, "middle"),
        ("K", 1.0, "middle"),
        ("L", 1.0, "middle"),
        (";", 1.0, "right"),
        ("'", 1.0, "right"),
        ("Enter", 2.2, "right"),
        ("", 0.8, None),
        ("4", 1.0, "right"),
        ("5", 1.0, "right"),
        ("6", 1.0, "right"),
        ("", 1.0, None),
    ],
    [
        ("Shift", 2.2, "left"),
        ("Z", 1.0, "left"),
        ("X", 1.0, "left"),
        ("C", 1.0, "left"),
        ("V", 1.0, "left"),
        ("B", 1.0, "left"),
        ("N", 1.0, "middle"),
        ("M", 1.0, "middle"),
        (",", 1.0, "middle"),
        (".", 1.0, "right"),
        ("/", 1.0, "right"),
        ("Shift", 2.8, "right"),
        ("", 0.8, None),
        ("1", 1.0, "right"),
        ("2", 1.0, "right"),
        ("3", 1.0, "right"),
        ("Enter", 1.0, "right"),
    ],
    [
        ("Ctrl", 1.3, "left"),
        ("Fn", 1.1, "left"),
        ("Win", 1.1, "left"),
        ("Alt", 1.1, "left"),
        ("Space", 2.5, "left"),
        ("Space", 2.5, "middle"),
        ("Space", 2.5, "right"),
        ("Alt", 1.1, "right"),
        ("Menu", 1.1, "right"),
        ("Ctrl", 1.3, "right"),
        ("", 0.8, None),
        ("0", 2.0, "right"),
        (".", 1.0, "right"),
    ],
]

ROW_OFFSETS = [0, 12, 20, 30, 40]
HEX_PATTERN = re.compile(r"^#?[0-9a-fA-F]{6}$")

BG_APP = "#09111f"
BG_PANEL = "#111d33"
BG_PANEL_ALT = "#162540"
FG_MAIN = "#dbe8ff"
FG_MUTED = "#9eb4d8"
ACCENT = "#3b8bf3"
ACCENT_DARK = "#2d6dc0"
SEAM_COLOR = "#4a607d"
SEAM_GLOW = "#314861"
FONT_UI = "DejaVu Sans"
FONT_MONO = "DejaVu Sans Mono"
VOICE_COLOR_PATH = ["blue", "sky", "green", "yellow", "orange", "red", "purple"]
VOICE_FRAME_SECONDS = 0.12
VOICE_SAMPLE_RATE = 16000
VOICE_SAMPLES_PER_FRAME = int(VOICE_SAMPLE_RATE * VOICE_FRAME_SECONDS)
VOICE_AUTO_SOURCE = "Auto detect (try all)"
VOICE_ATTACK_ALPHA = 0.66
VOICE_RELEASE_ALPHA_ACTIVE = 0.18
VOICE_RELEASE_ALPHA_SILENT = 0.06
VOICE_BAR_DECAY = 0.90
VOICE_SILENCE_HOLD_FRAMES = max(2, int(1.35 / VOICE_FRAME_SECONDS))
VOICE_MIN_SEND_INTERVAL_ACTIVE = 0.08
VOICE_MIN_SEND_INTERVAL_SILENT = 0.18
VOICE_COMPAT_COLORS = ["red", "orange", "yellow", "green", "sky", "blue", "purple", "white"]


def normalize_hex(value: str) -> str | None:
    value = value.strip()
    if not HEX_PATTERN.match(value):
        return None
    if value.startswith("#"):
        return value.lower()
    return "#" + value.lower()


def darken(hex_color: str, ratio: float) -> str:
    r = int(hex_color[1:3], 16)
    g = int(hex_color[3:5], 16)
    b = int(hex_color[5:7], 16)
    r = int(max(0, min(255, r * ratio)))
    g = int(max(0, min(255, g * ratio)))
    b = int(max(0, min(255, b * ratio)))
    return f"#{r:02x}{g:02x}{b:02x}"


def lighten(hex_color: str, ratio: float) -> str:
    r = int(hex_color[1:3], 16)
    g = int(hex_color[3:5], 16)
    b = int(hex_color[5:7], 16)
    r = int(max(0, min(255, r + (255 - r) * ratio)))
    g = int(max(0, min(255, g + (255 - g) * ratio)))
    b = int(max(0, min(255, b + (255 - b) * ratio)))
    return f"#{r:02x}{g:02x}{b:02x}"


def blend_hex(hex_a: str, hex_b: str, ratio: float) -> str:
    ratio = max(0.0, min(1.0, ratio))
    ar = int(hex_a[1:3], 16)
    ag = int(hex_a[3:5], 16)
    ab = int(hex_a[5:7], 16)
    br = int(hex_b[1:3], 16)
    bg = int(hex_b[3:5], 16)
    bb = int(hex_b[5:7], 16)
    rr = int(ar + ((br - ar) * ratio))
    rg = int(ag + ((bg - ag) * ratio))
    rb = int(ab + ((bb - ab) * ratio))
    return f"#{rr:02x}{rg:02x}{rb:02x}"


class SimpleTooltip:
    def __init__(self, widget: tk.Widget, text: str) -> None:
        self.widget = widget
        self.text = text
        self.tip_window: tk.Toplevel | None = None
        self.widget.bind("<Enter>", self._show, add="+")
        self.widget.bind("<Leave>", self._hide, add="+")

    def _show(self, _event: tk.Event[tk.Widget]) -> None:
        if self.tip_window is not None:
            return
        x = self.widget.winfo_rootx() + 16
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        tip = tk.Toplevel(self.widget)
        tip.wm_overrideredirect(True)
        tip.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            tip,
            text=self.text,
            bg="#0a1628",
            fg="#dbe8ff",
            padx=8,
            pady=4,
            bd=1,
            relief=tk.SOLID,
            font=(FONT_UI, 9),
        )
        label.pack()
        self.tip_window = tip

    def _hide(self, _event: tk.Event[tk.Widget]) -> None:
        if self.tip_window is None:
            return
        self.tip_window.destroy()
        self.tip_window = None


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="MSIKLM GUI launcher")
    parser.add_argument("--as-root", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument(
        "--no-elevate",
        action="store_true",
        help="Run the GUI without automatic privilege elevation.",
    )
    return parser.parse_args()


def relaunch_as_root(args: argparse.Namespace) -> None:
    if args.no_elevate:
        return
    if os.geteuid() == 0:
        return
    if args.as_root:
        raise SystemExit("Failed to launch elevated instance.")

    display = os.environ.get("DISPLAY") or os.environ.get("WAYLAND_DISPLAY")
    if not display:
        raise SystemExit(
            "No graphical display detected for elevated relaunch. "
            "Start from a desktop session or run with --no-elevate."
        )

    script = os.path.abspath(__file__)
    relaunch_cmd = [sys.executable, script, "--as-root"]

    preserve_env: list[str] = []
    for var in (
        "DISPLAY",
        "WAYLAND_DISPLAY",
        "XAUTHORITY",
        "XDG_RUNTIME_DIR",
        "DBUS_SESSION_BUS_ADDRESS",
        "HOME",
        "PATH",
    ):
        val = os.environ.get(var)
        if val:
            preserve_env.append(f"{var}={val}")

    if shutil.which("pkexec"):
        code = subprocess.call(["pkexec", "env", *preserve_env, *relaunch_cmd])
        if code == 0:
            raise SystemExit(0)
    if shutil.which("sudo"):
        code = subprocess.call(["sudo", "-E", *relaunch_cmd])
        if code == 0:
            raise SystemExit(0)

    raise SystemExit(
        "Could not relaunch with elevated privileges. Install/configure pkexec or sudo, or use --no-elevate."
    )


class MSIKLMGui(tk.Tk):
    def __init__(self, launched_as_root: bool) -> None:
        super().__init__()
        self.title("MSIKLM GUI")
        self.geometry("1420x900")
        self.minsize(1260, 780)
        self.configure(bg=BG_APP)
        self.launched_as_root = launched_as_root

        self.zone_color_mode: dict[str, tk.StringVar] = {}
        self.zone_custom_hex: dict[str, tk.StringVar] = {}
        self.zone_include: dict[str, tk.BooleanVar] = {}
        self.zone_swatch: dict[str, tk.Canvas] = {}
        self._updating_optional = False
        self.apply_button: ttk.Button | None = None
        self.mode_button: ttk.Button | None = None
        self.test_button: ttk.Button | None = None
        self.voice_meter: tk.Canvas | None = None
        self.voice_source_combo: ttk.Combobox | None = None

        self.use_brightness = tk.BooleanVar(value=False)
        self.use_mode = tk.BooleanVar(value=False)
        self.compat_mode = tk.BooleanVar(value=True)
        self.brightness = tk.StringVar(value="high")
        self.mode = tk.StringVar(value="normal")
        self.voice_mode_enabled = tk.BooleanVar(value=False)
        self.voice_source_var = tk.StringVar(value=VOICE_AUTO_SOURCE)
        self.voice_gain = tk.DoubleVar(value=5.0)
        self.voice_threshold = tk.DoubleVar(value=0.03)
        self.voice_gain_label_var = tk.StringVar(value="5.00")
        self.voice_threshold_label_var = tk.StringVar(value="0.030")
        self.voice_status_var = tk.StringVar(value="Voice mode is off.")
        self.voice_active_source_var = tk.StringVar(value="Active source: none")
        self.preview_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Ready.")
        self.binary_var = tk.StringVar(value="")
        self._voice_gain_value = float(self.voice_gain.get())
        self._voice_threshold_value = float(self.voice_threshold.get())
        self._voice_thread: threading.Thread | None = None
        self._voice_stop_event = threading.Event()
        self._voice_palette_phase = 0.0
        self._voice_last_payload = ""
        self._voice_preview_hex = {zone: "#000000" for zone in PRIMARY_ZONES}
        self._voice_sources: list[tuple[str, list[str]]] = []
        self._voice_active_source = "none"
        self._voice_smoothed_level = 0.0
        self._voice_silence_frames = 0
        self._voice_last_send_ts = 0.0
        self._tooltips: list[SimpleTooltip] = []
        self._closing = False

        self._setup_styles()
        self._build_ui()
        self._set_defaults()
        self._refresh_binary_label()
        self._update_command_preview()
        self._redraw_keyboard()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _setup_styles(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", font=(FONT_UI, 10))
        style.configure("App.TFrame", background=BG_APP)
        style.configure("Panel.TFrame", background=BG_PANEL)
        style.configure("PanelAlt.TFrame", background=BG_PANEL_ALT)
        style.configure("Card.TLabelframe", background=BG_PANEL, foreground=FG_MAIN, bordercolor="#243654")
        style.configure("Card.TLabelframe.Label", background=BG_PANEL, foreground=FG_MAIN)
        style.configure("RightCard.TLabelframe", background=BG_PANEL_ALT, foreground=FG_MAIN, bordercolor="#2b4365")
        style.configure("RightCard.TLabelframe.Label", background=BG_PANEL_ALT, foreground=FG_MAIN)
        style.configure("TLabel", background=BG_PANEL, foreground=FG_MAIN)
        style.configure("Muted.TLabel", background=BG_PANEL, foreground=FG_MUTED)
        style.configure("Title.TLabel", background=BG_APP, foreground=lighten(ACCENT, 0.4), font=(FONT_UI, 21, "bold"))
        style.configure("Subtitle.TLabel", background=BG_APP, foreground="#7fa2d3", font=(FONT_UI, 10))
        style.configure("PanelAlt.TLabel", background=BG_PANEL_ALT, foreground=FG_MAIN)
        style.configure("PanelAltMuted.TLabel", background=BG_PANEL_ALT, foreground="#89a7d2")
        style.configure("Accent.TButton", background=ACCENT, foreground="#ffffff", borderwidth=0, padding=8)
        style.map(
            "Accent.TButton",
            background=[("active", lighten(ACCENT, 0.1)), ("pressed", ACCENT_DARK)],
        )
        style.configure("Ghost.TButton", background="#1a2c49", foreground="#d5e5ff", borderwidth=1, padding=8)
        style.map(
            "Ghost.TButton",
            background=[("active", "#213557"), ("pressed", "#182b46")],
        )
        style.configure("TButton", padding=8)
        style.configure("TCheckbutton", background=BG_PANEL, foreground=FG_MAIN)
        style.map("TCheckbutton", background=[("active", BG_PANEL)])
        style.configure("TEntry", fieldbackground="#0f1a2d", foreground=FG_MAIN)
        style.configure("TCombobox", fieldbackground="#0f1a2d", foreground=FG_MAIN, arrowsize=14)

    def _build_ui(self) -> None:
        root = ttk.Frame(self, style="App.TFrame", padding=14)
        root.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(root, style="App.TFrame")
        header.pack(fill=tk.X, pady=(0, 10))
        header_left = ttk.Frame(header, style="App.TFrame")
        header_left.pack(side=tk.LEFT, anchor=tk.W)
        ttk.Label(header_left, text="MSIKLM Lighting Studio", style="Title.TLabel").pack(anchor=tk.W)
        ttk.Label(
            header_left,
            text="Modern 3-zone keyboard lighting control with reliability fallback.",
            style="Subtitle.TLabel",
        ).pack(anchor=tk.W, pady=(2, 0))

        privilege = "root" if os.geteuid() == 0 else "user"
        privilege_color = "#5dd39e" if os.geteuid() == 0 else "#f2b84c"
        self.privilege_label = tk.Label(
            header,
            text=f"session: {privilege}",
            bg="#0d1d33",
            fg=privilege_color,
            font=(FONT_UI, 10, "bold"),
            padx=12,
            pady=5,
            bd=1,
            relief=tk.SOLID,
            highlightthickness=1,
            highlightbackground=darken(privilege_color, 0.65),
        )
        self.privilege_label.pack(side=tk.RIGHT)

        body = ttk.Frame(root, style="App.TFrame")
        body.pack(fill=tk.BOTH, expand=True)

        left = ttk.Frame(body, style="Panel.TFrame", padding=12)
        right = ttk.Frame(body, style="PanelAlt.TFrame", padding=14, width=500)
        right.pack(side=tk.RIGHT, fill=tk.Y, padx=(12, 0))
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        right.pack_propagate(False)

        self.canvas = tk.Canvas(left, width=940, height=460, bg="#0a1628", highlightthickness=0)
        self.canvas.pack(fill=tk.X, expand=False)

        ttk.Label(
            left,
            text=(
                "3-zone keyboard split: LEFT | MIDDLE | RIGHT. "
                "On GT60-class models, numpad keys belong to RIGHT. "
                "Optional non-keyboard zones are shown as badges."
            ),
            style="Muted.TLabel",
            wraplength=940,
        ).pack(fill=tk.X, pady=(8, 6))
        ttk.Label(left, textvariable=self.binary_var, style="Muted.TLabel").pack(fill=tk.X, pady=(0, 8))

        preview_frame = ttk.LabelFrame(left, text="Command Preview", style="Card.TLabelframe", padding=8)
        preview_frame.pack(fill=tk.X)
        ttk.Label(preview_frame, textvariable=self.preview_var).pack(anchor=tk.W)

        log_frame = ttk.LabelFrame(left, text="Output", style="Card.TLabelframe", padding=8)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(8, 0))
        self.log_text = tk.Text(
            log_frame,
            height=10,
            wrap=tk.WORD,
            bg="#08101f",
            fg="#cfe0ff",
            insertbackground="#cfe0ff",
            relief=tk.FLAT,
            font=(FONT_MONO, 9),
            state=tk.DISABLED,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
        ttk.Label(log_frame, textvariable=self.status_var, style="Muted.TLabel").pack(anchor=tk.W, pady=(6, 0))

        zones_frame = ttk.LabelFrame(right, text="Zone Colors", style="RightCard.TLabelframe", padding=10)
        zones_frame.pack(fill=tk.X)
        ttk.Label(zones_frame, text="Use").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Label(zones_frame, text="Zone").grid(row=0, column=1, sticky=tk.W, padx=(0, 8))
        ttk.Label(zones_frame, text="Preview").grid(row=0, column=2, sticky=tk.W, padx=(0, 8))
        ttk.Label(zones_frame, text="Color").grid(row=0, column=3, sticky=tk.W, padx=(0, 8))
        ttk.Label(zones_frame, text="Hex").grid(row=0, column=4, sticky=tk.W, padx=(0, 8))
        ttk.Label(zones_frame, text="Pick").grid(row=0, column=5, sticky=tk.W, padx=(0, 8))
        ttk.Label(
            zones_frame,
            text="Keyboard note: RIGHT includes numpad on GT60 2OD.",
            style="PanelAltMuted.TLabel",
        ).grid(row=1, column=1, columnspan=5, sticky=tk.W, pady=(2, 6))

        row = 2
        for zone in ALL_ZONES:
            mode_var = tk.StringVar(value="blue" if zone in PRIMARY_ZONES else "off")
            hex_var = tk.StringVar(value="#0000ff" if zone in PRIMARY_ZONES else "#000000")
            include_var = tk.BooleanVar(value=zone in PRIMARY_ZONES)
            self.zone_color_mode[zone] = mode_var
            self.zone_custom_hex[zone] = hex_var
            self.zone_include[zone] = include_var

            check = ttk.Checkbutton(zones_frame, variable=include_var, command=lambda z=zone: self._on_include_toggle(z))
            check.grid(row=row, column=0, sticky=tk.W, padx=(0, 8), pady=3)

            zone_label = ttk.Label(zones_frame, text=zone.upper(), style="PanelAltMuted.TLabel")
            zone_label.grid(row=row, column=1, sticky=tk.W, padx=(0, 8), pady=4)
            if zone == "right":
                tip_text = "RIGHT zone includes alphanumeric right cluster and numpad keys on GT60 2OD."
                self._tooltips.append(SimpleTooltip(zone_label, tip_text))
                self._tooltips.append(SimpleTooltip(check, tip_text))

            swatch = tk.Canvas(zones_frame, width=18, height=18, bg=BG_PANEL_ALT, highlightthickness=0)
            swatch.grid(row=row, column=2, sticky=tk.W, padx=(0, 8), pady=4)
            self.zone_swatch[zone] = swatch

            combo = ttk.Combobox(
                zones_frame,
                values=COLOR_CHOICES,
                textvariable=mode_var,
                state="readonly",
                width=9,
            )
            combo.grid(row=row, column=3, sticky=tk.EW, padx=(0, 8), pady=4)
            combo.bind("<<ComboboxSelected>>", lambda _evt, z=zone: self._on_zone_changed(z))

            entry = ttk.Entry(zones_frame, textvariable=hex_var, width=8)
            entry.grid(row=row, column=4, sticky=tk.EW, padx=(0, 4), pady=4)
            entry.bind("<KeyRelease>", lambda _evt, z=zone: self._on_zone_changed(z))

            pick = ttk.Button(zones_frame, text="Pick", command=lambda z=zone: self._pick_color(z), width=4)
            pick.grid(row=row, column=5, sticky=tk.W, pady=4)
            row += 1

        zones_frame.grid_columnconfigure(1, minsize=82)
        zones_frame.grid_columnconfigure(3, weight=1, minsize=102)
        zones_frame.grid_columnconfigure(4, weight=1, minsize=86)
        zones_frame.grid_columnconfigure(5, minsize=56)

        opts = ttk.LabelFrame(right, text="Apply Behavior", style="RightCard.TLabelframe", padding=10)
        opts.pack(fill=tk.X, pady=(10, 0))

        ttk.Checkbutton(
            opts,
            text="Compatibility mode (recommended)",
            variable=self.compat_mode,
            command=self._update_command_preview,
        ).grid(row=0, column=0, columnspan=2, sticky=tk.W, pady=2)

        ttk.Checkbutton(
            opts,
            text="Include brightness explicitly",
            variable=self.use_brightness,
            command=self._update_command_preview,
        ).grid(row=1, column=0, sticky=tk.W, pady=2)
        brightness_combo = ttk.Combobox(
            opts,
            values=["rgb", "high", "medium", "low", "off"],
            textvariable=self.brightness,
            state="readonly",
            width=10,
        )
        brightness_combo.grid(row=1, column=1, sticky=tk.W, padx=(8, 0), pady=2)
        brightness_combo.bind("<<ComboboxSelected>>", lambda _evt: self._update_command_preview())

        ttk.Checkbutton(
            opts,
            text="Include mode",
            variable=self.use_mode,
            command=self._update_command_preview,
        ).grid(row=2, column=0, sticky=tk.W, pady=2)
        mode_combo = ttk.Combobox(
            opts,
            values=["normal", "gaming", "breathe", "demo", "wave"],
            textvariable=self.mode,
            state="readonly",
            width=10,
        )
        mode_combo.grid(row=2, column=1, sticky=tk.W, padx=(8, 0), pady=2)
        mode_combo.bind("<<ComboboxSelected>>", lambda _evt: self._update_command_preview())

        voice = ttk.LabelFrame(right, text="Voice Mode", style="RightCard.TLabelframe", padding=10)
        voice.pack(fill=tk.X, pady=(10, 0))

        ttk.Checkbutton(
            voice,
            text="Enable microphone reactive mode",
            variable=self.voice_mode_enabled,
            command=self._on_voice_toggle,
        ).grid(row=0, column=0, columnspan=3, sticky=tk.W, pady=(0, 6))

        ttk.Label(voice, text="Input", style="PanelAltMuted.TLabel").grid(row=1, column=0, sticky=tk.W)
        self.voice_source_combo = ttk.Combobox(
            voice,
            values=[VOICE_AUTO_SOURCE],
            textvariable=self.voice_source_var,
            state="readonly",
            width=26,
        )
        self.voice_source_combo.grid(row=1, column=1, sticky=tk.EW, padx=(8, 8))
        ttk.Button(voice, text="Refresh", style="Ghost.TButton", command=self._refresh_voice_sources).grid(
            row=1,
            column=2,
            sticky=tk.E,
        )

        ttk.Label(voice, text="Gain", style="PanelAltMuted.TLabel").grid(row=2, column=0, sticky=tk.W, pady=(4, 0))
        gain_scale = ttk.Scale(
            voice,
            from_=1.0,
            to=8.0,
            variable=self.voice_gain,
            command=lambda _value: self._on_voice_settings_changed(),
            orient=tk.HORIZONTAL,
        )
        gain_scale.grid(row=2, column=1, sticky=tk.EW, padx=(8, 8), pady=(4, 0))
        ttk.Label(voice, textvariable=self.voice_gain_label_var, style="PanelAltMuted.TLabel").grid(row=2, column=2, sticky=tk.E, pady=(4, 0))

        ttk.Label(voice, text="Threshold", style="PanelAltMuted.TLabel").grid(row=3, column=0, sticky=tk.W, pady=(4, 0))
        threshold_scale = ttk.Scale(
            voice,
            from_=0.005,
            to=0.2,
            variable=self.voice_threshold,
            command=lambda _value: self._on_voice_settings_changed(),
            orient=tk.HORIZONTAL,
        )
        threshold_scale.grid(row=3, column=1, sticky=tk.EW, padx=(8, 8), pady=(4, 0))
        ttk.Label(voice, textvariable=self.voice_threshold_label_var, style="PanelAltMuted.TLabel").grid(
            row=3,
            column=2,
            sticky=tk.E,
            pady=(4, 0),
        )

        self.voice_meter = tk.Canvas(voice, width=420, height=36, bg="#0d1930", highlightthickness=1, highlightbackground="#2c4365")
        self.voice_meter.grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=(8, 4))
        ttk.Label(voice, textvariable=self.voice_status_var, style="PanelAltMuted.TLabel", wraplength=430).grid(
            row=5,
            column=0,
            columnspan=3,
            sticky=tk.W,
        )
        ttk.Label(voice, textvariable=self.voice_active_source_var, style="PanelAltMuted.TLabel", wraplength=430).grid(
            row=6,
            column=0,
            columnspan=3,
            sticky=tk.W,
            pady=(2, 0),
        )
        voice.grid_columnconfigure(1, weight=1)

        actions = ttk.LabelFrame(right, text="Actions", style="RightCard.TLabelframe", padding=10)
        actions.pack(fill=tk.X, pady=(10, 0))
        self.apply_button = ttk.Button(actions, text="Apply Colors", style="Accent.TButton", command=self._apply_colors)
        self.apply_button.pack(fill=tk.X, pady=2)
        self.mode_button = ttk.Button(actions, text="Apply Mode Only", style="Ghost.TButton", command=self._apply_mode_only)
        self.mode_button.pack(fill=tk.X, pady=2)
        self.test_button = ttk.Button(actions, text="Test Connection", style="Ghost.TButton", command=self._test_connection)
        self.test_button.pack(fill=tk.X, pady=2)
        ttk.Button(actions, text="Show CLI Help", style="Ghost.TButton", command=self._show_help).pack(fill=tk.X, pady=2)

    def _set_defaults(self) -> None:
        for zone in PRIMARY_ZONES:
            self.zone_color_mode[zone].set("blue")
            self.zone_custom_hex[zone].set("#0000ff")
            self.zone_include[zone].set(True)
        for zone in OPTIONAL_ZONES:
            self.zone_color_mode[zone].set("off")
            self.zone_custom_hex[zone].set("#000000")
            self.zone_include[zone].set(False)
        self._refresh_zone_swatches()
        self._on_voice_settings_changed()
        self._refresh_voice_sources()
        self._draw_voice_meter([0.0, 0.0, 0.0])

    def _refresh_binary_label(self) -> None:
        binary = self._resolve_msiklm()
        if binary:
            self.binary_var.set(f"msiklm binary: {binary}")
        else:
            self.binary_var.set("msiklm binary: not found")

    def _on_optional_toggle(self, changed_zone: str) -> None:
        if self._updating_optional:
            return

        self._updating_optional = True
        try:
            idx = OPTIONAL_ZONES.index(changed_zone)
            enabled = self.zone_include[changed_zone].get()
            if enabled:
                for zone in OPTIONAL_ZONES[:idx]:
                    self.zone_include[zone].set(True)
            else:
                for zone in OPTIONAL_ZONES[idx:]:
                    self.zone_include[zone].set(False)
        finally:
            self._updating_optional = False

        self._on_zone_changed(changed_zone)

    def _on_include_toggle(self, changed_zone: str) -> None:
        if changed_zone in OPTIONAL_ZONES:
            self._on_optional_toggle(changed_zone)
            return
        self._on_zone_changed(changed_zone)

    def _on_voice_settings_changed(self) -> None:
        self._voice_gain_value = float(self.voice_gain.get())
        self._voice_threshold_value = float(self.voice_threshold.get())
        self.voice_gain_label_var.set(f"{self._voice_gain_value:.2f}")
        self.voice_threshold_label_var.set(f"{self._voice_threshold_value:.3f}")

    def _on_zone_changed(self, _zone: str) -> None:
        self._refresh_zone_swatches()
        self._update_command_preview()
        self._redraw_keyboard()

    def _refresh_zone_swatches(self) -> None:
        for zone, swatch in self.zone_swatch.items():
            fill = self._zone_visual_color(zone)
            swatch.delete("all")
            swatch.create_oval(2, 2, 16, 16, fill=fill, outline=darken(fill, 0.4), width=1)

    def _pick_color(self, zone: str) -> None:
        initial = self._zone_hex_or_fallback(zone)
        chosen = colorchooser.askcolor(color=initial, parent=self)[1]
        if chosen:
            self.zone_custom_hex[zone].set(chosen.lower())
            self.zone_color_mode[zone].set("custom")
            self._on_zone_changed(zone)

    def _zone_hex_or_fallback(self, zone: str) -> str:
        mode = self.zone_color_mode[zone].get()
        if mode == "custom":
            normalized = normalize_hex(self.zone_custom_hex[zone].get())
            return normalized if normalized else "#ffffff"
        return PRESET_COLORS.get(mode, "#000000")

    def _zone_cli_value(self, zone: str) -> tuple[str, bool]:
        if not self.zone_include[zone].get():
            return "off", False

        mode = self.zone_color_mode[zone].get()
        if mode == "custom":
            normalized = normalize_hex(self.zone_custom_hex[zone].get())
            if not normalized:
                raise ValueError(f"Zone '{zone}' has an invalid custom hex value.")
            return normalized, True
        if mode in ("none", "off"):
            return "off", False
        if mode in PRESET_COLORS:
            return mode, False
        raise ValueError(f"Zone '{zone}' has an unsupported color selection '{mode}'.")

    def _active_zones(self) -> list[str]:
        zones = list(PRIMARY_ZONES)
        last_optional = -1
        for idx, zone in enumerate(OPTIONAL_ZONES):
            if self.zone_include[zone].get():
                last_optional = idx
        if last_optional >= 0:
            zones.extend(OPTIONAL_ZONES[: last_optional + 1])
        return zones

    def _build_apply_args(self) -> list[str]:
        zones = self._active_zones()
        colors: list[str] = []
        has_custom = False
        for zone in zones:
            value, is_custom = self._zone_cli_value(zone)
            colors.append(value)
            has_custom = has_custom or is_custom

        args = [",".join(colors)]
        brightness: str | None = None
        if self.use_brightness.get():
            brightness = self.brightness.get()
        elif self.compat_mode.get() and not has_custom:
            # This route mirrors the known stable command path on affected devices.
            brightness = "high"

        if has_custom and brightness in ("high", "medium", "low"):
            raise ValueError(
                "Custom RGB values cannot be combined with high/medium/low brightness in MSIKLM. "
                "Use 'rgb' or disable explicit brightness."
            )

        if brightness:
            args.append(brightness)
        if self.use_mode.get():
            args.append(self.mode.get())
        return args

    def _set_action_buttons_state(self, enabled: bool) -> None:
        state = tk.NORMAL if enabled else tk.DISABLED
        for widget in (self.apply_button, self.mode_button, self.test_button):
            if widget is not None:
                widget.configure(state=state)

    @staticmethod
    def _arecord_cmd(device: str | None = None) -> list[str]:
        cmd = ["arecord", "-q"]
        if device:
            cmd.extend(["-D", device])
        cmd.extend(
            [
                "-f",
                "S16_LE",
                "-c",
                "1",
                "-r",
                str(VOICE_SAMPLE_RATE),
                "-t",
                "raw",
            ]
        )
        return cmd

    @staticmethod
    def _parec_cmd(device: str | None = None) -> list[str]:
        cmd = [
            "parec",
            "--raw",
            "--format=s16le",
            "--rate",
            str(VOICE_SAMPLE_RATE),
            "--channels",
            "1",
        ]
        if device:
            cmd.extend(["--device", device])
        return cmd

    def _detect_pulse_sources(self) -> list[tuple[str, list[str]]]:
        if not shutil.which("parec"):
            return []

        sources: list[tuple[str, list[str]]] = [("Pulse default", self._parec_cmd())]
        if not shutil.which("pactl"):
            return sources

        source_names: list[str] = []
        default_source = ""
        try:
            proc = subprocess.run(
                ["pactl", "get-default-source"],
                text=True,
                capture_output=True,
                timeout=5,
                check=False,
            )
            default_source = (proc.stdout or "").strip()
        except Exception:  # pylint: disable=broad-except
            default_source = ""

        try:
            proc = subprocess.run(
                ["pactl", "list", "short", "sources"],
                text=True,
                capture_output=True,
                timeout=5,
                check=False,
            )
        except Exception:  # pylint: disable=broad-except
            return sources

        for raw_line in (proc.stdout or "").splitlines():
            parts = raw_line.strip().split()
            if len(parts) < 2:
                continue
            source_name = parts[1]
            source_names.append(source_name)

        deduped: list[str] = []
        seen: set[str] = set()
        if default_source and default_source in source_names:
            deduped.append(default_source)
            seen.add(default_source)
        for source_name in source_names:
            if source_name in seen:
                continue
            deduped.append(source_name)
            seen.add(source_name)

        non_monitor = [name for name in deduped if not name.endswith(".monitor")]
        monitors = [name for name in deduped if name.endswith(".monitor")]
        source_names = non_monitor + monitors
        for source_name in source_names:
            sources.append((f"Pulse {source_name}", self._parec_cmd(source_name)))
        return sources

    def _detect_alsa_sources(self) -> list[tuple[str, list[str]]]:
        if not shutil.which("arecord"):
            return []

        sources: list[tuple[str, list[str]]] = [("ALSA default", self._arecord_cmd())]
        try:
            proc = subprocess.run(
                ["arecord", "-l"],
                text=True,
                capture_output=True,
                timeout=5,
                check=False,
            )
        except Exception:  # pylint: disable=broad-except
            return sources

        seen_devices: set[str] = set()
        for line in (proc.stdout or "").splitlines():
            match = re.search(r"card\s+(\d+):.*device\s+(\d+):", line, re.IGNORECASE)
            if not match:
                continue
            card_id = match.group(1)
            device_id = match.group(2)
            device_name = f"plughw:{card_id},{device_id}"
            if device_name in seen_devices:
                continue
            seen_devices.add(device_name)
            sources.append((f"ALSA card {card_id} device {device_id}", self._arecord_cmd(device_name)))
        return sources

    def _detect_voice_sources(self) -> list[tuple[str, list[str]]]:
        candidates = self._detect_pulse_sources() + self._detect_alsa_sources()
        unique: list[tuple[str, list[str]]] = []
        seen_cmds: set[tuple[str, ...]] = set()
        for label, cmd in candidates:
            key = tuple(cmd)
            if key in seen_cmds:
                continue
            seen_cmds.add(key)
            unique.append((label, cmd))
        return unique

    def _refresh_voice_sources(self) -> None:
        previous = self.voice_source_var.get()
        self._voice_sources = self._detect_voice_sources()

        values = [VOICE_AUTO_SOURCE]
        values.extend(label for label, _cmd in self._voice_sources)
        if self.voice_source_combo is not None:
            self.voice_source_combo.configure(values=values)

        if previous in values:
            self.voice_source_var.set(previous)
        else:
            self.voice_source_var.set(VOICE_AUTO_SOURCE)

        if self.voice_mode_enabled.get():
            return
        self._voice_active_source = "none"
        self.voice_active_source_var.set("Active source: none")
        if self._voice_sources:
            self.voice_status_var.set(f"Voice mode is off. Detected {len(self._voice_sources)} input source(s).")
        else:
            self.voice_status_var.set("Voice mode is off. No compatible microphone source detected.")

    def _selected_voice_sources(self) -> list[tuple[str, list[str]]]:
        selected = self.voice_source_var.get()
        if selected == VOICE_AUTO_SOURCE:
            return list(self._voice_sources)
        return [(label, cmd) for label, cmd in self._voice_sources if label == selected]

    def _on_voice_toggle(self) -> None:
        if self.voice_mode_enabled.get():
            self._start_voice_mode()
            return
        self._stop_voice_mode(push_off=True)

    def _start_voice_mode(self) -> None:
        if os.geteuid() != 0:
            messagebox.showerror(
                "Voice Mode Requires Root",
                "Voice mode streams continuous writes to the keyboard and requires an elevated GUI session.",
            )
            self.voice_mode_enabled.set(False)
            return

        if self._voice_thread and self._voice_thread.is_alive():
            return

        self._refresh_voice_sources()
        capture_sources = self._selected_voice_sources()
        if not capture_sources:
            messagebox.showerror(
                "Microphone Not Detected",
                "No compatible microphone source was detected.\n"
                "Install alsa-utils (arecord) and/or pulseaudio-utils (parec), then click Refresh.",
            )
            self.voice_mode_enabled.set(False)
            return

        if not self._resolve_msiklm():
            messagebox.showerror("MSIKLM Not Found", "Could not find 'msiklm'. Build/install it first.")
            self.voice_mode_enabled.set(False)
            return
        probe_ok, probe_err = self._run_msiklm_quiet(["help"], root_required=False)
        if not probe_ok:
            messagebox.showerror(
                "MSIKLM Runtime Error",
                "Voice mode cannot control the keyboard because the msiklm binary failed to start.\n"
                f"Details: {probe_err}",
            )
            self.voice_mode_enabled.set(False)
            return

        self._voice_stop_event.clear()
        self._voice_palette_phase = 0.0
        self._voice_last_payload = "off,off,off"
        self._voice_smoothed_level = 0.0
        self._voice_silence_frames = 0
        self._voice_last_send_ts = 0.0
        self._voice_active_source = "probing..."
        self.voice_active_source_var.set("Active source: probing...")
        self.voice_status_var.set("Voice mode active. Probing microphone source...")
        self.status_var.set("Voice mode active.")
        self._set_action_buttons_state(False)
        self._update_command_preview()
        listed_sources = ", ".join(label for label, _cmd in capture_sources)
        self._append_log(f"Voice source candidates: {listed_sources}")
        if self.compat_mode.get():
            self._append_log("Voice mode output path: compatibility (named colors + high brightness).")
        else:
            self._append_log("Voice mode output path: custom RGB.")

        self._voice_thread = threading.Thread(
            target=self._voice_worker,
            args=(capture_sources,),
            daemon=True,
        )
        self._voice_thread.start()

    def _stop_voice_mode(self, push_off: bool) -> None:
        self._voice_stop_event.set()
        thread = self._voice_thread
        if thread and thread.is_alive():
            thread.join(timeout=1.5)
        self._voice_thread = None

        self.voice_status_var.set("Voice mode is off.")
        self._voice_active_source = "none"
        self.voice_active_source_var.set("Active source: none")
        self.status_var.set("Ready.")
        self._set_action_buttons_state(True)

        for zone in PRIMARY_ZONES:
            self._voice_preview_hex[zone] = "#000000"
        self._draw_voice_meter([0.0, 0.0, 0.0])
        self._refresh_zone_swatches()
        self._redraw_keyboard()
        self._update_command_preview()

        if push_off:
            self._run_msiklm_quiet(["off,off,off"], root_required=True)

    def _voice_worker(self, capture_sources: list[tuple[str, list[str]]]) -> None:
        process: subprocess.Popen[bytes] | None = None
        bars = [0.0, 0.0, 0.0]
        ended_with_error = ""
        chunk_bytes = VOICE_SAMPLES_PER_FRAME * 2

        try:
            selected_label = ""
            for source_label, capture_cmd in capture_sources:
                if self._voice_stop_event.is_set():
                    break
                try:
                    process = subprocess.Popen(
                        capture_cmd,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.DEVNULL,
                        bufsize=0,
                    )
                except Exception as exc:  # pylint: disable=broad-except
                    ended_with_error = f"{source_label}: failed to open stream ({exc})"
                    process = None
                    continue

                if process.stdout is None:
                    ended_with_error = f"{source_label}: stream unavailable."
                    self._terminate_process(process)
                    process = None
                    continue

                probe_data = process.stdout.read(chunk_bytes)
                if probe_data:
                    selected_label = source_label
                    if not self._closing:
                        self.after(0, lambda lbl=source_label: self._set_active_voice_source(lbl))
                    try:
                        self._process_voice_chunk(probe_data, bars)
                    except RuntimeError as exc:
                        ended_with_error = f"{source_label}: {exc}"
                        self._terminate_process(process)
                        process = None
                        continue
                    break

                self._terminate_process(process)
                process = None
                ended_with_error = f"{source_label}: no audio stream data."

            if process is None:
                if not ended_with_error:
                    ended_with_error = "No microphone source produced an audio stream."
                return

            while not self._voice_stop_event.is_set():
                if process.stdout is None:
                    ended_with_error = f"{selected_label}: stream unavailable."
                    break
                data = process.stdout.read(chunk_bytes)
                if not data:
                    ended_with_error = f"{selected_label}: stream ended."
                    break
                try:
                    self._process_voice_chunk(data, bars)
                except RuntimeError as exc:
                    ended_with_error = f"{selected_label}: {exc}"
                    break
        finally:
            self._terminate_process(process)
            if not self._closing:
                self.after(0, lambda msg=ended_with_error: self._voice_worker_finished(msg))

    def _set_active_voice_source(self, source_label: str) -> None:
        self._voice_active_source = source_label
        self.voice_active_source_var.set(f"Active source: {source_label}")
        self.voice_status_var.set(f"Voice mode active on {source_label}.")
        self._append_log(f"Voice source in use: {source_label}")

    def _process_voice_chunk(self, data: bytes, bars: list[float]) -> None:
        raw_level = self._pcm16_rms(data)
        level = min(1.0, raw_level * self._voice_gain_value)
        threshold = self._voice_threshold_value
        if level > self._voice_smoothed_level:
            alpha = VOICE_ATTACK_ALPHA
        else:
            if level > threshold:
                alpha = VOICE_RELEASE_ALPHA_ACTIVE
            else:
                alpha = VOICE_RELEASE_ALPHA_SILENT
        self._voice_smoothed_level += (level - self._voice_smoothed_level) * alpha
        smooth_level = self._voice_smoothed_level

        if smooth_level <= threshold:
            self._voice_silence_frames += 1
        else:
            self._voice_silence_frames = 0

        colors, new_bars = self._voice_frame_colors(smooth_level, threshold, bars, self._voice_silence_frames)
        bars[0] = new_bars[0]
        bars[1] = new_bars[1]
        bars[2] = new_bars[2]
        if self.compat_mode.get():
            compat_colors: list[str] = []
            for color in colors:
                if color == "off":
                    compat_colors.append("off")
                else:
                    compat_colors.append(self._nearest_voice_compat_color(color))
            cmd_args = [",".join(compat_colors), "high"]
            payload_key = "|".join(cmd_args)
        else:
            cmd_args = [",".join(colors)]
            payload_key = cmd_args[0]

        now = time.monotonic()
        send_interval = VOICE_MIN_SEND_INTERVAL_ACTIVE if smooth_level > threshold else VOICE_MIN_SEND_INTERVAL_SILENT
        should_send = (
            payload_key != self._voice_last_payload
            and (
                (now - self._voice_last_send_ts) >= send_interval
                or self._voice_last_payload == ""
            )
        )

        if should_send:
            ok, err = self._run_msiklm_quiet(cmd_args, root_required=True)
            if not ok:
                raise RuntimeError(err or "Voice write failed.")
            self._voice_last_payload = payload_key
            self._voice_last_send_ts = now
        if not self._closing:
            self.after(0, lambda c=colors, b=bars.copy(), l=smooth_level: self._apply_voice_frame(c, b, l))

    @staticmethod
    def _terminate_process(process: subprocess.Popen[bytes] | None) -> None:
        if process is None:
            return
        try:
            process.terminate()
        except Exception:  # pylint: disable=broad-except
            pass
        try:
            process.wait(timeout=0.8)
        except Exception:  # pylint: disable=broad-except
            try:
                process.kill()
            except Exception:  # pylint: disable=broad-except
                pass

    def _voice_worker_finished(self, error_message: str) -> None:
        if self._closing:
            return
        was_enabled = self.voice_mode_enabled.get()
        self.voice_mode_enabled.set(False)
        self._stop_voice_mode(push_off=was_enabled)
        if was_enabled and error_message:
            self.voice_status_var.set(f"Voice mode stopped: {error_message}")
            self.status_var.set("Voice mode failed.")
            self._append_log(error_message)

    @staticmethod
    def _pcm16_rms(data: bytes) -> float:
        if len(data) < 2:
            return 0.0
        clipped = data[: len(data) - (len(data) % 2)]
        if not clipped:
            return 0.0
        samples = array.array("h")
        samples.frombytes(clipped)
        if sys.byteorder != "little":
            samples.byteswap()
        if not samples:
            return 0.0
        energy = 0
        for value in samples:
            energy += int(value) * int(value)
        rms = (energy / len(samples)) ** 0.5
        return rms / 32768.0

    def _palette_hex(self, phase: float) -> str:
        palette_size = len(VOICE_COLOR_PATH)
        wrapped = phase % palette_size
        idx = int(wrapped)
        next_idx = (idx + 1) % palette_size
        ratio = wrapped - idx
        left_hex = PRESET_COLORS[VOICE_COLOR_PATH[idx]]
        right_hex = PRESET_COLORS[VOICE_COLOR_PATH[next_idx]]
        return blend_hex(left_hex, right_hex, ratio)

    @staticmethod
    def _hex_to_rgb(hex_color: str) -> tuple[int, int, int]:
        return (
            int(hex_color[1:3], 16),
            int(hex_color[3:5], 16),
            int(hex_color[5:7], 16),
        )

    def _nearest_voice_compat_color(self, hex_color: str) -> str:
        r, g, b = self._hex_to_rgb(hex_color)
        best_name = "blue"
        best_dist = 1 << 62
        for name in VOICE_COMPAT_COLORS:
            cr, cg, cb = self._hex_to_rgb(PRESET_COLORS[name])
            dr = r - cr
            dg = g - cg
            db = b - cb
            dist = (dr * dr) + (dg * dg) + (db * db)
            if dist < best_dist:
                best_dist = dist
                best_name = name
        return best_name

    def _voice_frame_colors(
        self,
        level: float,
        threshold: float,
        previous_bars: list[float],
        silence_frames: int,
    ) -> tuple[list[str], list[float]]:
        if level <= threshold:
            bars = [value * VOICE_BAR_DECAY for value in previous_bars]
            if silence_frames >= VOICE_SILENCE_HOLD_FRAMES and max(bars) < 0.035:
                self._voice_palette_phase = (self._voice_palette_phase + 0.05) % len(VOICE_COLOR_PATH)
                return ["off", "off", "off"], [0.0, 0.0, 0.0]
            self._voice_palette_phase = (self._voice_palette_phase + 0.04) % len(VOICE_COLOR_PATH)
        else:
            active = (level - threshold) / max(0.001, 1.0 - threshold)
            active = max(0.0, min(1.0, active))
            self._voice_palette_phase = (self._voice_palette_phase + 0.09 + (active * 0.28)) % len(VOICE_COLOR_PATH)

            bars = [0.0, 0.0, 0.0]
            bars[0] = active
            bars[1] = max(active * 0.66, previous_bars[0] * 0.84)
            bars[2] = max(active * 0.44, previous_bars[1] * 0.82)

        colors: list[str] = []
        for idx, bar in enumerate(bars):
            if bar < 0.02:
                colors.append("off")
                continue
            phase = self._voice_palette_phase + (idx * 0.58)
            base = self._palette_hex(phase)
            strength = 0.16 + (0.76 * bar)
            colors.append(blend_hex("#000000", base, strength))
        return colors, bars

    def _apply_voice_frame(self, colors: list[str], bars: list[float], level: float) -> None:
        for zone, color in zip(PRIMARY_ZONES, colors):
            if color == "off":
                self._voice_preview_hex[zone] = "#000000"
            else:
                self._voice_preview_hex[zone] = color
        self.voice_status_var.set(
            f"Voice mode active on {self._voice_active_source}. "
            f"Input level: {level:.2f} (threshold {self._voice_threshold_value:.2f})"
        )
        self._draw_voice_meter(bars)
        self._refresh_zone_swatches()
        self._redraw_keyboard()

    def _draw_voice_meter(self, bars: list[float]) -> None:
        if self.voice_meter is None:
            return
        canvas = self.voice_meter
        canvas.delete("all")
        width = int(canvas.cget("width"))
        height = int(canvas.cget("height"))
        pad = 6
        gap = 6
        bar_w = (width - (2 * pad) - (2 * gap)) / 3.0
        for idx, value in enumerate(bars):
            x1 = pad + idx * (bar_w + gap)
            x2 = x1 + bar_w
            y2 = height - pad
            y1 = pad + ((1.0 - max(0.0, min(1.0, value))) * (height - (2 * pad)))
            phase = self._voice_palette_phase + (idx * 0.58)
            fill = self._palette_hex(phase) if value > 0.03 else "#1a2740"
            canvas.create_rectangle(x1, y1, x2, y2, fill=fill, outline=darken(fill, 0.55), width=1)
            canvas.create_text((x1 + x2) / 2, y2 - 2, text=PRIMARY_ZONES[idx].upper(), fill="#adc4e8", anchor=tk.S)

    def _set_log(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.insert("1.0", text)
        self.log_text.configure(state=tk.DISABLED)

    def _append_log(self, text: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, text + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _resolve_msiklm(self) -> str | None:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        repo_root = os.path.dirname(script_dir)
        candidates = [
            os.path.join(repo_root, "msiklm"),
            "/usr/local/bin/msiklm",
            shutil.which("msiklm"),
        ]
        for candidate in candidates:
            if candidate and os.path.isfile(candidate) and os.access(candidate, os.X_OK):
                return candidate
        return None

    def _build_command_attempts(self, executable: str, args: list[str], root_required: bool) -> list[list[str]]:
        if root_required and os.geteuid() != 0:
            attempts: list[list[str]] = []
            if shutil.which("pkexec"):
                attempts.append(["pkexec", executable, *args])
            if shutil.which("sudo"):
                attempts.append(["sudo", executable, *args])
            if attempts:
                return attempts
        return [[executable, *args]]

    def _run_msiklm_quiet(self, args: list[str], root_required: bool = True) -> tuple[bool, str]:
        executable = self._resolve_msiklm()
        if not executable:
            return False, "Could not find 'msiklm'."

        attempts = self._build_command_attempts(executable, args, root_required)
        first_error = ""
        retry_count = 3 if self.compat_mode.get() and args and args[0] not in ("help", "list", "test") else 1

        for cmd in attempts:
            for attempt_idx in range(retry_count):
                try:
                    proc = subprocess.run(
                        cmd,
                        text=True,
                        capture_output=True,
                        timeout=20,
                        check=False,
                    )
                except Exception as exc:  # pylint: disable=broad-except
                    text = f"{type(exc).__name__}: {exc}"
                    if not first_error:
                        first_error = text
                    if attempt_idx + 1 < retry_count:
                        time.sleep(0.06)
                    continue

                if proc.returncode == 0:
                    return True, ""

                output = ((proc.stdout or "") + (proc.stderr or "")).strip()
                if output and not first_error:
                    first_error = output
                if not first_error:
                    first_error = f"Command exited with code {proc.returncode}."
                if attempt_idx + 1 < retry_count:
                    time.sleep(0.06)

        return False, first_error or "Command failed."

    def _run_msiklm(self, args: list[str], root_required: bool = True) -> None:
        executable = self._resolve_msiklm()
        self._refresh_binary_label()
        if not executable:
            messagebox.showerror(
                "MSIKLM Not Found",
                "Could not find 'msiklm'. Build/install it first (for example: make && sudo make install).",
            )
            return

        attempts = self._build_command_attempts(executable, args, root_required)
        self.status_var.set("Running command...")
        self._set_log("")
        first_error = ""

        for cmd in attempts:
            retry_count = 1
            if self.compat_mode.get() and cmd[0] == executable and args and args[0] not in ("help", "list", "test"):
                retry_count = 3

            for attempt_idx in range(retry_count):
                cmd_text = " ".join(shlex.quote(part) for part in cmd)
                if retry_count > 1:
                    self._append_log(f"$ {cmd_text}   [attempt {attempt_idx + 1}/{retry_count}]")
                else:
                    self._append_log(f"$ {cmd_text}")

                try:
                    proc = subprocess.run(
                        cmd,
                        text=True,
                        capture_output=True,
                        timeout=45,
                        check=False,
                    )
                except Exception as exc:  # pylint: disable=broad-except
                    error_text = f"{type(exc).__name__}: {exc}"
                    self._append_log(error_text)
                    if not first_error:
                        first_error = error_text
                    break

                output = (proc.stdout or "") + (proc.stderr or "")
                if output.strip():
                    self._append_log(output.rstrip())

                if proc.returncode == 0:
                    self.status_var.set("Success.")
                    return

                if not first_error:
                    first_error = f"Command exited with code {proc.returncode}."
                self._append_log(f"Command exited with code {proc.returncode}.")

                if attempt_idx + 1 < retry_count:
                    time.sleep(0.12)

        self.status_var.set("Command failed.")
        messagebox.showerror("Command Failed", first_error or "Unknown error while running msiklm.")

    def _apply_colors(self) -> None:
        if self.voice_mode_enabled.get():
            messagebox.showinfo("Voice Mode Active", "Disable voice mode before applying manual colors.")
            return
        try:
            args = self._build_apply_args()
        except ValueError as err:
            messagebox.showerror("Invalid Input", str(err))
            return
        self._run_msiklm(args, root_required=True)
        self._update_command_preview()

    def _apply_mode_only(self) -> None:
        if self.voice_mode_enabled.get():
            messagebox.showinfo("Voice Mode Active", "Disable voice mode before applying mode-only commands.")
            return
        self._run_msiklm([self.mode.get()], root_required=True)

    def _test_connection(self) -> None:
        if self.voice_mode_enabled.get():
            messagebox.showinfo("Voice Mode Active", "Disable voice mode before running a connection test.")
            return
        self._run_msiklm(["test"], root_required=True)

    def _show_help(self) -> None:
        self._run_msiklm(["help"], root_required=False)

    def _update_command_preview(self) -> None:
        if self.voice_mode_enabled.get():
            if self.compat_mode.get():
                self.preview_var.set("voice mode active: compatibility stream (named colors + high brightness)")
            else:
                self.preview_var.set("voice mode active: microphone-reactive custom RGB stream (left,middle,right)")
            return
        try:
            args = self._build_apply_args()
            prefix = "msiklm" if os.geteuid() == 0 else "sudo msiklm"
            self.preview_var.set(prefix + " " + " ".join(shlex.quote(arg) for arg in args))
        except ValueError as err:
            self.preview_var.set(f"Invalid: {err}")

    def _zone_visual_color(self, zone: str) -> str:
        if self.voice_mode_enabled.get() and zone in PRIMARY_ZONES:
            return darken(self._voice_preview_hex[zone], 0.86)
        base = self._zone_hex_or_fallback(zone)
        if not self.zone_include[zone].get():
            return darken(base, 0.18)
        return darken(base, 0.86)

    def _round_rect(self, x1: float, y1: float, x2: float, y2: float, radius: float, **kwargs: object) -> int:
        points = [
            x1 + radius,
            y1,
            x2 - radius,
            y1,
            x2,
            y1,
            x2,
            y1 + radius,
            x2,
            y2 - radius,
            x2,
            y2,
            x2 - radius,
            y2,
            x1 + radius,
            y2,
            x1,
            y2,
            x1,
            y2 - radius,
            x1,
            y1 + radius,
            x1,
            y1,
        ]
        return self.canvas.create_polygon(points, smooth=True, splinesteps=18, **kwargs)

    def _draw_key(
        self,
        x: float,
        y: float,
        w: float,
        h: float,
        label: str,
        zone: str | None,
        row_idx: int,
        zone_bounds: dict[str, list[float]],
        zone_rows: dict[str, dict[int, list[float]]],
    ) -> None:
        if zone is None:
            return

        fill = self._zone_visual_color(zone)
        stroke = darken(fill, 0.48)
        shadow = darken(fill, 0.2)
        self._round_rect(x + 1, y + 2, x + w + 1, y + h + 2, 7, fill=shadow, outline="")
        self._round_rect(x, y, x + w, y + h, 7, fill=fill, outline=stroke, width=1)
        if label and label != "Space":
            self.canvas.create_text(
                x + (w / 2.0),
                y + (h / 2.0),
                text=label,
                fill=lighten(fill, 0.82),
                font=(FONT_UI, 8),
            )

        if zone in PRIMARY_ZONES:
            bounds = zone_bounds[zone]
            bounds[0] = min(bounds[0], x)
            bounds[1] = min(bounds[1], y)
            bounds[2] = max(bounds[2], x + w)
            bounds[3] = max(bounds[3], y + h)

            row_bounds = zone_rows[zone].get(row_idx)
            if row_bounds is None:
                zone_rows[zone][row_idx] = [x, y, x + w, y + h]
            else:
                row_bounds[0] = min(row_bounds[0], x)
                row_bounds[1] = min(row_bounds[1], y)
                row_bounds[2] = max(row_bounds[2], x + w)
                row_bounds[3] = max(row_bounds[3], y + h)

    def _draw_optional_zone_badges(self) -> None:
        badges = [
            ("logo", 808, 376, 102, 26, "Logo"),
            ("front_left", 196, 376, 118, 24, "Front L"),
            ("front_right", 618, 376, 118, 24, "Front R"),
            ("mouse", 808, 330, 102, 34, "Mouse"),
        ]
        for zone, x, y, w, h, label in badges:
            fill = self._zone_visual_color(zone)
            stroke = darken(fill, 0.55)
            self._round_rect(x, y, x + w, y + h, 8, fill=darken(fill, 0.88), outline=stroke, width=1)
            self.canvas.create_text(
                x + (w / 2),
                y + (h / 2),
                text=label,
                fill=lighten(fill, 0.62),
                font=(FONT_UI, 8, "bold"),
            )

    def _draw_primary_zone_split(
        self,
        bounds: dict[str, list[float]],
        zone_rows: dict[str, dict[int, list[float]]],
    ) -> None:
        colors = {"left": "#5e86b8", "middle": "#6993bd", "right": "#749fbe"}

        for zone in PRIMARY_ZONES:
            x1, y1, x2, y2 = bounds[zone]
            if x2 <= x1:
                continue

            chip_cx = (x1 + x2) / 2
            chip_y = y1 - 30
            chip_w = 118
            self._round_rect(
                chip_cx - (chip_w / 2),
                chip_y - 4,
                chip_cx + (chip_w / 2),
                chip_y + 20,
                10,
                fill=darken(colors[zone], 0.28),
                outline=darken(colors[zone], 0.78),
                width=1,
            )
            self.canvas.create_text(
                chip_cx,
                chip_y + 10,
                text=f"{zone.upper()} ZONE",
                fill=colors[zone],
                font=(FONT_UI, 9, "bold"),
            )

        def seam_polyline(zone_a: str, zone_b: str) -> tuple[list[float], float]:
            rows = sorted(set(zone_rows[zone_a]).intersection(zone_rows[zone_b]))
            if not rows:
                return [], 0.0

            # Each row produces one vertical seam segment, then we connect segments only in row gaps.
            # tuple: (seam_x, top_y, bottom_y, min_allowed_x, max_allowed_x)
            segments: list[tuple[float, float, float, float, float]] = []
            narrowest_gap = 9999.0
            inset = 5
            for row in rows:
                ax1, ay1, ax2, ay2 = zone_rows[zone_a][row]
                bx1, by1, bx2, by2 = zone_rows[zone_b][row]
                seam_x = (ax2 + bx1) / 2
                top_y = min(ay1, by1) + inset
                bottom_y = max(ay2, by2) - inset
                if bottom_y <= top_y:
                    top_y = min(ay1, by1) + 2
                    bottom_y = max(ay2, by2) - 2

                # Keep seam inside the actual inter-zone gap for this row.
                min_allowed = ax2 + 2
                max_allowed = bx1 - 2
                if max_allowed < min_allowed:
                    min_allowed = max_allowed = seam_x
                seam_x = max(min_allowed, min(max_allowed, seam_x))
                narrowest_gap = min(narrowest_gap, max(0.0, bx1 - ax2))

                segments.append((seam_x, top_y, bottom_y, min_allowed, max_allowed))

            # Damp sharp horizontal jumps (mostly visible on middle|right seam near numpad/enter area).
            smooth_segments: list[tuple[float, float, float]] = []
            for idx, (x, top, bottom, min_allowed, max_allowed) in enumerate(segments):
                if idx == 0 or idx == len(segments) - 1:
                    smooth_x = x
                else:
                    prev_x = segments[idx - 1][0]
                    next_x = segments[idx + 1][0]
                    smooth_x = (prev_x + (2.0 * x) + next_x) / 4.0
                smooth_x = max(min_allowed, min(max_allowed, smooth_x))
                smooth_segments.append((smooth_x, top, bottom))

            points: list[float] = []
            x0, y0_top, y0_bottom = smooth_segments[0]
            points.extend([x0, y0_top, x0, y0_bottom])
            for idx in range(1, len(smooth_segments)):
                prev_x, _prev_top, prev_bottom = smooth_segments[idx - 1]
                x, top, bottom = smooth_segments[idx]
                bridge_y = (prev_bottom + top) / 2.0
                points.extend([prev_x, bridge_y, x, bridge_y, x, top, x, bottom])

            if narrowest_gap == 9999.0:
                narrowest_gap = 0.0
            return points, narrowest_gap

        for seam, gap_width in (
            seam_polyline("left", "middle"),
            seam_polyline("middle", "right"),
        ):
            if len(seam) < 4:
                continue
            glow_width = max(1.8, min(3.0, gap_width - 1.2))
            core_width = max(1.0, glow_width - 1.2)
            self.canvas.create_line(
                seam,
                fill=SEAM_GLOW,
                width=glow_width,
                capstyle=tk.ROUND,
                joinstyle=tk.ROUND,
            )
            self.canvas.create_line(
                seam,
                fill=SEAM_COLOR,
                width=core_width,
                capstyle=tk.ROUND,
                joinstyle=tk.ROUND,
            )

    def _redraw_keyboard(self) -> None:
        self.canvas.delete("all")

        self._round_rect(18, 18, 948, 410, 22, fill="#0f1b30", outline="#2a4164", width=2)
        self._round_rect(28, 24, 938, 56, 14, fill="#0d1a2d", outline="#223656", width=1)
        self.canvas.create_text(
            483,
            28,
            text="Keyboard Zone Layout",
            fill=FG_MAIN,
            font=(FONT_UI, 11, "bold"),
        )

        unit = 31
        key_h = 42
        gap = 4
        start_x = 44
        start_y = 74
        zone_bounds = {zone: [9999.0, 9999.0, -1.0, -1.0] for zone in PRIMARY_ZONES}
        zone_rows = {zone: {} for zone in PRIMARY_ZONES}

        for row_idx, row in enumerate(KEY_LAYOUT):
            x = start_x + ROW_OFFSETS[row_idx]
            y = start_y + row_idx * (key_h + gap)
            for label, width_units, zone in row:
                width = width_units * unit
                if zone is None:
                    x += width + gap
                    continue
                self._draw_key(x, y, width, key_h, label, zone, row_idx, zone_bounds, zone_rows)
                x += width + gap

        self._draw_primary_zone_split(zone_bounds, zone_rows)
        self._draw_optional_zone_badges()

        legend_y = 422
        x = 36
        for zone in ALL_ZONES:
            fill = self._zone_visual_color(zone)
            self._round_rect(x, legend_y, x + 24, legend_y + 14, 4, fill=fill, outline=darken(fill, 0.5), width=1)
            status = "" if zone in PRIMARY_ZONES or self.zone_include[zone].get() else " (off)"
            self.canvas.create_text(
                x + 30,
                legend_y + 7,
                text=zone + status,
                fill=FG_MUTED,
                anchor=tk.W,
                font=(FONT_UI, 8),
            )
            x += 128

    def _on_close(self) -> None:
        self._closing = True
        self.voice_mode_enabled.set(False)
        self._stop_voice_mode(push_off=False)
        self.destroy()


def main() -> None:
    args = parse_args()
    relaunch_as_root(args)
    app = MSIKLMGui(launched_as_root=args.as_root or os.geteuid() == 0)
    app.mainloop()


if __name__ == "__main__":
    main()
