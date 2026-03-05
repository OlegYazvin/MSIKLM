#!/usr/bin/env python3
"""Modern GUI wrapper for MSIKLM with explicit 3-zone keyboard mapping."""

from __future__ import annotations

import argparse
import os
import re
import shlex
import shutil
import subprocess
import sys
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
FONT_UI = "DejaVu Sans"
FONT_MONO = "DejaVu Sans Mono"


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
        self.geometry("1360x900")
        self.minsize(1220, 780)
        self.configure(bg=BG_APP)
        self.launched_as_root = launched_as_root

        self.zone_color_mode: dict[str, tk.StringVar] = {}
        self.zone_custom_hex: dict[str, tk.StringVar] = {}
        self.zone_include: dict[str, tk.BooleanVar] = {}
        self.zone_swatch: dict[str, tk.Canvas] = {}
        self._updating_optional = False

        self.use_brightness = tk.BooleanVar(value=False)
        self.use_mode = tk.BooleanVar(value=False)
        self.compat_mode = tk.BooleanVar(value=True)
        self.brightness = tk.StringVar(value="high")
        self.mode = tk.StringVar(value="normal")
        self.preview_var = tk.StringVar(value="")
        self.status_var = tk.StringVar(value="Ready.")
        self.binary_var = tk.StringVar(value="")

        self._setup_styles()
        self._build_ui()
        self._set_defaults()
        self._refresh_binary_label()
        self._update_command_preview()
        self._redraw_keyboard()

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
        style.configure("TLabel", background=BG_PANEL, foreground=FG_MAIN)
        style.configure("Muted.TLabel", background=BG_PANEL, foreground=FG_MUTED)
        style.configure("Title.TLabel", background=BG_APP, foreground=lighten(ACCENT, 0.4), font=(FONT_UI, 21, "bold"))
        style.configure("Subtitle.TLabel", background=BG_APP, foreground="#7fa2d3", font=(FONT_UI, 10))
        style.configure("Accent.TButton", background=ACCENT, foreground="#ffffff", borderwidth=0, padding=8)
        style.map(
            "Accent.TButton",
            background=[("active", lighten(ACCENT, 0.1)), ("pressed", ACCENT_DARK)],
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
        right = ttk.Frame(body, style="PanelAlt.TFrame", padding=12)
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        right.pack(side=tk.LEFT, fill=tk.Y, padx=(12, 0))

        self.canvas = tk.Canvas(left, width=980, height=460, bg="#0a1628", highlightthickness=0)
        self.canvas.pack(fill=tk.X, expand=False)

        ttk.Label(
            left,
            text="3-zone split (denoted on keyboard): LEFT | MIDDLE | RIGHT. Optional zones are shown as badges.",
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

        zones_frame = ttk.LabelFrame(right, text="Zone Colors", style="Card.TLabelframe", padding=8)
        zones_frame.pack(fill=tk.X)
        ttk.Label(zones_frame, text="Use").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
        ttk.Label(zones_frame, text="Zone").grid(row=0, column=1, sticky=tk.W, padx=(0, 8))
        ttk.Label(zones_frame, text="Preview").grid(row=0, column=2, sticky=tk.W, padx=(0, 8))
        ttk.Label(zones_frame, text="Color").grid(row=0, column=3, sticky=tk.W, padx=(0, 8))
        ttk.Label(zones_frame, text="Hex").grid(row=0, column=4, sticky=tk.W, padx=(0, 8))

        row = 1
        for zone in ALL_ZONES:
            mode_var = tk.StringVar(value="blue" if zone in PRIMARY_ZONES else "off")
            hex_var = tk.StringVar(value="#0000ff" if zone in PRIMARY_ZONES else "#000000")
            include_var = tk.BooleanVar(value=zone in PRIMARY_ZONES)
            self.zone_color_mode[zone] = mode_var
            self.zone_custom_hex[zone] = hex_var
            self.zone_include[zone] = include_var

            check_state = tk.DISABLED if zone in PRIMARY_ZONES else tk.NORMAL
            check_cmd = None if zone in PRIMARY_ZONES else (lambda z=zone: self._on_optional_toggle(z))
            check = ttk.Checkbutton(zones_frame, variable=include_var, state=check_state, command=check_cmd)
            check.grid(row=row, column=0, sticky=tk.W, padx=(0, 8), pady=3)

            ttk.Label(zones_frame, text=zone.upper(), style="Muted.TLabel").grid(row=row, column=1, sticky=tk.W, padx=(0, 8), pady=3)

            swatch = tk.Canvas(zones_frame, width=18, height=18, bg=BG_PANEL, highlightthickness=0)
            swatch.grid(row=row, column=2, sticky=tk.W, padx=(0, 8), pady=3)
            self.zone_swatch[zone] = swatch

            combo = ttk.Combobox(
                zones_frame,
                values=COLOR_CHOICES,
                textvariable=mode_var,
                state="readonly",
                width=10,
            )
            combo.grid(row=row, column=3, sticky=tk.W, padx=(0, 8), pady=3)
            combo.bind("<<ComboboxSelected>>", lambda _evt, z=zone: self._on_zone_changed(z))

            entry = ttk.Entry(zones_frame, textvariable=hex_var, width=9)
            entry.grid(row=row, column=4, sticky=tk.W, padx=(0, 4), pady=3)
            entry.bind("<KeyRelease>", lambda _evt, z=zone: self._on_zone_changed(z))

            pick = ttk.Button(zones_frame, text="Pick", command=lambda z=zone: self._pick_color(z), width=6)
            pick.grid(row=row, column=4, sticky=tk.W, pady=3)
            row += 1

        opts = ttk.LabelFrame(right, text="Apply Behavior", style="Card.TLabelframe", padding=8)
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

        actions = ttk.LabelFrame(right, text="Actions", style="Card.TLabelframe", padding=8)
        actions.pack(fill=tk.X, pady=(10, 0))
        ttk.Button(actions, text="Apply Colors", style="Accent.TButton", command=self._apply_colors).pack(fill=tk.X, pady=2)
        ttk.Button(actions, text="Apply Mode Only", command=self._apply_mode_only).pack(fill=tk.X, pady=2)
        ttk.Button(actions, text="Test Connection", command=self._test_connection).pack(fill=tk.X, pady=2)
        ttk.Button(actions, text="Show CLI Help", command=self._show_help).pack(fill=tk.X, pady=2)

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
        try:
            args = self._build_apply_args()
        except ValueError as err:
            messagebox.showerror("Invalid Input", str(err))
            return
        self._run_msiklm(args, root_required=True)
        self._update_command_preview()

    def _apply_mode_only(self) -> None:
        self._run_msiklm([self.mode.get()], root_required=True)

    def _test_connection(self) -> None:
        self._run_msiklm(["test"], root_required=True)

    def _show_help(self) -> None:
        self._run_msiklm(["help"], root_required=False)

    def _update_command_preview(self) -> None:
        try:
            args = self._build_apply_args()
            prefix = "msiklm" if os.geteuid() == 0 else "sudo msiklm"
            self.preview_var.set(prefix + " " + " ".join(shlex.quote(arg) for arg in args))
        except ValueError as err:
            self.preview_var.set(f"Invalid: {err}")

    def _zone_visual_color(self, zone: str) -> str:
        base = self._zone_hex_or_fallback(zone)
        if zone in OPTIONAL_ZONES and not self.zone_include[zone].get():
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
                font=("Sans", 8),
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
            ("logo", 515, 32, 90, 30, "Logo"),
            ("front_left", 220, 356, 110, 24, "Front L"),
            ("front_right", 640, 356, 110, 24, "Front R"),
            ("mouse", 830, 184, 95, 36, "Mouse"),
        ]
        for zone, x, y, w, h, label in badges:
            fill = self._zone_visual_color(zone)
            stroke = darken(fill, 0.45)
            self._round_rect(x, y, x + w, y + h, 8, fill=fill, outline=stroke, width=2)
            self.canvas.create_text(x + (w / 2), y + (h / 2), text=label, fill=lighten(fill, 0.75), font=("Sans", 8, "bold"))

    def _draw_primary_zone_split(
        self,
        bounds: dict[str, list[float]],
        zone_rows: dict[str, dict[int, list[float]]],
    ) -> None:
        colors = {"left": "#72a7ff", "middle": "#7ac3ff", "right": "#89d9ff"}

        for zone in PRIMARY_ZONES:
            x1, y1, x2, y2 = bounds[zone]
            if x2 <= x1:
                continue

            rows = sorted(zone_rows[zone].keys())
            if rows:
                for row in rows:
                    rx1, ry1, rx2, ry2 = zone_rows[zone][row]
                    self._round_rect(
                        rx1 - 7,
                        ry1 - 7,
                        rx2 + 7,
                        ry2 + 7,
                        10,
                        fill="",
                        outline=darken(colors[zone], 0.8),
                        width=1,
                    )

            self.canvas.create_text(
                (x1 + x2) / 2,
                y1 - 22,
                text=f"{zone.upper()} ZONE",
                fill=colors[zone],
                font=("Sans", 9, "bold"),
            )

        def seam_polyline(zone_a: str, zone_b: str) -> list[float]:
            rows = sorted(set(zone_rows[zone_a]).intersection(zone_rows[zone_b]))
            if not rows:
                return []

            margin = 9
            points: list[float] = []
            prev_x = 0.0
            prev_bottom = 0.0
            for idx, row in enumerate(rows):
                ax1, ay1, ax2, ay2 = zone_rows[zone_a][row]
                bx1, by1, bx2, by2 = zone_rows[zone_b][row]
                seam_x = (ax2 + bx1) / 2
                top_y = min(ay1, by1) - margin
                bottom_y = max(ay2, by2) + margin

                if idx == 0:
                    points.extend([seam_x, top_y, seam_x, bottom_y])
                else:
                    connector_y = (prev_bottom + top_y) / 2.0
                    points.extend(
                        [
                            prev_x,
                            connector_y,
                            seam_x,
                            connector_y,
                            seam_x,
                            top_y,
                            seam_x,
                            bottom_y,
                        ]
                    )

                prev_x = seam_x
                prev_bottom = bottom_y

            return points

        for seam in (seam_polyline("left", "middle"), seam_polyline("middle", "right")):
            if len(seam) < 4:
                continue
            self.canvas.create_line(
                seam,
                fill="#7aa9e6",
                width=3,
                capstyle=tk.ROUND,
                joinstyle=tk.ROUND,
            )

    def _redraw_keyboard(self) -> None:
        self.canvas.delete("all")

        self._round_rect(18, 18, 948, 410, 22, fill="#0f1b30", outline="#2a4164", width=2)
        self.canvas.create_text(
            483,
            34,
            text="Keyboard Zone Layout",
            fill=FG_MAIN,
            font=("Sans", 11, "bold"),
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
                font=("Sans", 8),
            )
            x += 128


def main() -> None:
    args = parse_args()
    relaunch_as_root(args)
    app = MSIKLMGui(launched_as_root=args.as_root or os.geteuid() == 0)
    app.mainloop()


if __name__ == "__main__":
    main()
