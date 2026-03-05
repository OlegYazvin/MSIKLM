#!/bin/sh

# Install MSIKLM CLI + GUI launcher with desktop integration.
# Targets common modern distros (apt, dnf, yum, pacman, zypper).

set -eu

SKIP_DEPS=0
NO_PANEL_PIN=0
FORCE_USER_INSTALL=0

while [ $# -gt 0 ]; do
    case "$1" in
        --skip-deps)
            SKIP_DEPS=1
            ;;
        --no-panel-pin)
            NO_PANEL_PIN=1
            ;;
        --user-only)
            FORCE_USER_INSTALL=1
            ;;
        -h|--help)
            cat <<'EOF'
Usage: ./install_ui.sh [options]

Options:
  --skip-deps     Do not install distro packages.
  --no-panel-pin  Do not add the launcher to Cinnamon panel favorites.
  --user-only     Install binaries into ~/.local/bin only.
  -h, --help      Show this help.
EOF
            exit 0
            ;;
        *)
            echo "Unknown option: $1" >&2
            exit 1
            ;;
    esac
    shift
done

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)
cd "$SCRIPT_DIR"

if [ ! -f "gui/msiklm_gui.py" ] || [ ! -f "Makefile" ]; then
    echo "Run this script from the MSIKLM repository root." >&2
    exit 1
fi

run_as_root() {
    if [ "$(id -u)" -eq 0 ]; then
        "$@"
    elif command -v sudo >/dev/null 2>&1; then
        sudo "$@"
    else
        return 1
    fi
}

try_install_packages() {
    if [ "$SKIP_DEPS" -eq 1 ]; then
        echo "[deps] skipped."
        return 0
    fi

    echo "[deps] installing required packages..."
    if command -v apt >/dev/null 2>&1; then
        run_as_root apt install -y gcc make libhidapi-dev python3-tk desktop-file-utils
        return $?
    fi
    if command -v dnf >/dev/null 2>&1; then
        run_as_root dnf install -y gcc make hidapi-devel python3-tkinter desktop-file-utils
        return $?
    fi
    if command -v yum >/dev/null 2>&1; then
        run_as_root yum install -y gcc make hidapi-devel python3-tkinter desktop-file-utils
        return $?
    fi
    if command -v pacman >/dev/null 2>&1; then
        run_as_root pacman -S --noconfirm gcc make hidapi tk desktop-file-utils
        return $?
    fi
    if command -v zypper >/dev/null 2>&1; then
        run_as_root zypper --non-interactive install gcc make libhidapi-devel python3-tk desktop-file-utils
        return $?
    fi

    echo "[deps] no supported package manager found. Install deps manually." >&2
    return 1
}

echo "[build] compiling msiklm..."
make

try_install_packages || true

if [ "$FORCE_USER_INSTALL" -eq 1 ]; then
    BIN_DIR="$HOME/.local/bin"
elif [ -w /usr/local/bin ]; then
    BIN_DIR="/usr/local/bin"
else
    BIN_DIR="$HOME/.local/bin"
fi

mkdir -p "$BIN_DIR"

install_file() {
    src="$1"
    dst="$2"
    if [ -w "$(dirname "$dst")" ]; then
        install -m 0755 "$src" "$dst"
    else
        run_as_root install -m 0755 "$src" "$dst"
    fi
}

echo "[install] installing binaries into $BIN_DIR..."
install_file "$SCRIPT_DIR/msiklm" "$BIN_DIR/msiklm"
install_file "$SCRIPT_DIR/gui/msiklm_gui.py" "$BIN_DIR/msiklm-gui"

APP_DIR="$HOME/.local/share/applications"
mkdir -p "$APP_DIR"
DESKTOP_FILE="$APP_DIR/msiklm-gui.desktop"

cat >"$DESKTOP_FILE" <<EOF
[Desktop Entry]
Type=Application
Version=1.0
Name=MSIKLM Lighting Studio
Comment=MSI keyboard lighting GUI
Exec=$BIN_DIR/msiklm-gui
Icon=input-keyboard
Terminal=false
Categories=Settings;Utility;
StartupNotify=true
Keywords=msi;keyboard;rgb;lighting;
EOF

chmod 0644 "$DESKTOP_FILE"
echo "[desktop] wrote $DESKTOP_FILE"

if command -v update-desktop-database >/dev/null 2>&1; then
    update-desktop-database "$APP_DIR" >/dev/null 2>&1 || true
fi

add_to_cinnamon_favorites() {
    if [ "$NO_PANEL_PIN" -eq 1 ]; then
        return 0
    fi
    if ! command -v gsettings >/dev/null 2>&1; then
        echo "[panel] gsettings not available; skipping panel pin."
        return 0
    fi
    if ! gsettings writable org.cinnamon favorite-apps >/dev/null 2>&1; then
        echo "[panel] Cinnamon favorite-apps key not writable; skipping panel pin."
        return 0
    fi

    current=$(gsettings get org.cinnamon favorite-apps 2>/dev/null || echo "[]")
    updated=$(python3 - "$current" <<'PY'
import ast
import sys

raw = sys.argv[1]
try:
    items = ast.literal_eval(raw)
except Exception:
    items = []
if not isinstance(items, list):
    items = []
entry = "msiklm-gui.desktop"
if entry not in items:
    items.append(entry)
print(str(items).replace('"', "'"))
PY
)
    gsettings set org.cinnamon favorite-apps "$updated"
    echo "[panel] added msiklm launcher to Cinnamon favorites."
}

add_to_cinnamon_favorites

if ! command -v msiklm-gui >/dev/null 2>&1; then
    echo
    echo "[note] '$BIN_DIR' is not in PATH for this shell. You can still launch from menu/panel."
fi

echo
echo "Installation complete."
echo "Launcher: $DESKTOP_FILE"
echo "Binary:   $BIN_DIR/msiklm-gui"
