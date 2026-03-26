#!/bin/sh

# Install MSIKLM CLI + GUI launcher with desktop integration.
# Targets common modern distros (apt, dnf, yum, pacman, zypper).

set -eu

SKIP_DEPS=0
NO_PANEL_PIN=0
FORCE_USER_INSTALL=0
ENABLE_PASSWORDLESS_SUDO=0

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
        --passwordless-sudo)
            ENABLE_PASSWORDLESS_SUDO=1
            ;;
        -h|--help)
            cat <<'EOF'
Usage: ./install_ui.sh [options]

Options:
  --skip-deps     Do not install distro packages.
  --no-panel-pin  Do not add the launcher to Cinnamon panel favorites.
  --user-only     Install binaries into ~/.local/bin only.
  --passwordless-sudo  Allow current user to run msiklm-gui/msiklm via sudo without password.
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
    elif command -v doas >/dev/null 2>&1; then
        doas "$@"
    elif command -v pkexec >/dev/null 2>&1; then
        pkexec "$@"
    else
        echo "No privilege escalation tool found (sudo/doas/pkexec)." >&2
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
        run_as_root apt install -y gcc make libhidapi-dev python3-tk desktop-file-utils alsa-utils pulseaudio-utils
        return $?
    fi
    if command -v dnf >/dev/null 2>&1; then
        run_as_root dnf install -y gcc make hidapi-devel python3-tkinter desktop-file-utils alsa-utils pulseaudio-utils
        return $?
    fi
    if command -v yum >/dev/null 2>&1; then
        run_as_root yum install -y gcc make hidapi-devel python3-tkinter desktop-file-utils alsa-utils pulseaudio-utils
        return $?
    fi
    if command -v pacman >/dev/null 2>&1; then
        run_as_root pacman -S --noconfirm gcc make hidapi tk desktop-file-utils alsa-utils libpulse
        return $?
    fi
    if command -v zypper >/dev/null 2>&1; then
        run_as_root zypper --non-interactive install gcc make libhidapi-devel python3-tk desktop-file-utils alsa-utils pulseaudio-utils
        return $?
    fi

    echo "[deps] no supported package manager found. Install deps manually." >&2
    return 1
}

echo "[build] compiling msiklm..."
make

try_install_packages || true

if [ "$ENABLE_PASSWORDLESS_SUDO" -eq 1 ] && [ "$FORCE_USER_INSTALL" -eq 1 ]; then
    echo "--passwordless-sudo cannot be combined with --user-only." >&2
    exit 1
fi

if [ "$ENABLE_PASSWORDLESS_SUDO" -eq 1 ]; then
    BIN_DIR="/usr/local/bin"
elif [ "$FORCE_USER_INSTALL" -eq 1 ]; then
    BIN_DIR="$HOME/.local/bin"
elif [ -w /usr/local/bin ]; then
    BIN_DIR="/usr/local/bin"
else
    BIN_DIR="$HOME/.local/bin"
fi

if [ ! -d "$BIN_DIR" ]; then
    if [ -w "$(dirname "$BIN_DIR")" ]; then
        mkdir -p "$BIN_DIR"
    else
        run_as_root mkdir -p "$BIN_DIR"
    fi
fi

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

configure_passwordless_sudo() {
    if [ "$ENABLE_PASSWORDLESS_SUDO" -ne 1 ]; then
        return 0
    fi

    if [ "$BIN_DIR" != "/usr/local/bin" ]; then
        echo "[sudoers] passwordless mode requires installation to /usr/local/bin (do not use --user-only)." >&2
        return 1
    fi

    if ! command -v sudo >/dev/null 2>&1; then
        echo "[sudoers] sudo is not available; skipping passwordless setup." >&2
        return 0
    fi

    target_user="${SUDO_USER:-$(id -un)}"
    if [ -z "$target_user" ] || [ "$target_user" = "root" ]; then
        echo "[sudoers] could not determine a non-root target user; skipping." >&2
        return 0
    fi

    gui_bin="$BIN_DIR/msiklm-gui"
    cli_bin="$BIN_DIR/msiklm"
    case "$gui_bin$cli_bin" in
        *[[:space:]]*)
            echo "[sudoers] binary path contains whitespace; skipping automatic sudoers rule." >&2
            return 1
            ;;
    esac

    rule_file="/etc/sudoers.d/msiklm-gui-$target_user"
    tmp_file=$(mktemp)
    cat >"$tmp_file" <<EOF
# Managed by install_ui.sh for MSIKLM GUI passwordless relaunch.
$target_user ALL=(root) NOPASSWD: $cli_bin
$target_user ALL=(root) NOPASSWD: $gui_bin --as-root
EOF

    if command -v visudo >/dev/null 2>&1; then
        if ! run_as_root visudo -cf "$tmp_file" >/dev/null 2>&1; then
            echo "[sudoers] generated rule failed validation; skipping." >&2
            rm -f "$tmp_file"
            return 1
        fi
    fi

    run_as_root install -m 0440 "$tmp_file" "$rule_file"
    rm -f "$tmp_file"
    echo "[sudoers] installed $rule_file"
}

configure_passwordless_sudo

APP_DIR="$HOME/.local/share/applications"
mkdir -p "$APP_DIR"
DESKTOP_FILE="$APP_DIR/msiklm-gui.desktop"

cat >"$DESKTOP_FILE" <<EOF
[Desktop Entry]
Type=Application
Version=1.0
Name=MSIKLM Control Center
Comment=MSI keyboard lighting control
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
