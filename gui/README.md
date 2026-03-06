# MSIKLM GUI

This folder contains a Tkinter GUI for `msiklm`:

```bash
sudo apt install python3-tk alsa-utils pulseaudio-utils
python3 gui/msiklm_gui.py
```

The GUI auto-relaunches itself as root (`sudo` first, then `pkexec`) so you should authenticate once
when opening the app, not on every color change.

For passwordless launch after a one-time setup, use:

```bash
./install_ui.sh --passwordless-sudo
```

This installs a scoped sudoers rule for the current user (`msiklm` and
`msiklm-gui --as-root` only). It requires system install to `/usr/local/bin`
(do not use `--user-only`).
Run the installer as your normal user; it will request elevation only when needed.

The GUI can apply:
- keyboard zones: `left`, `middle`, `right`
- optional non-keyboard zones: `logo`, `front_left`, `front_right`, `mouse`

It includes an optional compatibility mode (off by default) that improves reliability on
keyboards that stop responding after one color write by using the known stable command path for
named colors and retrying writes.

Voice mode is available while the GUI is open. It uses microphone input (`arecord`) to drive a
sequential 3-zone gradient visualizer (`left -> middle -> right`) and ignores manual color
selection while enabled. If input is below threshold, all keyboard zones are forced to `off`.
If the default microphone does not work, use the built-in input source selector and `Refresh`
to detect available ALSA/Pulse sources.

## Zone-To-Key Outline Used In The GUI

MSIKLM itself exposes 3 keyboard zones (`left`, `middle`, `right`) and does not provide a per-key
matrix map. The GUI uses the common SteelSeries 3-zone split that aligns with the historical
MSI/SteelSeries Linux tools:

- `left`: key cluster from `Esc` through roughly `5`, `T`, `G`, `B`
- `middle`: central cluster from roughly `6` through `0`, `Y`..`L`, `N`..`,`
- `right`: remaining keyboard keys on the right side (including numpad/right cluster)

Sources consulted for this split:
- MSIKLM zone order and capabilities: https://github.com/Gibtnix/MSIKLM
- msi-keyboard regions (`left`, `middle`, `right`): https://github.com/stephenlacy/msi-keyboard
- msi-keyboard GUI visual region split: https://github.com/stephenlacy/msi-keyboard-gui

Because laptop keyboard geometries vary by model, this should be treated as a best-fit zone
outline, not a strict per-key hardware guarantee for every MSI device.
