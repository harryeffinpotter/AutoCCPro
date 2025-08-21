# CapCut Bypass Pro (One‑Click)

CapCut Bypass Pro automates a reliable “one‑click” flow to replace a clip and prep your timeline without requiring CapCut Pro. It controls CapCut’s UI directly and enforces your shortcut bindings so you don’t have to.

## What it does

- Focuses CapCut and saves your project
- Un‑compounds any previously compounded segments, then compounds all into a single clip
- Triggers your Pre‑process action
- Finds the newest non‑alpha mp4 in your CapCut drafts and replaces the clip via the file dialog
- Saves again and gives you a status update (with a chime) – no popups

## Requirements (shortcuts)

The app includes an “Install Config” button that patches CapCut’s shortcut JSONs to ensure the following bindings exist (added if missing):

- replaceFragment: Ctrl+L
- precompileCombination: Ctrl+P
- segmentCombination: Alt+G
- selectAll: Ctrl+A

Install flow: auto‑save project, close CapCut, patch JSONs in `%LOCALAPPDATA%\CapCut\User Data\Config\Shortcut`, then relaunch CapCut. Backups (`.bak`) are created for any modified files.

If you prefer manual setup, make sure you set the same bindings in CapCut.

## Run the app

1) Double‑click the EXE (see Build below) or run from source with Python installed.
2) First run: click “Install Config” (recommended). It will relaunch CapCut to load shortcuts.
3) Click “Bypass Pro”. Watch status at the bottom: it shows steps like Uncompounding…, Compounding…, PRE‑PROCESSING…, Pasting path…, Saving…
4) When it says Done. Export in CapCut., you can export from CapCut.

## Build a standalone EXE

From a Python environment with deps:

```bash
pip install -U pyinstaller psutil pywinauto pyperclip
pyinstaller --clean --noconfirm --onefile --windowed \
  --name CapCutBypass \
  --icon=icon.ico \
  --add-data "icon.ico;." \
  app_gui.py
```

Output: `dist/CapCutBypass.exe`

## Dev notes

- Core logic lives in `capcut.py` (UI automation)
- GUI lives in `app_gui.py` (dark theme, neon blue outline buttons)
- The app prefers UI replacement and avoids template edits
- Status is communicated via a callback (`capcut.set_status_callback`)

## Troubleshooting

- If “Install Config” says no JSONs were patched: open CapCut once to generate shortcut files, then run Install again
- If CapCut doesn’t relaunch: ensure `%LOCALAPPDATA%\CapCut\Apps\CapCut.exe` exists; the app also tries common Program Files paths and PATH fallback
- If the Replace dialog doesn’t accept the path: the app writes directly into the filename field; ensure the dialog is not blocked by another window

## License

MIT


