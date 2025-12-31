import os, time, psutil, subprocess, ctypes
from pathlib import Path
from pywinauto import Desktop, Application, mouse
import pyperclip as cb
from pywinauto.keyboard import send_keys
import win32gui
import win32process

# -------- input blocking --------
import threading
import atexit

# Global failsafe - ensure input is ALWAYS unblocked on exit
def _emergency_unblock():
    try:
        ctypes.windll.user32.BlockInput(False)
    except Exception:
        pass
atexit.register(_emergency_unblock)

def block_input(block: bool = True) -> bool:
    """Block/unblock user input (keyboard+mouse). Requires admin or UAC elevation.
    Returns True if successful, False otherwise."""
    try:
        # Properly define the function signature
        BlockInput = ctypes.windll.user32.BlockInput
        BlockInput.argtypes = [ctypes.c_bool]
        BlockInput.restype = ctypes.c_bool
        result = BlockInput(block)
        return bool(result)
    except Exception:
        return False

class BlockingOverlay:
    """Transparent overlay window shown during input blocking."""
    def __init__(self, text="CapCut Bypass Pro - Blocking Input..."):
        self.root = None
        self.text = text

    def show(self):
        import tkinter as tk
        try:
            self.root = tk.Tk()
            self.root.withdraw()  # Hide while setting up

            # Transparent, always-on-top, no decorations
            self.root.overrideredirect(True)
            self.root.attributes('-topmost', True)
            self.root.attributes('-alpha', 0.95)
            # Transparent color key (magenta = transparent)
            self.root.configure(bg='#010101')
            self.root.attributes('-transparentcolor', '#010101')

            # Get screen dimensions
            screen_w = self.root.winfo_screenwidth()
            screen_h = self.root.winfo_screenheight()

            # Canvas for outlined text
            canvas = tk.Canvas(self.root, bg='#010101', highlightthickness=0)
            canvas.pack(fill=tk.BOTH, expand=True)

            text = self.text
            font = ("Segoe UI", 22, "bold")

            # Calculate size needed
            canvas.update_idletasks()
            test_id = canvas.create_text(0, 0, text=text, font=font)
            bbox = canvas.bbox(test_id)
            canvas.delete(test_id)

            text_w = bbox[2] - bbox[0] + 60
            text_h = bbox[3] - bbox[1] + 30

            # Position at bottom center
            x = (screen_w - text_w) // 2
            y = screen_h - text_h - 120

            self.root.geometry(f"{text_w}x{text_h}+{x}+{y}")
            canvas.config(width=text_w, height=text_h)

            cx, cy = text_w // 2, text_h // 2

            # Draw outline (dark shadow in multiple directions)
            for dx, dy in [(-2,-2), (-2,2), (2,-2), (2,2), (-2,0), (2,0), (0,-2), (0,2)]:
                canvas.create_text(cx+dx, cy+dy, text=text, font=font, fill='#000000')

            # Draw main text (white)
            canvas.create_text(cx, cy, text=text, font=font, fill='#ffffff')

            self.root.deiconify()  # Show
            self.root.update()
        except Exception as e:
            print(f"[!] Overlay failed: {e}")
            self.root = None

    def hide(self):
        try:
            if self.root:
                self.root.destroy()
                self.root = None
        except Exception:
            pass

def show_timed_overlay(text: str, duration: float = 3.0):
    """Show an overlay with text for up to `duration` seconds, then hide it.
    Runs in background thread so it doesn't block."""
    def _show():
        overlay = BlockingOverlay(text)
        overlay.show()
        time.sleep(duration)
        overlay.hide()
    threading.Thread(target=_show, daemon=True).start()

class InputBlocker:
    """Context manager to block input during critical automation steps."""
    TIMEOUT_S = 10  # Failsafe: auto-unblock after this many seconds

    def __init__(self):
        self.blocked = False
        self.overlay = None
        self._released = threading.Event()
        self._watchdog = None

    def _watchdog_thread(self):
        """Auto-unblock if timeout expires without normal release."""
        if not self._released.wait(timeout=self.TIMEOUT_S):
            # Timeout! Force unblock
            print("[⚠️] FAILSAFE: Input blocked too long, forcing unblock!")
            block_input(False)
            if self.overlay:
                try:
                    self.overlay.hide()
                except Exception:
                    pass

    def __enter__(self):
        self._released.clear()

        # Show overlay first
        self.overlay = BlockingOverlay()
        self.overlay.show()

        # Start watchdog timer
        self._watchdog = threading.Thread(target=self._watchdog_thread, daemon=True)
        self._watchdog.start()

        self.blocked = block_input(True)
        if self.blocked:
            print("[🔒] Input BLOCKED")
        else:
            print("[⚠️] Input blocking FAILED - not running as admin?")
        return self

    def __exit__(self, *args):
        # Signal watchdog that we released normally
        self._released.set()

        block_input(False)
        if self.blocked:
            print("[🔓] Input UNBLOCKED")

        # Hide overlay
        if self.overlay:
            self.overlay.hide()

# -------- config --------
CAPCUT_EXE = "capcut.exe"
# Saved main window handle for reliable refocus between actions
CAPCUT_MAIN_HANDLE = None
DASH_STATUS_CB = None
CONFIRM_CB = None

def set_status_callback(callback):
    """Install a status callback callable(str)."""
    global DASH_STATUS_CB
    DASH_STATUS_CB = callback

def _status(msg: str):
    try:
        if DASH_STATUS_CB:
            DASH_STATUS_CB(msg)
    except Exception:
        pass

def set_confirm_callback(callback):
    """Install a confirm callback callable(str)->bool."""
    global CONFIRM_CB
    CONFIRM_CB = callback

def _confirm(msg: str) -> bool:
    try:
        if CONFIRM_CB:
            return bool(CONFIRM_CB(msg))
    except Exception:
        pass
    return False
WAIT_WINDOW_S = 5  # Window detection timeout
SEARCH_TIMEOUT_S = 86400 * 7  # 1 week - effectively infinite, let it cook
GRACE_S = 5  # filesystem time cushion

DRAFT_ROOTS = [
    os.path.expandvars(r"%LOCALAPPDATA%\CapCut\User Data\Projects\com.lveditor.draft"),
    os.path.expandvars(r"%USERPROFILE%\Documents\CapCut\Projects\CapCut Drafts"),
]


# -------- window helpers --------
def focus_capcut_or_fail(timeout=WAIT_WINDOW_S):
    t0 = time.time()

    # Find CapCut process FIRST using psutil (super fast)
    capcut_processes = []
    for p in psutil.process_iter(attrs=["name", "pid", "memory_info"]):
        try:
            if p.info.get("name", "").lower() == CAPCUT_EXE:
                mem = p.info.get("memory_info").rss if p.info.get("memory_info") else 0
                capcut_processes.append((p.info["pid"], mem))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    if not capcut_processes:
        raise RuntimeError("CapCut.exe not running. Open CapCut first.")

    # Sort by memory usage (highest first) and pick the top one
    capcut_processes.sort(key=lambda x: x[1], reverse=True)
    main_pid = capcut_processes[0][0]
    capcut_pids = [main_pid]

    # Use win32gui to enumerate windows (MUCH faster than UIA)
    def enum_windows_callback(hwnd, results):
        if not win32gui.IsWindowVisible(hwnd):
            return
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid in capcut_pids:
                title = win32gui.GetWindowText(hwnd)
                title_lower = title.lower()

                # Must have "capcut" in title, but NOT "bypass"
                if "bypass" in title_lower:
                    return

                if "capcut" in title_lower:
                    results.append(hwnd)
        except Exception:
            pass

    hwnds = []
    win32gui.EnumWindows(enum_windows_callback, hwnds)

    # Remove duplicates
    hwnds = list(set(hwnds))

    # Convert HWNDs to pywinauto window objects
    wins = []
    for hwnd in hwnds:
        try:
            spec = Desktop(backend="uia").window(handle=hwnd)
            wins.append(spec.wrapper_object())
        except Exception:
            pass

    if wins:
        # Pick the biggest window
        def get_size(window):
            try:
                r = window.rectangle()
                return r.width() * r.height()
            except Exception:
                return 0

        best = max(wins, key=get_size)

        try:
            spec = Desktop(backend="uia").window(handle=best.handle)
            spec.wait("visible ready", timeout=2)
            wrapper = spec.wrapper_object()
            wrapper.set_focus()
            time.sleep(0.5)  # Give window time to actually accept focus

            # remember main window handle
            global CAPCUT_MAIN_HANDLE
            CAPCUT_MAIN_HANDLE = best.handle
            return wrapper
        except Exception:
            # If we fail to focus this window, try the next biggest
            if len(wins) > 1:
                for alt_win in sorted(wins, key=get_size, reverse=True)[1:]:
                    try:
                        spec = Desktop(backend="uia").window(handle=alt_win.handle)
                        spec.wait("visible ready", timeout=2)
                        wrapper = spec.wrapper_object()
                        wrapper.set_focus()
                        time.sleep(0.5)
                        CAPCUT_MAIN_HANDLE = alt_win.handle
                        return wrapper
                    except Exception:
                        continue

    raise RuntimeError("CapCut.exe window not found. Open your project first.")

def refocus_capcut_if_possible() -> bool:
    """Best-effort refocus to the saved CapCut window. Safe to call often."""
    global CAPCUT_MAIN_HANDLE
    try:
        if CAPCUT_MAIN_HANDLE:
            spec = Desktop(backend="uia").window(handle=CAPCUT_MAIN_HANDLE)
            spec.wait("visible ready", timeout=2)
            spec.wrapper_object().set_focus()
            return True
    except Exception:
        pass
    return False

def select_project_by_name(project_name: str, timeout=15):
    """Try to open a project tile by its visible name (folder name) on the CapCut Home screen."""
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            win = focus_capcut_or_fail(5)
            spec = Desktop(backend="uia").window(handle=win.handle)
            # Search any descendant with matching title text
            matches = spec.descendants(title=project_name)
            if not matches:
                # Sometimes the clickable element is the parent of the text node; try broader search
                texts = spec.descendants(control_type="Text")
                for t in texts:
                    if getattr(t, "window_text", lambda: "")() == project_name:
                        matches = [t]
                        break
            if matches:
                elem = matches[0]
                rect = elem.rectangle()
                # Use mouse to avoid blocked keyboard
                mouse.double_click(coords=(rect.left + 10, rect.top + 10))
                time.sleep(1.0)
                return True
            # Fallback: approximate the first tile area near bottom of the blue hero banner
            rectw = win.rectangle()
            mouse.double_click(coords=(rectw.left + 350, rectw.top + 300))
            time.sleep(0.8)
            return True
        except Exception:
            time.sleep(0.3)
    return False

# Attempt to open the first project tile under the 'Projects' header
def open_most_recent_project(timeout=15):
    t0 = time.time()
    while time.time() - t0 < timeout:
        try:
            win = focus_capcut_or_fail(5)
            spec = Desktop(backend="uia").window(handle=win.handle)
            # Find the 'Projects' header to anchor the grid area
            headers = [e for e in spec.descendants(control_type="Text") if "project" in getattr(e, "window_text", lambda: "")().lower()]
            anchor_y = None
            if headers:
                # Pick the one most likely the section title (small height)
                hdr = sorted(headers, key=lambda e: e.rectangle().top)[0]
                anchor_y = hdr.rectangle().bottom
            # Gather clickable tiles below the anchor
            candidates = []
            for e in spec.descendants():
                try:
                    rect = e.rectangle()
                    if rect.width() < 120 or rect.height() < 90:
                        continue
                    if anchor_y is not None and rect.top < anchor_y:
                        continue
                    ct = getattr(e, "control_type", None)
                    if ct in ("ListItem", "Pane", "Custom", "Button"):
                        candidates.append(e)
                except Exception:
                    continue
            if candidates:
                # Choose the left-most, top-most large tile
                cand = sorted(candidates, key=lambda e: (e.rectangle().top, e.rectangle().left))[0]
                try:
                    cand.wrapper_object().double_click_input()
                except Exception:
                    Desktop(backend="uia").window(handle=cand.handle).wrapper_object().double_click_input()
                time.sleep(1.0)
                return True
            # Fallback: keyboard heuristic
            send_keys("{HOME}"); time.sleep(0.1); send_keys("{TAB 5}"); time.sleep(0.1); send_keys("{ENTER}")
            time.sleep(0.8)
            return True
        except Exception:
            time.sleep(0.3)
    return False

# -------- app control --------
def restart_capcut_and_focus():
    exe_path = None
    # Capture running instances and exe path
    for p in psutil.process_iter(attrs=["name", "exe"]):
        try:
            if p.info.get("name", "").lower() == CAPCUT_EXE:
                exe_path = p.info.get("exe") or exe_path
                p.terminate()
        except Exception:
            pass
    # Wait for shutdown
    t0 = time.time()
    while time.time() - t0 < 10:
        if not any((q.info.get("name", "").lower() == CAPCUT_EXE) for q in psutil.process_iter(attrs=["name"])):
            break
        time.sleep(0.2)
    # Relaunch
    launch = exe_path or CAPCUT_EXE
    # Use ShellExecute to avoid job inheritance so the app won't close when this script exits
    try:
        os.startfile(launch)
    except Exception:
        # Fallback to cmd start if ShellExecute is unavailable for some reason
        subprocess.Popen(["cmd", "/c", "start", "", launch], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    # Focus when ready
    focus_capcut_or_fail()

# -------- timeline ops --------
def compound_and_save():
    _status("Compounding all clips…")
    with InputBlocker():
        # Select everything
        refocus_capcut_if_possible(); send_keys("^a"); time.sleep(0.5)
        # Compound all into one: Alt+G
        refocus_capcut_if_possible(); send_keys("%g"); time.sleep(1.0)
        # Save
        refocus_capcut_if_possible(); send_keys("^s"); time.sleep(0.5)

def trigger_preprocess_shortcut_and_mark(extra_settle_time: float = 0.0):
    _status("PRE-PROCESSING – PLEASE WAIT…")
    with InputBlocker():
        # Select the compound clip first
        refocus_capcut_if_possible(); time.sleep(0.5 + extra_settle_time)
        send_keys("^a"); time.sleep(0.3 + extra_settle_time)
        # Trigger pre-process
        refocus_capcut_if_possible(); time.sleep(0.5 + extra_settle_time)
        send_keys("^+{F11}")                      # Ctrl+Shift+F11 Pre-process shortcut
        time.sleep(0.3 + extra_settle_time)
        refocus_capcut_if_possible(); time.sleep(0.3 + extra_settle_time)
        send_keys("^s")
    return time.time()                   # timestamp AFTER sending the shortcut

# -------- dialog helpers --------
def find_open_file_dialog(timeout: float = 5.0):
    """Locate a standard file-open dialog in a language-agnostic way.
    Heuristic: top-level visible window that contains at least one Edit and
    a Button whose text suggests confirmation (Open/OK/Select/Browse)."""
    confirm_words = {"open", "ok", "select", "choose", "browse", "attach", "insert"}
    t0 = time.time()
    while time.time() - t0 < timeout:
        for w in Desktop(backend="uia").windows(visible_only=True, top_level_only=True):
            try:
                spec = Desktop(backend="uia").window(handle=w.handle)
                edits = spec.descendants(control_type="Edit")
                if not edits:
                    continue
                buttons = spec.descendants(control_type="Button")
                has_confirm = any((b.window_text() or "").lower() in confirm_words for b in buttons)
                if has_confirm or buttons:
                    return spec
            except Exception:
                continue
        time.sleep(0.15)
    return None

def find_filename_edit_in_dialog(dialog_spec):
    """Find the filename input box in a file dialog.
    Returns the Edit control most likely to be the filename field."""
    try:
        edits = dialog_spec.descendants(control_type="Edit")
        if not edits:
            return None

        # Strategy 1: Look for an Edit with automation_id or name suggesting filename
        filename_hints = ["filename", "file name", "1148", "1001"]  # common IDs
        for edit in edits:
            try:
                auto_id = getattr(edit, "automation_id", lambda: "")()
                name = getattr(edit, "window_text", lambda: "")()
                if any(hint in auto_id.lower() for hint in filename_hints):
                    return edit
                if any(hint in name.lower() for hint in filename_hints):
                    return edit
            except Exception:
                continue

        # Strategy 2: Return the bottom-most single-line Edit (filename box is usually at bottom)
        single_line_edits = []
        for edit in edits:
            try:
                rect = edit.rectangle()
                # Single-line edits are typically wider than tall
                if rect.width() > rect.height() * 2:
                    single_line_edits.append(edit)
            except Exception:
                continue

        if single_line_edits:
            # Return the one with the highest Y coordinate (bottom-most)
            return max(single_line_edits, key=lambda e: e.rectangle().bottom)

        # Strategy 3: Just return the last Edit (often the filename box)
        return edits[-1] if edits else None
    except Exception:
        return None

# -------- UI replace clip (no relaunch) --------
def replace_clip_via_open_dialog(clip_path: Path, shortcut="^l", settle_s: float = 0.05):
    """Assumes the timeline has a clip selected; sends your Replace Clip shortcut (e.g., Ctrl+L),
    then pastes the absolute path into the OS file picker and presses Enter, finally saves.
    """
    focus_capcut_or_fail()
    abs_path = str(clip_path)

    with InputBlocker():
        _status("Selecting clip…")
        # Ensure the only clip is selected, then open Replace dialog
        refocus_capcut_if_possible(); send_keys("^a")
        time.sleep(0.15)
        _status("Opening Replace dialog…")
        refocus_capcut_if_possible(); send_keys(shortcut)
        time.sleep(0.3)

        # Give dialog time to fully open
        time.sleep(0.5)

        # Just use keyboard shortcuts - WAY more reliable
        send_keys("%n")  # Alt+N jumps to "File name:" field
        time.sleep(0.15)
        send_keys("^a")  # Select all
        time.sleep(0.1)
        cb.copy(abs_path)
        send_keys("^v")  # Paste
        time.sleep(0.2)
        send_keys("{ENTER}")  # Open file dialog
        time.sleep(0.8)  # Wait for replace confirmation dialog
        send_keys("{ENTER}")  # Confirm "Replace Clip"
        time.sleep(0.8)  # Wait for replacement to complete
        _status("Saving…")
        send_keys("^s")
    return

    # OLD CODE BELOW (kept as fallback)
    dlg = find_open_file_dialog(timeout=3.0)
    if dlg is not None:
        _status("Pasting path…")
        success = False

        try:
            # Strategy 1: Find the correct filename Edit control using our smart helper
            filename_edit = find_filename_edit_in_dialog(dlg)

            if filename_edit:
                try:
                    # Focus dialog first
                    dlg.wrapper_object().set_focus()
                    time.sleep(0.1)

                    # Get the wrapper and try multiple focus methods
                    ew = filename_edit.wrapper_object()

                    # Method 1: Direct focus
                    try:
                        ew.set_focus()
                        time.sleep(0.05)
                    except Exception:
                        pass

                    # Method 2: Click on the control to ensure it's focused
                    try:
                        ew.click_input()
                        time.sleep(0.05)
                    except Exception:
                        pass

                    # Method 3: Select all existing text first (ensures focus)
                    try:
                        send_keys("^a")
                        time.sleep(0.05)
                    except Exception:
                        pass

                    # Now try to set the text directly
                    try:
                        ew.set_edit_text(abs_path)
                        success = True
                    except Exception:
                        # If set_edit_text fails, use clipboard paste
                        cb.copy(abs_path)
                        send_keys("^v")
                        success = True

                except Exception as e:
                    # Log but continue to fallbacks
                    pass
        except Exception:
            pass

        # Strategy 2: Keyboard navigation fallback
        if not success:
            try:
                # Focus the dialog
                dlg.wrapper_object().set_focus()
                time.sleep(0.1)

                # Use Ctrl+L or Alt+N to jump to filename field (common shortcuts)
                # Try Alt+N first (File name field accelerator key in many dialogs)
                send_keys("%n")
                time.sleep(0.1)

                # Clear any existing text and paste
                send_keys("^a")
                cb.copy(abs_path)
                send_keys("^v")
                success = True
            except Exception:
                pass

        # Strategy 3: Ultimate fallback - just paste wherever focus is
        if not success:
            try:
                dlg.wrapper_object().set_focus()
                cb.copy(abs_path)
                send_keys("^v")
            except Exception:
                pass

        # Click Open button or press Enter
        time.sleep(0.2)
        try:
            buttons = [e for e in dlg.descendants(control_type="Button")
                      if (e.window_text() or "").lower() in ("open", "select", "ok", "choose", "insert", "attach")]
            if buttons:
                buttons[0].wrapper_object().click_input()
            else:
                send_keys("{ENTER}")
        except Exception:
            send_keys("{ENTER}")
    else:
        # Fallback: dialog not found, try blind paste
        _status("Pasting path (fallback mode)…")
        try:
            # Try Alt+N to jump to filename field
            send_keys("%n")
            time.sleep(0.1)
            send_keys("^a")
            cb.copy(abs_path)
            send_keys("^v")
        except Exception:
            cb.copy(abs_path)
            send_keys("^v")
        time.sleep(0.2)
        send_keys("{ENTER}")

    # Allow media import/replace and save
    time.sleep(1.0)
    _status("Saving…")
    send_keys("^s")

# -------- draft helpers --------
def list_all_draft_roots():
    return [Path(r) for r in DRAFT_ROOTS if r and Path(r).exists()]

def pick_active_draft_dir():
    # choose the draft with most recent draft_content.json
    candidates = []
    for root in list_all_draft_roots():
        for d in root.iterdir():
            dc = d / "draft_content.json"
            if dc.exists():
                candidates.append((dc.stat().st_mtime, d))
    if not candidates:
        raise RuntimeError("No CapCut draft found.")
    candidates.sort(reverse=True)
    return candidates[0][1]

def backup_active_draft():
    """Backup all drafts modified within 5 minutes of the most recent. Returns list of (original, backup) tuples."""
    import shutil

    # Save project first
    print("[*] Saving project before backup...")
    with InputBlocker():
        refocus_capcut_if_possible()
        send_keys("^s")
        time.sleep(1.0)

    # Find ALL drafts across all locations with their modification times
    all_drafts = []
    for root in list_all_draft_roots():
        for d in root.iterdir():
            dc = d / "draft_content.json"
            if dc.exists():
                try:
                    mtime = dc.stat().st_mtime
                    all_drafts.append((d, mtime))
                except Exception:
                    pass

    if not all_drafts:
        print("[!] No drafts found")
        return []

    # Sort by modification time (most recent first)
    all_drafts.sort(key=lambda x: x[1], reverse=True)
    most_recent_time = all_drafts[0][1]

    # Backup all drafts within 5 minutes (300 seconds) of the most recent
    backups = []
    for draft_dir, mtime in all_drafts:
        if (most_recent_time - mtime) <= 300:  # Within 5 minutes
            print(f"[*] Backing up recent draft: {draft_dir.name} (modified {int(most_recent_time - mtime)}s ago)")
            backup_dir = Path(str(draft_dir) + "-backup")

            # Remove old backup if exists
            if backup_dir.exists():
                shutil.rmtree(backup_dir)

            # Copy to backup
            shutil.copytree(draft_dir, backup_dir)
            backups.append((draft_dir, backup_dir))

    print(f"[✓] Backed up {len(backups)} draft(s)")
    return backups

def restore_draft_from_backup(backup_list):
    """Restore drafts from backup list. Closes CapCut, replaces each backed up draft."""
    import shutil

    if not backup_list:
        raise RuntimeError("No backups to restore")

    print("[*] Closing CapCut...")
    # Close CapCut
    for p in psutil.process_iter(attrs=["name"]):
        try:
            if p.info.get("name", "").lower() == CAPCUT_EXE:
                p.terminate()
        except Exception:
            pass

    # Wait for shutdown
    t0 = time.time()
    while time.time() - t0 < 10:
        if not any((q.info.get("name", "").lower() == CAPCUT_EXE) for q in psutil.process_iter(attrs=["name"])):
            break
        time.sleep(0.2)

    # Restore each backup
    for original, backup in backup_list:
        if not backup.exists():
            print(f"[!] Warning: Backup not found: {backup}")
            continue

        print(f"[*] Deleting working draft: {original}")
        if original.exists():
            shutil.rmtree(original)

        print(f"[*] Restoring backup: {backup} -> {original}")
        shutil.move(str(backup), str(original))

    print(f"[✓] Restored {len(backup_list)} draft(s) successfully!")
    print("[*] You can now reopen CapCut and continue editing.")

# -------- file search / temp-to-final resolution --------
def snapshot_existing_mp4s() -> dict[str, float]:
    """Capture a baseline of all existing .mp4 files and their mtimes BEFORE triggering pre-process.
    Returns dict mapping file path -> mtime."""
    baseline = {}
    for root in list_all_draft_roots():
        try:
            for p in root.rglob("*.mp4"):
                try:
                    baseline[str(p)] = p.stat().st_mtime
                except OSError:
                    pass
        except Exception:
            pass
    return baseline

def newest_nonalpha_mp4_since_anywhere(pre_ts: float, timeout_s: int = SEARCH_TIMEOUT_S, grace_s: int = GRACE_S, baseline: dict[str, float] | None = None) -> Path:
    """
    Search the ENTIRE drafts tree for a NEW .mp4 file.
    If baseline is provided, a file is considered "new" if:
      - It doesn't exist in baseline (completely new file), OR
      - Its mtime is newer than what was recorded in baseline
    If no baseline, falls back to mtime >= pre_ts - grace.
    Excludes files ending with 'alpha.mp4'. Returns the Path (may be a _temp.mp4).
    """
    deadline = time.time() + timeout_s
    threshold = pre_ts - grace_s
    best_path, best_m = None, -1.0

    while time.time() < deadline:
        for root in list_all_draft_roots():
            for p in root.rglob("*.mp4"):
                name = p.name.lower()
                if name.endswith("alpha.mp4"):   # your explicit rule
                    continue
                try:
                    m = p.stat().st_mtime
                except OSError:
                    continue
                # Determine if this file is "new" (created/modified after we started)
                file_path = str(p)
                if baseline is not None:
                    # File is new if it's not in baseline OR has newer mtime than baseline
                    baseline_mtime = baseline.get(file_path)
                    is_new = (baseline_mtime is None) or (m > baseline_mtime)
                else:
                    # Fallback: use timestamp threshold
                    is_new = (m >= threshold)
                if is_new and m > best_m:
                    best_path, best_m = p, m
        if best_path:
            _status(f"Found processed clip: {best_path}")
            return best_path
        time.sleep(0.4)

    raise TimeoutError("No post-Preprocess .mp4 found under drafts.")

def resolve_temp_to_final(temp_or_final_path: Path, timeout_s: int = 86400) -> Path:
    """
    If the discovered file endswith '_temp.mp4', wait until the same name WITHOUT '_temp'
    appears in the same folder, then return that final path. Otherwise return the given path.
    NO size checks—pure rename/appearance check as requested.
    """
    p = Path(temp_or_final_path)
    name = p.name
    parent = p.parent

    if not name.lower().endswith("_temp.mp4"):
        return p  # already final

    # compute the intended final filename (strip the last occurrence of '_temp' before .mp4)
    final_name = name[:-9]  # remove '_temp.mp4' (len 9)
    final_path = parent / final_name

    print(f"[temp] waiting for final file to appear: {final_path}")
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        if final_path.exists():
            print(f"[final ready] {final_path}")
            return final_path
        time.sleep(0.5)

    # if final never showed up, fall back to the temp path (your call)
    # but here we’ll raise to avoid using a temp artifact
    raise TimeoutError(f"Final file not found for temp: {p}")

## pycapcut fallback removed – UI replace flow is the single source of truth

# -------- main --------
def main(do_backup=False):
    """
    Main bypass flow.

    Args:
        do_backup: If True, backup project before bypass. Returns (original, backup) paths.
                   If False, returns None.
    """
    backup_paths = None

    if do_backup:
        backup_paths = backup_active_draft()

    focus_capcut_or_fail()
    print("[*] CapCut focused.")

    compound_and_save()
    print("[*] Compound created & saved.")

    # CRITICAL: Snapshot existing files BEFORE triggering pre-process
    # This catches fast pre-processing on small projects that finish before we start looking
    _status("Preparing baseline…")
    print("[*] Snapshotting existing files (baseline)...")
    baseline = snapshot_existing_mp4s()
    print(f"[*] Baseline captured: {len(baseline)} existing .mp4 files")

    pre_ts = trigger_preprocess_shortcut_and_mark()
    print("[*] Pre-process triggered. Watching for NEW files...")

    # Show pre-processing overlay for 3 seconds
    show_timed_overlay("CapCut Bypass Pro - Pre-processing...", 3.0)

    # Check if temp file OR final file appears within 10 seconds, retry if not
    # Using baseline comparison to catch fast pre-processing
    max_retries = 2
    for retry in range(max_retries):
        # Poll for new files every 0.5s for up to 10 seconds
        found_new = False
        _status("Watching for output…")
        for check_iteration in range(20):  # 20 * 0.5s = 10 seconds
            time.sleep(0.5)
            # Check for any NEW .mp4 file (temp or final) not in baseline
            for root in list_all_draft_roots():
                try:
                    for mp4_file in root.rglob("*.mp4"):
                        try:
                            file_path = str(mp4_file)
                            mtime = mp4_file.stat().st_mtime
                            baseline_mtime = baseline.get(file_path)
                            # File is new if not in baseline OR has newer mtime
                            if baseline_mtime is None or mtime > baseline_mtime:
                                found_new = True
                                is_temp = mp4_file.name.lower().endswith("_temp.mp4")
                                print(f"[✓] {'Temp' if is_temp else 'Final'} file detected after {(check_iteration + 1) * 0.5}s")
                                break
                        except Exception:
                            pass
                    if found_new:
                        break
                except Exception:
                    pass

            if found_new:
                break

        if found_new:
            break

        if retry < max_retries - 1:
            print(f"[!] No new file after 10s, retrying pre-process (attempt {retry + 2}/{max_retries})...")
            _status(f"Retrying pre-process (attempt {retry + 2})…")
            # Re-snapshot baseline before retry
            baseline = snapshot_existing_mp4s()
            # Retry with extra settle time to let UI fully respond
            pre_ts = trigger_preprocess_shortcut_and_mark(extra_settle_time=0.5)
        else:
            print(f"[!] No new file after {max_retries} attempts, continuing search anyway...")
            _status("No new file detected, waiting for output…")

    newest = newest_nonalpha_mp4_since_anywhere(pre_ts, baseline=baseline)
    print(f"[*] Candidate: {newest}")

    final_clip = resolve_temp_to_final(newest)
    print(f"[+] Using final clip: {final_clip}")

    # Replace via UI
    replace_clip_via_open_dialog(final_clip)
    print("[✓] Clip replaced. Opening export dialog...")

    # Open export dialog
    with InputBlocker():
        time.sleep(0.5)
        refocus_capcut_if_possible()
        send_keys("^e")
    print("[✓] Export dialog opened. Configure settings and export.")

    # Show done overlay
    show_timed_overlay("CapCut Bypass Pro - Done! Export when ready.", 4.0)

    return backup_paths

if __name__ == "__main__":
    main()
