import os, time, psutil, subprocess
from pathlib import Path
from pywinauto import Desktop, Application, mouse
import pyperclip as cb
from pywinauto.keyboard import send_keys

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
WAIT_WINDOW_S = 30
SEARCH_TIMEOUT_S = 900
GRACE_S = 5  # filesystem time cushion

DRAFT_ROOTS = [
    os.path.expandvars(r"%LOCALAPPDATA%\CapCut\User Data\Projects\com.lveditor.draft"),
    os.path.expandvars(r"%USERPROFILE%\Documents\CapCut\Projects\CapCut Drafts"),
]

# Optional extra search dirs (e.g., motion blur cache)
EXTRA_SEARCH_DIRS = [
    os.path.expandvars(r"%LOCALAPPDATA%\CapCut\User Data\Cache\MotionBlurCache"),
]
_extra_env = os.environ.get("CAPCUT_EXTRA_SEARCH_DIRS", "")
if _extra_env:
    for _p in _extra_env.split(";"):
        if _p.strip():
            EXTRA_SEARCH_DIRS.append(os.path.expandvars(_p.strip()))
_mb_env = os.environ.get("CAPCUT_MOTIONBLUR_CACHE", "")
if _mb_env:
    EXTRA_SEARCH_DIRS.append(os.path.expandvars(_mb_env))

# -------- window helpers --------
def focus_capcut_or_fail(timeout=WAIT_WINDOW_S):
    t0 = time.time()
    while time.time() - t0 < timeout:
        wins = []
        for w in Desktop(backend="uia").windows(visible_only=True, top_level_only=True):
            try:
                if psutil.Process(w.element_info.process_id).name().lower() == CAPCUT_EXE:
                    wins.append(w)
            except Exception:
                pass
        if wins:
            # Heuristic: prefer the actual editor window over the Home/Projects hub
            def score(window):
                try:
                    spec = Desktop(backend="uia").window(handle=window.handle)
                    # Cheap checks first
                    r = window.rectangle()
                    s = (r.width() * r.height()) / 1_000_000.0  # size factor
                    txts = []
                    try:
                        txts = [e.window_text().lower() for e in spec.descendants(control_type="Text")[:200]]
                    except Exception:
                        pass
                    # Signals of Home
                    is_home = any(k in txts for k in [
                        "create project", "projects", "ai video maker", "script to video"
                    ])
                    # Signals of Editor
                    is_editor = any(k in txts for k in [
                        "timeline", "export", "media", "canvas", "track", "preview"
                    ])
                    base = 0.0
                    if is_editor:
                        base += 10.0
                    if is_home:
                        base -= 5.0
                    return base + s
                except Exception:
                    return 0

            best = max(wins, key=score)
            Application(backend="uia").connect(process=best.element_info.process_id)
            spec = Desktop(backend="uia").window(handle=best.handle)  # WindowSpecification
            spec.wait("visible ready", timeout=5)
            wrapper = spec.wrapper_object()
            wrapper.set_focus()
            # remember main window handle
            global CAPCUT_MAIN_HANDLE
            CAPCUT_MAIN_HANDLE = best.handle
            return wrapper
        time.sleep(0.2)
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
    _status("Uncompounding any compounded clips…")
    # Select everything
    refocus_capcut_if_possible(); send_keys("^a"); time.sleep(0.2)
    _status("Compounding all clips…")
    # Compound all into one: Alt+G
    refocus_capcut_if_possible(); send_keys("%g"); time.sleep(0.2)
    # Save
    refocus_capcut_if_possible(); send_keys("^s"); time.sleep(0.2)

def trigger_preprocess_shortcut_and_mark():
    _status("PRE-PROCESSING – PLEASE WAIT…")
    refocus_capcut_if_possible(); send_keys("^p")                      # your Pre-process shortcut
    time.sleep(0.3)
    refocus_capcut_if_possible(); send_keys("^s")
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

# -------- UI replace clip (no relaunch) --------
def replace_clip_via_open_dialog(clip_path: Path, shortcut="^l", settle_s: float = 0.05):
    """Assumes the timeline has a clip selected; sends your Replace Clip shortcut (e.g., Ctrl+L),
    then pastes the absolute path into the OS file picker and presses Enter, finally saves.
    """
    focus_capcut_or_fail()
    _status("Selecting clip…")
    # Ensure the only clip is selected, then open Replace dialog
    refocus_capcut_if_possible(); send_keys("^a")
    time.sleep(0.15)
    _status("Opening Replace dialog…")
    refocus_capcut_if_possible(); send_keys(shortcut)
    # Do not wait long; immediately start polling for the dialog

    # Prefer controlling the Open dialog directly to avoid send_keys escaping issues
    abs_path = str(clip_path)
    dlg = find_open_file_dialog(timeout=2.0)
    if dlg is not None:
        try:
            # Find filename edit and set text
            edits = dlg.descendants(control_type="Edit")
            if edits:
                # Focus the dialog and the edit control explicitly
                try:
                    dlg.wrapper_object().set_focus()
                except Exception:
                    pass
                ew = edits[0].wrapper_object()
                try:
                    ew.set_focus()
                except Exception:
                    pass
                _status("Pasting path…")
                ew.set_edit_text(abs_path)
            else:
                # Clipboard paste into focused control
                _status("Pasting path…")
                try:
                    dlg.wrapper_object().set_focus()
                except Exception:
                    pass
                cb.copy(abs_path); send_keys("^v")
            # Click Open/确定
            buttons = [e for e in dlg.descendants(control_type="Button") if (e.window_text() or "").lower() in ("open", "select", "ok", "choose", "insert", "attach")]
            if buttons:
                buttons[0].wrapper_object().click_input()
            else:
                send_keys("{ENTER}")
        except Exception:
            cb.copy(abs_path); send_keys("^v"); time.sleep(0.1); send_keys("{ENTER}")
    else:
        # Fallback: clipboard paste into current focus
        try:
            _status("Pasting path…")
            cb.copy(abs_path); send_keys("^v")
        except Exception:
            # Last resort: type (may mangle special chars)
            send_keys(abs_path)
        time.sleep(0.1)
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

# -------- file search / temp-to-final resolution --------
def newest_nonalpha_mp4_since_anywhere(pre_ts: float, timeout_s: int = SEARCH_TIMEOUT_S, grace_s: int = GRACE_S) -> Path:
    """
    Search the ENTIRE drafts tree for the newest .mp4 whose mtime >= pre_ts - grace,
    and whose filename does NOT end with 'alpha.mp4'. Returns the Path (may be a _temp.mp4).
    """
    normal_roots = list_all_draft_roots()
    extra_roots: list[Path] = []
    for extra in EXTRA_SEARCH_DIRS:
        try:
            p = Path(extra)
            if p.exists():
                extra_roots.append(p)
        except Exception:
            pass
    # by default search both; after Motion Blur confirmation, restrict to extra only
    restrict_to_extra = False
    deadline = time.time() + timeout_s
    threshold = pre_ts - grace_s
    best_path, best_m = None, -1.0

    warned_stuck = False
    def _check_temp_progress_or_final(check_duration: float, threshold_ts: float):
        """Wait check_duration seconds, then decide: (is_progress, final_path_or_None).
        - If a final non-alpha mp4 appears after waiting, return (False, final_path)
        - Else if newest temp file grew in size, return (True, None)
        - Else return (False, None)
        """
        try:
            # pick newest temp
            newest_temp, newest_m = None, -1.0
            for root in roots:
                for p in Path(root).rglob("*_temp.mp4"):
                    try:
                        m = p.stat().st_mtime
                    except OSError:
                        continue
                    if m > newest_m:
                        newest_temp, newest_m = p, m
            s0 = None
            if newest_temp is not None:
                try:
                    s0 = newest_temp.stat().st_size
                except OSError:
                    s0 = None
            # wait
            time.sleep(check_duration)
            # re-scan for a final candidate that appeared meanwhile
            best_path2, best_m2 = None, -1.0
            for root in roots:
                for p in Path(root).rglob("*.mp4"):
                    name = p.name.lower()
                    if name.endswith("alpha.mp4"):
                        continue
                    try:
                        m = p.stat().st_mtime
                    except OSError:
                        continue
                    if m >= threshold_ts and m > best_m2:
                        best_path2, best_m2 = p, m
            if best_path2 is not None:
                return (False, best_path2)
            # otherwise, see if temp grew
            if newest_temp is not None and s0 is not None:
                try:
                    s1 = newest_temp.stat().st_size
                    if s1 > s0:
                        return (True, None)
                except OSError:
                    pass
            return (False, None)
        except Exception:
            return (False, None)
    while time.time() < deadline:
        search_roots = (extra_roots if restrict_to_extra else (normal_roots + extra_roots))
        for root in search_roots:
            for p in root.rglob("*.mp4"):
                name = p.name.lower()
                if name.endswith("alpha.mp4"):   # your explicit rule
                    continue
                try:
                    m = p.stat().st_mtime
                except OSError:
                    continue
                if m >= threshold and m > best_m:
                    best_path, best_m = p, m
        if best_path:
            _status(f"Found processed clip: {best_path}")
            return best_path
        # After 120s with no output, check for temp activity; only warn if no temp growth is observed
        if not warned_stuck and (time.time() - pre_ts) > 120:
            warned_stuck = True
            # If a normal temp file exists, assume in-progress without prompting
            try:
                has_normal_temp = any(True for r in normal_roots for _ in Path(r).rglob("*_temp.mp4"))
            except Exception:
                has_normal_temp = False
            if has_normal_temp:
                _status("Pre-processing in progress (temp file present)…")
                time.sleep(1.0)
                continue
            in_progress, final_now = _check_temp_progress_or_final(10.0, threshold)
            if final_now is not None:
                _status(f"Found processed clip: {final_now}")
                return final_now
            if in_progress:
                _status("Pre-processing in progress (temp file is growing)…")
            else:
                _status("Still waiting (>2m) and no temp file growth.")
                if _confirm(
                    "We noticed no Pre-processing activity. Does your compound clip show 'Waiting to preprocess'?\n\n"
                    "This typically means the workflow is corrupted. The best workaround is to select the clip and apply a 1% Motion Blur.\n\n"
                    "To continue with this solution simply select the clip, then select \"Video\" tab then \"Basic\" tab underneath that, scroll down to Motion blur, enable it, and set it to 1%. It will immediately begin applying the motion blur.\n\n"
                    "Click Yes once you've applied Motion Blur, and the automation will continue."
                ):
                    _status("Motion Blur acknowledged. Monitoring MotionBlurCache for output…")
                    restrict_to_extra = True
        time.sleep(0.4)

    raise TimeoutError("No post-Preprocess .mp4 found under drafts.")

def resolve_temp_to_final(temp_or_final_path: Path, timeout_s: int = 600) -> Path:
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
def main():
    focus_capcut_or_fail()
    print("[*] CapCut focused.")

    compound_and_save()
    print("[*] Compound created & saved.")

    pre_ts = trigger_preprocess_shortcut_and_mark()
    print("[*] Pre-process triggered. Searching drafts for newest non-alpha .mp4...")

    newest = newest_nonalpha_mp4_since_anywhere(pre_ts)
    print(f"[*] Candidate: {newest}")

    final_clip = resolve_temp_to_final(newest)
    print(f"[+] Using final clip: {final_clip}")

    # Replace via UI
    replace_clip_via_open_dialog(final_clip)
    print("[✓] Clip replaced via UI. Export manually in CapCut.")

if __name__ == "__main__":
    main()
