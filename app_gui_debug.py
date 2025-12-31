import os, sys, json, shutil, re, ctypes
import winsound
import threading
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

# -------- Auto-elevation --------
def is_admin():
    """Check if running with admin privileges."""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False

def run_as_admin():
    """Re-launch the current script with admin privileges."""
    try:
        if getattr(sys, 'frozen', False):
            script = sys.executable
            params = ' '.join(sys.argv[1:])
        else:
            script = sys.executable
            params = f'"{sys.argv[0]}"'
            if len(sys.argv) > 1:
                params += ' ' + ' '.join(sys.argv[1:])
        ret = ctypes.windll.shell32.ShellExecuteW(None, "runas", script, params, None, 1)
        return ret > 32
    except Exception:
        return False

def ensure_admin():
    """Ensure we're running as admin, re-launch if needed."""
    if not is_admin():
        if run_as_admin():
            sys.exit(0)

# Reuse your existing logic
import capcut

# Global debug log
DEBUG_LOG = []

def debug(msg):
	print(f"[DEBUG] {msg}")
	DEBUG_LOG.append(msg)
	if len(DEBUG_LOG) > 100:
		DEBUG_LOG.pop(0)

def run_bypass_async(button: tk.Button, status_var: tk.StringVar, status_label: tk.Label) -> None:
	debug("run_bypass_async called")
	def _task():
		try:
			debug("bypass thread starting")
			# reset status style to normal while running
			try:
				status_label.config(fg="#e6edf3", font=("Segoe UI", 10))
			except Exception:
				pass
			status_var.set("Running…")
			capcut.main()
			status_var.set("Done. Export in CapCut.")
			try:
				status_label.config(fg="#22ff66", font=("Segoe UI", 10, "bold"))
			except Exception:
				pass
			try:
				winsound.MessageBeep(winsound.MB_ICONASTERISK)
			except Exception:
				pass
			debug("bypass thread completed successfully")
		except Exception as exc:
			debug(f"bypass thread exception: {exc}")
			import traceback
			traceback.print_exc()
			status_var.set(f"Error: {exc}")
		finally:
			button.config(state=tk.NORMAL)

	# disable button and run in background
	button.config(state=tk.DISABLED)
	debug("Starting bypass thread")
	threading.Thread(target=_task, daemon=True).start()


def main():
	debug("main() called")
	# Colors
	DARK_BG = "#0f1116"
	DARK_2 = "#161b22"
	TEXT = "#e6edf3"
	ACCENT = "#00b7ff"  # sky blue neon
	ACCENT_HOVER = "#33c7ff"

	# Paths and state
	def get_app_dir() -> Path:
		if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
			return Path(sys._MEIPASS)  # type: ignore[attr-defined]
		return Path(__file__).resolve().parent

	app_dir = get_app_dir()
	local_state_dir = Path(os.path.expandvars(r"%LOCALAPPDATA%\CapCutBypass"))
	local_state_dir.mkdir(parents=True, exist_ok=True)
	state_file = local_state_dir / "state.json"
	shortcut_dir = Path(os.path.expandvars(r"%LOCALAPPDATA%\CapCut\User Data\Config\Shortcut"))
	combined_src = app_dir / "combined.json"
	combined_dst = shortcut_dir / "combined.json"

	debug(f"app_dir: {app_dir}")
	debug(f"state_file: {state_file}")
	debug(f"shortcut_dir: {shortcut_dir}")

	# Required shortcuts to enforce across all files
	required_shortcuts = {
		"replaceFragment": ["Ctrl+L"],
		"selectAll": ["Ctrl+A"],
		"precompileCombination": ["Ctrl+P"],
		"segmentCombination": ["Alt+G"],
	}

	def _patch_json_node(node, mapping: dict) -> int:
		replaced = 0
		if isinstance(node, dict):
			# If this dict looks like a CapCut preset entry, enforce keys under its 'sequence'
			seq = node.get("sequence")
			if isinstance(seq, dict):
				for key, desired in mapping.items():
					if seq.get(key) != desired:
						seq[key] = desired
						replaced += 1
			# Also handle any direct key matches at arbitrary nesting
			for k, v in list(node.items()):
				if k in mapping and v != mapping[k]:
					node[k] = mapping[k]
					replaced += 1
				replaced += _patch_json_node(node[k], mapping)
		elif isinstance(node, list):
			for i in range(len(node)):
				replaced += _patch_json_node(node[i], mapping)
		return replaced

	def patch_shortcuts_in_folder(folder: Path) -> int:
		total = 0
		if not folder.exists():
			debug(f"shortcut folder doesn't exist: {folder}")
			return 0
		debug(f"patching shortcuts in: {folder}")
		for fp in folder.glob("*.json"):
			raw = None
			try:
				raw = fp.read_text(encoding="utf-8")
				data = json.loads(raw)
				changed = _patch_json_node(data, required_shortcuts)
				did_write = False
				if changed > 0:
					try:
						shutil.copy2(fp, fp.with_suffix(fp.suffix + ".bak"))
					except Exception:
						pass
					fp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
					did_write = True
					total += changed
					debug(f"patched {fp.name}: {changed} changes")
				# If JSON patch found nothing, try regex on raw text to be robust to non-strict JSON
				if not did_write:
					patterns = {
						"replaceFragment": '\\[.*?\\]',
						"selectAll": '\\[.*?\\]',
						"precompileCombination": '\\[.*?\\]',
						"segmentCombination": '\\[.*?\\]',
					}
					replacements = {
						"replaceFragment": '["Ctrl+L"]',
						"selectAll": '["Ctrl+A"]',
						"precompileCombination": '["Ctrl+P"]',
						"segmentCombination": '["Alt+G"]',
					}
					new_text = raw
					file_changes = 0
					for key, arr_pat in patterns.items():
						pat = re.compile(r'("' + re.escape(key) + r'"\s*:\s*)' + arr_pat, flags=re.DOTALL)
						new_text, n = pat.subn(r'\1' + replacements[key], new_text)
						file_changes += n
					if file_changes > 0:
						try:
							shutil.copy2(fp, fp.with_suffix(fp.suffix + ".bak"))
						except Exception:
							pass
						fp.write_text(new_text, encoding="utf-8")
						total += file_changes
						debug(f"regex-patched {fp.name}: {file_changes} changes")
			except Exception as e:
				debug(f"failed to patch {fp}: {e}")
				# If even reading failed, skip file
				continue
		return total

	def read_state() -> dict:
		try:
			return json.loads(state_file.read_text(encoding="utf-8"))
		except Exception:
			return {}

	def write_state(data: dict) -> None:
		try:
			state_file.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
			debug(f"wrote state: {data}")
		except Exception as e:
			debug(f"failed to write state: {e}")

	def is_config_installed() -> bool:
		# Only rely on our own state flag so first run always prompts
		st = read_state()
		result = bool(st.get("config_installed"))
		debug(f"is_config_installed: {result} (state={st})")
		return result

	def show_install_warning(title_prefix: str = "Install") -> bool:
		debug(f"show_install_warning called with title_prefix='{title_prefix}'")
		msg = (
			"Warning: Clicking Yes will automatically edit your CapCut shortcuts to make the bypass work.\n\n"
			"If you have custom keybinds you want to keep, set the following manually instead:\n\n"
			"Required custom hotkeys:\n"
			"  Ctrl+L = Replace Clip\n"
			"  Ctrl+P = Preprocess Clip\n\n"
			"Required default hotkeys:\n"
			"  Alt+G  = Create Compound Clip\n"
			"  Ctrl+A = Select All\n\n"
			f"The installer will patch all *.json in: {shortcut_dir}\n\n"
			"Note: This will auto-save your project, close CapCut, apply changes, then relaunch CapCut."
		)
		try:
			result = messagebox.askyesno(f"{title_prefix}: CapCut Shortcut Config", msg)
			debug(f"show_install_warning returned: {result}")
			return result
		except Exception as e:
			debug(f"show_install_warning exception: {e}")
			import traceback
			traceback.print_exc()
			return False

	def install_config() -> bool:
		debug("install_config called")
		try:
			# 1) Save and close CapCut if running
			try:
				import psutil
				from pywinauto.keyboard import send_keys as _send_keys
				# Try to save via UI if a window is present
				try:
					debug("trying to focus CapCut")
					capcut.focus_capcut_or_fail(0.5)
					debug("sending Ctrl+S")
					_send_keys("^s")
					import time as _time
					_time.sleep(0.2)
				except Exception as e:
					debug(f"failed to save CapCut: {e}")
				# Then terminate existing processes
				debug("terminating CapCut processes")
				for p in psutil.process_iter(attrs=["name", "exe"]):
					try:
						if (p.info.get("name", "") or "").lower() == "capcut.exe":
							p.terminate()
							debug(f"terminated PID {p.pid}")
					except Exception:
						pass
				# Wait for shutdown up to ~8s
				import time as _time
				t0 = _time.time()
				while _time.time() - t0 < 8:
					if not any((q.info.get("name", "") or "").lower() == "capcut.exe" for q in psutil.process_iter(attrs=["name"])):
						break
					_time.sleep(0.2)
				debug("CapCut closed")
			except Exception as e:
				debug(f"exception during CapCut close: {e}")

			# 2) Apply shortcut patches
			shortcut_dir.mkdir(parents=True, exist_ok=True)
			debug("patching shortcuts...")
			changes = patch_shortcuts_in_folder(shortcut_dir)
			debug(f"total changes: {changes}")
			# If no files were present, optionally fall back to copying provided combined.json
			if changes == 0:
				if combined_src.exists():
					debug(f"copying {combined_src} to {combined_dst}")
					shutil.copy2(combined_src, combined_dst)
					changes = 1
				else:
					debug("no shortcuts found, showing info")
					messagebox.showinfo(
						"Install",
						"No existing CapCut shortcut JSONs were found to patch. Open CapCut once to generate them, then rerun Install."
					)
					return False
			st = read_state(); st["config_installed"] = True; write_state(st)
			messagebox.showinfo("Install", f"Shortcut config updated. Changes applied: {changes}\nRelaunching CapCut to load new shortcuts…")
			# 3) Relaunch CapCut detached (resolve actual exe path if possible)
			debug("relaunching CapCut")
			try:
				import subprocess, psutil
				# Try to find a previous exe path from processes (if any still listed)
				launch_path = None
				for p in psutil.process_iter(attrs=["name", "exe"]):
					try:
						if (p.info.get("name", "") or "").lower() == "capcut.exe" and p.info.get("exe"):
							launch_path = p.info.get("exe"); break
					except Exception:
						pass
				candidates = [
					os.path.expandvars(r"%LOCALAPPDATA%\CapCut\Apps\CapCut.exe"),
				]
				if launch_path and os.path.exists(launch_path):
					candidates.append(launch_path)
				candidates.extend([
					os.path.expandvars(r"%LOCALAPPDATA%\Programs\CapCut\CapCut.exe"),
					os.path.expandvars(r"%PROGRAMFILES%\CapCut\CapCut.exe"),
					os.path.expandvars(r"%PROGRAMFILES(X86)%\CapCut\CapCut.exe"),
					os.path.expandvars(r"%LOCALAPPDATA%\CapCut\CapCut.exe"),
				])
				for path in candidates:
					if path and os.path.exists(path):
						try:
							debug(f"launching: {path}")
							os.startfile(path)
							break
						except Exception as e:
							debug(f"failed to launch {path}: {e}")
							continue
				else:
					# Last resort: rely on PATH
					try:
						debug("launching CapCut from PATH")
						os.startfile("CapCut")
					except Exception as e:
						debug(f"startfile failed: {e}")
						subprocess.Popen(["cmd", "/c", "start", "", "CapCut"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
			except Exception as e:
				debug(f"exception during relaunch: {e}")
			debug("install_config completed successfully")
			return True
		except Exception as e:
			debug(f"install_config exception: {e}")
			import traceback
			traceback.print_exc()
			messagebox.showerror("Install", f"Failed to install config: {e}")
			return False

	def run_install_config_async(button: tk.Button):
		"""Run install_config in a background thread to avoid freezing the GUI."""
		debug("run_install_config_async called")
		def _task():
			try:
				debug("install_config thread starting")
				install_config()
				debug("install_config thread completed")
			except Exception as e:
				debug(f"install_config thread exception: {e}")
				import traceback
				traceback.print_exc()
			finally:
				button.config(state=tk.NORMAL)
				debug("button re-enabled")
		button.config(state=tk.DISABLED)
		debug("button disabled, starting thread")
		threading.Thread(target=_task, daemon=True).start()

	def ensure_config_before_run() -> bool:
		debug("ensure_config_before_run called")
		if is_config_installed():
			debug("config is installed, proceeding")
			return True
		debug("config not installed, showing warning")
		if show_install_warning("First run"):
			debug("user clicked Yes, installing config synchronously")
			return install_config()
		debug("user clicked No")
		messagebox.showwarning("Config Required", "Please click 'Install Config' and try again.")
		return False

	root = tk.Tk()
	root.title("CapCut Bypass Pro [DEBUG]")
	root.geometry("520x280")
	root.minsize(500, 260)
	# Standard window chrome gives minimize and close
	root.configure(bg=DARK_BG)

	# Set window/taskbar icon from bundled icon.ico (works for frozen and non-frozen)
	try:
		icon_path = app_dir / "icon.ico"
		if icon_path.exists():
			root.iconbitmap(default=str(icon_path))
	except Exception:
		pass

	# Base font
	root.option_add("*Font", ("Segoe UI", 10))

	container = tk.Frame(root, padx=16, pady=16, bg=DARK_BG)
	container.pack(fill=tk.BOTH, expand=True)

	status_var = tk.StringVar(value="Ready.")
	debug_var = tk.StringVar(value="")

	def outlined_button(parent: tk.Widget, text: str, command):
		border = tk.Frame(parent, bg=ACCENT, bd=0)
		inner = tk.Button(
			border,
			text=text,
			font=("Segoe UI", 11, "bold"),
			bg=DARK_BG,
			fg=ACCENT,
			activebackground=DARK_2,
			activeforeground=ACCENT_HOVER,
			relief=tk.FLAT,
			bd=0,
			cursor="hand2",
			command=command,
		)
		inner.pack(padx=2, pady=2)
		def _hover(_e, on: bool):
			border.configure(bg=ACCENT_HOVER if on else ACCENT)
			inner.configure(fg=(ACCENT_HOVER if on else ACCENT))
		inner.bind("<Enter>", lambda e: _hover(e, True))
		inner.bind("<Leave>", lambda e: _hover(e, False))
		return border, inner

	title = tk.Label(container, text="CapCut Bypass Pro [DEBUG]", fg=TEXT, bg=DARK_BG, font=("Segoe UI", 12, "bold"))
	title.pack(anchor="w", pady=(0, 8))

	row = tk.Frame(container, bg=DARK_BG)
	row.pack(fill=tk.X)

	def on_install_click():
		debug("Install Config button clicked")
		try:
			if show_install_warning("Install"):
				debug("User clicked Yes, running install_config_async")
				run_install_config_async(btn_install)
			else:
				debug("User clicked No or dialog was cancelled")
		except Exception as e:
			debug(f"on_install_click exception: {e}")
			import traceback
			traceback.print_exc()

	btn_install_border, btn_install = outlined_button(row, "Install Config", on_install_click)
	btn_install_border.pack(side=tk.LEFT, padx=(0, 12))

	def update_status(msg: str):
		try:
			status_var.set(msg)
		except Exception:
			pass

	def ask_confirm(msg: str) -> bool:
		try:
			return bool(messagebox.askyesno("CapCut Bypass Pro", msg))
		except Exception:
			return False

	def on_bypass_click():
		debug("Bypass Pro button clicked")
		try:
			if ensure_config_before_run():
				debug("Config check passed, setting callbacks and running")
				capcut.set_status_callback(update_status)
				capcut.set_confirm_callback(ask_confirm)
				run_bypass_async(btn_bypass, status_var, status)
			else:
				debug("Config check failed, not running")
		except Exception as e:
			debug(f"on_bypass_click exception: {e}")
			import traceback
			traceback.print_exc()

	btn_bypass_border, btn_bypass = outlined_button(row, "Bypass Pro", on_bypass_click)
	btn_bypass_border.pack(side=tk.LEFT)

	status = tk.Label(container, textvariable=status_var, anchor="w", fg=TEXT, bg=DARK_BG)
	status.pack(fill=tk.X, pady=(12, 0))

	# Debug output
	debug_label = tk.Label(container, textvariable=debug_var, anchor="w", fg="#888", bg=DARK_BG, font=("Consolas", 8), justify=tk.LEFT)
	debug_label.pack(fill=tk.BOTH, expand=True, pady=(8, 0))

	def update_debug_display():
		if DEBUG_LOG:
			debug_var.set("\n".join(DEBUG_LOG[-5:]))
		root.after(100, update_debug_display)

	update_debug_display()

	# subtle border frame
	border = tk.Frame(container, bg=DARK_2, height=1)
	border.pack(fill=tk.X, pady=(8, 0))

	debug("GUI initialized")
	root.mainloop()


if __name__ == "__main__":
	ensure_admin()
	debug(f"Script started (Admin: {is_admin()})")
	main()
