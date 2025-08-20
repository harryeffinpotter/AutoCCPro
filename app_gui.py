import os, sys, json, shutil, re
import threading
import tkinter as tk
from tkinter import messagebox
from pathlib import Path

# Reuse your existing logic
import capcut


def run_bypass_async(button: tk.Button, status_var: tk.StringVar) -> None:
	def _task():
		try:
			status_var.set("Running…")
			capcut.main()
			status_var.set("Done. Export in CapCut.")
			messagebox.showinfo("Bypass Pro", "Finished. Export in CapCut.")
		except Exception as exc:
			status_var.set(f"Error: {exc}")
			messagebox.showerror("Bypass Pro", f"Error: {exc}")
		finally:
			button.config(state=tk.NORMAL)

	# disable button and run in background
	button.config(state=tk.DISABLED)
	threading.Thread(target=_task, daemon=True).start()


def main():
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
			return 0
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
			except Exception:
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
		except Exception:
			pass

	def is_config_installed() -> bool:
		# Only rely on our own state flag so first run always prompts
		st = read_state()
		return bool(st.get("config_installed"))

	def show_install_warning(title_prefix: str = "Install") -> bool:
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
		return messagebox.askyesno(f"{title_prefix}: CapCut Shortcut Config", msg)

	def install_config() -> bool:
		try:
			# 1) Save and close CapCut if running
			try:
				import psutil
				from pywinauto.keyboard import send_keys as _send_keys
				# Try to save via UI if a window is present
				try:
					capcut.focus_capcut_or_fail(3)
					_send_keys("^s")
					tk.Frame().after(200)  # noop delay
				except Exception:
					pass
				# Then terminate existing processes
				for p in psutil.process_iter(attrs=["name", "exe"]):
					try:
						if (p.info.get("name", "") or "").lower() == "capcut.exe":
							p.terminate()
					except Exception:
						pass
				# Wait for shutdown up to ~8s
				import time as _time
				t0 = _time.time()
				while _time.time() - t0 < 8:
					if not any((q.info.get("name", "") or "").lower() == "capcut.exe" for q in psutil.process_iter(attrs=["name"])):
						break
					_time.sleep(0.2)
			except Exception:
				pass

			# 2) Apply shortcut patches
			shortcut_dir.mkdir(parents=True, exist_ok=True)
			changes = patch_shortcuts_in_folder(shortcut_dir)
			# If no files were present, optionally fall back to copying provided combined.json
			if changes == 0:
				if combined_src.exists():
					shutil.copy2(combined_src, combined_dst)
					changes = 1
				else:
					messagebox.showinfo(
						"Install",
						"No existing CapCut shortcut JSONs were found to patch. Open CapCut once to generate them, then rerun Install."
					)
			st = read_state(); st["config_installed"] = True; write_state(st)
			messagebox.showinfo("Install", f"Shortcut config updated. Changes applied: {changes}\nRelaunching CapCut to load new shortcuts…")
			# 3) Relaunch CapCut detached (resolve actual exe path if possible)
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
							os.startfile(path)
							break
						except Exception:
							continue
				else:
					# Last resort: rely on PATH
					try:
						os.startfile("CapCut")
					except Exception:
						subprocess.Popen(["cmd", "/c", "start", "", "CapCut"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
			except Exception:
				pass
			return True
		except Exception as e:
			messagebox.showerror("Install", f"Failed to install config: {e}")
			return False

	def ensure_config_before_run() -> bool:
		if is_config_installed():
			return True
		if show_install_warning("First run"):
			return install_config()
		messagebox.showwarning("Config Required", "Please click 'Install Config' and try again.")
		return False

	root = tk.Tk()
	root.title("CapCut Bypass Pro")
	root.geometry("520x220")
	root.minsize(500, 200)
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

	title = tk.Label(container, text="CapCut Bypass Pro", fg=TEXT, bg=DARK_BG, font=("Segoe UI", 12, "bold"))
	title.pack(anchor="w", pady=(0, 8))

	row = tk.Frame(container, bg=DARK_BG)
	row.pack(fill=tk.X)

	btn_install_border, btn_install = outlined_button(row, "Install Config", lambda: install_config() if show_install_warning("Install") else None)
	btn_install_border.pack(side=tk.LEFT, padx=(0, 12))

	def update_status(msg: str):
		try:
			status_var.set(msg)
		except Exception:
			pass

	btn_bypass_border, btn_bypass = outlined_button(row, "Bypass Pro", lambda: (capcut.set_status_callback(update_status), run_bypass_async(btn_bypass, status_var)) if ensure_config_before_run() else None)
	btn_bypass_border.pack(side=tk.LEFT)

	status = tk.Label(container, textvariable=status_var, anchor="w", fg=TEXT, bg=DARK_BG)
	status.pack(fill=tk.X, pady=(12, 0))

	# subtle border frame
	border = tk.Frame(container, bg=DARK_2, height=1)
	border.pack(fill=tk.X, pady=(8, 0))

	root.mainloop()


if __name__ == "__main__":
	main()


