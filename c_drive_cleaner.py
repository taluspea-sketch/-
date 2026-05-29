"""
Windows C Drive Cleanup Tool
A safe, professional tool to reclaim disk space by removing known temp/cache files.
"""

import os
import sys
import shutil
import threading
import subprocess
import winreg
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, font as tkfont


# ---------------------------------------------------------------------------
# Constants / palette
# ---------------------------------------------------------------------------
BG_DARK   = "#1e1e2e"
BG_PANEL  = "#2a2a3e"
BG_CARD   = "#313147"
ACCENT    = "#7c6af7"
ACCENT2   = "#5a9cf8"
FG_WHITE  = "#e8e8f0"
FG_GRAY   = "#9090a8"
FG_GREEN  = "#50fa7b"
FG_RED    = "#ff5555"
FG_YELLOW = "#f1fa8c"
BTN_SCAN  = "#5a9cf8"
BTN_CLEAN = "#50fa7b"
BTN_FG    = "#1e1e2e"

HOVER_SCAN  = "#7ab3ff"
HOVER_CLEAN = "#70ffaa"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def human_size(num_bytes: int) -> str:
    """Return a human-readable size string."""
    for unit in ("B", "KB", "MB", "GB"):
        if abs(num_bytes) < 1024:
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024
    return f"{num_bytes:.1f} TB"


def safe_remove(path: Path) -> int:
    """
    Delete a file or directory tree.  Returns bytes freed, 0 on any error.
    Never raises.
    """
    freed = 0
    try:
        if path.is_file() or path.is_symlink():
            freed = path.stat().st_size
            path.unlink(missing_ok=True)
        elif path.is_dir():
            freed = sum(
                f.stat().st_size
                for f in path.rglob("*")
                if f.is_file()
            )
            shutil.rmtree(path, ignore_errors=True)
    except (PermissionError, OSError):
        pass
    return freed


def dir_size(path: Path) -> int:
    """Return total size of all files under *path* in bytes. Never raises."""
    total = 0
    try:
        for entry in path.rglob("*"):
            try:
                if entry.is_file():
                    total += entry.stat().st_size
            except (PermissionError, OSError):
                pass
    except (PermissionError, OSError):
        pass
    return total


def recycle_bin_size() -> int:
    """Estimate Recycle Bin size by reading $Recycle.Bin on all drives."""
    total = 0
    for drive_letter in "CDEFGHIJKLMNOPQRSTUVWXYZ":
        rb = Path(f"{drive_letter}:\\$Recycle.Bin")
        if rb.exists():
            total += dir_size(rb)
    return total


def empty_recycle_bin() -> int:
    """Use SHEmptyRecycleBin via shell32 to empty the Recycle Bin."""
    freed = recycle_bin_size()
    try:
        import ctypes
        ctypes.windll.shell32.SHEmptyRecycleBinW(None, None, 0x0007)
    except Exception:
        pass
    return freed


def get_browser_cache_paths() -> list[Path]:
    """Return known browser cache directories that exist on this machine."""
    local_app = Path(os.environ.get("LOCALAPPDATA", "C:\\Users\\Default\\AppData\\Local"))
    app_data  = Path(os.environ.get("APPDATA",      "C:\\Users\\Default\\AppData\\Roaming"))

    candidates = [
        # Chrome
        local_app / "Google" / "Chrome" / "User Data" / "Default" / "Cache",
        local_app / "Google" / "Chrome" / "User Data" / "Default" / "Code Cache",
        local_app / "Google" / "Chrome" / "User Data" / "Default" / "GPUCache",
        # Edge (Chromium)
        local_app / "Microsoft" / "Edge" / "User Data" / "Default" / "Cache",
        local_app / "Microsoft" / "Edge" / "User Data" / "Default" / "Code Cache",
        local_app / "Microsoft" / "Edge" / "User Data" / "Default" / "GPUCache",
        # Firefox
        local_app / "Mozilla" / "Firefox" / "Profiles",
        app_data  / "Mozilla" / "Firefox" / "Profiles",
    ]

    paths = []
    for p in candidates:
        if "Profiles" in p.parts and p.exists():
            # Firefox keeps a separate cache folder per profile
            for profile in p.iterdir():
                cache = profile / "cache2"
                if cache.exists():
                    paths.append(cache)
        elif p.exists():
            paths.append(p)
    return paths


def get_thumbnail_cache_paths() -> list[Path]:
    local_app = Path(os.environ.get("LOCALAPPDATA", "C:\\Users\\Default\\AppData\\Local"))
    return [p for p in [
        local_app / "Microsoft" / "Windows" / "Explorer",
    ] if p.exists()]


# ---------------------------------------------------------------------------
# Category definitions
# ---------------------------------------------------------------------------

class CleanCategory:
    """
    Encapsulates one cleanup category.
    scan()  -> fills self.size_bytes
    clean() -> deletes files, returns bytes freed
    """

    def __init__(self, key: str, label: str, description: str, paths_fn=None, custom_fn=None):
        self.key         = key
        self.label       = label
        self.description = description
        self._paths_fn   = paths_fn   # callable -> list[Path]
        self._custom_fn  = custom_fn  # callable -> int  (handles its own deletion)
        self.size_bytes  = 0
        self.error       = None

    def _get_paths(self) -> list[Path]:
        if self._paths_fn:
            try:
                return self._paths_fn()
            except Exception as exc:
                self.error = str(exc)
        return []

    def scan(self) -> int:
        self.error = None
        if self._custom_fn and self.key == "recycle":
            self.size_bytes = recycle_bin_size()
        else:
            total = 0
            for p in self._get_paths():
                if p.is_file():
                    try:
                        total += p.stat().st_size
                    except OSError:
                        pass
                else:
                    total += dir_size(p)
            self.size_bytes = total
        return self.size_bytes

    def clean(self, progress_cb=None) -> int:
        if self.key == "recycle":
            return empty_recycle_bin()

        freed = 0
        paths = self._get_paths()
        for i, p in enumerate(paths):
            freed += safe_remove(p)
            if progress_cb:
                progress_cb(i + 1, len(paths))
        return freed


def build_categories() -> list[CleanCategory]:
    temp_env  = Path(os.environ.get("TEMP", "C:\\Windows\\Temp"))
    win_temp  = Path("C:\\Windows\\Temp")
    wu_cache  = Path("C:\\Windows\\SoftwareDistribution\\Download")
    prefetch  = Path("C:\\Windows\\Prefetch")

    def temp_paths():
        paths = []
        for base in {temp_env, win_temp}:
            if base.exists():
                paths.extend(base.iterdir())
        return paths

    def wu_paths():
        return list(wu_cache.iterdir()) if wu_cache.exists() else []

    def prefetch_paths():
        return [p for p in prefetch.iterdir() if p.suffix.upper() == ".PF"] if prefetch.exists() else []

    return [
        CleanCategory(
            key="temp",
            label="Temporary Files",
            description="%TEMP% and C:\\Windows\\Temp",
            paths_fn=temp_paths,
        ),
        CleanCategory(
            key="recycle",
            label="Recycle Bin",
            description="Files waiting in the Recycle Bin",
            custom_fn=True,
        ),
        CleanCategory(
            key="wupdate",
            label="Windows Update Cache",
            description="C:\\Windows\\SoftwareDistribution\\Download",
            paths_fn=wu_paths,
        ),
        CleanCategory(
            key="prefetch",
            label="Prefetch Files",
            description="C:\\Windows\\Prefetch  (*.pf)",
            paths_fn=prefetch_paths,
        ),
        CleanCategory(
            key="browser",
            label="Browser Caches",
            description="Chrome, Edge, Firefox cache folders",
            paths_fn=get_browser_cache_paths,
        ),
        CleanCategory(
            key="thumbs",
            label="Thumbnail Cache",
            description="Windows Explorer thumbnail database",
            paths_fn=get_thumbnail_cache_paths,
        ),
    ]


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------

class CleanerApp(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("C Drive Cleanup Tool")
        self.configure(bg=BG_DARK)
        self.resizable(False, False)

        # Center on screen
        w, h = 720, 660
        sw = self.winfo_screenwidth()
        sh = self.winfo_screenheight()
        self.geometry(f"{w}x{h}+{(sw-w)//2}+{(sh-h)//2}")

        self.categories   = build_categories()
        self.check_vars   = {}   # key -> BooleanVar
        self.size_labels  = {}   # key -> StringVar
        self._scanning    = False
        self._cleaning    = False

        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        # ---- Header ----
        header = tk.Frame(self, bg=ACCENT, pady=14)
        header.pack(fill="x")

        title_fnt  = tkfont.Font(family="Segoe UI", size=18, weight="bold")
        sub_fnt    = tkfont.Font(family="Segoe UI", size=9)

        tk.Label(header, text="  C Drive Cleanup Tool",
                 font=title_fnt, fg=FG_WHITE, bg=ACCENT).pack(side="left")
        tk.Label(header, text="Safe. Fast. Professional.",
                 font=sub_fnt, fg="#d0d0f0", bg=ACCENT).pack(side="left", padx=8, pady=4)

        # ---- Main body ----
        body = tk.Frame(self, bg=BG_DARK)
        body.pack(fill="both", expand=True, padx=20, pady=12)

        # Category list card
        card = tk.Frame(body, bg=BG_PANEL, bd=0, relief="flat",
                        highlightbackground=BG_CARD, highlightthickness=1)
        card.pack(fill="both", expand=True)

        # Column headers
        hdr = tk.Frame(card, bg=BG_CARD, pady=6)
        hdr.pack(fill="x")
        hdr_fnt = tkfont.Font(family="Segoe UI", size=9, weight="bold")
        tk.Label(hdr, text="  Category", font=hdr_fnt,
                 fg=FG_GRAY, bg=BG_CARD, width=26, anchor="w").pack(side="left")
        tk.Label(hdr, text="Description", font=hdr_fnt,
                 fg=FG_GRAY, bg=BG_CARD, width=36, anchor="w").pack(side="left")
        tk.Label(hdr, text="Size", font=hdr_fnt,
                 fg=FG_GRAY, bg=BG_CARD, width=10, anchor="e").pack(side="left")

        # Separator
        tk.Frame(card, bg=BG_CARD, height=1).pack(fill="x")

        # Category rows
        row_fnt  = tkfont.Font(family="Segoe UI", size=10)
        desc_fnt = tkfont.Font(family="Segoe UI", size=9)

        for i, cat in enumerate(self.categories):
            var = tk.BooleanVar(value=True)
            self.check_vars[cat.key] = var

            sv = tk.StringVar(value="—")
            self.size_labels[cat.key] = sv

            row_bg = BG_PANEL if i % 2 == 0 else BG_CARD
            row = tk.Frame(card, bg=row_bg, pady=8)
            row.pack(fill="x")

            cb = tk.Checkbutton(
                row, variable=var, bg=row_bg,
                activebackground=row_bg, fg=FG_WHITE,
                selectcolor=BG_DARK, cursor="hand2",
                relief="flat", bd=0,
            )
            cb.pack(side="left", padx=(8, 0))

            tk.Label(row, text=cat.label, font=row_fnt,
                     fg=FG_WHITE, bg=row_bg, width=22, anchor="w").pack(side="left")
            tk.Label(row, text=cat.description, font=desc_fnt,
                     fg=FG_GRAY, bg=row_bg, width=34, anchor="w").pack(side="left")

            tk.Label(row, textvariable=sv, font=row_fnt,
                     fg=ACCENT2, bg=row_bg, width=10, anchor="e").pack(side="left")

        # Total bar
        total_frame = tk.Frame(card, bg=BG_CARD, pady=8)
        total_frame.pack(fill="x")
        tk.Frame(total_frame, bg=ACCENT, height=1, width=660).pack(pady=(0, 6))
        tk.Label(total_frame, text="  Total estimated size:",
                 font=hdr_fnt, fg=FG_GRAY, bg=BG_CARD).pack(side="left", padx=4)
        self.total_var = tk.StringVar(value="—")
        tk.Label(total_frame, textvariable=self.total_var,
                 font=tkfont.Font(family="Segoe UI", size=11, weight="bold"),
                 fg=FG_YELLOW, bg=BG_CARD).pack(side="right", padx=16)

        # ---- Status / progress area ----
        status_frame = tk.Frame(body, bg=BG_DARK, pady=10)
        status_frame.pack(fill="x")

        self.status_var = tk.StringVar(value="Ready. Click Scan to calculate disk usage.")
        tk.Label(status_frame, textvariable=self.status_var,
                 font=tkfont.Font(family="Segoe UI", size=9),
                 fg=FG_GRAY, bg=BG_DARK, anchor="w").pack(fill="x")

        self.progress = ttk.Progressbar(
            status_frame, orient="horizontal", mode="determinate",
            maximum=100, value=0,
            style="Custom.Horizontal.TProgressbar",
        )
        self.progress.pack(fill="x", pady=(4, 0))

        # ---- Buttons ----
        btn_frame = tk.Frame(body, bg=BG_DARK, pady=6)
        btn_frame.pack(fill="x")

        btn_fnt = tkfont.Font(family="Segoe UI", size=11, weight="bold")

        self.scan_btn = self._make_button(
            btn_frame, "  Scan", BTN_SCAN, HOVER_SCAN, BTN_FG,
            self._on_scan, btn_fnt,
        )
        self.scan_btn.pack(side="left", padx=(0, 10))

        self.clean_btn = self._make_button(
            btn_frame, "  Clean Selected", BTN_CLEAN, HOVER_CLEAN, BTN_FG,
            self._on_clean, btn_fnt,
        )
        self.clean_btn.pack(side="left")

        self.select_all_btn = self._make_button(
            btn_frame, "Select All", BG_CARD, BG_PANEL, FG_GRAY,
            self._select_all,
            tkfont.Font(family="Segoe UI", size=9),
            pad=(6, 4),
        )
        self.select_all_btn.pack(side="right", padx=(0, 6))

        self.deselect_btn = self._make_button(
            btn_frame, "Deselect All", BG_CARD, BG_PANEL, FG_GRAY,
            self._deselect_all,
            tkfont.Font(family="Segoe UI", size=9),
            pad=(6, 4),
        )
        self.deselect_btn.pack(side="right", padx=(0, 6))

        # ---- Result label ----
        self.result_var = tk.StringVar(value="")
        tk.Label(body, textvariable=self.result_var,
                 font=tkfont.Font(family="Segoe UI", size=11, weight="bold"),
                 fg=FG_GREEN, bg=BG_DARK, anchor="w").pack(fill="x", pady=(4, 0))

        # Style progress bar
        style = ttk.Style(self)
        style.theme_use("default")
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor=BG_CARD,
            background=ACCENT,
            thickness=8,
            bordercolor=BG_CARD,
            lightcolor=ACCENT,
            darkcolor=ACCENT,
        )

    def _make_button(self, parent, text, bg, hover_bg, fg, command, fnt,
                     pad=(12, 8)):
        btn = tk.Label(
            parent, text=text, font=fnt,
            fg=fg, bg=bg, cursor="hand2",
            padx=pad[0], pady=pad[1],
            relief="flat", bd=0,
        )
        btn.bind("<Button-1>", lambda e: command())
        btn.bind("<Enter>",    lambda e: btn.configure(bg=hover_bg))
        btn.bind("<Leave>",    lambda e: btn.configure(bg=bg))
        return btn

    # ------------------------------------------------------------------
    # Button handlers
    # ------------------------------------------------------------------

    def _select_all(self):
        for v in self.check_vars.values():
            v.set(True)

    def _deselect_all(self):
        for v in self.check_vars.values():
            v.set(False)

    def _on_scan(self):
        if self._scanning or self._cleaning:
            return
        self._scanning = True
        self.result_var.set("")
        self.status_var.set("Scanning…")
        self.progress["value"] = 0
        for sv in self.size_labels.values():
            sv.set("…")
        self.total_var.set("…")
        threading.Thread(target=self._scan_worker, daemon=True).start()

    def _scan_worker(self):
        n = len(self.categories)
        for i, cat in enumerate(self.categories):
            self.after(0, self.status_var.set,
                       f"Scanning: {cat.label}…")
            cat.scan()
            self.after(0, self.size_labels[cat.key].set, human_size(cat.size_bytes))
            self.after(0, self._update_progress, (i + 1) / n * 100)

        total = sum(c.size_bytes for c in self.categories)
        self.after(0, self.total_var.set, human_size(total))
        self.after(0, self.status_var.set,
                   f"Scan complete. Found {human_size(total)} of cleanable data.")
        self.after(0, self._update_progress, 0)
        self._scanning = False

    def _on_clean(self):
        if self._scanning or self._cleaning:
            return

        selected = [c for c in self.categories if self.check_vars[c.key].get()]
        if not selected:
            messagebox.showinfo("Nothing selected",
                                "Please check at least one category to clean.")
            return

        total_est = sum(c.size_bytes for c in selected)
        msg = (
            f"You are about to delete files from {len(selected)} categor"
            f"{'y' if len(selected)==1 else 'ies'} "
            f"(~{human_size(total_est)}).\n\n"
            "This cannot be undone (except for Recycle Bin items already there).\n\n"
            "Proceed?"
        )
        if not messagebox.askyesno("Confirm Cleanup", msg, icon="warning"):
            return

        self._cleaning = True
        self.result_var.set("")
        self.status_var.set("Cleaning…")
        self.progress["value"] = 0
        threading.Thread(
            target=self._clean_worker, args=(selected,), daemon=True
        ).start()

    def _clean_worker(self, selected: list[CleanCategory]):
        total_freed = 0
        n = len(selected)

        for i, cat in enumerate(selected):
            self.after(0, self.status_var.set, f"Cleaning: {cat.label}…")

            def progress_cb(done, total, _i=i, _n=n):
                pct = (_i / _n + done / (total * _n)) * 100
                self.after(0, self._update_progress, pct)

            freed = cat.clean(progress_cb=progress_cb)
            total_freed += freed
            cat.size_bytes = max(0, cat.size_bytes - freed)
            self.after(0, self.size_labels[cat.key].set, human_size(cat.size_bytes))
            self.after(0, self._update_progress, (i + 1) / n * 100)

        new_total = sum(c.size_bytes for c in self.categories)
        self.after(0, self.total_var.set, human_size(new_total))
        self.after(0, self.status_var.set,
                   "Cleanup complete.")
        self.after(0, self._update_progress, 0)
        self.after(0, self.result_var.set,
                   f"  Freed {human_size(total_freed)} of disk space.")
        self._cleaning = False

    def _update_progress(self, value: float):
        self.progress["value"] = value


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    # Warn if not on Windows (the tool will still open but paths won't exist)
    if sys.platform != "win32":
        print("[WARNING] This tool is designed for Windows. "
              "Paths and Win32 APIs will not work on this platform.")

    app = CleanerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
