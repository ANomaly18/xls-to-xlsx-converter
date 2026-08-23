"""XLS to XLSX Converter - Desktop GUI (tkinter dark theme, minimalist).

Converts .xls files to .xlsx preserving all formatting and content.
Supports drag & drop, file selection, batch conversion.
"""

from __future__ import annotations

import logging
import os
import queue
import threading
import concurrent.futures
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from xls_to_xlsx_converter.constants import (
    APP_TITLE,
    APP_VERSION,
    FOOTER_TEXT,
    OUTPUT_SUFFIX,
    resource_path,
)
from xls_to_xlsx_converter.converter import (
    convert_file,
    decode_with_fallback,
    detect_file_format,
    identify_file_type,
)

# ------------------------------------------------------------------
# Constants
# ------------------------------------------------------------------

LOG_POLL_INTERVAL_MS = 120
CONVERSION_TIMEOUT_SEC = 120
WINDOW_GEOMETRY = "860x640"
WINDOW_MIN_SIZE = (720, 500)

# Keyboard shortcuts
SHORTCUT_OPEN = "<Control-o>"
SHORTCUT_CONVERT = "<Control-Return>"

# ------------------------------------------------------------------
# Logging setup
# ------------------------------------------------------------------


def setup_logging() -> logging.Logger:
    """Configure logging to write errors to a local file."""
    logger = logging.getLogger("xls_converter")
    if not logger.handlers:
        logger.setLevel(logging.DEBUG)
        log_dir = Path(os.environ.get("LOCALAPPDATA", Path.home())) / "XlsToXlsxConverter" / "logs"
        log_dir.mkdir(parents=True, exist_ok=True)
        log_file = log_dir / "converter.log"
        fh = logging.FileHandler(str(log_file), encoding="utf-8")
        fh.setLevel(logging.DEBUG)
        formatter = logging.Formatter(
            "%(asctime)s [%(levelname)s] %(name)s: %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        fh.setFormatter(formatter)
        logger.addHandler(fh)
        logger.propagate = True
    return logger


logger = setup_logging()


# ------------------------------------------------------------------
# Color palette - refined minimalist dark theme
# ------------------------------------------------------------------

class Colors:
    """Minimalist dark theme palette — fewer colors, more impact."""
    # Surfaces
    BG_BASE = "#0E0E10"        # Main background
    BG_SURFACE = "#18181B"     # Cards, panels
    BG_ELEVATED = "#1F1F23"    # Hover, elevated elements
    BG_INPUT = "#131316"       # Inputs, code areas

    # Borders
    BORDER = "#27272A"
    BORDER_FOCUS = "#3B82F6"

    # Text
    TEXT_PRIMARY = "#FAFAFA"
    TEXT_SECONDARY = "#A1A1AA"
    TEXT_MUTED = "#52525B"

    # Accent
    ACCENT = "#3B82F6"
    ACCENT_HOVER = "#2563EB"
    ACCENT_DIM = "#1A2744"  # approx rgba(59, 130, 246, 0.1) on #0E0E10
    ACCENT_TEXT = "#60A5FA"

    # States
    SUCCESS = "#10B981"
    WARNING = "#F59E0B"
    ERROR = "#EF4444"

    # Drop zone
    DROP_BG = "#131316"
    DROP_ACTIVE_BG = "#0F1729"
    DROP_BORDER = "#27272A"
    DROP_ACTIVE_BORDER = "#3B82F6"


# ------------------------------------------------------------------
# Style definitions
# ------------------------------------------------------------------

def setup_styles() -> None:
    """Configure all ttk styles for the minimalist dark theme."""
    style = ttk.Style()
    style.theme_use("clam")

    # Frames
    style.configure("TFrame", background=Colors.BG_BASE)
    style.configure("Surface.TFrame", background=Colors.BG_SURFACE, borderwidth=0)
    style.configure("Card.TFrame", background=Colors.BG_SURFACE, borderwidth=1, relief="solid")

    # Labels
    style.configure(
        "Title.TLabel",
        background=Colors.BG_BASE,
        foreground=Colors.TEXT_PRIMARY,
        font=("Segoe UI", 20, "bold"),
    )
    style.configure(
        "Subtitle.TLabel",
        background=Colors.BG_BASE,
        foreground=Colors.TEXT_MUTED,
        font=("Segoe UI", 9),
    )
    style.configure(
        "SectionTitle.TLabel",
        background=Colors.BG_SURFACE,
        foreground=Colors.TEXT_SECONDARY,
        font=("Segoe UI", 8, "bold"),
    )
    style.configure(
        "Body.TLabel",
        background=Colors.BG_BASE,
        foreground=Colors.TEXT_SECONDARY,
        font=("Segoe UI", 9),
    )
    style.configure(
        "Status.TLabel",
        background=Colors.BG_BASE,
        foreground=Colors.TEXT_SECONDARY,
        font=("Segoe UI", 9),
    )
    style.configure(
        "Badge.TLabel",
        background=Colors.ACCENT_DIM,
        foreground=Colors.ACCENT_TEXT,
        font=("Segoe UI Semibold", 8),
    )
    style.configure(
        "Success.TLabel",
        background=Colors.BG_BASE,
        foreground=Colors.SUCCESS,
        font=("Segoe UI", 9),
    )
    style.configure(
        "Error.TLabel",
        background=Colors.BG_BASE,
        foreground=Colors.ERROR,
        font=("Segoe UI", 9),
    )
    style.configure(
        "Footer.TLabel",
        background=Colors.BG_BASE,
        foreground=Colors.TEXT_MUTED,
        font=("Segoe UI", 8),
    )

    # Buttons
    style.configure(
        "Primary.TButton",
        background=Colors.ACCENT,
        foreground="#FFFFFF",
        bordercolor=Colors.ACCENT,
        font=("Segoe UI Semibold", 9),
        padding=(28, 8),
    )
    style.map(
        "Primary.TButton",
        background=[("active", Colors.ACCENT_HOVER), ("disabled", "#1A2744")],
        foreground=[("disabled", Colors.TEXT_MUTED)],
    )
    style.configure(
        "Secondary.TButton",
        background=Colors.BG_ELEVATED,
        foreground=Colors.TEXT_SECONDARY,
        bordercolor=Colors.BORDER,
        font=("Segoe UI", 9),
        padding=(16, 8),
    )
    style.map(
        "Secondary.TButton",
        background=[("active", "#2A2A30")],
        foreground=[("disabled", Colors.TEXT_MUTED)],
    )
    style.configure(
        "Ghost.TButton",
        background="transparent",
        foreground=Colors.TEXT_MUTED,
        bordercolor="transparent",
        font=("Segoe UI", 8),
        padding=(8, 4),
    )
    style.map(
        "Ghost.TButton",
        background=[("active", Colors.BG_ELEVATED)],
        foreground=[("active", Colors.TEXT_SECONDARY)],
    )
    style.configure(
        "Danger.TButton",
        background="transparent",
        foreground=Colors.ERROR,
        bordercolor="transparent",
        font=("Segoe UI", 8),
        padding=(8, 4),
    )
    style.map(
        "Danger.TButton",
        background=[("active", "#2A1515")],
    )

    # Progressbar — configure built-in style (custom style needs explicit layout in clam)
    style.layout("TProgressbar", style.layout("Horizontal.TProgressbar"))
    style.configure(
        "Horizontal.TProgressbar",
        troughcolor=Colors.BG_INPUT,
        background=Colors.ACCENT,
        bordercolor=Colors.BG_INPUT,
        lightcolor=Colors.ACCENT,
        darkcolor=Colors.ACCENT,
    )

    # Treeview
    style.configure(
        "FileTree.Treeview",
        background=Colors.BG_SURFACE,
        foreground=Colors.TEXT_PRIMARY,
        fieldbackground=Colors.BG_SURFACE,
        borderwidth=0,
        font=("Segoe UI", 9),
        rowheight=32,
    )
    style.map(
        "FileTree.Treeview",
        background=[("selected", "#1A2744")],
        foreground=[("selected", Colors.TEXT_PRIMARY)],
    )
    style.configure(
        "FileTree.Treeview.Heading",
        background=Colors.BG_SURFACE,
        foreground=Colors.TEXT_MUTED,
        borderwidth=0,
        font=("Segoe UI Semibold", 8),
        padding=(12, 8),
    )
    style.map(
        "FileTree.Treeview.Heading",
        background=[("active", Colors.BG_ELEVATED)],
    )

    # Scrollbar
    style.configure(
        "Minimal.Vertical.TScrollbar",
        troughcolor=Colors.BG_BASE,
        bordercolor=Colors.BORDER,
        darkcolor=Colors.BORDER,
        lightcolor=Colors.BORDER,
        arrowcolor=Colors.TEXT_MUTED,
        gripcount=0,
    )

    # Entry
    style.configure(
        "Minimal.TEntry",
        fieldbackground=Colors.BG_INPUT,
        foreground=Colors.TEXT_SECONDARY,
        bordercolor=Colors.BORDER,
        font=("Segoe UI", 9),
    )


# ------------------------------------------------------------------
# Application class
# ------------------------------------------------------------------

class XlsToXlsxApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.files: list[Path] = []
        self.file_statuses: dict[str, str] = {}
        self.status_var = tk.StringVar(value="Siap")
        self.progress_var = tk.DoubleVar(value=0)
        self.progress_label_var = tk.StringVar(value="")
        self.log_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_flag = threading.Event()
        self.drop_zone_active = False
        self.output_folder: str | None = None

        self.configure_root()
        setup_styles()
        self.build_layout()
        self.root.after(LOG_POLL_INTERVAL_MS, self.drain_log_queue)

    # ------------------------------------------------------------------
    # Window setup
    # ------------------------------------------------------------------

    def configure_root(self) -> None:
        self.root.title(APP_TITLE)
        self.root.geometry(WINDOW_GEOMETRY)
        self.root.minsize(*WINDOW_MIN_SIZE)
        self.root.configure(bg=Colors.BG_BASE)
        self.set_window_icon()
        self.root.bind(SHORTCUT_OPEN, lambda e: self.browse_files())
        self.root.bind(SHORTCUT_CONVERT, lambda e: self.start_conversion() if self.files and not (self.worker and self.worker.is_alive()) else None)

    def set_window_icon(self) -> None:
        ico_path = resource_path("assets", "app.ico")
        png_path = resource_path("assets", "app.png")
        try:
            if ico_path.exists():
                self.root.iconbitmap(str(ico_path))
            if png_path.exists():
                self.icon_image = tk.PhotoImage(file=str(png_path))
                self.root.iconphoto(True, self.icon_image)
        except tk.TclError:
            pass

    # ------------------------------------------------------------------
    # Layout
    # ------------------------------------------------------------------

    def build_layout(self) -> None:
        # Main scrollable container
        main = ttk.Frame(self.root, style="TFrame", padding=(24, 20, 24, 16))
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)

        # Row weights for expandable sections
        main.rowconfigure(3, weight=1)  # File list expands
        main.rowconfigure(5, weight=1)  # Log expands

        self._build_header(main, row=0)
        self._build_drop_zone(main, row=1)
        self._build_actions(main, row=2)
        self._build_file_list(main, row=3)
        self._build_output_section(main, row=4)
        self._build_progress(main, row=5)
        self._build_log_panel(main, row=6)
        self._build_status_bar(main, row=7)

    # --- Header ---

    def _build_header(self, parent: ttk.Frame, row: int) -> None:
        hdr = ttk.Frame(parent, style="TFrame", padding=(0, 0, 0, 20))
        hdr.grid(row=row, column=0, sticky="ew")

        # Left: Title + subtitle
        left = ttk.Frame(hdr, style="TFrame")
        left.pack(side="left")

        ttk.Label(left, text="XLS → XLSX", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            left,
            text="Konverter Excel — cepat, akurat, batch processing.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        # Right: File count badge
        self.badge_var = tk.StringVar(value="0 file")
        self.badge = tk.Label(
            left,
            textvariable=self.badge_var,
            bg="#1A2744",
            fg="#60A5FA",
            font=("Segoe UI Semibold", 8),
            padx=10,
            pady=3,
        )
        self.badge.pack(side="left", padx=(16, 0))

    # --- Drop Zone ---

    def _build_drop_zone(self, parent: ttk.Frame, row: int) -> None:
        """Clean drop zone using canvas for dashed border, no emoji clutter."""
        self.drop_frame = ttk.Frame(parent, style="TFrame")
        self.drop_frame.grid(row=row, column=0, sticky="ew", pady=(0, 16))
        self.drop_frame.columnconfigure(0, weight=1)

        # Canvas for precise dashed border + centered content
        self.drop_canvas = tk.Canvas(
            self.drop_frame,
            bg=Colors.DROP_BG,
            highlightthickness=0,
            height=80,
            cursor="hand2",
        )
        self.drop_canvas.pack(fill="x")

        self._draw_drop_zone()

        # Events
        self.drop_canvas.bind("<Button-1>", lambda e: self.browse_files())
        self.drop_canvas.bind("<Enter>", self._on_drop_enter)
        self.drop_canvas.bind("<Leave>", self._on_drop_leave)
        self.drop_canvas.bind("<Configure>", self._center_drop_content)

    def _draw_drop_zone(self, active: bool = False) -> None:
        """Draw or redraw the drop zone content."""
        self.drop_canvas.delete("all")
        w = self.drop_canvas.winfo_width() or 800
        h = self.drop_canvas.winfo_height() or 80

        border_color = Colors.DROP_ACTIVE_BORDER if active else Colors.DROP_BORDER
        bg = Colors.DROP_ACTIVE_BG if active else Colors.DROP_BG
        text_color = Colors.ACCENT_TEXT if active else Colors.TEXT_SECONDARY

        self.drop_canvas.configure(bg=bg)

        # Dashed border
        self.drop_canvas.create_rectangle(
            2, 2, w - 2, h - 2,
            outline=border_color,
            width=1.5,
            dash=(5, 4),
        )

        # Centered content
        cy = h / 2
        self.drop_canvas.create_text(
            w / 2, cy - 14,
            text="Seret file .xls ke sini",
            fill=text_color,
            font=("Segoe UI", 11),
        )
        self.drop_canvas.create_text(
            w / 2, cy + 6,
            text="atau klik untuk memilih",
            fill=Colors.TEXT_MUTED,
            font=("Segoe UI", 9),
        )

    def _center_drop_content(self, event: tk.Event | None = None) -> None:
        self._draw_drop_zone(self.drop_zone_active)

    def _on_drop_enter(self, event: tk.Event) -> None:
        if not self.drop_zone_active:
            self.drop_zone_active = True
            self._draw_drop_zone(active=True)

    def _on_drop_leave(self, event: tk.Event) -> None:
        if self.drop_zone_active:
            self.drop_zone_active = False
            self._draw_drop_zone(active=False)

    # --- Action Buttons ---

    def _build_actions(self, parent: ttk.Frame, row: int) -> None:
        row_frame = ttk.Frame(parent, style="TFrame", padding=(0, 0, 0, 16))
        row_frame.grid(row=row, column=0, sticky="ew")
        row_frame.columnconfigure(0, weight=1)

        # Left: Browse
        self.browse_button = ttk.Button(
            row_frame, text="Pilih File", style="Secondary.TButton",
            command=self.browse_files,
        )
        self.browse_button.pack(side="left")

        # Right: Convert
        btn_right = ttk.Frame(row_frame, style="TFrame")
        btn_right.pack(side="right")

        self.convert_button = ttk.Button(
            btn_right, text="Konversi", style="Primary.TButton",
            command=self.start_conversion,
        )
        self.convert_button.pack(side="left", padx=(8, 0))
        self.convert_button.configure(state="disabled")

        self.cancel_button = ttk.Button(
            btn_right, text="Batal", style="Secondary.TButton",
            command=self.cancel_conversion,
        )
        self.cancel_button.pack(side="left", padx=(8, 0))
        self.cancel_button.pack_forget()

    # --- File List ---

    def _build_file_list(self, parent: ttk.Frame, row: int) -> None:
        self.file_section = ttk.Frame(parent, style="Surface.TFrame")
        self.file_section.grid(row=row, column=0, sticky="nsew", pady=(0, 16))
        self.file_section.columnconfigure(0, weight=1)
        self.file_section.rowconfigure(1, weight=1)

        # Section header
        header = ttk.Frame(self.file_section, style="Surface.TFrame", padding=(0, 0, 0, 8))
        header.grid(row=0, column=0, sticky="ew")

        self.list_title = ttk.Label(
            header, text="File dipilih", style="SectionTitle.TLabel",
        )
        self.list_title.pack(side="left")

        self.clear_button = ttk.Button(
            header, text="Hapus semua", style="Danger.TButton",
            command=self.clear_files,
        )
        self.clear_button.pack(side="right")
        self.clear_button.configure(state="disabled")

        # Treeview
        columns = ("name", "type", "status")
        self.file_tree = ttk.Treeview(
            self.file_section,
            columns=columns,
            show="headings",
            height=4,
            style="FileTree.Treeview",
        )

        self.file_tree.heading("name", text="Nama File", anchor="w")
        self.file_tree.heading("type", text="Format", anchor="w")
        self.file_tree.heading("status", text="Status", anchor="e")

        self.file_tree.column("name", width=400, anchor="w", minwidth=180)
        self.file_tree.column("type", width=80, anchor="w", minwidth=60)
        self.file_tree.column("status", width=100, anchor="e", minwidth=80)

        # Scrollbar
        sb = ttk.Scrollbar(
            self.file_section, orient="vertical",
            command=self.file_tree.yview,
            style="Minimal.Vertical.TScrollbar",
        )
        self.file_tree.configure(yscrollcommand=sb.set)

        self.file_tree.grid(row=1, column=0, sticky="nsew")
        sb.grid(row=1, column=1, sticky="ns")

        # Status colors
        self.file_tree.tag_configure("success", foreground=Colors.SUCCESS)
        self.file_tree.tag_configure("error", foreground=Colors.ERROR)
        self.file_tree.tag_configure("warning", foreground=Colors.WARNING)

        # Placeholder
        self.file_tree.insert("", "end", values=("Belum ada file", "", ""))

    # --- Output Section ---

    def _build_output_section(self, parent: ttk.Frame, row: int) -> None:
        sec = ttk.Frame(parent, style="Surface.TFrame", padding=(12, 10))
        sec.grid(row=row, column=0, sticky="ew", pady=(0, 16))
        sec.columnconfigure(0, weight=1)

        inner = ttk.Frame(sec, style="Surface.TFrame")
        inner.grid(row=0, column=0, sticky="ew")
        inner.columnconfigure(1, weight=1)

        ttk.Label(inner, text="Folder output", style="SectionTitle.TLabel").grid(
            row=0, column=0, sticky="w", padx=(0, 12)
        )

        self.output_var = tk.StringVar(value="Sama dengan file asal")
        self.output_entry = tk.Entry(
            inner,
            textvariable=self.output_var,
            bg=Colors.BG_INPUT,
            fg=Colors.TEXT_MUTED,
            insertbackground=Colors.TEXT_PRIMARY,
            relief="flat",
            font=("Segoe UI", 9),
            state="readonly",
        )
        self.output_entry.grid(row=0, column=1, sticky="ew", padx=(0, 12), ipady=4)

        self.output_browse = ttk.Button(
            inner, text="Pilih", style="Secondary.TButton",
            command=self.browse_output,
        )
        self.output_browse.grid(row=0, column=2)

    # --- Progress ---

    def _build_progress(self, parent: ttk.Frame, row: int) -> None:
        self.progress_frame = ttk.Frame(parent, style="TFrame", padding=(0, 0, 0, 12))
        self.progress_frame.grid(row=row, column=0, sticky="ew")
        self.progress_frame.columnconfigure(0, weight=1)
        self.progress_frame.grid_remove()

        # Progress text
        self.progress_label = ttk.Label(
            self.progress_frame,
            textvariable=self.progress_label_var,
            style="Body.TLabel",
        )
        self.progress_label.grid(row=0, column=0, sticky="w", pady=(0, 6))

        # Bar
        self.progress = ttk.Progressbar(
            self.progress_frame,
            style="Horizontal.TProgressbar",
            mode="determinate",
            variable=self.progress_var,
        )
        self.progress.grid(row=1, column=0, sticky="ew")

    # --- Log Panel ---

    def _build_log_panel(self, parent: ttk.Frame, row: int) -> None:
        sec = ttk.Frame(parent, style="Surface.TFrame")
        sec.grid(row=row, column=0, sticky="nsew", pady=(0, 12))
        sec.columnconfigure(0, weight=1)
        sec.rowconfigure(1, weight=1)

        header = ttk.Frame(sec, style="Surface.TFrame", padding=(12, 10, 12, 6))
        header.grid(row=0, column=0, sticky="ew")

        ttk.Label(header, text="Log", style="SectionTitle.TLabel").pack(side="left")

        log_inner = ttk.Frame(sec, style="Surface.TFrame")
        log_inner.grid(row=1, column=0, sticky="nsew", padx=(0, 0))
        log_inner.columnconfigure(0, weight=1)
        log_inner.rowconfigure(0, weight=1)

        self.log_text = tk.Text(
            log_inner,
            bg=Colors.BG_INPUT,
            fg=Colors.TEXT_SECONDARY,
            insertbackground=Colors.TEXT_PRIMARY,
            relief="flat",
            wrap="word",
            font=("Consolas", 9),
            padx=12,
            pady=8,
            state="normal",
            height=6,
        )
        self.log_text.bind("<Key>", self._log_read_only)
        self.log_text.grid(row=0, column=0, sticky="nsew")

        sb = ttk.Scrollbar(
            log_inner, orient="vertical", command=self.log_text.yview,
            style="Minimal.Vertical.TScrollbar",
        )
        sb.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=sb.set)

        # Tags
        self.log_text.tag_configure("success", foreground=Colors.SUCCESS)
        self.log_text.tag_configure("error", foreground=Colors.ERROR)
        self.log_text.tag_configure("warning", foreground=Colors.WARNING)
        self.log_text.tag_configure("info", foreground=Colors.TEXT_SECONDARY)
        self.log_text.tag_configure("muted", foreground=Colors.TEXT_MUTED)

    # --- Status Bar ---

    def _build_status_bar(self, parent: ttk.Frame, row: int) -> None:
        bar = ttk.Frame(parent, style="TFrame", padding=(0, 4, 0, 0))
        bar.grid(row=row, column=0, sticky="ew")
        bar.columnconfigure(0, weight=1)
        bar.columnconfigure(1, weight=1)

        ttk.Label(bar, textvariable=self.status_var, style="Status.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(bar, text=FOOTER_TEXT, style="Footer.TLabel").grid(
            row=0, column=1, sticky="e"
        )

    # ------------------------------------------------------------------
    # Log read-only helper
    # ------------------------------------------------------------------

    def _log_read_only(self, event: tk.Event) -> str | None:
        """Allow copy (Ctrl+C, Ctrl+A) but block all other key input."""
        allowed = {"c", "a"}
        if event.state & 0x4:
            if event.keysym.lower() in allowed:
                return None
            return "break"
        return "break"

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def browse_files(self) -> None:
        if self.worker and self.worker.is_alive():
            return
        files = filedialog.askopenfilenames(
            title="Pilih file .xls",
            filetypes=[("Excel 97-2003 (.xls)", "*.xls"), ("Semua file", "*.*")],
        )
        if not files:
            return
        self.add_files([Path(f) for f in files])

    def add_files(self, new_files: list[Path]) -> None:
        existing = {str(p) for p in self.files}
        added = 0
        for f in new_files:
            if str(f) not in existing and f.suffix.lower() == ".xls":
                self.files.append(f)
                self.file_statuses[f.name] = "Menunggu"
                existing.add(str(f))
                added += 1

        if added > 0:
            self.refresh_file_list()
            n = len(self.files)
            self.status_var.set(f"{n} file siap dikonversi")

    def browse_output(self) -> None:
        folder = filedialog.askdirectory(title="Pilih folder output")
        if folder:
            self.output_folder = folder
            self.output_var.set(folder)
            self.output_entry.configure(fg=Colors.TEXT_PRIMARY)

    def clear_files(self) -> None:
        self.files.clear()
        self.file_statuses.clear()
        self.refresh_file_list()
        self.status_var.set("Siap")

    def refresh_file_list(self) -> None:
        for item in self.file_tree.get_children():
            self.file_tree.delete(item)

        count = len(self.files)
        self.badge_var.set(f"{count} file")

        tree_h = min(max(count, 1), 6)
        self.file_tree.configure(height=tree_h)

        if count == 0:
            self.list_title.configure(text="File dipilih")
            self.convert_button.configure(state="disabled")
            self.clear_button.configure(state="disabled")
            self.file_tree.insert("", "end", values=("Belum ada file", "", ""))
        else:
            self.list_title.configure(text=f"{count} file")
            self.convert_button.configure(state="normal")
            self.clear_button.configure(state="normal")

            for filepath in self.files:
                status = self.file_statuses.get(filepath.name, "Menunggu")
                item_id = self.file_tree.insert("", "end", values=(
                    filepath.name,
                    "XLS",
                    status,
                ))
                if status == "Berhasil":
                    self.file_tree.item(item_id, tags=("success",))
                elif status in ("Gagal", "Ditolak"):
                    self.file_tree.item(item_id, tags=("error",))
                elif status == "Dibatalkan":
                    self.file_tree.item(item_id, tags=("warning",))

    def start_conversion(self) -> None:
        if not self.files:
            messagebox.showwarning("Tidak ada file", "Pilih file .xls terlebih dahulu.")
            return
        if self.worker and self.worker.is_alive():
            return

        self.log_text.delete("1.0", "end")
        self.cancel_flag.clear()

        output_dir = self.output_folder
        if output_dir and not Path(output_dir).is_dir():
            messagebox.showwarning("Folder tidak valid", "Pilih folder output yang valid.")
            return

        self.set_running_state(True)
        self.append_log(f"Konversi {len(self.files)} file...")
        self.append_log("")

        self.worker = threading.Thread(
            target=self.worker_run,
            args=(list(self.files), output_dir),
            daemon=True,
        )
        self.worker.start()

    def worker_run(self, files: list[Path], output_dir: str | None) -> None:
        total = len(files)
        for idx, filepath in enumerate(files):
            if self.cancel_flag.is_set():
                self.log_queue.put(("log", ("\nDibatalkan.", "warning")))
                self.log_queue.put(("status", (filepath.name, "Dibatalkan")))
                self.log_queue.put(("done", None))
                return

            progress_pct = (idx / total) * 100
            self.log_queue.put(("progress", progress_pct))
            self.log_queue.put(("status", (filepath.name, "Memproses...")))

            try:
                if output_dir:
                    output_path = Path(output_dir) / f"{filepath.stem}{OUTPUT_SUFFIX}"
                else:
                    output_path = filepath.with_name(f"{filepath.stem}{OUTPUT_SUFFIX}")

                detect_msg = f"[{idx + 1}/{total}] {filepath.name}"
                self.log_queue.put(("log", (detect_msg, "info")))

                file_format = detect_file_format(filepath)
                self.log_queue.put(("log", (f"  Format: {file_format.upper()}", "muted")))

                if file_format == "html":
                    from bs4 import BeautifulSoup
                    with open(filepath, "rb") as f:
                        raw = f.read()
                    html_content = decode_with_fallback(raw)
                    soup = BeautifulSoup(html_content, "lxml")
                    tables = soup.find_all("table")
                    if tables:
                        target_table = max(tables, key=lambda t: len(t.find_all("tr")))
                        file_type = identify_file_type(soup, target_table)
                        if file_type:
                            self.log_queue.put(("log", (f"  Tipe: {file_type}", "muted")))
                        else:
                            self.log_queue.put(("log", ("  Tipe: Tidak dikenali", "warning")))
                    else:
                        self.log_queue.put(("log", ("  Tipe: HTML (tanpa tabel)", "warning")))
                else:
                    self.log_queue.put(("log", ("  Tipe: Binary XLS", "muted")))

                def progress_msg(msg: str) -> None:
                    self.log_queue.put(("log", (f"  {msg}", "muted")))

                def _do_convert():
                    return convert_file(filepath, output_path, progress_callback=progress_msg)

                with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
                    future = executor.submit(_do_convert)
                    result = future.result(timeout=CONVERSION_TIMEOUT_SEC)

                size_kb = os.path.getsize(result) / 1024
                self.log_queue.put(("log", (f"  ✓ Berhasil ({size_kb:.1f} KB)", "success")))
                self.log_queue.put(("log", ("", "info")))
                self.log_queue.put(("status", (filepath.name, "Berhasil")))

            except concurrent.futures.TimeoutError:
                self.log_queue.put(("log", (f"  ✗ Timeout: Konversi melebihi 120 detik", "error")))
                self.log_queue.put(("log", ("", "info")))
                self.log_queue.put(("status", (filepath.name, "Gagal")))
            except FileExistsError as exc:
                self.log_queue.put(("log", (f"  ✗ Tidak ditimpa: {exc}", "warning")))
                self.log_queue.put(("log", ("", "info")))
                self.log_queue.put(("status", (filepath.name, "Sudah ada")))
            except ValueError as exc:
                self.log_queue.put(("log", (f"  ✗ Ditolak: {exc}", "error")))
                self.log_queue.put(("log", ("", "info")))
                self.log_queue.put(("status", (filepath.name, "Ditolak")))
            except Exception as exc:
                self.log_queue.put(("log", (f"  ✗ Gagal: {exc}", "error")))
                self.log_queue.put(("log", ("", "info")))
                self.log_queue.put(("status", (filepath.name, "Gagal")))

        self.log_queue.put(("progress", 100))
        self.log_queue.put(("done", None))

    def drain_log_queue(self) -> None:
        try:
            while True:
                event, payload = self.log_queue.get_nowait()
                if event == "log":
                    msg, level = payload if isinstance(payload, tuple) else (payload, "info")
                    self.append_log(str(msg), level)
                elif event == "status":
                    filename, status = payload
                    self.file_statuses[filename] = status
                    self.refresh_file_list()
                elif event == "progress":
                    self.progress_var.set(payload)
                    current = int(self.progress_var.get())
                    self.progress_label_var.set(f"{current}%")
                elif event == "done":
                    self.set_running_state(False)
                    self.status_var.set("Selesai")
                    self.append_log("\nSelesai.")
                    messagebox.showinfo("Selesai", "Semua file selesai dikonversi.")
        except queue.Empty:
            pass
        finally:
            self.root.after(LOG_POLL_INTERVAL_MS, self.drain_log_queue)

    def append_log(self, message: str, level: str = "info") -> None:
        self.log_text.insert("end", f"{message}\n", level)
        self.log_text.see("end")

    def set_running_state(self, is_running: bool) -> None:
        if is_running:
            self.status_var.set("Mengonversi...")
            self.progress_frame.grid()
            self.progress_var.set(0)
            self.progress_label_var.set("0%")
            self.browse_button.configure(state="disabled")
            self.convert_button.pack_forget()
            self.cancel_button.pack(side="left", padx=(8, 0))
            self.clear_button.configure(state="disabled")
            self.output_browse.configure(state="disabled")
        else:
            self.progress.stop()
            self.browse_button.configure(state="normal")
            self.cancel_button.pack_forget()
            self.convert_button.pack(side="left", padx=(8, 0))
            self.convert_button.configure(state="normal" if self.files else "disabled")
            self.clear_button.configure(state="normal" if self.files else "disabled")
            self.output_browse.configure(state="normal")

    def cancel_conversion(self) -> None:
        self.cancel_flag.set()
        self.cancel_button.configure(state="disabled")
        self.status_var.set("Membatalkan...")


# ------------------------------------------------------------------
# Entry point
# ------------------------------------------------------------------

def main() -> None:
    root = tk.Tk()
    XlsToXlsxApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
