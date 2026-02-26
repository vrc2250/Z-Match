"""
Z-Match
---------------------------------
Matches recorder metadata (CSV) with transmitter audio files (CSV)
using timecode + userbits, then renames/copies files and generates a PDF report.
"""

import io
import os
import re
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk


# ---------------------------------------------------------------------------
# CheckboxTreeview — Treeview with a togglable checkbox column
# ---------------------------------------------------------------------------

class CheckboxTreeview(ttk.Treeview):
    """
    A ttk.Treeview subclass that draws a ☐/☑ checkbox as the first visible
    column.  Check-state is stored per iid in self._checked (set of iids).

    Public API:
        check(iid)   / uncheck(iid)
        is_checked(iid) -> bool
        checked_items() -> list[str]   # iids of all checked rows
        set_all(state: bool)           # check or uncheck every visible row
    """

    BOX_CHECKED   = "☑"
    BOX_UNCHECKED = "☐"
    CB_COL        = "__cb__"

    def __init__(self, parent, columns, **kw):
        all_cols = (self.CB_COL, *columns)
        super().__init__(parent, columns=all_cols, **kw)

        self._checked: set[str] = set()

        # Larger font for the checkbox symbols
        style = ttk.Style()
        style.configure("Checkbox.Treeview", font=("Arial", 14), rowheight=26)
        self.configure(style="Checkbox.Treeview")

        self.heading(self.CB_COL, text=self.BOX_UNCHECKED, command=self._toggle_all)
        self.column(self.CB_COL, width=40, minwidth=40, stretch=False, anchor="center")

        self.bind("<ButtonRelease-1>", self._on_click)

    # ------------------------------------------------------------------
    # Public helpers
    # ------------------------------------------------------------------

    def check(self, iid: str) -> None:
        self._checked.add(iid)
        self._redraw(iid)

    def uncheck(self, iid: str) -> None:
        self._checked.discard(iid)
        self._redraw(iid)

    def is_checked(self, iid: str) -> bool:
        return iid in self._checked

    def checked_items(self) -> list[str]:
        """Return iids of every currently-checked row (in tree order)."""
        return [iid for iid in self.get_children() if iid in self._checked]

    def set_all(self, state: bool) -> None:
        """Check or uncheck every visible row and sync the header."""
        for iid in self.get_children():
            if state:
                self._checked.add(iid)
            else:
                self._checked.discard(iid)
            self._redraw(iid)
        self._sync_header()

    # ------------------------------------------------------------------
    # Overrides
    # ------------------------------------------------------------------

    def insert(self, parent, index, iid=None, **kw):
        """Insert a row defaulting to checked, then sync header."""
        values = list(kw.pop("values", ()))
        kw["values"] = [self.BOX_CHECKED, *values]
        result = super().insert(parent, index, iid=iid, **kw)
        actual_iid = result if result else iid
        self._checked.add(actual_iid)
        self._sync_header()
        return result

    def delete(self, *items) -> None:
        targets = items if items else self.get_children()
        for iid in targets:
            self._checked.discard(iid)
        super().delete(*targets)
        self._sync_header()

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------

    def reset(self) -> None:
        """Nuke all rows and internal state, leaving a clean empty tree."""
        self._checked.clear()
        for iid in self.get_children():
            super().delete(iid)
        self.heading(self.CB_COL, text=self.BOX_UNCHECKED)

    def _redraw(self, iid: str) -> None:
        """Update the checkbox symbol for a single row."""
        symbol = self.BOX_CHECKED if iid in self._checked else self.BOX_UNCHECKED
        row_vals = list(self.item(iid, "values"))
        if row_vals:
            row_vals[0] = symbol
            self.item(iid, values=row_vals)

    def _sync_header(self) -> None:
        """Set header symbol based purely on current _checked state."""
        all_iids = self.get_children()
        all_on = bool(all_iids) and all(iid in self._checked for iid in all_iids)
        self.heading(self.CB_COL, text=self.BOX_CHECKED if all_on else self.BOX_UNCHECKED)

    def _toggle_all(self) -> None:
        """Header clicked: if every row is checked → uncheck all, else check all."""
        all_iids = self.get_children()
        all_on = bool(all_iids) and all(iid in self._checked for iid in all_iids)
        self.set_all(not all_on)

    def _on_click(self, event: tk.Event) -> None:
        """Row clicked: toggle that row's checkbox. Ignore header/column clicks."""
        # identify_region tells us if the click landed on a cell row vs heading
        region = self.identify_region(event.x, event.y)
        if region != "cell":
            return  # heading clicks are handled by command= binding, don't double-fire
        iid = self.identify_row(event.y)
        if not iid:
            return
        if iid in self._checked:
            self.uncheck(iid)
        else:
            self.check(iid)
        self._sync_header()

try:
    from fpdf import FPDF
    FPDF_AVAILABLE = True
except ImportError:
    FPDF_AVAILABLE = False


# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

APP_TITLE = "Z-Match"
APP_GEOMETRY = "1250x850"
DEFAULT_FPS = 23.98
DEFAULT_OFFSET_SECONDS = 3.0

MATCH_COLORS: list[dict] = [
    {"bg": "#c8e6c9", "fg": "black"},
    {"bg": "#bbdefb", "fg": "black"},
    {"bg": "#fff9c4", "fg": "black"},
    {"bg": "#f8bbd0", "fg": "black"},
    {"bg": "#e1bee7", "fg": "black"},
    {"bg": "#ffccbc", "fg": "black"},
]


# ---------------------------------------------------------------------------
# Data layer
# ---------------------------------------------------------------------------

@dataclass
class MatchResult:
    """Represents a single matched recorder/transmitter pair."""
    recorder_idx: int
    transmitter_idx: int
    new_filename: str
    original_filename: str
    offset_seconds: float
    fps_mismatch: bool = False          # True when recorder/transmitter FPS differ
    fps_recorder: float = DEFAULT_FPS
    fps_transmitter: float = DEFAULT_FPS


def timecode_to_seconds(tc: str, fps: float = DEFAULT_FPS) -> float:
    """Convert a timecode string (HH:MM:SS:FF) to total seconds."""
    try:
        h, m, s, f = map(float, re.split(r'[:;.]', str(tc)))
        return h * 3600 + m * 60 + s + f / fps
    except Exception:
        return 0.0


def parse_csv(path: str, list_num: int) -> tuple[pd.DataFrame, str]:
    """
    Read and normalise a recorder or transmitter CSV file.

    Returns:
        (DataFrame, sound_roll_string)  — sound_roll is empty for transmitter files.
    """
    with open(path, "r", encoding="latin1") as fh:
        content = fh.read()

    lines = content.splitlines()
    sound_roll = "---"

    if list_num == 1:
        for line in lines:
            if "FolderName" in line:
                value = line.split("FolderName", 1)[1]
                value = value.replace("=", "").replace(":", "").strip()
                sound_roll = value.split(",")[0].strip()
                break

    header_idx = next(i for i, line in enumerate(lines) if "FileID" in line)
    df = pd.read_csv(
        io.StringIO("\n".join(lines[header_idx:])),
        skipinitialspace=True,
        dtype=str,
    )
    df.columns = [str(c).replace('"', "").strip() for c in df.columns]
    df = df.map(
        lambda x: str(x).replace('"', "").strip()
        if isinstance(x, str) and str(x).lower() != "nan"
        else ""
    )
    return df, sound_roll


def find_matches(
    recorder_df: pd.DataFrame,
    transmitter_df: pd.DataFrame,
    recorder_indices: list[int],
    track_col: str,
    offset_limit: float,
) -> list[MatchResult]:
    """Return all recorder/transmitter pairs that agree on userbits and timecode."""
    results: list[MatchResult] = []

    for rec_idx in recorder_indices:
        rec_row = recorder_df.loc[rec_idx]
        fps_rec = float(rec_row.get("FrameRate", DEFAULT_FPS))
        tc_rec = timecode_to_seconds(rec_row.get("TimeCode"), fps_rec)
        ub_rec = str(rec_row.get("UserBits"))
        char = rec_row.get(track_col, "")

        for trans_idx, trans_row in transmitter_df.iterrows():
            fps_trans = float(trans_row.get("FrameRate", DEFAULT_FPS) or DEFAULT_FPS)
            tc_trans = timecode_to_seconds(trans_row.get("TimeCode"), fps_trans)
            ub_match = ub_rec == str(trans_row.get("UserBits"))
            tc_match = abs(tc_rec - tc_trans) <= offset_limit

            if ub_match and tc_match:
                orig = str(trans_row.get("FileName", "")).strip()
                ext = os.path.splitext(orig)[1]
                new_name = (
                    f"{rec_row.get('Scene')}-T{rec_row.get('Take')}-{char}{ext}"
                )
                results.append(
                    MatchResult(
                        recorder_idx=rec_idx,
                        transmitter_idx=trans_idx,
                        new_filename=new_name,
                        original_filename=orig,
                        offset_seconds=abs(tc_rec - tc_trans),
                        fps_mismatch=(fps_rec != fps_trans),
                        fps_recorder=fps_rec,
                        fps_transmitter=fps_trans,
                    )
                )
    return results


def copy_and_rename_files(
    matches: list[MatchResult],
    source_dir: str,
    dest_dir: str,
    actual_files_map: Optional[dict[str, str]] = None,
) -> tuple[list[MatchResult], list[str]]:
    """
    Copy matched files to dest_dir with their new names.

    Args:
        actual_files_map: Optional dict of {lowercase_name: actual_name} for
                          case-insensitive lookup. Built from os.listdir if not provided.

    Returns:
        (successful_matches, error_messages)
    """
    os.makedirs(dest_dir, exist_ok=True)
    successful: list[MatchResult] = []
    errors: list[str] = []

    # Build case-insensitive map if not supplied
    if actual_files_map is None:
        actual_files_map = {f.lower(): f for f in os.listdir(source_dir)}

    for match in matches:
        # Resolve the real filename on disk (handles .WAV vs .wav etc.)
        actual_name = actual_files_map.get(match.original_filename.lower())
        if not actual_name:
            errors.append(f"NOT FOUND on disk: {match.original_filename}")
            continue

        src = os.path.join(source_dir, actual_name)
        dst = os.path.join(dest_dir, match.new_filename)

        try:
            shutil.copy2(src, dst)
            successful.append(match)
        except OSError as exc:
            errors.append(f"COPY FAILED: {actual_name}\n  → {exc}")

    return successful, errors


def generate_pdf_report(
    dest_dir: str,
    folder_name: str,
    matches: list[MatchResult],
    recorder_df: pd.DataFrame,
    transmitter_df: pd.DataFrame,
) -> None:
    """Write a landscape A4 PDF sound report to dest_dir."""
    if not FPDF_AVAILABLE:
        return

    try:
        pdf = FPDF(orientation="L", unit="mm", format="A4")
        pdf.add_page()

        # Title
        pdf.set_font("helvetica", "B", 16)
        pdf.cell(0, 10, f"Sound Report for '{folder_name}'", ln=True, align="C")
        pdf.ln(5)

        # Header row
        pdf.set_font("helvetica", "B", 10)
        pdf.set_fill_color(220, 220, 220)
        headers = [
            ("New Filename", 75),
            ("Old Filename", 75),
            ("Transmitter TC", 40),
            ("Userbits", 40),
            ("Offset", 30),
        ]
        for label, width in headers:
            pdf.cell(width, 8, label, border=1, align="C", fill=True)
        pdf.ln()

        # Data rows
        pdf.set_font("helvetica", "", 9)
        for match in matches:
            trans_row = transmitter_df.loc[match.transmitter_idx]
            pdf.cell(75, 7, match.new_filename, border=1)
            pdf.cell(75, 7, match.original_filename, border=1)
            pdf.cell(40, 7, str(trans_row.get("TimeCode", "")), border=1, align="C")
            pdf.cell(40, 7, str(trans_row.get("UserBits", "")), border=1, align="C")
            pdf.cell(30, 7, f"{match.offset_seconds:.2f}s", border=1, align="C")
            pdf.ln()

        report_path = os.path.join(dest_dir, f"{folder_name}.pdf")
        pdf.output(report_path)

    except Exception as exc:
        print(f"PDF generation error: {exc}")


def identify_active_tracks(df: pd.DataFrame) -> tuple[list[str], dict[str, str]]:
    """
    Scan recorder DataFrame for populated track/name columns.

    Returns:
        (active_cols, display_name -> column_name mapping)
    """
    track_cols = [
        c for c in df.columns
        if (c.startswith("Name") or c.startswith("Track")) and c != "Tracks"
    ]
    active_cols: list[str] = []
    mapping: dict[str, str] = {}

    for col in track_cols:
        populated = [n for n in df[col].unique() if n != ""]
        if not populated:
            continue
        active_cols.append(col)
        display = populated[0] if len(populated) == 1 else col
        if display in mapping:
            display = f"{display} ({col})"
        mapping[display] = col

    return active_cols, mapping


# ---------------------------------------------------------------------------
# UI helpers
# ---------------------------------------------------------------------------

def build_treeview(
    parent: tk.Widget,
    columns: tuple[str, ...],
    use_checkboxes: bool = False,
) -> ttk.Treeview:
    """
    Create a scrollable Treeview and configure match-colour tags.

    Args:
        use_checkboxes: If True, returns a CheckboxTreeview with a leading
                        ☐/☑ column.  The header acts as Select All toggle.
    """
    if use_checkboxes:
        tree: ttk.Treeview = CheckboxTreeview(parent, columns=columns, show="headings")
    else:
        tree = ttk.Treeview(parent, columns=columns, show="headings")

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=200 if "Names" in col else 110)

    scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    tree.pack(side="left", expand=True, fill="both")
    scrollbar.pack(side="right", fill="y")

    for i, color in enumerate(MATCH_COLORS):
        tree.tag_configure(f"match_{i}", background=color["bg"], foreground=color["fg"])

    return tree


def scroll_tree_to_item(tree: ttk.Treeview, item_index: int) -> None:
    """Scroll a Treeview so that item_index is near the top."""
    total = len(tree.get_children())
    if total > 0:
        tree.yview_moveto(item_index / total)


# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------

class AudioRenamerApp:
    """Main application window for Z-Match."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.geometry(APP_GEOMETRY)

        # State
        self.recorder_df: Optional[pd.DataFrame] = None
        self.transmitter_df: Optional[pd.DataFrame] = None
        self.matches: list[MatchResult] = []
        self.filtered_recorder_indices: list[int] = []
        self.track_mapping: dict[str, str] = {}
        self.active_track_cols: list[str] = []
        self.transmitter_source_dir: Optional[str] = None

        # Paths
        self.home_dir = str(Path.home())
        self.downloads_dir = str(Path.home() / "Downloads")

        # Tk variables
        self.sound_roll_var = tk.StringVar(value="---")
        self.time_offset_var = tk.DoubleVar(value=DEFAULT_OFFSET_SECONDS)

        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self) -> None:
        self._build_top_bar()
        self._build_main_pane()

    def _build_top_bar(self) -> None:
        """Configuration bar at the top of the window."""
        bar = tk.LabelFrame(self.root, text="Configuration & Recorder Metadata")
        bar.pack(pady=10, fill="x", padx=10)

        # File loaders
        tk.Button(bar, text="Load Recorder CSV", command=lambda: self._load_file(1), width=18).grid(
            row=0, column=0, rowspan=2, padx=2, pady=10
        )
        tk.Button(bar, text="Load Transmitter CSV", command=lambda: self._load_file(2), width=18).grid(
            row=0, column=1, rowspan=2, padx=2
        )

        # Offset spinner
        tk.Label(bar, text="Offset(s):").grid(row=0, column=2, rowspan=2, padx=2)
        tk.Spinbox(bar, from_=0, to=15, increment=1.0, textvariable=self.time_offset_var, width=4).grid(
            row=0, column=3, rowspan=2
        )

        # Sound roll display
        tk.Label(bar, text="Sound Roll:").grid(row=0, column=4, sticky="e", padx=5)
        tk.Label(bar, textvariable=self.sound_roll_var, font=("Arial", 11, "bold"), fg="#D32F2F").grid(
            row=0, column=5, sticky="w"
        )

        # Track selector
        tk.Label(bar, text="Track:").grid(row=1, column=4, sticky="e", padx=5)
        self.track_combo = ttk.Combobox(bar, width=25, state="readonly")
        self.track_combo.grid(row=1, column=5, sticky="w")
        self.track_combo.bind("<<ComboboxSelected>>", self._on_track_selected)

        # Action buttons
        self.match_btn = tk.Button(
            bar, text="Match & Preview", command=self._compare_and_preview,
            bg="#2196F3", fg="white", width=14, state="disabled",
        )
        self.match_btn.grid(row=0, column=6, rowspan=2, padx=10)

        self.rename_btn = tk.Button(
            bar, text="COMMIT & SAVE RENAMED", command=self._execute_rename,
            bg="#4CAF50", fg="white", width=22, state="disabled",
        )
        self.rename_btn.grid(row=0, column=7, rowspan=2, padx=5)

    def _build_main_pane(self) -> None:
        """Vertically split pane: two source lists on top, preview below."""
        paned = tk.PanedWindow(self.root, orient=tk.VERTICAL)
        paned.pack(expand=True, fill="both", padx=10)

        # --- Source lists ---
        list_frame = tk.Frame(paned)

        self.recorder_frame = tk.LabelFrame(list_frame, text="Recorder Metadata")
        self.recorder_frame.pack(side="left", expand=True, fill="both", padx=2)
        self.recorder_tree = build_treeview(
            self.recorder_frame,
            ("TimeCode", "UserBits", "Scene", "Take", "All Tracks / Names"),
            use_checkboxes=True,
        )

        self.transmitter_frame = tk.LabelFrame(list_frame, text="Transmitter Files")
        self.transmitter_frame.pack(side="left", expand=True, fill="both", padx=2)
        self.transmitter_tree = build_treeview(
            self.transmitter_frame, ("TimeCode", "UserBits", "FileName")
        )

        paned.add(list_frame)

        # --- Preview pane ---
        preview_frame = tk.LabelFrame(paned, text="Proposed Filename Changes")
        self.preview_frame_ref = preview_frame

        self.preview_tree = ttk.Treeview(
            preview_frame,
            columns=("Old", "New", "Status", "Offset", "FPS", "Warning"),
            show="headings",
        )
        for col, label, width, anchor, stretch in [
            ("Old",     "Original Filename", 300, "w",      True),
            ("New",     "New Filename",      300, "w",      True),
            ("Status",  "ID",                 50, "center", False),
            ("Offset",  "Offset",             70, "center", False),
            ("FPS",     "FPS (R / T)",        90, "center", False),
            ("Warning", "⚠ FPS",              60, "center", False),
        ]:
            self.preview_tree.heading(col, text=label)
            self.preview_tree.column(col, width=width, anchor=anchor, stretch=stretch)

        # Tag for FPS mismatch rows in the preview
        self.preview_tree.tag_configure(
            "fps_warn", background="#FFF3CD", foreground="#7B4F00"
        )

        self.preview_tree.pack(expand=True, fill="both")
        paned.add(preview_frame)

    # ------------------------------------------------------------------
    # File loading
    # ------------------------------------------------------------------

    def _load_file(self, list_num: int) -> None:
        label = "Recorder" if list_num == 1 else "Transmitter"
        path = filedialog.askopenfilename(
            initialdir=self.home_dir,
            title=f"Select {label} CSV File",
            filetypes=[("CSV files", "*.csv")],
        )
        if not path:
            return

        try:
            df, sound_roll = parse_csv(path, list_num)
        except Exception as exc:
            messagebox.showerror("Error", f"Could not parse file:\n{exc}")
            return

        if list_num == 1:
            self.recorder_df = df
            self.sound_roll_var.set(sound_roll)
            self._setup_track_selector(df)
        else:
            self.transmitter_df = df
            self.transmitter_source_dir = os.path.dirname(path)
            self._populate_tree(
                self.transmitter_tree, df, ["TimeCode", "UserBits", "FileName"]
            )
            self.transmitter_frame.config(text=f"Transmitter Files (Count: {len(df)})")

        self._update_match_button_state()

    # ------------------------------------------------------------------
    # Track selector
    # ------------------------------------------------------------------

    def _setup_track_selector(self, df: pd.DataFrame) -> None:
        self.active_track_cols, self.track_mapping = identify_active_tracks(df)
        display_options = ["Select Track"] + list(self.track_mapping.keys())
        self.track_combo["values"] = display_options
        self.track_combo.current(0)
        self._refresh_recorder_list()

    def _on_track_selected(self, _event=None) -> None:
        self._refresh_recorder_list()
        self._update_match_button_state()

    def _refresh_recorder_list(self) -> None:
        if self.recorder_df is None:
            return

        # Preserve existing check state before clearing the tree.
        # Keys are iid strings; True = checked, False = unchecked.
        # Rows not yet seen default to checked when first inserted.
        saved_checks: dict[str, bool] = {
            iid: self.recorder_tree.is_checked(iid)
            for iid in self.recorder_tree.get_children()
        }

        selection = self.track_combo.get()
        self.recorder_tree.delete(*self.recorder_tree.get_children())
        self.filtered_recorder_indices = []

        if selection == "Select Track":
            for idx, row in self.recorder_df.iterrows():
                names = ", ".join(
                    row.get(c, "") for c in self.active_track_cols if row.get(c, "")
                )
                self.recorder_tree.insert(
                    "", "end",
                    values=[row.get("TimeCode"), row.get("UserBits"), row.get("Scene"), row.get("Take"), names],
                    iid=str(idx),
                )
                self.filtered_recorder_indices.append(idx)
            self.recorder_frame.config(text=f"Recorder Metadata (Total Takes: {len(self.recorder_df)})")
        else:
            col = self.track_mapping[selection]
            count = 0
            for idx, row in self.recorder_df.iterrows():
                char = row.get(col, "")
                if char:
                    self.recorder_tree.insert(
                        "", "end",
                        values=[row.get("TimeCode"), row.get("UserBits"), row.get("Scene"), row.get("Take"), char],
                        iid=str(idx),
                    )
                    self.filtered_recorder_indices.append(idx)
                    count += 1
            self.recorder_frame.config(text=f"Recorder Metadata ({selection} Count: {count})")

        # Restore saved check states. Rows that existed before keep their state;
        # brand-new rows (not in saved_checks) stay checked (insert default).
        for iid in self.recorder_tree.get_children():
            if iid in saved_checks and not saved_checks[iid]:
                self.recorder_tree.uncheck(iid)
        self.recorder_tree._sync_header()

    # ------------------------------------------------------------------
    # Matching
    # ------------------------------------------------------------------

    def _compare_and_preview(self) -> None:
        if self.recorder_df is None or self.transmitter_df is None:
            return

        track_display = self.track_combo.get()
        track_col = self.track_mapping.get(track_display, "")

        # Clear previous results
        self.matches = []
        self.preview_tree.delete(*self.preview_tree.get_children())
        for tree in (self.recorder_tree, self.transmitter_tree):
            for item in tree.get_children():
                tree.item(item, tags=())

        # Only match against rows the user has checked (iids are pandas df indices)
        checked_iids = (
            self.recorder_tree.checked_items()
            if isinstance(self.recorder_tree, CheckboxTreeview)
            else [str(i) for i in self.filtered_recorder_indices]
        )
        # All visible rows are valid candidates — no secondary filter needed
        checked_indices = [int(iid) for iid in checked_iids]

        if not checked_indices:
            messagebox.showwarning("No Rows Selected",
                                   "Please check at least one row in the Recorder list.")
            return

        self.matches = find_matches(
            recorder_df=self.recorder_df,
            transmitter_df=self.transmitter_df,
            recorder_indices=checked_indices,
            track_col=track_col,
            offset_limit=self.time_offset_var.get(),
        )

        first_rec_idx = first_trans_idx = None
        mismatch_count = 0

        # Tag for orange highlight on transmitter tree
        self.transmitter_tree.tag_configure(
            "fps_mismatch", background="#FF8F00", foreground="white"
        )

        for match_num, match in enumerate(self.matches):
            tag = f"match_{match_num % len(MATCH_COLORS)}"
            rec_iid = str(match.recorder_idx)
            trans_iid = str(match.transmitter_idx)

            if self.recorder_tree.exists(rec_iid):
                self.recorder_tree.item(rec_iid, tags=(tag,))

            if self.transmitter_tree.exists(trans_iid):
                # If FPS mismatch, override the colour tag with the warning one
                if match.fps_mismatch:
                    self.transmitter_tree.item(trans_iid, tags=("fps_mismatch",))
                else:
                    self.transmitter_tree.item(trans_iid, tags=(tag,))

            if first_rec_idx is None:
                first_rec_idx = self.recorder_tree.index(rec_iid)
                first_trans_idx = self.transmitter_tree.index(trans_iid)

            fps_label = f"{match.fps_recorder:.2f} / {match.fps_transmitter:.2f}"
            warning_label = "⚠ MISMATCH" if match.fps_mismatch else "✓ OK"
            preview_tag = "fps_warn" if match.fps_mismatch else ""

            self.preview_tree.insert(
                "", "end",
                values=(
                    match.original_filename,
                    match.new_filename,
                    str(match_num + 1),
                    f"{match.offset_seconds:.2f}s",
                    fps_label,
                    warning_label,
                ),
                tags=(preview_tag,) if preview_tag else (),
            )

            if match.fps_mismatch:
                mismatch_count += 1

        # Scroll to first match
        if first_rec_idx is not None:
            scroll_tree_to_item(self.recorder_tree, first_rec_idx)
        if first_trans_idx is not None:
            scroll_tree_to_item(self.transmitter_tree, first_trans_idx)

        count = len(self.matches)
        self.preview_frame_ref.config(text=f"Proposed Filename Changes (Matches Found: {count})")
        self.rename_btn.config(state="normal" if self.matches else "disabled")

        # Build summary message
        summary = f"Found {count} match(es) for: {track_display}"
        if mismatch_count:
            summary += (
                f"\n\n⚠  {mismatch_count} match(es) have a FRAME RATE MISMATCH "
                f"between Recorder and Transmitter.\n\n"
                f"Timecodes were still compared after converting both to seconds, "
                f"but the match may be unreliable. Review highlighted rows (orange in "
                f"Transmitter list, yellow in Preview) before committing the rename."
            )
            messagebox.showwarning("Match Complete — FPS Warning", summary)
        else:
            messagebox.showinfo("Done", summary)

    # ------------------------------------------------------------------
    # Rename / copy
    # ------------------------------------------------------------------

    def _execute_rename(self) -> None:
        if not self.transmitter_source_dir or not os.path.exists(self.transmitter_source_dir):
            messagebox.showerror("Error", "Source folder not found. Please reload Transmitter CSV.")
            return

        track_col = self.track_mapping.get(self.track_combo.get(), "")
        names_seen: list[str] = []
        for match in self.matches:
            name = str(self.recorder_df.loc[match.recorder_idx].get(track_col, "")).strip()
            if name and name not in names_seen:
                names_seen.append(name)

        char_segment = re.sub(r"[^\w\-]", "-", "-".join(names_seen) or "Unknown")
        folder_name = f"SR{self.sound_roll_var.get()}_{char_segment}"

        dest_parent = filedialog.askdirectory(
            title=f"DESTINATION: Create Folder [{folder_name}] In:",
            initialdir=self.downloads_dir,
        )
        if not dest_parent:
            return

        target_dir = os.path.join(dest_parent, folder_name)

        # --- Diagnostic: check for filename mismatches before copying ---
        actual_files = {f.lower(): f for f in os.listdir(self.transmitter_source_dir)}
        looking_for = [m.original_filename for m in self.matches]
        not_on_disk = [fn for fn in looking_for if fn.lower() not in actual_files]

        if not_on_disk:
            sample = "\n".join(not_on_disk[:10])
            disk_sample = "\n".join(list(actual_files.values())[:10])
            messagebox.showwarning(
                "Filename Mismatch Detected",
                f"{len(not_on_disk)} of {len(looking_for)} expected file(s) were NOT found on disk.\n\n"
                f"CSV expects:\n{sample}\n\n"
                f"Files actually on disk (first 10):\n{disk_sample}\n\n"
                f"Check for extension differences (e.g. .WAV vs .wav) or "
                f"extra characters in the FileName column of the Transmitter CSV."
            )

        # Case-insensitive copy — match disk filename regardless of case
        copied, errors = copy_and_rename_files(
            self.matches, self.transmitter_source_dir, target_dir,
            actual_files_map=actual_files,
        )

        # Surface any copy errors immediately so the user knows what failed
        if errors:
            error_summary = "\n".join(errors[:20])  # cap at 20 lines
            if len(errors) > 20:
                error_summary += f"\n...and {len(errors) - 20} more."
            messagebox.showerror(
                "Copy Errors",
                f"{len(errors)} file(s) could not be copied:\n\n{error_summary}\n\n"
                f"Tip: Make sure the audio files are in the same folder as the Transmitter CSV:\n"
                f"{self.transmitter_source_dir}"
            )

        if copied and FPDF_AVAILABLE:
            generate_pdf_report(target_dir, folder_name, copied, self.recorder_df, self.transmitter_df)
        elif not FPDF_AVAILABLE:
            messagebox.showwarning("Missing Library", "fpdf2 not installed — PDF report skipped.")

        if copied:
            # --- Larger completion dialog ---
            dlg = tk.Toplevel(self.root)
            dlg.withdraw()  # hide until positioned to prevent flash in corner
            dlg.title("Commit Complete")
            dlg.resizable(False, False)
            pad = dict(padx=20, pady=10)
            tk.Label(dlg, text="✅  Files Saved Successfully",
                     font=("Arial", 14, "bold"), fg="#2E7D32").pack(**pad)
            tk.Label(dlg, text=f"{len(copied)} of {len(self.matches)} file(s) copied and renamed.",
                     font=("Arial", 11)).pack(padx=20)
            tk.Label(dlg, text="Destination folder:", font=("Arial", 10, "bold")).pack(padx=20, pady=(12,0))
            path_var = tk.StringVar(value=target_dir)
            path_entry = tk.Entry(dlg, textvariable=path_var, font=("Arial", 10),
                                  state="readonly", width=70, relief="sunken")
            path_entry.pack(padx=20, pady=(2,10))
            tk.Button(dlg, text="OK", command=dlg.destroy,
                      bg="#4CAF50", fg="white", font=("Arial", 11), width=12).pack(pady=(0,16))
            # Calculate true size while hidden using reqwidth/reqheight
            # (winfo_width returns 1 for withdrawn windows even after update_idletasks)
            self.root.update_idletasks()
            dlg.update_idletasks()
            root_x = self.root.winfo_rootx()
            root_y = self.root.winfo_rooty()
            root_w = self.root.winfo_width()
            root_h = self.root.winfo_height()
            dlg_w  = dlg.winfo_reqwidth()
            dlg_h  = dlg.winfo_reqheight()
            mx = root_x + (root_w - dlg_w) // 2
            my = root_y + (root_h - dlg_h) // 2
            dlg.geometry(f"{dlg_w}x{dlg_h}+{mx}+{my}")
            dlg.deiconify()  # reveal already-positioned window
            dlg.grab_set()   # grab after deiconify so it's visible when modal
            self.root.wait_window(dlg)
        else:
            messagebox.showerror("Nothing Copied",
                f"No files were copied. The audio files need to be in the same folder as the Transmitter CSV:\n\n"
                f"{self.transmitter_source_dir}\n\n"
                f"If your audio files are elsewhere, reload the Transmitter CSV from that folder.")

        if copied:
            self.rename_btn.config(state="disabled")

            # Reset transmitter tree — clear all match highlights and tags
            for iid in self.transmitter_tree.get_children():
                self.transmitter_tree.item(iid, tags=())

            # Reset preview pane
            self.preview_tree.delete(*self.preview_tree.get_children())
            self.preview_frame_ref.config(text="Proposed Filename Changes")

            # Hard reset recorder checkboxes — wipe state then repopulate all checked
            self.recorder_tree.reset()
            self._refresh_recorder_list()
            self.recorder_tree.set_all(True)

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _update_match_button_state(self) -> None:
        ready = (
            self.recorder_df is not None
            and self.transmitter_df is not None
            and self.track_combo.get() not in ("", "Select Track")
        )
        self.match_btn.config(state="normal" if ready else "disabled")

    @staticmethod
    def _populate_tree(tree: ttk.Treeview, df: pd.DataFrame, cols: list[str]) -> None:
        tree.delete(*tree.get_children())
        for idx, row in df.iterrows():
            tree.insert("", "end", values=[row.get(c, "") for c in cols], iid=str(idx))


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    AudioRenamerApp(root)
    root.mainloop()
