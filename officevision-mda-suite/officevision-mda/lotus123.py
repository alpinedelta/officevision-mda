#!/usr/bin/env python3
"""
Lotus 1-2-3 Replica with Google Sheets Sync for IBM OfficeVision/400 (OV/400).
Designed specifically for the 90x43 Monochrome MDA Terminal (TERM=vt100).

Features:
  - Authentic Lotus 1-2-3 3-line Control Panel with bold reverse-video Mode Indicator
    (READY, VALUE, LABEL, EDIT, MENU, WAIT, FILES, GRAPH, HELP)
  - Classic Lotus 1-2-3 Two-Line Slash (/) Command Menu with single-letter hotkeys:
    Worksheet, Range, Copy, Move, File, Print, Graph, Data, Sheets, Quit
  - Multi-sheet Release 3 tabs (A:Budget, B:Summary, ...)
  - Dual Lotus 1-2-3 (@SUM(A1..B5), +B2*C2) & Google Sheets (=SUM(A1:B5)) formulas
  - Direct Google Sheets Sync (Pull / Push / Browse / Link / AutoSync) sharing
    OV/400's Google OAuth & rclone configuration (CFGDRV)
  - Monochrome MDA ASCII Bar & Line Charts (/Graph)
  - CSV / multi-sheet XLSX Import & Export + Print to OV/400 Word Processor
"""

import copy
import curses
import datetime
import os
import sys
from typing import Any, Callable, Dict, List, Optional, Tuple

from calc_engine import (
    MAX_COLS,
    MAX_ROWS,
    Workbook,
    Worksheet,
    col_to_idx,
    export_workbook_to_xlsx_bytes,
    export_worksheet_to_ascii,
    export_worksheet_to_csv,
    format_cell_display,
    format_cell_ref,
    idx_to_col,
    import_csv_to_worksheet,
    import_xlsx_bytes_to_workbook,
    is_formula_raw,
    parse_cell_ref,
    parse_range_str,
    shift_formula,
    try_parse_number,
)
from gdrive import GoogleDriveManager
from gsheets import GoogleSheetsManager, SheetEntry, extract_spreadsheet_id
from wordproc import GoogleDriveSetupScreen, draw_5250_box, read_key, safe_addstr


# =============================================================================
# Classic Lotus 1-2-3 Two-Line Slash Menu Specification
# =============================================================================
# Each node: (label, hotkey, description, children_or_action_str)

MENU_TREE: List[Tuple[str, str, str, Any]] = [
    (
        "Worksheet",
        "W",
        "Global  Insert  Delete  Column  Rename-Sheet  Titles  Erase  Status",
        [
            (
                "Global",
                "G",
                "Set default worksheet format, column width, or recalculation",
                [
                    ("Format", "F", "Set global number display format", "ws_global_fmt"),
                    ("Column-Width", "C", "Set default column width (3-48)", "ws_global_width"),
                    ("Recalc", "R", "Toggle Automatic vs Manual recalculation", "ws_toggle_recalc"),
                ],
            ),
            (
                "Insert",
                "I",
                "Insert a blank Row, Column, or new Sheet tab",
                [
                    ("Row", "R", "Insert blank row above current cell", "ws_ins_row"),
                    ("Column", "C", "Insert blank column left of current cell", "ws_ins_col"),
                    ("Sheet", "S", "Add a new worksheet tab to workbook", "ws_ins_sheet"),
                ],
            ),
            (
                "Delete",
                "D",
                "Delete current Row, Column, or Sheet tab",
                [
                    ("Row", "R", "Delete entire current row", "ws_del_row"),
                    ("Column", "C", "Delete entire current column", "ws_del_col"),
                    ("Sheet", "S", "Delete current worksheet tab", "ws_del_sheet"),
                ],
            ),
            (
                "Column",
                "C",
                "Set-Width  Reset-Width  Auto-Fit current column",
                [
                    ("Set-Width", "S", "Set width of current column", "ws_col_set"),
                    ("Reset-Width", "R", "Reset current column to global default width", "ws_col_reset"),
                    ("Auto-Fit", "A", "Auto-fit current column width to contents", "ws_col_autofit"),
                ],
            ),
            ("Rename-Sheet", "R", "Rename the active worksheet tab", "ws_rename_sheet"),
            (
                "Titles",
                "T",
                "Freeze top rows or left columns while scrolling",
                [
                    ("Both", "B", "Freeze rows above and columns left of cursor", "ws_titles_both"),
                    ("Horizontal", "H", "Freeze rows above cursor", "ws_titles_horiz"),
                    ("Vertical", "V", "Freeze columns left of cursor", "ws_titles_vert"),
                    ("Clear", "C", "Unfreeze all title rows and columns", "ws_titles_clear"),
                ],
            ),
            ("Erase", "E", "Erase all cells in the active worksheet", "ws_erase"),
            ("Status", "S", "Display workbook & Google Sheets sync status", "ws_status"),
        ],
    ),
    (
        "Range",
        "R",
        "Format  Label  Erase  Value  Transpose  Name",
        [
            (
                "Format",
                "F",
                "General  Fixed  Currency  Percent  Comma  Scientific  Date  Reset",
                [
                    ("General", "G", "General numeric format (G)", "rng_fmt_G"),
                    ("Fixed", "F", "Fixed decimal places (F2)", "rng_fmt_F2"),
                    ("Currency", "C", "Currency with $ and commas (C2)", "rng_fmt_C2"),
                    ("Dollars", "$", "Whole-dollar currency (C0)", "rng_fmt_C0"),
                    ("Percent", "P", "Percentage format (P1)", "rng_fmt_P1"),
                    ("Comma", ",", "Comma thousands separator (,2)", "rng_fmt_,2"),
                    ("Scientific", "S", "Scientific notation (S2)", "rng_fmt_S2"),
                    ("Date", "D", "Date display format (D)", "rng_fmt_D"),
                    ("Reset", "R", "Reset range to global worksheet format", "rng_fmt_reset"),
                ],
            ),
            (
                "Label",
                "L",
                "Align text labels: Left (')  Right (\")  Center (^)",
                [
                    ("Left", "L", "Left-align labels in range (')", "rng_lbl_left"),
                    ("Right", "R", 'Right-align labels in range (")', "rng_lbl_right"),
                    ("Center", "C", "Center-align labels in range (^)", "rng_lbl_center"),
                ],
            ),
            ("Erase", "E", "Erase contents of a cell or range", "rng_erase"),
            ("Value", "V", "Convert formulas in range to static values", "rng_value"),
            ("Transpose", "T", "Copy and transpose rows/columns of a range", "rng_transpose"),
            (
                "Name",
                "N",
                "Create, Delete, or List named ranges",
                [
                    ("Create", "C", "Assign a name to a range for formulas", "rng_name_create"),
                    ("Delete", "D", "Delete a named range", "rng_name_delete"),
                    ("List", "L", "View all named ranges in workbook", "rng_name_list"),
                ],
            ),
        ],
    ),
    ("Copy", "C", "Copy a cell or range (shifts relative formula references)", "op_copy"),
    ("Move", "M", "Move a cell or range to another location", "op_move"),
    (
        "File",
        "F",
        "Retrieve  Save  New  Import  Export  List",
        [
            ("Retrieve", "R", "Open a local workbook or Google Sheet", "file_list"),
            ("Save", "S", "Save workbook locally (+ sync if linked to Google Sheets)", "file_save"),
            ("New", "N", "Create a new blank Lotus 1-2-3 workbook", "file_new"),
            ("Import", "I", "Import a .csv or .xlsx spreadsheet file", "file_import"),
            (
                "Export",
                "E",
                "Export workbook/sheet to .csv or .xlsx",
                [
                    ("CSV", "C", "Export active sheet to CSV file", "file_export_csv"),
                    ("XLSX", "X", "Export entire multi-tab workbook to .xlsx", "file_export_xlsx"),
                ],
            ),
            ("List", "L", "Browse local workbooks and Google Sheets", "file_list"),
        ],
    ),
    (
        "Print",
        "P",
        "Preview  OV400-Doc  Text-File",
        [
            ("Preview", "P", "Preview formatted ASCII worksheet table", "print_preview"),
            ("OV400-Doc", "O", "Send ASCII table to OV/400 Word Processing documents", "print_ov400"),
            ("Text-File", "T", "Export formatted ASCII table to a .txt file", "print_txt"),
        ],
    ),
    (
        "Graph",
        "G",
        "Bar-Chart  Line-Chart  (Monochrome MDA ASCII Graphs)",
        [
            ("Bar", "B", "Render a Monochrome ASCII Bar Chart from a range", "graph_bar"),
            ("Line", "L", "Render a Monochrome ASCII Line Chart from a range", "graph_line"),
        ],
    ),
    (
        "Data",
        "D",
        "Fill  Sort",
        [
            ("Fill", "F", "Fill a range with a Start/Step/Stop numeric series", "data_fill"),
            ("Sort", "S", "Sort a range ascending or descending by column", "data_sort"),
        ],
    ),
    (
        "Sheets",
        "S",
        "Pull  Push  Browse  Link  Config  AutoSync  (Google Sheets)",
        [
            ("Pull", "P", "Pull latest tabs & formulas from linked Google Sheet", "gs_pull"),
            ("Push", "U", "Push workbook to Google Sheets (creates if unlinked)", "gs_push"),
            ("Browse", "B", "Browse & open spreadsheets from your Google Drive", "gs_browse"),
            ("Link", "L", "Link current workbook to an existing Google Sheet URL/ID", "gs_link"),
            ("Config", "C", "Configure Google Drive & Sheets OAuth / rclone (CFGDRV)", "gs_config"),
            ("AutoSync", "A", "Toggle automatic push to Google Sheets on Save", "gs_autosync"),
        ],
    ),
    (
        "Quit",
        "Q",
        "End Lotus 1-2-3 session and return to OfficeVision/400",
        [
            ("No", "N", "Do not quit; return to READY mode", "quit_no"),
            ("Yes", "Y", "Quit Lotus 1-2-3 and return to OV/400 Main Menu", "quit_yes"),
        ],
    ),
]


class Lotus123App:
    """Classic Lotus 1-2-3 Release 2.x/3.x TUI with Google Sheets Sync for MDA."""

    def __init__(
        self,
        gdrive: GoogleDriveManager,
        system_name: str = "AS400",
        current_user: str = "QUSER",
        launch_external_cb: Optional[Callable] = None,
    ):
        self.gdrive = gdrive
        self.gsheets = GoogleSheetsManager(gdrive)
        self.system_name = system_name
        self.current_user = current_user
        self.launch_external_cb = launch_external_cb

        # Active workbook & entry
        entries = self.gsheets.list_entries()
        if entries:
            self.current_entry: SheetEntry = entries[0]
            self.wb, self.status_message = self.gsheets.load_workbook(self.current_entry)
        else:
            self.current_entry, self.wb = self.gsheets.create_workbook_entry("UNTITLED")
            self.status_message = "Created new workbook 'UNTITLED'."

        # Cursor & viewport
        self.cur_row: int = 0
        self.cur_col: int = 0
        self.view_top: int = 0
        self.view_left: int = 0

        # Mode state: READY, VALUE, LABEL, EDIT, MENU, HELP
        self.mode: str = "READY"
        self.edit_buf: str = ""
        self.edit_pos: int = 0

        # Slash menu stack: list of (items_list, selected_index)
        self.menu_stack: List[Tuple[List[Tuple[str, str, str, Any]], int]] = []

        self.running: bool = True

    @property
    def ws(self) -> Worksheet:
        return self.wb.active_sheet

    # =========================================================================
    # Main Event Loop
    # =========================================================================
    def run(self, stdscr):
        self.running = True
        self.mode = "READY"
        self.menu_stack = []

        while self.running:
            max_y, max_x = stdscr.getmaxyx()
            self._ensure_cursor_visible(max_y, max_x)
            stdscr.erase()
            self._draw_screen(stdscr, max_y, max_x)
            stdscr.refresh()

            ch = read_key(stdscr, restore_timeout=500)
            if ch == -1:
                continue

            if self.mode == "MENU":
                self._handle_menu_key(stdscr, ch)
            elif self.mode in ("VALUE", "LABEL", "EDIT"):
                self._handle_edit_key(stdscr, ch)
            else:
                self._handle_ready_key(stdscr, ch, max_y, max_x)

    # =========================================================================
    # Viewport Management
    # =========================================================================
    def _grid_rows_visible(self, max_y: int) -> int:
        # Rows 0-2: control panel, Row 3: column headers, Last 2 rows: status + fkeys
        return max(1, max_y - 6)

    def _visible_cols(self, max_x: int) -> List[int]:
        """Computes which columns fit across `max_x` (gutter=4 chars)."""
        gutter = 4
        avail = max_x - gutter
        cols: List[int] = []
        used = 0
        # Frozen columns first
        fcols = min(self.ws.freeze_cols, 10)
        for c in range(fcols):
            w = self.ws.get_col_width(c)
            if used + w > avail and cols:
                break
            cols.append(c)
            used += w
        # Scrollable columns starting at max(view_left, fcols)
        c = max(self.view_left, fcols)
        while c < MAX_COLS and used < avail:
            w = self.ws.get_col_width(c)
            if used + w > avail and cols:
                # Allow partial last column if at least 4 chars remain
                if avail - used >= 4:
                    cols.append(c)
                break
            cols.append(c)
            used += w
            c += 1
        return cols

    def _ensure_cursor_visible(self, max_y: int, max_x: int):
        self.cur_row = max(0, min(MAX_ROWS - 1, self.cur_row))
        self.cur_col = max(0, min(MAX_COLS - 1, self.cur_col))

        frows = min(self.ws.freeze_rows, 10)
        fcols = min(self.ws.freeze_cols, 10)
        vis_rows = self._grid_rows_visible(max_y)
        scroll_rows = max(1, vis_rows - frows)

        if self.cur_row >= frows:
            if self.cur_row < self.view_top:
                self.view_top = self.cur_row
            elif self.cur_row >= self.view_top + scroll_rows:
                self.view_top = self.cur_row - scroll_rows + 1
            self.view_top = max(frows, self.view_top)

        if self.cur_col >= fcols:
            if self.cur_col < self.view_left:
                self.view_left = self.cur_col
            else:
                # Advance view_left until cur_col fits completely
                gutter = 4
                avail = max_x - gutter
                frozen_w = sum(self.ws.get_col_width(c) for c in range(fcols))
                scroll_avail = max(8, avail - frozen_w)
                while self.view_left < self.cur_col:
                    total = sum(
                        self.ws.get_col_width(c) for c in range(self.view_left, self.cur_col + 1)
                    )
                    if total <= scroll_avail:
                        break
                    self.view_left += 1
            self.view_left = max(fcols, self.view_left)

    # =========================================================================
    # Monochrome MDA Lotus 1-2-3 Screen Renderer
    # =========================================================================
    def _draw_screen(self, win, max_y: int, max_x: int):
        try:
            curses.curs_set(1 if self.mode in ("VALUE", "LABEL", "EDIT") else 0)
        except curses.error:
            pass

        # ---- ROW 0: Cell Address, Format/Width, Raw Content, Mode Badge ----
        addr = format_cell_ref(self.cur_row, self.cur_col)
        cell = self.ws.get_cell(self.cur_row, self.cur_col)
        meta_parts = [f"{addr}:"]
        if cell and cell.fmt:
            meta_parts.append(f"({cell.fmt})")
        if self.cur_col in self.ws.col_widths:
            meta_parts.append(f"[W{self.ws.col_widths[self.cur_col]}]")
        raw_disp = cell.raw if cell else ""
        left_info = " ".join(meta_parts) + (" " + raw_disp if raw_disp else "")

        badge = f" {self.mode:<5} "
        badge_x = max(0, max_x - len(badge) - 1)
        safe_addstr(win, 0, 0, left_info[: max(0, badge_x - 1)], curses.A_BOLD)
        safe_addstr(win, 0, badge_x, badge, curses.A_REVERSE | curses.A_BOLD)

        # ---- ROW 1 & ROW 2: Slash Menu OR Edit Buffer OR Sheet Tabs ----
        cursor_y, cursor_x = 1, 0
        if self.mode == "MENU" and self.menu_stack:
            items, sel_idx = self.menu_stack[-1]
            x = 1
            for idx, (lbl, _hk, _desc, _sub) in enumerate(items):
                if idx == sel_idx:
                    safe_addstr(win, 1, x, f" {lbl} ", curses.A_REVERSE | curses.A_BOLD)
                else:
                    # Bold the first letter (hotkey) like Lotus 1-2-3
                    safe_addstr(win, 1, x, " ")
                    safe_addstr(win, 1, x + 1, lbl[0], curses.A_BOLD | curses.A_UNDERLINE)
                    safe_addstr(win, 1, x + 2, lbl[1:] + " ")
                x += len(lbl) + 2
                if x >= max_x - 2:
                    break

            _, _, sel_desc, _ = items[sel_idx]
            safe_addstr(win, 2, 1, sel_desc[: max_x - 2])

        elif self.mode in ("VALUE", "LABEL", "EDIT"):
            prompt = f"{self.mode}> "
            safe_addstr(win, 1, 0, prompt, curses.A_BOLD)
            avail = max(10, max_x - len(prompt) - 1)
            scroll_off = max(0, self.edit_pos - avail + 2)
            visible_buf = self.edit_buf[scroll_off : scroll_off + avail]
            safe_addstr(win, 1, len(prompt), visible_buf)
            cursor_y = 1
            cursor_x = len(prompt) + (self.edit_pos - scroll_off)
            safe_addstr(
                win,
                2,
                1,
                "Enter/Arrows=Accept   Esc=Cancel   Formulas: =SUM(A1:B5) or @SUM(A1..B5)",
            )
        else:
            # Row 1: Workbook & Google Sheets Sync Banner
            sync_badge = self._cloud_badge()
            auto_str = "ON" if self.gdrive.config.get("auto_sync_on_save", True) else "OFF"
            banner = f"1-2-3 | {self.current_entry.name} | {sync_badge} | AutoSync:{auto_str}"
            safe_addstr(win, 1, 1, banner[: max_x - 2])

            # Row 2: Lotus 1-2-3 Release 3 Sheet Tabs Bar
            tx = 1
            for i, sh in enumerate(self.wb.sheets):
                letter = chr(ord("A") + (i % 26))
                tab_txt = f" {letter}:{sh.name} "
                if i == self.wb.active_sheet_idx:
                    safe_addstr(win, 2, tx, tab_txt, curses.A_REVERSE | curses.A_BOLD)
                else:
                    safe_addstr(win, 2, tx, tab_txt, curses.A_UNDERLINE)
                tx += len(tab_txt) + 1
                if tx >= max_x - 16:
                    break
            tab_hint = "F6=NextTab"
            if tx + len(tab_hint) + 2 < max_x:
                safe_addstr(win, 2, max_x - len(tab_hint) - 1, tab_hint)

        # ---- ROW 3: Reverse-Video Column Header Frame ----
        sheet_ltr = chr(ord("A") + (self.wb.active_sheet_idx % 26))
        safe_addstr(win, 3, 0, f" {sheet_ltr}  ", curses.A_REVERSE | curses.A_BOLD)

        vis_cols = self._visible_cols(max_x)
        col_x_map: List[Tuple[int, int, int]] = []  # (col_idx, start_x, render_width)
        cx = 4
        for col_idx in vis_cols:
            w = self.ws.get_col_width(col_idx)
            rw = min(w, max_x - cx)
            if rw <= 0:
                break
            col_x_map.append((col_idx, cx, rw))
            hdr = idx_to_col(col_idx).center(rw)[:rw]
            attr = (
                curses.A_REVERSE | curses.A_BOLD | curses.A_UNDERLINE
                if col_idx == self.cur_col
                else curses.A_REVERSE
            )
            safe_addstr(win, 3, cx, hdr, attr)
            cx += rw
        if cx < max_x:
            safe_addstr(win, 3, cx, " " * (max_x - cx), curses.A_REVERSE)

        # ---- ROWS 4 .. max_y-3: Spreadsheet Grid ----
        vis_rows_count = self._grid_rows_visible(max_y)
        frows = min(self.ws.freeze_rows, 10)
        row_indices: List[int] = []
        for r in range(min(frows, vis_rows_count)):
            row_indices.append(r)
        rem_rows = vis_rows_count - len(row_indices)
        start_scroll_r = max(self.view_top, frows)
        for i in range(rem_rows):
            r = start_scroll_r + i
            if r < MAX_ROWS:
                row_indices.append(r)

        for screen_offset, r in enumerate(row_indices):
            sy = 4 + screen_offset
            r_attr = curses.A_REVERSE | curses.A_BOLD if r == self.cur_row else curses.A_REVERSE
            safe_addstr(win, sy, 0, f"{r + 1:>3} ", r_attr)

            # Build cell text with classic Lotus 1-2-3 label overflow into adjacent empty cells
            overflow_text = ""
            for col_idx, x_pos, rw in col_x_map:
                c_obj = self.ws.get_cell(r, col_idx)
                is_cur = (r == self.cur_row and col_idx == self.cur_col)
                c_attr = curses.A_REVERSE | curses.A_BOLD if is_cur else curses.A_NORMAL

                if c_obj is not None and (c_obj.raw != "" or c_obj.value is not None):
                    overflow_text = ""
                    # Check for left-aligned text overflow
                    if (
                        not c_obj.error
                        and isinstance(c_obj.value, str)
                        and not c_obj.raw.startswith(("\\", '"', "^"))
                        and len(c_obj.value) > rw
                    ):
                        disp = c_obj.value[:rw]
                        overflow_text = c_obj.value[rw:]
                    else:
                        disp = format_cell_display(c_obj, self.ws.global_format, rw)
                elif overflow_text:
                    disp = overflow_text[:rw].ljust(rw)
                    overflow_text = overflow_text[rw:]
                else:
                    disp = " " * rw

                # Visual freeze indicator underline on last frozen row
                if frows > 0 and r == frows - 1 and not is_cur:
                    c_attr |= curses.A_UNDERLINE

                safe_addstr(win, sy, x_pos, disp, c_attr)

        # ---- BOTTOM STATUS BAR (max_y - 2) ----
        status_y = max_y - 2
        now_str = datetime.datetime.now().strftime("%d-%b-%y %H:%M")

        flags: List[str] = []
        if self.wb.dirty:
            flags.append("MOD")
        if self.wb.has_circular_ref:
            flags.append("CIRC")
        if not self.wb.auto_recalc:
            flags.append("CALC")
        if self.current_entry.gsheet_id:
            flags.append("GSYNC")
        if self.wb.undo_stack:
            flags.append("UNDO")
        flags_str = " ".join(f"[{fl}]" for fl in flags)

        left_status = f" {now_str} | {self.status_message}"
        right_x = max(1, max_x - len(flags_str) - 1)
        safe_addstr(win, status_y, 0, left_status[: max(0, right_x - 1)], curses.A_BOLD)
        if flags_str:
            safe_addstr(win, status_y, right_x, flags_str, curses.A_REVERSE)

        # ---- FUNCTION KEY LEGEND (max_y - 1) ----
        fkey_line = " / Menu  F1=Help  F2=Edit  F3=Quit  F4=Files/Sheets  F5=GoTo  F6=Tab  F9=Calc  F10=Sync "
        safe_addstr(win, max_y - 1, 0, fkey_line[: max_x - 1])

        if self.mode in ("VALUE", "LABEL", "EDIT"):
            try:
                win.move(cursor_y, min(cursor_x, max_x - 2))
            except curses.error:
                pass

    def _cloud_badge(self) -> str:
        backend = self.gsheets.get_backend()
        if backend == "local":
            return "GSheets: LOCAL (F10=Setup)"
        if self.current_entry.gsheet_id:
            st = "MODIFIED" if (self.wb.dirty or self.current_entry.dirty) else "SYNCED"
            return f"GSheets: {st}"
        return f"GSheets: UNLINKED ({backend.upper()})"

    # =========================================================================
    # READY Mode Key Handling
    # =========================================================================
    def _handle_ready_key(self, stdscr, ch: int, max_y: int, max_x: int):
        # Slash '/' opens classic Lotus 1-2-3 Command Menu!
        if ch == ord("/"):
            self.mode = "MENU"
            self.menu_stack = [(MENU_TREE, 0)]
            return

        # F1 = Help
        if ch == curses.KEY_F1:
            self._modal_help(stdscr)
            return

        # F2 = Edit current cell
        if ch == curses.KEY_F2:
            self.mode = "EDIT"
            self.edit_buf = self.ws.get_raw(self.cur_row, self.cur_col)
            self.edit_pos = len(self.edit_buf)
            return

        # F3 = Quit (prompt if modified)
        if ch == curses.KEY_F3:
            self._confirm_quit(stdscr)
            return

        # F4 = Workbook & Google Sheets Browser
        if ch == curses.KEY_F4:
            self._screen_workbook_browser(stdscr)
            return

        # F5 = GoTo cell / named range
        if ch == curses.KEY_F5:
            self._prompt_goto(stdscr)
            return

        # F6 = Switch to next sheet tab
        if ch == curses.KEY_F6:
            if len(self.wb.sheets) > 1:
                self.wb.active_sheet_idx = (self.wb.active_sheet_idx + 1) % len(self.wb.sheets)
                self.cur_row, self.cur_col, self.view_top, self.view_left = 0, 0, 0, 0
                self.status_message = f"Switched to tab '{self.ws.name}'."
            else:
                self.status_message = "Only 1 tab. Use /Worksheet Insert Sheet (/WIS) to add tabs."
            return

        # F9 = Recalculate
        if ch == curses.KEY_F9:
            self.wb.recalculate()
            self.status_message = "Worksheet recalculated."
            return

        # F10 = Quick Google Sheets Sync Menu
        if ch == curses.KEY_F10:
            self._quick_gsheets_sync(stdscr)
            return

        # Ctrl+S = Quick Save (+ AutoSync)
        if ch == 19:
            self._action_save(stdscr)
            return

        # Ctrl+Z = Undo, Ctrl+Y = Redo
        if ch == 26:
            if self.wb.undo():
                self.status_message = "Undo complete."
            else:
                self.status_message = "Nothing to undo."
            return
        if ch == 25:
            if self.wb.redo():
                self.status_message = "Redo complete."
            else:
                self.status_message = "Nothing to redo."
            return

        # Navigation keys
        vis_rows = self._grid_rows_visible(max_y)
        if ch == curses.KEY_UP:
            self.cur_row = max(0, self.cur_row - 1)
            return
        if ch == curses.KEY_DOWN:
            self.cur_row = min(MAX_ROWS - 1, self.cur_row + 1)
            return
        if ch == curses.KEY_LEFT:
            self.cur_col = max(0, self.cur_col - 1)
            return
        if ch == curses.KEY_RIGHT:
            self.cur_col = min(MAX_COLS - 1, self.cur_col + 1)
            return
        if ch == 9:  # Tab -> jump right 1 page (or 4 cols)
            self.cur_col = min(MAX_COLS - 1, self.cur_col + 4)
            return
        if ch == curses.KEY_BTAB:  # Shift+Tab -> jump left 4 cols
            self.cur_col = max(0, self.cur_col - 4)
            return
        if ch == curses.KEY_PPAGE:
            self.cur_row = max(0, self.cur_row - max(1, vis_rows - 2))
            return
        if ch == curses.KEY_NPAGE:
            self.cur_row = min(MAX_ROWS - 1, self.cur_row + max(1, vis_rows - 2))
            return
        if ch == curses.KEY_HOME:
            self.cur_row, self.cur_col = 0, 0
            self.view_top = self.ws.freeze_rows
            self.view_left = self.ws.freeze_cols
            return
        if ch == curses.KEY_END:
            max_r, max_c = self.ws.get_bounds()
            if max_r >= 0:
                self.cur_row, self.cur_col = max_r, max_c
            return

        # Delete / Backspace clears current cell
        if ch in (curses.KEY_DC, curses.KEY_BACKSPACE, 127, 8):
            if self.ws.get_raw(self.cur_row, self.cur_col):
                self.wb.push_undo()
                self.ws.set_raw(self.cur_row, self.cur_col, "")
                self.wb.dirty = True
                if self.wb.auto_recalc:
                    self.wb.recalculate()
                self.status_message = f"Cleared {format_cell_ref(self.cur_row, self.cur_col)}."
            return

        # Enter in READY mode enters EDIT mode on current cell
        if ch in (10, 13, curses.KEY_ENTER):
            self.mode = "EDIT"
            self.edit_buf = self.ws.get_raw(self.cur_row, self.cur_col)
            self.edit_pos = len(self.edit_buf)
            return

        # Direct printable typing -> VALUE or LABEL mode (authentic Lotus 1-2-3!)
        if 32 <= ch <= 126:
            c_char = chr(ch)
            if c_char in "0123456789.+-=(@$":
                self.mode = "VALUE"
            else:
                self.mode = "LABEL"
            self.edit_buf = c_char
            self.edit_pos = 1

    # =========================================================================
    # VALUE / LABEL / EDIT Mode Key Handling
    # =========================================================================
    def _commit_edit(self):
        self.wb.push_undo()
        self.ws.set_raw(self.cur_row, self.cur_col, self.edit_buf)
        self.wb.dirty = True
        if self.wb.auto_recalc:
            self.wb.recalculate()
        self.mode = "READY"
        self.edit_buf = ""
        self.edit_pos = 0

    def _handle_edit_key(self, stdscr, ch: int):
        # Esc or F12 cancels editing
        if ch in (27, curses.KEY_F12):
            self.mode = "READY"
            self.edit_buf = ""
            self.edit_pos = 0
            self.status_message = "Edit cancelled."
            return

        # Enter commits and moves down 1 row (or stays? Lotus keeps or moves; moving down on Enter is super smooth!)
        if ch in (10, 13, curses.KEY_ENTER):
            self._commit_edit()
            return

        # Up / Down arrows commit immediately and move cursor (classic Lotus 1-2-3 rapid entry!)
        if ch == curses.KEY_UP:
            self._commit_edit()
            self.cur_row = max(0, self.cur_row - 1)
            return
        if ch == curses.KEY_DOWN:
            self._commit_edit()
            self.cur_row = min(MAX_ROWS - 1, self.cur_row + 1)
            return
        if ch == 9:  # Tab commits and moves right
            self._commit_edit()
            self.cur_col = min(MAX_COLS - 1, self.cur_col + 1)
            return

        # In-line cursor navigation
        if ch == curses.KEY_LEFT:
            self.edit_pos = max(0, self.edit_pos - 1)
            return
        if ch == curses.KEY_RIGHT:
            self.edit_pos = min(len(self.edit_buf), self.edit_pos + 1)
            return
        if ch == curses.KEY_HOME:
            self.edit_pos = 0
            return
        if ch == curses.KEY_END:
            self.edit_pos = len(self.edit_buf)
            return
        if ch in (curses.KEY_BACKSPACE, 127, 8):
            if self.edit_pos > 0:
                self.edit_buf = self.edit_buf[: self.edit_pos - 1] + self.edit_buf[self.edit_pos :]
                self.edit_pos -= 1
            return
        if ch == curses.KEY_DC:
            if self.edit_pos < len(self.edit_buf):
                self.edit_buf = self.edit_buf[: self.edit_pos] + self.edit_buf[self.edit_pos + 1 :]
            return

        if 32 <= ch <= 126:
            self.edit_buf = self.edit_buf[: self.edit_pos] + chr(ch) + self.edit_buf[self.edit_pos :]
            self.edit_pos += 1

    # =========================================================================
    # Classic Lotus 1-2-3 Slash Menu Navigation
    # =========================================================================
    def _handle_menu_key(self, stdscr, ch: int):
        if not self.menu_stack:
            self.mode = "READY"
            return

        items, sel_idx = self.menu_stack[-1]

        # Esc / F12 steps back one menu level (or exits MENU to READY)
        if ch in (27, curses.KEY_F12):
            self.menu_stack.pop()
            if not self.menu_stack:
                self.mode = "READY"
            return

        if ch in (curses.KEY_LEFT, curses.KEY_BTAB):
            self.menu_stack[-1] = (items, (sel_idx - 1) % len(items))
            return
        if ch in (curses.KEY_RIGHT, 9, ord(" ")):
            self.menu_stack[-1] = (items, (sel_idx + 1) % len(items))
            return

        if ch in (10, 13, curses.KEY_ENTER):
            self._activate_menu_item(stdscr, items[sel_idx])
            return

        # Single-letter Lotus hotkey (e.g. /FS, /FR, /WCS, /QY)
        if 32 <= ch <= 126:
            key_u = chr(ch).upper()
            for idx, item in enumerate(items):
                if item[1].upper() == key_u:
                    self.menu_stack[-1] = (items, idx)
                    self._activate_menu_item(stdscr, item)
                    return

    def _activate_menu_item(self, stdscr, item: Tuple[str, str, str, Any]):
        _label, _hk, _desc, target = item
        if isinstance(target, list):
            self.menu_stack.append((target, 0))
            return
        # Leaf command: close menu and execute
        self.menu_stack = []
        self.mode = "READY"
        self._dispatch_action(stdscr, str(target))

    # =========================================================================
    # Slash Menu Action Dispatcher
    # =========================================================================
    def _dispatch_action(self, stdscr, action: str):
        cur_ref = format_cell_ref(self.cur_row, self.cur_col)

        # ---- Quit ----
        if action == "quit_no":
            self.status_message = "Returned to READY mode."
            return
        if action == "quit_yes":
            self._confirm_quit(stdscr)
            return

        # ---- Worksheet Commands ----
        if action == "ws_global_fmt":
            val = self._prompt_input(
                stdscr,
                "Global Format (G, F2, C0, C2, P1, ,2, S2, D):",
                self.ws.global_format,
            )
            if val is not None and val.strip():
                self.wb.push_undo()
                self.ws.global_format = val.strip().upper()
                self.wb.dirty = True
                self.status_message = f"Global format set to {self.ws.global_format}."
            return

        if action == "ws_global_width":
            val = self._prompt_input(
                stdscr,
                "Global Default Column Width (3-48):",
                str(self.ws.default_col_width),
            )
            if val and val.strip().isdigit():
                self.ws.default_col_width = max(3, min(48, int(val.strip())))
                self.wb.dirty = True
                self.status_message = f"Default column width set to {self.ws.default_col_width}."
            return

        if action == "ws_toggle_recalc":
            self.wb.auto_recalc = not self.wb.auto_recalc
            if self.wb.auto_recalc:
                self.wb.recalculate()
            self.status_message = (
                "Recalculation set to AUTOMATIC."
                if self.wb.auto_recalc
                else "Recalculation set to MANUAL (press F9 to recalc)."
            )
            return

        if action == "ws_ins_row":
            self.wb.push_undo()
            self.ws.insert_row(self.cur_row)
            self.wb.dirty = True
            self.wb.recalculate()
            self.status_message = f"Inserted row at row {self.cur_row + 1}."
            return

        if action == "ws_del_row":
            self.wb.push_undo()
            self.ws.delete_row(self.cur_row)
            self.wb.dirty = True
            self.wb.recalculate()
            self.status_message = f"Deleted row {self.cur_row + 1}."
            return

        if action == "ws_ins_col":
            self.wb.push_undo()
            self.ws.insert_col(self.cur_col)
            self.wb.dirty = True
            self.wb.recalculate()
            self.status_message = f"Inserted column at {idx_to_col(self.cur_col)}."
            return

        if action == "ws_del_col":
            self.wb.push_undo()
            self.ws.delete_col(self.cur_col)
            self.wb.dirty = True
            self.wb.recalculate()
            self.status_message = f"Deleted column {idx_to_col(self.cur_col)}."
            return

        if action == "ws_ins_sheet":
            name = self._prompt_input(
                stdscr, "New Sheet Tab Name:", f"Sheet{len(self.wb.sheets) + 1}"
            )
            if name and name.strip():
                self.wb.push_undo()
                ws_new = self.wb.add_sheet(name.strip())
                self.cur_row, self.cur_col, self.view_top, self.view_left = 0, 0, 0, 0
                self.status_message = f"Added sheet tab '{ws_new.name}'."
            return

        if action == "ws_del_sheet":
            if self._prompt_yes_no(stdscr, f"Delete sheet tab '{self.ws.name}'?"):
                self.wb.push_undo()
                ok, msg = self.wb.delete_active_sheet()
                self.cur_row, self.cur_col, self.view_top, self.view_left = 0, 0, 0, 0
                self.status_message = msg
            return

        if action == "ws_rename_sheet":
            new_nm = self._prompt_input(stdscr, "Rename Active Sheet Tab:", self.ws.name)
            if new_nm and new_nm.strip():
                self.wb.push_undo()
                self.ws.name = new_nm.strip()[:31]
                self.wb.dirty = True
                self.status_message = f"Renamed sheet tab to '{self.ws.name}'."
            return

        if action == "ws_col_set":
            cur_w = self.ws.get_col_width(self.cur_col)
            val = self._prompt_input(
                stdscr, f"Column {idx_to_col(self.cur_col)} Width (3-48):", str(cur_w)
            )
            if val and val.strip().isdigit():
                self.ws.set_col_width(self.cur_col, int(val.strip()))
                self.wb.dirty = True
                self.status_message = f"Column {idx_to_col(self.cur_col)} width set to {self.ws.get_col_width(self.cur_col)}."
            return

        if action == "ws_col_reset":
            self.ws.col_widths.pop(self.cur_col, None)
            self.wb.dirty = True
            self.status_message = f"Column {idx_to_col(self.cur_col)} reset to default width."
            return

        if action == "ws_col_autofit":
            w = self.ws.auto_fit_col(self.cur_col)
            self.wb.dirty = True
            self.status_message = f"Column {idx_to_col(self.cur_col)} auto-fit to {w} chars."
            return

        if action == "ws_titles_both":
            self.ws.freeze_rows = self.cur_row
            self.ws.freeze_cols = self.cur_col
            self.status_message = f"Frozen {self.cur_row} row(s) and {self.cur_col} col(s)."
            return
        if action == "ws_titles_horiz":
            self.ws.freeze_rows = self.cur_row
            self.status_message = f"Frozen top {self.cur_row} row(s)."
            return
        if action == "ws_titles_vert":
            self.ws.freeze_cols = self.cur_col
            self.status_message = f"Frozen left {self.cur_col} column(s)."
            return
        if action == "ws_titles_clear":
            self.ws.freeze_rows = 0
            self.ws.freeze_cols = 0
            self.status_message = "Cleared frozen titles."
            return

        if action == "ws_erase":
            if self._prompt_yes_no(stdscr, f"Erase ALL cells on '{self.ws.name}'?"):
                self.wb.push_undo()
                self.ws.cells.clear()
                self.wb.dirty = True
                self.wb.recalculate()
                self.status_message = f"Erased worksheet '{self.ws.name}'."
            return

        if action == "ws_status":
            self._modal_status(stdscr)
            return

        # ---- Range Commands ----
        if action.startswith("rng_fmt_"):
            code = action[len("rng_fmt_") :]
            fmt_val = "" if code == "reset" else code
            rstr = self._prompt_input(stdscr, f"Format ({code}) Range (e.g. A1..D10):", cur_ref)
            if rstr:
                rng = parse_range_str(rstr)
                if not rng:
                    self.status_message = f"Invalid range: {rstr}"
                    return
                self.wb.push_undo()
                for r in range(rng[0], rng[2] + 1):
                    for c in range(rng[1], rng[3] + 1):
                        self.ws.set_format(r, c, fmt_val)
                self.wb.dirty = True
                self.status_message = f"Applied format '{code}' to {rstr}."
            return

        if action.startswith("rng_lbl_"):
            prefix = {"rng_lbl_left": "'", "rng_lbl_right": '"', "rng_lbl_center": "^"}[action]
            rstr = self._prompt_input(stdscr, "Align Labels in Range:", cur_ref)
            if rstr:
                rng = parse_range_str(rstr)
                if not rng:
                    self.status_message = f"Invalid range: {rstr}"
                    return
                self.wb.push_undo()
                for r in range(rng[0], rng[2] + 1):
                    for c in range(rng[1], rng[3] + 1):
                        cell = self.ws.get_cell(r, c)
                        if cell and cell.raw and not is_formula_raw(cell.raw) and try_parse_number(cell.raw) is None:
                            stripped = cell.raw[1:] if cell.raw[0] in ("'", '"', "^") else cell.raw
                            cell.raw = prefix + stripped
                self.wb.dirty = True
                self.wb.recalculate()
                self.status_message = f"Aligned labels in {rstr}."
            return

        if action == "rng_erase":
            rstr = self._prompt_input(stdscr, "Erase Range (e.g. A1..C5):", cur_ref)
            if rstr:
                rng = parse_range_str(rstr)
                if not rng:
                    self.status_message = f"Invalid range: {rstr}"
                    return
                self.wb.push_undo()
                for r in range(rng[0], rng[2] + 1):
                    for c in range(rng[1], rng[3] + 1):
                        self.ws.clear_cell(r, c)
                self.wb.dirty = True
                self.wb.recalculate()
                self.status_message = f"Erased range {rstr}."
            return

        if action == "rng_value":
            rstr = self._prompt_input(stdscr, "Convert Formulas to Values in Range:", cur_ref)
            if rstr:
                rng = parse_range_str(rstr)
                if not rng:
                    self.status_message = f"Invalid range: {rstr}"
                    return
                self.wb.push_undo()
                for r in range(rng[0], rng[2] + 1):
                    for c in range(rng[1], rng[3] + 1):
                        cell = self.ws.get_cell(r, c)
                        if cell and is_formula_raw(cell.raw) and not cell.error and cell.value is not None:
                            cell.raw = str(cell.value)
                self.wb.dirty = True
                self.wb.recalculate()
                self.status_message = f"Converted formulas in {rstr} to literal values."
            return

        if action == "rng_transpose":
            src = self._prompt_input(stdscr, "Transpose FROM Range:", cur_ref)
            if not src:
                return
            dst = self._prompt_input(stdscr, "Transpose TO Top-Left Cell:", cur_ref)
            if not dst:
                return
            sr = parse_range_str(src)
            dp = parse_cell_ref(dst)
            if not sr or not dp:
                self.status_message = "Invalid source or destination range."
                return
            self.wb.push_undo()
            snapshot = []
            for r in range(sr[0], sr[2] + 1):
                row_snap = []
                for c in range(sr[1], sr[3] + 1):
                    cell = self.ws.get_cell(r, c)
                    row_snap.append((cell.raw, cell.fmt) if cell else ("", ""))
                snapshot.append(row_snap)
            for dr, row_snap in enumerate(snapshot):
                for dc, (raw_v, fmt_v) in enumerate(row_snap):
                    nr, nc = dp[0] + dc, dp[1] + dr
                    if nr < MAX_ROWS and nc < MAX_COLS:
                        self.ws.set_raw(nr, nc, raw_v, fmt_v)
            self.wb.dirty = True
            self.wb.recalculate()
            self.status_message = f"Transposed {src} to {dst}."
            return

        if action == "rng_name_create":
            nm = self._prompt_input(stdscr, "Range Name (e.g. SALES):", "")
            if not nm or not nm.strip():
                return
            rstr = self._prompt_input(stdscr, f"Range for '{nm.strip().upper()}':", cur_ref)
            if rstr and parse_range_str(rstr):
                self.wb.named_ranges[nm.strip().upper()] = (self.ws.name, rstr.strip().upper())
                self.wb.dirty = True
                self.wb.recalculate()
                self.status_message = f"Created named range {nm.strip().upper()} = {rstr.upper()}."
            return

        if action == "rng_name_delete":
            nm = self._prompt_input(stdscr, "Delete Range Name:", "")
            if nm and nm.strip().upper() in self.wb.named_ranges:
                del self.wb.named_ranges[nm.strip().upper()]
                self.wb.dirty = True
                self.wb.recalculate()
                self.status_message = f"Deleted named range {nm.strip().upper()}."
            else:
                self.status_message = "Named range not found."
            return

        if action == "rng_name_list":
            lines = ["Named Ranges in Workbook:", ""]
            if not self.wb.named_ranges:
                lines.append("  (No named ranges defined yet. Use /Range Name Create.)")
            else:
                for k, (sn, rs) in sorted(self.wb.named_ranges.items()):
                    lines.append(f"  {k:<16} -> {sn}!{rs}")
            self._modal_text_viewer(stdscr, "Named Ranges", lines)
            return

        # ---- Copy & Move ----
        if action in ("op_copy", "op_move"):
            is_move = (action == "op_move")
            verb = "Move" if is_move else "Copy"
            src = self._prompt_input(stdscr, f"{verb} FROM Range (e.g. A1..B5):", cur_ref)
            if not src:
                return
            dst = self._prompt_input(stdscr, f"{verb} TO Cell or Range:", cur_ref)
            if not dst:
                return
            sr = parse_range_str(src)
            dr = parse_range_str(dst)
            if not sr or not dr:
                self.status_message = "Invalid source or target range."
                return
            self.wb.push_undo()
            h = sr[2] - sr[0] + 1
            w = sr[3] - sr[1] + 1
            # Snapshot source cells
            src_data = []
            for r in range(sr[0], sr[2] + 1):
                row_d = []
                for c in range(sr[1], sr[3] + 1):
                    cell = self.ws.get_cell(r, c)
                    row_d.append((cell.raw, cell.fmt) if cell else ("", ""))
                src_data.append(row_d)

            if is_move:
                for r in range(sr[0], sr[2] + 1):
                    for c in range(sr[1], sr[3] + 1):
                        self.ws.clear_cell(r, c)

            # Support classic Lotus 1-2-3 single-cell copy into a multi-cell target range!
            if not is_move and h == 1 and w == 1:
                raw_v, fmt_v = src_data[0][0]
                for tr in range(dr[0], dr[2] + 1):
                    for tc in range(dr[1], dr[3] + 1):
                        shifted = shift_formula(raw_v, tr - sr[0], tc - sr[1])
                        self.ws.set_raw(tr, tc, shifted, fmt_v)
            else:
                for ro in range(h):
                    for co in range(w):
                        tr = dr[0] + ro
                        tc = dr[1] + co
                        if tr < MAX_ROWS and tc < MAX_COLS:
                            raw_v, fmt_v = src_data[ro][co]
                            shifted = raw_v if is_move else shift_formula(raw_v, tr - (sr[0] + ro), tc - (sr[1] + co))
                            self.ws.set_raw(tr, tc, shifted, fmt_v)

            self.wb.dirty = True
            self.wb.recalculate()
            self.status_message = f"{verb} complete: {src} -> {dst}."
            return

        # ---- File Commands ----
        if action == "file_save":
            self._action_save(stdscr)
            return
        if action == "file_list":
            self._screen_workbook_browser(stdscr)
            return
        if action == "file_new":
            name = self._prompt_input(stdscr, "New Workbook Name:", "NEW_SHEET")
            if name and name.strip():
                entry, wb = self.gsheets.create_workbook_entry(name.strip())
                self.current_entry = entry
                self.wb = wb
                self.cur_row, self.cur_col, self.view_top, self.view_left = 0, 0, 0, 0
                self.status_message = f"Created new workbook '{entry.name}'."
            return
        if action == "file_import":
            path = self._prompt_input(stdscr, "Import CSV or XLSX Path:", "~/")
            if path and path.strip():
                self._action_import_file(os.path.expanduser(path.strip()))
            return
        if action == "file_export_csv":
            default_p = os.path.expanduser(f"~/{self.current_entry.name}.csv")
            path = self._prompt_input(stdscr, "Export Active Sheet to CSV Path:", default_p)
            if path and path.strip():
                try:
                    with open(os.path.expanduser(path.strip()), "w", encoding="utf-8") as f:
                        f.write(export_worksheet_to_csv(self.ws, formulas=False))
                    self.status_message = f"Exported CSV to {path.strip()}."
                except Exception as exc:
                    self.status_message = f"CSV export error: {exc}"
            return
        if action == "file_export_xlsx":
            default_p = os.path.expanduser(f"~/{self.current_entry.name}.xlsx")
            path = self._prompt_input(stdscr, "Export Workbook to XLSX Path:", default_p)
            if path and path.strip():
                try:
                    with open(os.path.expanduser(path.strip()), "wb") as f:
                        f.write(export_workbook_to_xlsx_bytes(self.wb))
                    self.status_message = f"Exported multi-tab XLSX to {path.strip()}."
                except Exception as exc:
                    self.status_message = f"XLSX export error: {exc}"
            return

        # ---- Print Commands ----
        if action == "print_preview":
            txt = export_worksheet_to_ascii(self.ws, title=self.current_entry.name)
            self._modal_text_viewer(stdscr, f"Print Preview: {self.current_entry.name}", txt.splitlines())
            return
        if action == "print_ov400":
            doc_name = self._prompt_input(
                stdscr, "Save Report to OV/400 Documents As:", f"{self.current_entry.name}_REPORT"
            )
            if doc_name and doc_name.strip():
                txt = export_worksheet_to_ascii(self.ws, title=self.current_entry.name)
                md_text = f"# {self.current_entry.name} ({self.ws.name})\n\n```\n{txt}```\n"
                self.gdrive.create_document(
                    name=doc_name.strip(),
                    description=f"Lotus 1-2-3 Table from {self.current_entry.name}",
                    initial_text=md_text,
                )
                self.status_message = f"Saved report '{doc_name.strip()}' to OV/400 Documents (Option 6/9)."
            return
        if action == "print_txt":
            default_p = os.path.expanduser(f"~/{self.current_entry.name}.txt")
            path = self._prompt_input(stdscr, "Export ASCII Report to Path:", default_p)
            if path and path.strip():
                try:
                    with open(os.path.expanduser(path.strip()), "w", encoding="utf-8") as f:
                        f.write(export_worksheet_to_ascii(self.ws, title=self.current_entry.name))
                    self.status_message = f"Exported ASCII report to {path.strip()}."
                except Exception as exc:
                    self.status_message = f"Print export error: {exc}"
            return

        # ---- Graph Commands ----
        if action in ("graph_bar", "graph_line"):
            self._action_graph(stdscr, chart_type="BAR" if action == "graph_bar" else "LINE")
            return

        # ---- Data Commands ----
        if action == "data_fill":
            self._action_data_fill(stdscr)
            return
        if action == "data_sort":
            self._action_data_sort(stdscr)
            return

        # ---- Google Sheets Sync Commands ----
        if action == "gs_pull":
            self._action_gsheets_pull(stdscr)
            return
        if action == "gs_push":
            self._action_gsheets_push(stdscr)
            return
        if action == "gs_browse":
            self._screen_workbook_browser(stdscr, auto_refresh=True)
            return
        if action == "gs_link":
            self._action_gsheets_link(stdscr)
            return
        if action == "gs_config":
            setup = GoogleDriveSetupScreen(
                self.gdrive,
                system_name=self.system_name,
                launch_external_cb=self.launch_external_cb,
            )
            setup.run(stdscr)
            self.status_message = f"Google Sheets backend: {self.gsheets.get_status_summary()}"
            return
        if action == "gs_autosync":
            cur = bool(self.gdrive.config.get("auto_sync_on_save", True))
            self.gdrive.config["auto_sync_on_save"] = not cur
            self.gdrive.save_config()
            self.status_message = (
                "Auto-Sync on Save enabled." if not cur else "Auto-Sync on Save disabled."
            )
            return

    # =========================================================================
    # Core Actions (Save, Import, Fill, Sort, Graph, Google Sheets Sync)
    # =========================================================================
    def _set_wait_banner(self, stdscr, msg: str):
        old_mode = self.mode
        self.mode = "WAIT"
        self.status_message = msg
        max_y, max_x = stdscr.getmaxyx()
        stdscr.erase()
        self._draw_screen(stdscr, max_y, max_x)
        stdscr.refresh()
        self.mode = old_mode

    def _action_save(self, stdscr):
        self._set_wait_banner(stdscr, f"Saving '{self.current_entry.name}'...")
        ok, msg = self.gsheets.save_workbook(self.current_entry, self.wb, sync_cloud=True)
        self.status_message = msg

    def _action_gsheets_pull(self, stdscr):
        if self.gsheets.get_backend() == "local":
            self.status_message = "Google Drive/Sheets not configured. Opening setup (/Sheets Config)..."
            setup = GoogleDriveSetupScreen(
                self.gdrive,
                system_name=self.system_name,
                launch_external_cb=self.launch_external_cb,
            )
            setup.run(stdscr)
            if self.gsheets.get_backend() == "local":
                return

        if not self.current_entry.gsheet_id:
            self._action_gsheets_link(stdscr)
            return

        self._set_wait_banner(stdscr, f"Pulling '{self.current_entry.name}' from Google Sheets...")
        ok, msg, pulled_wb = self.gsheets.pull_spreadsheet(self.current_entry)
        if ok and pulled_wb is not None:
            self.wb = pulled_wb
            self.cur_row, self.cur_col, self.view_top, self.view_left = 0, 0, 0, 0
        self.status_message = msg

    def _action_gsheets_push(self, stdscr):
        if self.gsheets.get_backend() == "local":
            setup = GoogleDriveSetupScreen(
                self.gdrive,
                system_name=self.system_name,
                launch_external_cb=self.launch_external_cb,
            )
            setup.run(stdscr)
            if self.gsheets.get_backend() == "local":
                self.status_message = "Google Sheets not connected (still in Local Mode)."
                return

        self._set_wait_banner(stdscr, f"Pushing '{self.current_entry.name}' to Google Sheets...")
        # Save locally first
        self.gsheets.save_workbook(self.current_entry, self.wb, sync_cloud=False)
        ok, msg = self.gsheets.push_spreadsheet(self.current_entry, self.wb)
        self.status_message = msg

    def _action_gsheets_link(self, stdscr):
        cur_id = self.current_entry.gsheet_id or ""
        raw = self._prompt_input(stdscr, "Google Sheet URL or Spreadsheet ID:", cur_id)
        if not raw or not raw.strip():
            return
        gid = extract_spreadsheet_id(raw)
        if not gid:
            self.status_message = "Could not parse Google Sheet ID."
            return
        self.current_entry.gsheet_id = gid
        self.wb.gsheet_id = gid
        self.gsheets.save_index()
        self.status_message = f"Linked to Google Sheet ID: {gid}"
        if self.gsheets.get_backend() != "local":
            if self._prompt_yes_no(stdscr, "Pull data from this Google Sheet now?"):
                self._action_gsheets_pull(stdscr)

    def _quick_gsheets_sync(self, stdscr):
        if self.gsheets.get_backend() == "local":
            setup = GoogleDriveSetupScreen(
                self.gdrive,
                system_name=self.system_name,
                launch_external_cb=self.launch_external_cb,
            )
            setup.run(stdscr)
            self.status_message = f"Google Sheets status: {self.gsheets.get_status_summary()}"
            return

        # Open quick Sync menu directly on /Sheets
        self.mode = "MENU"
        sheets_node = [m for m in MENU_TREE if m[0] == "Sheets"][0]
        self.menu_stack = [(sheets_node[3], 0)]

    def _action_import_file(self, path: str):
        if not os.path.exists(path):
            self.status_message = f"File not found: {path}"
            return
        base = os.path.splitext(os.path.basename(path))[0] or "IMPORTED"
        try:
            if path.lower().endswith(".xlsx"):
                with open(path, "rb") as f:
                    wb = import_xlsx_bytes_to_workbook(f.read(), title=base)
            else:
                with open(path, "r", encoding="utf-8", errors="replace") as f:
                    ws = import_csv_to_worksheet(f.read(), name=base[:31])
                wb = Workbook(title=base)
                wb.sheets = [ws]
                wb.recalculate()
            entry, wb = self.gsheets.create_workbook_entry(base, wb=wb)
            self.current_entry = entry
            self.wb = wb
            self.cur_row, self.cur_col, self.view_top, self.view_left = 0, 0, 0, 0
            self.status_message = f"Imported '{os.path.basename(path)}' ({len(wb.sheets)} tab(s))."
        except Exception as exc:
            self.status_message = f"Import failed: {exc}"

    def _action_data_fill(self, stdscr):
        cur_ref = format_cell_ref(self.cur_row, self.cur_col)
        rstr = self._prompt_input(stdscr, "Fill Range (e.g. A1..A10):", f"{cur_ref}..{idx_to_col(self.cur_col)}{self.cur_row + 10}")
        if not rstr:
            return
        rng = parse_range_str(rstr)
        if not rng:
            self.status_message = f"Invalid fill range: {rstr}"
            return
        s_start = self._prompt_input(stdscr, "Start Value:", "1")
        s_step = self._prompt_input(stdscr, "Step Value:", "1")
        s_stop = self._prompt_input(stdscr, "Stop Value:", "8192")
        if s_start is None or s_step is None or s_stop is None:
            return
        try:
            val = float(s_start)
            step = float(s_step)
            stop = float(s_stop)
        except ValueError:
            self.status_message = "Start, Step, and Stop must be numeric."
            return

        self.wb.push_undo()
        count = 0
        for c in range(rng[1], rng[3] + 1):
            for r in range(rng[0], rng[2] + 1):
                if (step >= 0 and val > stop) or (step < 0 and val < stop):
                    break
                txt = str(int(val)) if val.is_integer() else str(val)
                self.ws.set_raw(r, c, txt)
                val += step
                count += 1
        self.wb.dirty = True
        self.wb.recalculate()
        self.status_message = f"Filled {count} cell(s) in {rstr}."

    def _action_data_sort(self, stdscr):
        max_r, max_c = self.ws.get_bounds()
        default_rng = f"A1..{idx_to_col(max(0, max_c))}{max(1, max_r + 1)}"
        rstr = self._prompt_input(stdscr, "Sort Data Range (exclude header row):", default_rng)
        if not rstr:
            return
        rng = parse_range_str(rstr)
        if not rng:
            self.status_message = f"Invalid range: {rstr}"
            return
        key_col_s = self._prompt_input(stdscr, "Primary Sort Key Column (e.g. B):", idx_to_col(self.cur_col))
        if not key_col_s:
            return
        try:
            key_c = col_to_idx(key_col_s.strip())
        except ValueError:
            self.status_message = f"Invalid column: {key_col_s}"
            return
        order = self._prompt_input(stdscr, "Sort Order (A=Ascending, D=Descending):", "A")
        if not order:
            return
        desc = order.strip().upper().startswith("D")

        self.wb.push_undo()
        rows_block = []
        for r in range(rng[0], rng[2] + 1):
            row_cells = []
            for c in range(rng[1], rng[3] + 1):
                cell = self.ws.get_cell(r, c)
                row_cells.append((cell.raw, cell.fmt) if cell else ("", ""))
            key_cell = self.ws.get_cell(r, key_c)
            kval = key_cell.value if key_cell else None
            if isinstance(kval, (int, float)) and not isinstance(kval, bool):
                sort_key = (0, float(kval), "")
            elif kval is None or kval == "":
                sort_key = (2, 0.0, "")
            else:
                sort_key = (1, 0.0, str(kval).lower())
            rows_block.append((sort_key, row_cells))

        rows_block.sort(key=lambda x: x[0], reverse=desc)
        for i, (_, row_cells) in enumerate(rows_block):
            r = rng[0] + i
            for j, (raw_v, fmt_v) in enumerate(row_cells):
                c = rng[1] + j
                self.ws.set_raw(r, c, raw_v, fmt_v)

        self.wb.dirty = True
        self.wb.recalculate()
        self.status_message = f"Sorted {rstr} by column {key_col_s.upper()} ({'DESC' if desc else 'ASC'})."

    def _action_graph(self, stdscr, chart_type: str = "BAR"):
        # Smart defaults (works right away on sample budget or user data!)
        default_x = "A6..A10" if self.ws.get_cell(5, 0) else "A2..A10"
        default_y = "E6..E10" if self.ws.get_cell(5, 4) else "B2..B10"
        x_str = self._prompt_input(stdscr, "X-Axis Labels Range:", default_x)
        if not x_str:
            return
        y_str = self._prompt_input(stdscr, "Y-Axis Values Range:", default_y)
        if not y_str:
            return

        xr = parse_range_str(x_str)
        yr = parse_range_str(y_str)
        if not xr or not yr:
            self.status_message = "Invalid graph range."
            return

        labels: List[str] = []
        for r in range(xr[0], xr[2] + 1):
            for c in range(xr[1], xr[3] + 1):
                cell = self.ws.get_cell(r, c)
                labels.append(str(cell.value) if (cell and cell.value is not None) else format_cell_ref(r, c))

        values: List[float] = []
        for r in range(yr[0], yr[2] + 1):
            for c in range(yr[1], yr[3] + 1):
                cell = self.ws.get_cell(r, c)
                if cell and isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                    values.append(float(cell.value))
                else:
                    values.append(0.0)

        n = min(len(labels), len(values))
        if n == 0:
            self.status_message = "No data points found in graph range."
            return
        labels = labels[:n]
        values = values[:n]

        cur_type = chart_type
        while True:
            max_y, max_x = stdscr.getmaxyx()
            stdscr.erase()
            title = f"LOTUS 1-2-3 GRAPH ({cur_type}) - {self.current_entry.name} [{y_str}]"
            safe_addstr(win := stdscr, 0, max(0, (max_x - len(title)) // 2), title[: max_x - 2], curses.A_BOLD | curses.A_REVERSE)
            safe_addstr(win, 1, 1, "=" * (max_x - 2))

            max_val = max(max(values), 1e-9)
            min_val = min(min(values), 0.0)

            if cur_type == "BAR":
                label_w = min(20, max(8, max(len(l) for l in labels) + 1))
                bar_max_w = max(10, max_x - label_w - 18)
                y = 3
                for lbl, val in zip(labels, values):
                    if y >= max_y - 4:
                        break
                    frac = max(0.0, val) / max_val if max_val > 0 else 0.0
                    blen = int(round(frac * bar_max_w))
                    bar_str = "#" * blen
                    safe_addstr(win, y, 2, f"{lbl[:label_w]:<{label_w}} |", curses.A_BOLD)
                    if blen > 0:
                        safe_addstr(win, y, 2 + label_w + 3, bar_str, curses.A_REVERSE)
                    val_txt = f" {val:,.2f}" if val != int(val) else f" {int(val):,}"
                    safe_addstr(win, y, 2 + label_w + 3 + blen + 1, val_txt)
                    y += 2 if (max_y - 8) // max(1, n) >= 2 else 1
            else:
                # Monochrome ASCII Line Chart
                plot_h = max(8, max_y - 10)
                plot_w = max(20, max_x - 16)
                span = max(max_val - min_val, 1e-9)
                grid = [[" " for _ in range(plot_w)] for _ in range(plot_h)]
                pts: List[Tuple[int, int]] = []
                for idx, val in enumerate(values):
                    px = int(round(idx * (plot_w - 1) / max(1, n - 1))) if n > 1 else plot_w // 2
                    py = plot_h - 1 - int(round((val - min_val) / span * (plot_h - 1)))
                    py = max(0, min(plot_h - 1, py))
                    pts.append((px, py))
                # Connect points with dots/lines
                for i in range(len(pts) - 1):
                    x0, y0 = pts[i]
                    x1, y1 = pts[i + 1]
                    steps = max(abs(x1 - x0), abs(y1 - y0), 1)
                    for s in range(steps + 1):
                        ix = int(round(x0 + (x1 - x0) * s / steps))
                        iy = int(round(y0 + (y1 - y0) * s / steps))
                        if 0 <= iy < plot_h and 0 <= ix < plot_w and grid[iy][ix] == " ":
                            grid[iy][ix] = "*"
                for px, py in pts:
                    grid[py][px] = "O"

                for r_i in range(plot_h):
                    y_val = max_val - (span * r_i / max(1, plot_h - 1))
                    safe_addstr(win, 3 + r_i, 1, f"{y_val:>10.1f} |")
                    safe_addstr(win, 3 + r_i, 13, "".join(grid[r_i]), curses.A_BOLD)
                safe_addstr(win, 3 + plot_h, 12, "+" + "-" * plot_w)
                # X labels summary
                x_leg = "  ".join(f"{l[:8]}={v:g}" for l, v in zip(labels, values))
                safe_addstr(win, 4 + plot_h, 2, x_leg[: max_x - 4])

            safe_addstr(
                win,
                max_y - 2,
                2,
                "B=Bar Chart   L=Line Chart   Esc/Enter/F12=Return to Worksheet",
                curses.A_BOLD,
            )
            stdscr.refresh()
            ch = read_key(stdscr)
            if ch in (ord("b"), ord("B")):
                cur_type = "BAR"
            elif ch in (ord("l"), ord("L")):
                cur_type = "LINE"
            elif ch in (27, 10, 13, curses.KEY_ENTER, curses.KEY_F12, curses.KEY_F3, ord("q"), ord("Q")):
                break

    # =========================================================================
    # Workbook & Google Sheets Browser Screen (F4 / /File Retrieve / /Sheets Browse)
    # =========================================================================
    def _screen_workbook_browser(self, stdscr, auto_refresh: bool = False):
        sel_idx = 0
        scroll = 0
        msg = f"Backend: {self.gsheets.get_status_summary()}  (F5=Refresh Google Sheets, F10=Setup)"

        if auto_refresh and self.gsheets.get_backend() != "local":
            ok, msg = self.gsheets.sync_spreadsheet_list()

        while True:
            entries = self.gsheets.list_entries()
            if entries:
                sel_idx = max(0, min(sel_idx, len(entries) - 1))
            else:
                sel_idx = 0

            max_y, max_x = stdscr.getmaxyx()
            stdscr.erase()

            safe_addstr(stdscr, 0, 1, "LOTUS123", curses.A_BOLD)
            title = "Workbooks & Google Sheets Manager"
            safe_addstr(stdscr, 0, (max_x - len(title)) // 2, title, curses.A_BOLD)
            safe_addstr(stdscr, 0, max_x - 8, " FILES ", curses.A_REVERSE | curses.A_BOLD)

            safe_addstr(
                stdscr,
                2,
                2,
                f"Active Cloud Sync: {self.gsheets.get_status_summary()}   |   Active: {self.current_entry.name}",
                curses.A_BOLD,
            )
            safe_addstr(stdscr, 3, 2, "Enter=Open   P=Pull from GSheets   U=Push to GSheets   D=Delete Local")

            # Table header
            hdr = f"  {'Workbook / Google Sheet Name':<34} {'Tabs':>4}  {'Cloud Status':<16} {'Updated':<16}"
            safe_addstr(stdscr, 5, 1, hdr[: max_x - 2], curses.A_REVERSE | curses.A_BOLD)

            list_h = max(4, max_y - 11)
            if sel_idx < scroll:
                scroll = sel_idx
            elif sel_idx >= scroll + list_h:
                scroll = sel_idx - list_h + 1

            if not entries:
                safe_addstr(stdscr, 7, 4, "(No workbooks found. Press F6 to create a new workbook.)")
            else:
                for i in range(list_h):
                    idx = scroll + i
                    if idx >= len(entries):
                        break
                    ent = entries[idx]
                    is_cur = (ent.sheet_id == self.current_entry.sheet_id)
                    marker = "*" if is_cur else " "
                    if ent.gsheet_id:
                        cst = "GSHEET (SYNCED)" if ent.synced and not ent.dirty else "GSHEET (DIRTY)"
                    else:
                        cst = "LOCAL ONLY"
                    row_txt = f"{marker} {ent.name[:34]:<34} {ent.tab_count:>4}  {cst:<16} {ent.updated_at[:16]:<16}"
                    attr = curses.A_REVERSE | curses.A_BOLD if idx == sel_idx else curses.A_NORMAL
                    safe_addstr(stdscr, 6 + i, 1, row_txt[: max_x - 2], attr)

            safe_addstr(
                stdscr,
                max_y - 3,
                1,
                "F5=Refresh Google Sheets   F6=New Workbook   F7=Open by GSheet URL/ID   F10=GDrive Setup",
            )
            safe_addstr(stdscr, max_y - 2, 1, "Enter=Open   P=Pull   U=Push   D=Delete Local   F12/Esc=Back to Grid")
            safe_addstr(stdscr, max_y - 1, 1, msg[: max_x - 2], curses.A_BOLD)
            stdscr.refresh()

            ch = read_key(stdscr)
            if ch == -1:
                continue
            if ch in (27, curses.KEY_F12, curses.KEY_F3):
                return
            if ch == curses.KEY_UP and entries:
                sel_idx = (sel_idx - 1) % len(entries)
                continue
            if ch == curses.KEY_DOWN and entries:
                sel_idx = (sel_idx + 1) % len(entries)
                continue

            # F5 = Refresh list from Google Drive / Sheets
            if ch == curses.KEY_F5:
                msg = "Fetching Google Sheets list..."
                safe_addstr(stdscr, max_y - 1, 1, msg.ljust(max_x - 2)[: max_x - 2], curses.A_BOLD)
                stdscr.refresh()
                ok, msg = self.gsheets.sync_spreadsheet_list()
                continue

            # F6 = New Workbook
            if ch == curses.KEY_F6:
                nm = self._prompt_input(stdscr, "New Workbook Name:", "NEW_SHEET")
                if nm and nm.strip():
                    entry, wb = self.gsheets.create_workbook_entry(nm.strip())
                    self.current_entry = entry
                    self.wb = wb
                    self.cur_row, self.cur_col, self.view_top, self.view_left = 0, 0, 0, 0
                    self.status_message = f"Created & opened '{entry.name}'."
                    return
                continue

            # F7 = Open / Link by Google Sheets URL or ID
            if ch == curses.KEY_F7:
                raw = self._prompt_input(stdscr, "Paste Google Sheet URL or ID:", "")
                if raw and raw.strip():
                    gid = extract_spreadsheet_id(raw)
                    nm = self._prompt_input(stdscr, "Local Workbook Name:", f"GSHEET_{gid[:6]}")
                    if nm and nm.strip():
                        entry, wb = self.gsheets.create_workbook_entry(nm.strip(), gsheet_id=gid)
                        if self.gsheets.get_backend() != "local":
                            ok, pmsg, pulled = self.gsheets.pull_spreadsheet(entry)
                            if ok and pulled is not None:
                                wb = pulled
                            msg = pmsg
                        else:
                            msg = f"Linked '{entry.name}' to {gid} (press F10 to sign in & pull)."
                        self.current_entry = entry
                        self.wb = wb
                        self.cur_row, self.cur_col, self.view_top, self.view_left = 0, 0, 0, 0
                        self.status_message = msg
                        return
                continue

            # F10 = Configure Google Drive & Sheets
            if ch == curses.KEY_F10:
                setup = GoogleDriveSetupScreen(
                    self.gdrive,
                    system_name=self.system_name,
                    launch_external_cb=self.launch_external_cb,
                )
                setup.run(stdscr)
                msg = f"Updated config. Status: {self.gsheets.get_status_summary()}"
                continue

            if not entries:
                continue
            target = entries[sel_idx]

            # Enter = Open workbook
            if ch in (10, 13, curses.KEY_ENTER, ord("1")):
                wb, lmsg = self.gsheets.load_workbook(target)
                self.current_entry = target
                self.wb = wb
                self.cur_row, self.cur_col, self.view_top, self.view_left = 0, 0, 0, 0
                self.status_message = lmsg
                return

            # P = Pull from Google Sheets
            if ch in (ord("p"), ord("P"), ord("2")):
                msg = f"Pulling '{target.name}' from Google Sheets..."
                safe_addstr(stdscr, max_y - 1, 1, msg.ljust(max_x - 2)[: max_x - 2], curses.A_BOLD)
                stdscr.refresh()
                ok, pmsg, pulled = self.gsheets.pull_spreadsheet(target)
                msg = pmsg
                if ok and pulled is not None:
                    self.current_entry = target
                    self.wb = pulled
                    self.cur_row, self.cur_col, self.view_top, self.view_left = 0, 0, 0, 0
                    self.status_message = pmsg
                    return
                continue

            # U = Push to Google Sheets
            if ch in (ord("u"), ord("U"), ord("3")):
                wb_to_push = self.wb if target.sheet_id == self.current_entry.sheet_id else self.gsheets.load_workbook(target)[0]
                msg = f"Pushing '{target.name}' to Google Sheets..."
                safe_addstr(stdscr, max_y - 1, 1, msg.ljust(max_x - 2)[: max_x - 2], curses.A_BOLD)
                stdscr.refresh()
                ok, pmsg = self.gsheets.push_spreadsheet(target, wb_to_push)
                msg = pmsg
                continue

            # D = Delete local copy
            if ch in (ord("d"), ord("D"), ord("4")):
                if len(entries) <= 1:
                    msg = "Cannot delete the last remaining workbook."
                    continue
                if self._prompt_yes_no(stdscr, f"Delete local workbook '{target.name}'?"):
                    ok, dmsg = self.gsheets.delete_entry(target)
                    msg = dmsg
                    if target.sheet_id == self.current_entry.sheet_id:
                        rem = self.gsheets.list_entries()
                        if rem:
                            self.current_entry = rem[0]
                            self.wb, _ = self.gsheets.load_workbook(self.current_entry)
                continue

    # =========================================================================
    # Modals & Prompts (GoTo, Input, Yes/No, Status, Text Viewer, Help)
    # =========================================================================
    def _prompt_goto(self, stdscr):
        val = self._prompt_input(stdscr, "GoTo Cell (e.g. E12, Summary!B4, or Named Range):", "")
        if not val or not val.strip():
            return
        target = val.strip()
        # Check named range
        if target.upper() in self.wb.named_ranges:
            sname, rstr = self.wb.named_ranges[target.upper()]
            for i, ws in enumerate(self.wb.sheets):
                if ws.name.upper() == sname.upper():
                    self.wb.active_sheet_idx = i
                    break
            rng = parse_range_str(rstr)
            if rng:
                self.cur_row, self.cur_col = rng[0], rng[1]
                self.status_message = f"Jumped to named range {target.upper()} ({sname}!{rstr})."
            return

        # Check sheet!cell
        if "!" in target:
            sn, cref = target.split("!", 1)
            sn = sn.strip().strip("'")
            for i, ws in enumerate(self.wb.sheets):
                if ws.name.upper() == sn.upper():
                    self.wb.active_sheet_idx = i
                    break
            target = cref

        parsed = parse_cell_ref(target)
        if parsed:
            self.cur_row, self.cur_col = parsed[0], parsed[1]
            self.status_message = f"Moved to {format_cell_ref(self.cur_row, self.cur_col)}."
        else:
            self.status_message = f"Invalid cell reference: {val}"

    def _prompt_input(self, stdscr, prompt: str, initial: str = "") -> Optional[str]:
        buf = initial
        pos = len(buf)
        try:
            curses.curs_set(1)
        except curses.error:
            pass
        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw = min(74, max_x - 4)
            bh = 7
            sy = max(3, (max_y - bh) // 2)
            sx = max(2, (max_x - bw) // 2)
            draw_5250_box(stdscr, sy, sx, bh, bw, "Lotus 1-2-3 Prompt")
            safe_addstr(stdscr, sy + 2, sx + 2, prompt[: bw - 4], curses.A_BOLD)
            field_w = bw - 6
            scroll_off = max(0, pos - field_w + 1)
            disp = buf[scroll_off : scroll_off + field_w]
            safe_addstr(stdscr, sy + 3, sx + 2, f"{disp:<{field_w}}", curses.A_REVERSE | curses.A_BOLD)
            safe_addstr(stdscr, sy + 5, sx + 2, "Enter=Confirm   Esc/F12=Cancel")
            try:
                stdscr.move(sy + 3, sx + 2 + (pos - scroll_off))
            except curses.error:
                pass
            stdscr.refresh()

            ch = read_key(stdscr)
            if ch == -1:
                continue
            if ch in (27, curses.KEY_F12):
                return None
            if ch in (10, 13, curses.KEY_ENTER):
                return buf
            if ch == curses.KEY_LEFT:
                pos = max(0, pos - 1)
            elif ch == curses.KEY_RIGHT:
                pos = min(len(buf), pos + 1)
            elif ch == curses.KEY_HOME:
                pos = 0
            elif ch == curses.KEY_END:
                pos = len(buf)
            elif ch in (curses.KEY_BACKSPACE, 127, 8):
                if pos > 0:
                    buf = buf[: pos - 1] + buf[pos:]
                    pos -= 1
            elif ch == curses.KEY_DC:
                if pos < len(buf):
                    buf = buf[:pos] + buf[pos + 1 :]
            elif 32 <= ch <= 126:
                buf = buf[:pos] + chr(ch) + buf[pos:]
                pos += 1

    def _prompt_yes_no(self, stdscr, question: str) -> bool:
        try:
            curses.curs_set(0)
        except curses.error:
            pass
        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw = min(68, max_x - 4)
            bh = 6
            sy = max(3, (max_y - bh) // 2)
            sx = max(2, (max_x - bw) // 2)
            draw_5250_box(stdscr, sy, sx, bh, bw, "Confirm")
            safe_addstr(stdscr, sy + 2, sx + 2, question[: bw - 4], curses.A_BOLD)
            safe_addstr(stdscr, sy + 4, sx + 2, "Press [Y] for Yes, or [N] / [Esc] for No.")
            stdscr.refresh()
            ch = read_key(stdscr)
            if ch in (ord("y"), ord("Y")):
                return True
            if ch in (ord("n"), ord("N"), 27, curses.KEY_F12):
                return False

    def _confirm_quit(self, stdscr):
        if self.wb.dirty:
            if self._prompt_yes_no(stdscr, f"Save changes to '{self.current_entry.name}' before exiting?"):
                self._action_save(stdscr)
        self.running = False

    def _modal_status(self, stdscr):
        max_r, max_c = self.ws.get_bounds()
        total_cells = sum(len(s.cells) for s in self.wb.sheets)
        lines = [
            f"Workbook Name    : {self.current_entry.name}",
            f"Local File       : {self.current_entry.local_path or '(Not saved yet)'}",
            f"Active Sheet Tab : {self.ws.name} ({self.wb.active_sheet_idx + 1} of {len(self.wb.sheets)})",
            f"Used Bounds      : {format_cell_ref(max(0, max_r), max(0, max_c))} ({total_cells} populated cells)",
            f"Global Format    : {self.ws.global_format}   Default Col Width: {self.ws.default_col_width}",
            f"Recalculation    : {'Automatic' if self.wb.auto_recalc else 'Manual'}   Circular Ref: {'YES' if self.wb.has_circular_ref else 'No'}",
            f"Frozen Titles    : {self.ws.freeze_rows} row(s), {self.ws.freeze_cols} col(s)",
            "",
            "--- Google Sheets Cloud Synchronization ---",
            f"Active Backend   : {self.gsheets.get_status_summary()}",
            f"Linked Sheet ID  : {self.current_entry.gsheet_id or '(Not linked — use /Sheets Push or /Sheets Link)'}",
            f"Auto-Sync Save   : {'Enabled' if self.gdrive.config.get('auto_sync_on_save', True) else 'Disabled'}",
        ]
        self._modal_text_viewer(stdscr, "Worksheet & Google Sheets Status", lines)

    def _modal_text_viewer(self, stdscr, title: str, lines: List[str]):
        try:
            curses.curs_set(0)
        except curses.error:
            pass
        scroll = 0
        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw = min(84, max_x - 4)
            bh = min(max_y - 4, max(12, len(lines) + 5))
            sy = max(1, (max_y - bh) // 2)
            sx = max(1, (max_x - bw) // 2)
            draw_5250_box(stdscr, sy, sx, bh, bw, title)

            body_h = bh - 4
            max_scroll = max(0, len(lines) - body_h)
            scroll = max(0, min(scroll, max_scroll))
            for i in range(body_h):
                idx = scroll + i
                if idx < len(lines):
                    safe_addstr(stdscr, sy + 2 + i, sx + 2, lines[idx][: bw - 4])

            safe_addstr(stdscr, sy + bh - 2, sx + 2, "Up/Down=Scroll   Enter/Esc/F12=Close", curses.A_BOLD)
            stdscr.refresh()
            ch = read_key(stdscr)
            if ch in (27, 10, 13, curses.KEY_ENTER, curses.KEY_F12, curses.KEY_F3, ord("q"), ord("Q")):
                return
            if ch == curses.KEY_UP:
                scroll = max(0, scroll - 1)
            elif ch == curses.KEY_DOWN:
                scroll = min(max_scroll, scroll + 1)

    def _modal_help(self, stdscr):
        lines = [
            "LOTUS 1-2-3 FOR OFFICEVISION/400 — QUICK REFERENCE",
            "",
            "Navigation & Editing:",
            "  * Arrow Keys / PgUp / PgDn / Home / End  Move cell cursor",
            "  * Type any number, '=', '@', or text     Enter VALUE or LABEL directly",
            "  * F2 or Enter                            Edit current cell in-line",
            "  * Del / Backspace                        Clear current cell",
            "  * Ctrl+Z / Ctrl+Y                        Multi-level Undo / Redo",
            "  * Ctrl+S                                 Save (+ AutoSync to Google Sheets)",
            "  * F5                                     GoTo cell, tab, or named range",
            "  * F6                                     Switch to next worksheet tab",
            "  * F9                                     Force full recalculation",
            "",
            "Classic Lotus 1-2-3 Slash (/) Commands:",
            "  * /  Opens 2-line menu (use Arrows+Enter OR single-letter hotkeys!)",
            "  * /W Worksheet (Insert/Delete Row/Col/Sheet, Column Width, Freeze Titles)",
            "  * /R Range     (Format C2/P1/F2, Align Label, Erase, Value, Transpose, Name)",
            "  * /C Copy      /M Move      /D Data (Fill series, Sort range)",
            "  * /F File      (Retrieve, Save, New, Import/Export CSV & XLSX)",
            "  * /P Print     (ASCII Preview, Send to OV/400 Word Processor, .txt)",
            "  * /G Graph     (Monochrome MDA ASCII Bar Chart & Line Chart)",
            "  * /S Sheets    (Pull, Push, Browse, Link, Google OAuth/rclone Config)",
            "  * /Q Quit      (Exit back to OV/400 Main Menu)",
            "",
            "Formulas (BOTH Lotus 1-2-3 & Google Sheets syntax work seamlessly!):",
            "  * @SUM(B6..B10)  or  =SUM(B6:B10)        @AVG(...) or =AVERAGE(...)",
            "  * +E6/$E$12      or  =E6/$E$12           =Budget!E12 (cross-tab refs)",
            "  * Labels: 'Left  \"Right  ^Center  \\- Repeating fill (e.g. \\- or \\=)",
            "  * 55+ functions: SUM, AVERAGE, MIN, MAX, IF, IFERROR, VLOOKUP, INDEX,",
            "    MATCH, COUNTIF, SUMIF, PMT, PV, FV, NPV, ROUND, LEFT, MID, CONCAT...",
            "",
            "Google Sheets Sync (F4 = Files/Sheets Browser, F10 = Sync Menu):",
            "  * Shares OV/400 Google OAuth / rclone credentials (CFGDRV).",
            "  * Translates @SUM(A1..B5) -> =SUM(A1:B5) automatically on Push!",
        ]
        self._modal_text_viewer(stdscr, "Lotus 1-2-3 Help (F1)", lines)


def run_standalone():
    """Allows launching Lotus 1-2-3 standalone outside OV/400."""
    if os.environ.get("MDA") == "1" or os.environ.get("TERM") in ("vt100", "wsvt25", "ansi", "dumb", ""):
        os.environ["LINES"] = os.environ.get("LINES", "43")
        os.environ["COLUMNS"] = os.environ.get("COLUMNS", "90")
        try:
            curses.use_env(True)
        except Exception:
            pass
        sys.stdout.write("\x1b[1;43r")
        sys.stdout.flush()

    gdrive = GoogleDriveManager()
    app = Lotus123App(gdrive)
    try:
        curses.wrapper(app.run)
    finally:
        if os.environ.get("MDA") == "1" or os.environ.get("TERM") == "vt100":
            sys.stdout.write("\x1b[r\x1b[1;43r\x1b[H\x1b[2J\x1b[0m\x1b[?25h")
        else:
            sys.stdout.write("\x1b[r\x1b[H\x1b[2J\x1b[0m\x1b[?25h")
        sys.stdout.flush()


if __name__ == "__main__":
    run_standalone()
