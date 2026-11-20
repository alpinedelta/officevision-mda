#!/usr/bin/env python3
"""
IBM OfficeVision/400 (OV/400) Word Processor & Google Drive Document Browser
Tailored for 90x43 Monochrome MDA Terminal (TERM=vt100).

Includes:
  - WRKDOC: Work with Documents in Folders (Option 6) with Google Drive sync
  - EDTDOC: DisplayWrite/400-style Fullscreen Word Processor (Option 9)
  - CFGDRV: Google Drive & Google Docs Connection Setup (F10)
  - ASPELL: Interactive 5250 Spell Checker using /usr/bin/aspell (F4)
"""

import curses
import datetime
import os
import re
import shutil
import subprocess
import textwrap
import urllib.parse
from typing import Callable, Dict, List, Optional, Tuple

from gdrive import DocumentEntry, GoogleDriveManager


ESC_SEQ_MAP: Dict[str, int] = {
    # Function keys F1-F20 (VT100 / VT220 / xterm / FabGL / Linux console variants)
    "OP": curses.KEY_F1,
    "[11~": curses.KEY_F1,
    "[[A": curses.KEY_F1,
    "OQ": curses.KEY_F2,
    "[12~": curses.KEY_F2,
    "[[B": curses.KEY_F2,
    "OR": curses.KEY_F3,
    "[13~": curses.KEY_F3,
    "[[C": curses.KEY_F3,
    "OS": curses.KEY_F4,
    "[14~": curses.KEY_F4,
    "[[D": curses.KEY_F4,
    "[15~": curses.KEY_F5,
    "Ot": curses.KEY_F5,
    "[[E": curses.KEY_F5,
    "[17~": curses.KEY_F6,
    "Ou": curses.KEY_F6,
    "[18~": curses.KEY_F7,
    "Ov": curses.KEY_F7,
    "[19~": curses.KEY_F8,
    "Ol": curses.KEY_F8,
    "[20~": curses.KEY_F9,
    "Ow": curses.KEY_F9,
    "[21~": curses.KEY_F10,
    "Ox": curses.KEY_F10,
    "[23~": curses.KEY_F11,
    "Oy": curses.KEY_F11,
    "[24~": curses.KEY_F12,
    "[25~": curses.KEY_F13,
    "[26~": curses.KEY_F14,
    "[28~": curses.KEY_F15,
    "[29~": curses.KEY_F16,
    "[31~": curses.KEY_F17,
    "[32~": curses.KEY_F18,
    "[33~": curses.KEY_F19,
    "[34~": curses.KEY_F20,
    # Cursor & navigation keys
    "[A": curses.KEY_UP,
    "OA": curses.KEY_UP,
    "[B": curses.KEY_DOWN,
    "OB": curses.KEY_DOWN,
    "[C": curses.KEY_RIGHT,
    "OC": curses.KEY_RIGHT,
    "[D": curses.KEY_LEFT,
    "OD": curses.KEY_LEFT,
    "[Z": curses.KEY_BTAB,
    "[1~": curses.KEY_HOME,
    "[7~": curses.KEY_HOME,
    "[H": curses.KEY_HOME,
    "OH": curses.KEY_HOME,
    "[4~": curses.KEY_END,
    "[8~": curses.KEY_END,
    "[F": curses.KEY_END,
    "OF": curses.KEY_END,
    "[2~": curses.KEY_IC,
    "[3~": curses.KEY_DC,
    "[5~": curses.KEY_PPAGE,
    "[6~": curses.KEY_NPAGE,
}


def read_key(stdscr, restore_timeout: int = 250) -> int:
    """
    Reads a keypress from curses and reliably decodes multi-byte VT100/VT220/FabGL
    escape sequences (such as F5-F12 \\x1b[15~ .. \\x1b[24~) so they never get
    misinterpreted as a bare ESC (27) or leak trailing '[17~' characters.
    """
    try:
        ch = stdscr.getch()
    except curses.error:
        return -1

    if ch == -1:
        return -1

    # Swallow trailing LF/NUL if Telnet sends CR+LF or CR+NUL on Enter
    if ch == 13:
        try:
            stdscr.timeout(10)
            nxt = stdscr.getch()
            if nxt not in (-1, 10, 0):
                curses.ungetch(nxt)
        except curses.error:
            pass
        finally:
            stdscr.timeout(restore_timeout)
        return 10

    if ch != 27:
        return ch

    # Received ESC (27): check if an escape sequence immediately follows
    try:
        stdscr.timeout(60)
        nxt = stdscr.getch()
    except curses.error:
        nxt = -1

    if nxt == -1 or nxt == 27:
        stdscr.timeout(restore_timeout)
        return 27

    seq = chr(nxt)
    try:
        if nxt == ord("["):
            stdscr.timeout(40)
            for _ in range(10):
                b = stdscr.getch()
                if b == -1:
                    break
                seq += chr(b)
                # Linux console F1-F5 start with '[[A'..'[[E'
                if seq == "[[":
                    continue
                if 64 <= b <= 126:
                    break
        elif nxt == ord("O"):
            stdscr.timeout(40)
            b = stdscr.getch()
            if b != -1:
                seq += chr(b)
    except curses.error:
        pass
    finally:
        stdscr.timeout(restore_timeout)

    return ESC_SEQ_MAP.get(seq, -1)


def safe_addstr(win, y: int, x: int, text: str, attr: int = curses.A_NORMAL):
    max_y, max_x = win.getmaxyx()
    if y < 0 or y >= max_y or x < 0 or x >= max_x:
        return
    avail = max_x - x
    if avail <= 0:
        return
    to_print = text[:avail]
    try:
        if y == max_y - 1 and len(to_print) == avail:
            to_print = to_print[:-1]
        win.addstr(y, x, to_print, attr)
    except curses.error:
        pass


def draw_5250_box(win, start_y: int, start_x: int, height: int, width: int, title: str = ""):
    max_y, max_x = win.getmaxyx()
    for r in range(height):
        y = start_y + r
        if y < 0 or y >= max_y:
            continue
        if r == 0 or r == height - 1:
            border = "+" + "=" * max(0, width - 2) + "+"
            safe_addstr(win, y, start_x, border, curses.A_BOLD)
        else:
            safe_addstr(win, y, start_x, "|" + " " * max(0, width - 2) + "|")
    if title:
        hdr = f" {title} "
        tx = start_x + max(2, (width - len(hdr)) // 2)
        safe_addstr(win, start_y, tx, hdr, curses.A_BOLD | curses.A_REVERSE)


# =============================================================================
# SCREEN: Work with Documents in Folders (WRKDOC - Option 6 & Option 9 Picker)
# =============================================================================
class DocumentManagerScreen:
    """
    IBM OfficeVision/400 'Work with Documents in Folders' (WRKDOC) screen.
    Browses local + Google Drive documents and launches the OV/400 Word Processor.
    """

    def __init__(
        self,
        gdrive: GoogleDriveManager,
        system_name: str = "AS400",
        current_user: str = "QUSER",
        launch_external_cb: Optional[Callable] = None,
    ):
        self.gdrive = gdrive
        self.system_name = system_name
        self.current_user = current_user
        self.launch_external_cb = launch_external_cb

        self.current_folder = "/My Drive"
        self.search_filter = ""
        self.cursor_idx = 0
        self.scroll_offset = 0
        self.opt_inputs: Dict[str, str] = {}
        self.status_message = ""
        self.show_help = False

    def run(self, stdscr, direct_editor: bool = False):
        """
        If direct_editor=True (Option 9 from OFCMAIN), opens a quick prompt to
        create/select a document or immediately open the Word Processor.
        """
        if direct_editor:
            docs = [e for e in self.gdrive.list_documents(self.current_folder) if not e.is_folder]
            choice = self._prompt_wordproc_entry(stdscr, docs)
            if choice is None:
                return
            editor = WordProcessorScreen(
                self.gdrive,
                choice,
                system_name=self.system_name,
                current_user=self.current_user,
            )
            editor.run(stdscr)
            self.status_message = editor.exit_status_message
            return

        while True:
            max_y, max_x = stdscr.getmaxyx()
            stdscr.erase()

            items = self._get_filtered_items()
            if self.cursor_idx >= len(items):
                self.cursor_idx = max(0, len(items) - 1)

            self._draw_wrkdoc(stdscr, max_y, max_x, items)
            if self.show_help:
                self._draw_help_modal(stdscr, max_y, max_x)

            stdscr.refresh()

            ch = read_key(stdscr)
            if ch == -1:
                continue

            if self.show_help:
                if ch in (27, 10, 13, ord(" "), curses.KEY_F3, curses.KEY_F12, ord("q")):
                    self.show_help = False
                continue

            if not self._handle_input(stdscr, ch, items, max_y):
                break

    def _get_filtered_items(self) -> List[DocumentEntry]:
        items = self.gdrive.list_documents(self.current_folder)
        if self.search_filter:
            q = self.search_filter.upper()
            items = [
                e
                for e in items
                if q in e.name.upper() or q in e.description.upper()
            ]
        return items

    def _prompt_wordproc_entry(
        self, stdscr, docs: List[DocumentEntry]
    ) -> Optional[DocumentEntry]:
        """
         IBM OV/400 'Word Processing' (Option 9) entry screen:
         Allows creating a new document, opening the last revised document,
         or choosing from existing documents.
        """
        sel_idx = 0
        new_name = ""
        new_desc = ""
        active_field = 0  # 0 = list selection, 1 = new doc name, 2 = new doc desc
        msg = "Select a document with [Up/Down]+[Enter], or press [F6] (or [N]) for New Document."

        while True:
            max_y, max_x = stdscr.getmaxyx()
            stdscr.erase()

            # Header
            safe_addstr(stdscr, 0, 1, "OVWP01", curses.A_BOLD)
            title = "Word Processing - Select or Create Document"
            safe_addstr(stdscr, 0, (max_x - len(title)) // 2, title, curses.A_BOLD)
            sys_hdr = f"System: {self.system_name}"
            safe_addstr(stdscr, 0, max_x - len(sys_hdr) - 2, sys_hdr)

            safe_addstr(stdscr, 1, 1, f"Folder . . . . :   {self.current_folder}")
            gstatus = f"Google Drive: {self.gdrive.get_status_summary()}"
            safe_addstr(stdscr, 1, max_x - len(gstatus) - 2, gstatus, curses.A_BOLD)

            # New Document Form at Top
            safe_addstr(stdscr, 3, 2, "Create New Document (or leave blank to open selected document below):")
            safe_addstr(stdscr, 4, 4, "New Document Name  . . . . : ")
            name_attr = curses.A_BOLD | curses.A_REVERSE if active_field == 1 else curses.A_UNDERLINE
            safe_addstr(stdscr, 4, 33, f"{new_name:<28}"[:28], name_attr)

            safe_addstr(stdscr, 5, 4, "Document Description . . . : ")
            desc_attr = curses.A_BOLD | curses.A_REVERSE if active_field == 2 else curses.A_UNDERLINE
            safe_addstr(stdscr, 5, 33, f"{new_desc:<42}"[:42], desc_attr)

            # Existing Documents List
            safe_addstr(stdscr, 7, 2, "Existing Documents in Folder:", curses.A_BOLD)
            hdr = f"   {'Document':<24} {'Description':<32} {'Revised':<15} {'Cloud':<10}"
            safe_addstr(stdscr, 8, 2, hdr[: max_x - 4], curses.A_UNDERLINE | curses.A_BOLD)

            list_rows = max(6, max_y - 14)
            start_idx = max(0, sel_idx - list_rows + 1) if sel_idx >= list_rows else 0

            cursor_y, cursor_x = 9, 2
            if not docs:
                safe_addstr(stdscr, 10, 4, "(No documents in folder yet. Type a New Document Name above and press Enter.)")
            else:
                for i in range(list_rows):
                    idx = start_idx + i
                    if idx >= len(docs):
                        break
                    d = docs[idx]
                    y = 9 + i
                    is_sel = (active_field == 0 and idx == sel_idx)
                    marker = "=>" if is_sel else "  "
                    cloud_tag = "GDoc Sync" if (d.synced and not d.dirty) else ("Modified" if d.dirty else "Local")
                    row_str = f"{marker} {d.name:<24.24} {d.description:<32.32} {d.updated_at:<15.15} {cloud_tag:<10}"
                    attr = curses.A_REVERSE | curses.A_BOLD if is_sel else curses.A_NORMAL
                    safe_addstr(stdscr, y, 2, row_str[: max_x - 4], attr)
                    if is_sel:
                        cursor_y, cursor_x = y, 2

            # Bottom keys
            safe_addstr(
                stdscr,
                max_y - 2,
                1,
                "Enter=Open/Create   Tab=Switch Field   F5=Sync GDrive   F6/N=New Doc   F10=GDrive Setup   F12=Cancel",
            )
            safe_addstr(stdscr, max_y - 1, 1, msg[: max_x - 2], curses.A_BOLD)

            if active_field == 1:
                cursor_y, cursor_x = 4, 33 + min(len(new_name), 27)
            elif active_field == 2:
                cursor_y, cursor_x = 5, 33 + min(len(new_desc), 41)

            try:
                stdscr.move(cursor_y, min(cursor_x, max_x - 2))
            except curses.error:
                pass

            stdscr.refresh()
            ch = read_key(stdscr)
            if ch == -1:
                continue

            if ch in (curses.KEY_F3, curses.KEY_F12, 27):
                if active_field in (1, 2) and (new_name or new_desc):
                    new_name = ""
                    new_desc = ""
                    active_field = 0
                    msg = "New document cancelled. Select an existing document or press [F6] for New Doc."
                    continue
                return None

            if ch == curses.KEY_F5:
                ok, sync_msg = self.gdrive.sync_all_from_gdrive()
                msg = sync_msg
                docs = [e for e in self.gdrive.list_documents(self.current_folder) if not e.is_folder]
                sel_idx = min(sel_idx, max(0, len(docs) - 1))
                continue

            if ch == curses.KEY_F10:
                cfg = GoogleDriveSetupScreen(self.gdrive, self.system_name, self.launch_external_cb)
                cfg.run(stdscr)
                docs = [e for e in self.gdrive.list_documents(self.current_folder) if not e.is_folder]
                msg = f"Google Drive status: {self.gdrive.get_status_summary()}"
                continue

            if ch == curses.KEY_F6 or (active_field == 0 and ch in (ord("n"), ord("N"), ord("1"))):
                if active_field != 1:
                    active_field = 1
                    msg = "Type a New Document Name and press [Enter] (or press [Enter] now for UNTITLED)."
                else:
                    doc_title = new_name.strip() or f"UNTITLED_{datetime.datetime.now().strftime('%H%M%S')}"
                    return self.gdrive.create_document(
                        name=doc_title,
                        description=new_desc.strip() or "OV/400 Document",
                        folder=self.current_folder,
                    )
                continue

            if ch == 9:  # Tab
                active_field = (active_field + 1) % 3
                continue

            if ch == curses.KEY_BTAB:
                active_field = (active_field - 1) % 3
                continue

            if active_field == 0:
                if ch == curses.KEY_UP and docs:
                    sel_idx = (sel_idx - 1) % len(docs)
                elif ch == curses.KEY_DOWN and docs:
                    sel_idx = (sel_idx + 1) % len(docs)
                elif ch in (10, 13, curses.KEY_ENTER):
                    if new_name.strip():
                        return self.gdrive.create_document(
                            name=new_name.strip(),
                            description=new_desc.strip() or "OV/400 Document",
                            folder=self.current_folder,
                        )
                    if docs:
                        return docs[sel_idx]
                    active_field = 1
            else:
                if ch in (10, 13, curses.KEY_ENTER):
                    doc_title = new_name.strip() or f"UNTITLED_{datetime.datetime.now().strftime('%H%M%S')}"
                    return self.gdrive.create_document(
                        name=doc_title,
                        description=new_desc.strip() or "OV/400 Document",
                        folder=self.current_folder,
                    )
                elif ch in (curses.KEY_BACKSPACE, 127, 8):
                    if active_field == 1 and new_name:
                        new_name = new_name[:-1]
                    elif active_field == 2 and new_desc:
                        new_desc = new_desc[:-1]
                elif 32 <= ch <= 126:
                    if active_field == 1 and len(new_name) < 28:
                        new_name += chr(ch)
                    elif active_field == 2 and len(new_desc) < 42:
                        new_desc += chr(ch)

    def _draw_wrkdoc(self, win, max_y: int, max_x: int, items: List[DocumentEntry]):
        try:
            curses.curs_set(1)
        except curses.error:
            pass

        # Row 0: Header
        safe_addstr(win, 0, 1, "WRKDOC", curses.A_BOLD)
        title = "Work with Documents in Folders"
        safe_addstr(win, 0, (max_x - len(title)) // 2, title, curses.A_BOLD)
        sys_hdr = f"System: {self.system_name}"
        safe_addstr(win, 0, max_x - len(sys_hdr) - 2, sys_hdr)

        # Row 1: Folder & Google Drive Status
        safe_addstr(win, 1, 1, f"Folder . . . . :   {self.current_folder}")
        gd_status = f"GDrive: {self.gdrive.get_status_summary()}"
        safe_addstr(win, 1, max_x - len(gd_status) - 2, gd_status, curses.A_BOLD)

        if self.search_filter:
            safe_addstr(win, 2, 1, f"Filter . . . . :   '{self.search_filter}' (Press F17 to clear)")

        # Instructions (Rows 3-4)
        safe_addstr(win, 3, 1, "Type options, press Enter.  (Or highlight a document and press Enter to Revise)")
        safe_addstr(
            win,
            4,
            2,
            "1=Create  2=Revise  3=Copy  4=Delete  5=Display  7=Rename  8=Sync/Info",
            curses.A_BOLD,
        )

        # Subfile Table Header (Row 6)
        hdr = f" Opt  {'Document / Folder':<24} {'Type':<6} {'Revised':<16} {'Sync':<8} {'Description':<22}"
        safe_addstr(win, 6, 1, hdr[: max_x - 2], curses.A_UNDERLINE | curses.A_BOLD)

        table_start_y = 7
        table_height = max(5, max_y - 12)

        if self.cursor_idx < self.scroll_offset:
            self.scroll_offset = self.cursor_idx
        elif self.cursor_idx >= self.scroll_offset + table_height:
            self.scroll_offset = self.cursor_idx - table_height + 1

        cursor_screen_y = table_start_y
        cursor_screen_x = 3

        if not items:
            safe_addstr(
                win,
                table_start_y + 1,
                4,
                "(No documents found. Press F6 to create a document or F5 to sync Google Drive.)",
            )
        else:
            for row_i in range(table_height):
                idx = self.scroll_offset + row_i
                if idx >= len(items):
                    break
                entry = items[idx]
                y = table_start_y + row_i
                is_cur = idx == self.cursor_idx

                opt_val = self.opt_inputs.get(entry.doc_id, "")
                opt_cell = f"_{opt_val}_" if len(opt_val) == 0 else f" {opt_val} "

                doc_type = "FLR" if entry.is_folder else ("GDOC" if entry.gdrive_id else "DOC")
                if entry.is_folder:
                    sync_str = "Folder"
                elif entry.synced and not entry.dirty:
                    sync_str = "Synced"
                elif entry.dirty and entry.gdrive_id:
                    sync_str = "Modified"
                else:
                    sync_str = "Local"

                name_disp = f"[{entry.name}]" if entry.is_folder else entry.name
                row_text = (
                    f" {opt_cell}  {name_disp:<24.24} {doc_type:<6} "
                    f"{entry.updated_at:<16.16} {sync_str:<8} {entry.description:<22.22}"
                )
                attr = curses.A_REVERSE | curses.A_BOLD if is_cur else curses.A_NORMAL
                safe_addstr(win, y, 1, row_text[: max_x - 2], attr)

                if is_cur:
                    cursor_screen_y = y
                    cursor_screen_x = 3

        # More / Bottom indicator
        more_y = table_start_y + table_height
        more_txt = "More..." if (self.scroll_offset + table_height < len(items)) else "Bottom"
        safe_addstr(win, more_y, max_x - len(more_txt) - 3, more_txt, curses.A_BOLD)

        # Function Keys & Status Bar
        fkey1_y = max_y - 3
        fkey2_y = max_y - 2
        msg_y = max_y - 1

        safe_addstr(
            win,
            fkey1_y,
            1,
            "F1=Help   F3=Exit   F5=Sync GDrive   F6=New Doc   F7=New Folder   F10=GDrive Setup",
        )
        safe_addstr(
            win,
            fkey2_y,
            1,
            "F12=Cancel/Up Folder   F17=Search/Filter   Enter=Revise/Open",
        )

        if self.status_message:
            safe_addstr(win, msg_y, 1, self.status_message[: max_x - 2], curses.A_BOLD)

        try:
            win.move(cursor_screen_y, cursor_screen_x)
        except curses.error:
            pass

    def _handle_input(
        self, stdscr, ch: int, items: List[DocumentEntry], max_y: int
    ) -> bool:
        if ch in (curses.KEY_F1, ord("?")):
            self.show_help = True
            return True

        if ch == curses.KEY_F3:
            return False

        if ch in (curses.KEY_F12, 27):
            if self.current_folder != "/My Drive":
                parent = os.path.dirname(self.current_folder.rstrip("/"))
                self.current_folder = parent if parent and parent != "/" else "/My Drive"
                self.cursor_idx = 0
                self.status_message = f"Returned to folder {self.current_folder}."
                return True
            return False

        if ch == curses.KEY_UP and items:
            self.cursor_idx = (self.cursor_idx - 1) % len(items)
            return True

        if ch == curses.KEY_DOWN and items:
            self.cursor_idx = (self.cursor_idx + 1) % len(items)
            return True

        if ch == curses.KEY_PPAGE and items:
            self.cursor_idx = max(0, self.cursor_idx - 10)
            return True

        if ch == curses.KEY_NPAGE and items:
            self.cursor_idx = min(len(items) - 1, self.cursor_idx + 10)
            return True

        # F5 (or S) = Sync with Google Drive
        if ch in (curses.KEY_F5, ord("s"), ord("S")):
            self.status_message = "Synchronizing with Google Drive..."
            stdscr.refresh()
            ok, msg = self.gdrive.sync_all_from_gdrive()
            self.status_message = msg
            return True

        # F6 (or 1 / N) = Create New Document
        if ch in (curses.KEY_F6, ord("1"), ord("n"), ord("N")):
            self._modal_create_document(stdscr)
            return True

        # F7 (or F) = Create New Folder
        if ch in (curses.KEY_F7, ord("f"), ord("F")):
            self._modal_create_folder(stdscr)
            return True

        # F10 (or G) = Configure Google Drive
        if ch in (curses.KEY_F10, ord("g"), ord("G")):
            cfg = GoogleDriveSetupScreen(
                self.gdrive, self.system_name, self.launch_external_cb
            )
            cfg.run(stdscr)
            self.status_message = f"Google Drive status: {self.gdrive.get_status_summary()}"
            return True

        # F17 (Shift+F5) or '/' = Search / Filter
        if ch in (curses.KEY_F17, ord("/")):
            self._modal_search_filter(stdscr)
            return True

        # Backspace clears option on current item
        if ch in (curses.KEY_BACKSPACE, 127, 8) and items:
            cur_id = items[self.cursor_idx].doc_id
            self.opt_inputs.pop(cur_id, None)
            return True

        # Option numbers 2, 3, 4, 5, 7, 8 typed on current item
        if items and ch in (ord("2"), ord("3"), ord("4"), ord("5"), ord("7"), ord("8")):
            cur_entry = items[self.cursor_idx]
            self.opt_inputs[cur_entry.doc_id] = chr(ch)
            return True

        # Enter = Execute pending option code(s) or open highlighted item
        if ch in (10, 13, curses.KEY_ENTER):
            if not items:
                self._modal_create_document(stdscr)
                return True

            cur_entry = items[self.cursor_idx]
            opt = self.opt_inputs.pop(cur_entry.doc_id, "").strip()
            if not opt:
                opt = "2"  # Default action on Enter is Revise/Open

            self._execute_item_option(stdscr, cur_entry, opt)
            return True

        return True

    def _execute_item_option(self, stdscr, entry: DocumentEntry, opt: str):
        if entry.is_folder and opt in ("2", "5"):
            self.current_folder = f"{self.current_folder.rstrip('/')}/{entry.name}"
            self.cursor_idx = 0
            self.status_message = f"Opened folder {self.current_folder}"
            return

        if opt in ("2", "5"):
            read_only = opt == "5"
            # If this is a remote Google Doc not yet downloaded locally, pull it first
            if entry.gdrive_id and (not entry.local_path or not os.path.exists(entry.local_path)):
                self.status_message = f"Downloading '{entry.name}' from Google Docs..."
                stdscr.refresh()
                self.gdrive.pull_document(entry)

            editor = WordProcessorScreen(
                self.gdrive,
                entry,
                system_name=self.system_name,
                current_user=self.current_user,
                read_only=read_only,
            )
            editor.run(stdscr)
            self.status_message = editor.exit_status_message
        elif opt == "3":
            self._modal_copy_document(stdscr, entry)
        elif opt == "4":
            self._modal_delete_document(stdscr, entry)
        elif opt == "7":
            self._modal_rename_document(stdscr, entry)
        elif opt == "8":
            self._modal_sync_and_details(stdscr, entry)
        else:
            self.status_message = f"CPD0030: Option '{opt}' is not valid."

    # =========================================================================
    # Modals for WRKDOC
    # =========================================================================
    def _prompt_two_fields(
        self,
        stdscr,
        title: str,
        label1: str,
        val1: str,
        label2: str,
        val2: str,
    ) -> Optional[Tuple[str, str]]:
        idx = 0
        fields = [val1, val2]
        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw, bh = min(68, max_x - 6), 11
            sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
            draw_5250_box(stdscr, sy, sx, bh, bw, title)

            safe_addstr(stdscr, sy + 2, sx + 3, f"{label1:<22}: ")
            a0 = curses.A_BOLD | curses.A_REVERSE if idx == 0 else curses.A_UNDERLINE
            safe_addstr(stdscr, sy + 2, sx + 27, f"{fields[0]:<34}"[:34], a0)

            safe_addstr(stdscr, sy + 4, sx + 3, f"{label2:<22}: ")
            a1 = curses.A_BOLD | curses.A_REVERSE if idx == 1 else curses.A_UNDERLINE
            safe_addstr(stdscr, sy + 4, sx + 27, f"{fields[1]:<34}"[:34], a1)

            safe_addstr(
                stdscr,
                sy + bh - 2,
                sx + 3,
                "Enter=Confirm   Tab=Next Field   F12/Esc=Cancel",
            )
            stdscr.refresh()

            ch = read_key(stdscr)
            if ch in (27, curses.KEY_F12, curses.KEY_F3):
                return None
            if ch in (9, curses.KEY_DOWN, curses.KEY_UP, curses.KEY_BTAB):
                idx = 1 - idx
            elif ch in (10, 13, curses.KEY_ENTER):
                return fields[0].strip(), fields[1].strip()
            elif ch in (curses.KEY_BACKSPACE, 127, 8):
                if fields[idx]:
                    fields[idx] = fields[idx][:-1]
            elif 32 <= ch <= 126 and len(fields[idx]) < 34:
                fields[idx] += chr(ch)

    def _modal_create_document(self, stdscr):
        res = self._prompt_two_fields(
            stdscr,
            "Create Document (CRTDOC)",
            "Document Name",
            "",
            "Description",
            "OfficeVision/400 Document",
        )
        if not res or not res[0]:
            self.status_message = "Create document cancelled."
            return
        entry = self.gdrive.create_document(
            name=res[0], description=res[1], folder=self.current_folder
        )
        editor = WordProcessorScreen(
            self.gdrive,
            entry,
            system_name=self.system_name,
            current_user=self.current_user,
        )
        editor.run(stdscr)
        self.status_message = editor.exit_status_message

    def _modal_create_folder(self, stdscr):
        res = self._prompt_two_fields(
            stdscr,
            "Create Folder (CRTFLR)",
            "Folder Name",
            "",
            "Parent Folder",
            self.current_folder,
        )
        if not res or not res[0]:
            return
        ok, msg = self.gdrive.create_folder(res[0], res[1] or self.current_folder)
        self.status_message = msg

    def _modal_copy_document(self, stdscr, entry: DocumentEntry):
        res = self._prompt_two_fields(
            stdscr,
            "Copy Document (CPYDOC)",
            "From Document",
            entry.name,
            "To New Name",
            f"{entry.name}_COPY",
        )
        if not res or not res[1]:
            return
        ok, msg = self.gdrive.copy_document(entry, res[1])
        self.status_message = msg

    def _modal_rename_document(self, stdscr, entry: DocumentEntry):
        res = self._prompt_two_fields(
            stdscr,
            "Rename Document (RNMDOC)",
            "New Document Name",
            entry.name,
            "New Description",
            entry.description,
        )
        if not res or not res[0]:
            return
        ok, msg = self.gdrive.rename_document(entry, res[0], res[1])
        self.status_message = msg

    def _modal_delete_document(self, stdscr, entry: DocumentEntry):
        max_y, max_x = stdscr.getmaxyx()
        bw, bh = min(64, max_x - 6), 9
        sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
        draw_5250_box(stdscr, sy, sx, bh, bw, "Confirm Delete of Document (DLTDOC)")
        safe_addstr(stdscr, sy + 2, sx + 3, f"Document . . : {entry.name}", curses.A_BOLD)
        safe_addstr(stdscr, sy + 3, sx + 3, f"Folder . . . : {entry.folder}")
        safe_addstr(
            stdscr,
            sy + 5,
            sx + 3,
            "Press [Y] or [Enter] to Delete, or [N]/[F12] to Cancel.",
            curses.A_BOLD,
        )
        stdscr.refresh()
        while True:
            ch = read_key(stdscr)
            if ch in (ord("y"), ord("Y"), 10, 13, curses.KEY_ENTER):
                ok, msg = self.gdrive.delete_document(entry)
                self.status_message = msg
                break
            if ch in (ord("n"), ord("N"), 27, curses.KEY_F12, curses.KEY_F3):
                self.status_message = "Delete cancelled."
                break

    def _modal_search_filter(self, stdscr):
        res = self._prompt_two_fields(
            stdscr,
            "Search / Filter Documents",
            "Search Text",
            self.search_filter,
            "Folder Scope",
            self.current_folder,
        )
        if res is not None:
            self.search_filter = res[0]
            self.cursor_idx = 0
            self.status_message = (
                f"Filter set to '{self.search_filter}'."
                if self.search_filter
                else "Search filter cleared."
            )

    def _modal_sync_and_details(self, stdscr, entry: DocumentEntry):
        msg = "Press [P] to Push to Google Docs, [G] to Pull from Google Docs, or [Enter] to close."
        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw, bh = min(72, max_x - 6), 14
            sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
            draw_5250_box(stdscr, sy, sx, bh, bw, "Document Details & Cloud Sync (DOCDTL)")

            lines = [
                f"Document ID . . . . : {entry.doc_id}",
                f"Document Name . . . : {entry.name}",
                f"Folder  . . . . . . : {entry.folder}",
                f"Description . . . . : {entry.description}",
                f"Google Drive ID . . : {entry.gdrive_id or '(Not yet uploaded)'}",
                f"Cloud Sync State  . : {'Synced' if (entry.synced and not entry.dirty) else 'Modified / Unsynced'}",
                f"Last Revised  . . . : {entry.updated_at} ({entry.size} bytes)",
                f"Local Cache Path  . : {entry.local_path}",
            ]
            for i, line in enumerate(lines):
                safe_addstr(stdscr, sy + 2 + i, sx + 3, line[: bw - 6])

            safe_addstr(stdscr, sy + bh - 2, sx + 3, msg[: bw - 6], curses.A_BOLD)
            stdscr.refresh()

            ch = read_key(stdscr)
            if ch in (10, 13, 27, curses.KEY_F12, curses.KEY_F3, ord(" ")):
                break
            if ch in (ord("p"), ord("P")):
                msg = "Pushing document to Google Docs..."
                stdscr.refresh()
                ok, res_msg = self.gdrive.push_document(entry)
                msg = res_msg
                self.status_message = res_msg
            elif ch in (ord("g"), ord("G")):
                msg = "Pulling document from Google Docs..."
                stdscr.refresh()
                ok, res_msg = self.gdrive.pull_document(entry)
                msg = res_msg
                self.status_message = res_msg

    def _draw_help_modal(self, win, max_y: int, max_x: int):
        bw, bh = min(74, max_x - 6), 18
        sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
        draw_5250_box(win, sy, sx, bh, bw, "Work with Documents (WRKDOC) - Help")
        help_lines = [
            "Using the OfficeVision/400 Document & Google Drive Browser:",
            "",
            "  * Highlight any document with [Up]/[Down] and press [Enter] to edit.",
            "  * Or type an IBM 5250 Option number in the 'Opt' column:",
            "      1 = Create a new document        2 = Revise (Edit) document",
            "      3 = Copy document                4 = Delete document",
            "      5 = Display (Read-Only)          7 = Rename document",
            "      8 = Document Details & Push/Pull Google Drive Sync",
            "",
            "Google Drive Integration:",
            "  * Press [F10] to configure Google Drive (OAuth API or rclone).",
            "  * Press [F5] to refresh and pull your latest Google Docs list.",
            "  * Inside the Word Processor, press [F2] to save & sync to Google Docs.",
            "",
            "Press [Enter], [Space], or [F12] to close Help.",
        ]
        for i, line in enumerate(help_lines):
            safe_addstr(win, sy + 2 + i, sx + 3, line[: bw - 6])


# =============================================================================
# SCREEN: IBM OfficeVision/400 Word Processor (EDTDOC - Option 9)
# =============================================================================
class WordProcessorScreen:
    """
    IBM OfficeVision/400 (DisplayWrite/400) Fullscreen Word Processor for 90x43 MDA.
    """

    def __init__(
        self,
        gdrive: GoogleDriveManager,
        entry: DocumentEntry,
        system_name: str = "AS400",
        current_user: str = "QUSER",
        read_only: bool = False,
    ):
        self.gdrive = gdrive
        self.entry = entry
        self.system_name = system_name
        self.current_user = current_user
        self.read_only = read_only

        raw_text = self.gdrive.read_document_text(self.entry)
        self.lines: List[str] = raw_text.splitlines() if raw_text else [""]
        if not self.lines:
            self.lines = [""]

        self.cy = 0
        self.cx = 0
        self.scroll_y = 0
        self.insert_mode = True
        self.formatted_view = True  # True = render bold/underline/reverse on MDA
        self.wrap_width = int(self.gdrive.config.get("page_width", 72))
        self.auto_wrap = True
        self.modified = False
        self.clipboard_line = ""

        # Undo / Redo stacks: list of (lines_copy, cy, cx)
        self.undo_stack: List[Tuple[List[str], int, int]] = []
        self.redo_stack: List[Tuple[List[str], int, int]] = []

        self.status_message = (
            f"Viewing '{self.entry.name}' (Read-Only)."
            if self.read_only
            else f"Editing '{self.entry.name}'. Press F1=Help, F2=Save/Sync, F3=Save&Exit."
        )
        self.exit_status_message = f"Returned from Word Processing ({self.entry.name})."
        self.show_help = False
        self.last_search = ""

    def _save_undo(self):
        snapshot = (list(self.lines), self.cy, self.cx)
        if self.undo_stack and self.undo_stack[-1][0] == snapshot[0]:
            return
        self.undo_stack.append(snapshot)
        if len(self.undo_stack) > 100:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def _undo(self):
        if not self.undo_stack:
            self.status_message = "Nothing to undo."
            return
        self.redo_stack.append((list(self.lines), self.cy, self.cx))
        prev_lines, prev_cy, prev_cx = self.undo_stack.pop()
        self.lines = prev_lines
        self.cy = min(prev_cy, len(self.lines) - 1)
        self.cx = min(prev_cx, len(self.lines[self.cy]))
        self.modified = True
        self.status_message = "Undo performed."

    def _redo(self):
        if not self.redo_stack:
            self.status_message = "Nothing to redo."
            return
        self.undo_stack.append((list(self.lines), self.cy, self.cx))
        next_lines, next_cy, next_cx = self.redo_stack.pop()
        self.lines = next_lines
        self.cy = min(next_cy, len(self.lines) - 1)
        self.cx = min(next_cx, len(self.lines[self.cy]))
        self.modified = True
        self.status_message = "Redo performed."

    def run(self, stdscr):
        while True:
            max_y, max_x = stdscr.getmaxyx()
            stdscr.erase()

            self._draw_editor(stdscr, max_y, max_x)
            if self.show_help:
                self._draw_help_modal(stdscr, max_y, max_x)

            stdscr.refresh()

            ch = read_key(stdscr)
            if ch == -1:
                continue

            if self.show_help:
                if ch in (27, 10, 13, ord(" "), curses.KEY_F3, curses.KEY_F12, ord("q")):
                    self.show_help = False
                continue

            if not self._handle_key(stdscr, ch, max_y, max_x):
                break

    def _draw_editor(self, win, max_y: int, max_x: int):
        try:
            curses.curs_set(1)
        except curses.error:
            pass

        # Calculate document stats
        full_text = "\n".join(self.lines)
        word_count = len(full_text.split())
        lines_per_page = 55
        cur_page = (self.cy // lines_per_page) + 1
        total_pages = max(1, ((len(self.lines) - 1) // lines_per_page) + 1)

        # Row 0: IBM DisplayWrite/400 Header Bar
        mode_tag = "READ-ONLY" if self.read_only else ("MODIFIED" if self.modified else ("GDRIVE SYNCED" if self.entry.synced else "SAVED"))
        left_hdr = f"EDTDOC   Document: {self.entry.name[:24]} [{mode_tag}]"
        right_hdr = f"Ln {self.cy + 1:04d}/{len(self.lines):04d}  Col {self.cx + 1:02d}  Pg {cur_page}/{total_pages}"
        safe_addstr(win, 0, 1, left_hdr, curses.A_BOLD)
        safe_addstr(win, 0, max(1, max_x - len(right_hdr) - 2), right_hdr, curses.A_BOLD)

        # Row 1: Metadata & Formatting Toolbar
        ins_str = "INS" if self.insert_mode else "OVR"
        view_str = "MDA-Rich" if self.formatted_view else "Raw-Text"
        wrap_str = f"Wrap:{self.wrap_width}" if self.auto_wrap else "Wrap:OFF"
        sub_hdr = (
            f"Folder: {self.entry.folder:<14.14} | Words: {word_count:<5d} | "
            f"Mode: {ins_str} | {wrap_str} | View(F11): {view_str}"
        )
        safe_addstr(win, 1, 1, sub_hdr[: max_x - 2])

        # Center paper canvas on 90-column MDA screen
        paper_w = min(max_x - 8, max(60, self.wrap_width + 4))
        gutter_x = max(1, (max_x - paper_w) // 2)
        text_start_x = gutter_x + 2

        # Row 2: IBM DisplayWrite/400 Scale / Margin Ruler Line
        ruler_chars = []
        for col in range(1, self.wrap_width + 1):
            if col == 1:
                ruler_chars.append("L")
            elif col == self.wrap_width:
                ruler_chars.append("R")
            elif col % 10 == 0:
                ruler_chars.append(str((col // 10) % 10))
            elif col % 5 == 0:
                ruler_chars.append("+")
            elif col in (5, 9, 17, 25, 33, 41, 49, 57, 65):
                ruler_chars.append("t")
            else:
                ruler_chars.append(".")
        ruler_str = "".join(ruler_chars)
        safe_addstr(win, 2, gutter_x, "[=" + ruler_str.ljust(paper_w - 4, "=") + "=]", curses.A_REVERSE)
        # Highlight current cursor column on the ruler
        if 0 <= self.cx < len(ruler_str):
            cur_rchar = ruler_str[self.cx]
            safe_addstr(
                win,
                2,
                text_start_x + self.cx,
                cur_rchar,
                curses.A_BOLD | curses.A_UNDERLINE,
            )

        # Document Viewport (Rows 3 .. max_y - 4)
        view_top = 3
        view_height = max(5, max_y - 6)

        if self.cy < self.scroll_y:
            self.scroll_y = self.cy
        elif self.cy >= self.scroll_y + view_height:
            self.scroll_y = self.cy - view_height + 1

        for row_offset in range(view_height):
            line_idx = self.scroll_y + row_offset
            screen_y = view_top + row_offset

            # Draw subtle left and right paper margins on 90-column screen
            safe_addstr(win, screen_y, gutter_x, "|")
            safe_addstr(win, screen_y, gutter_x + paper_w - 1, "|")

            if line_idx >= len(self.lines):
                if line_idx == len(self.lines):
                    eof_marker = "************* End of Document *************"
                    ex = text_start_x + max(0, (self.wrap_width - len(eof_marker)) // 2)
                    safe_addstr(win, screen_y, ex, eof_marker, curses.A_BOLD)
                continue

            line_str = self.lines[line_idx]

            # Page break marker check
            if line_str.strip() == "---":
                pb_label = f"--------------------[ Page Break ]--------------------"
                safe_addstr(
                    win,
                    screen_y,
                    text_start_x,
                    pb_label[: paper_w - 4].center(paper_w - 4, "-"),
                    curses.A_BOLD,
                )
                continue

            if self.formatted_view:
                self._render_formatted_mda_line(win, screen_y, text_start_x, line_str, paper_w - 4)
            else:
                safe_addstr(win, screen_y, text_start_x, line_str[: paper_w - 4])

        # Bottom Function Key Bar & Status Line
        fkey1_y = max_y - 3
        fkey2_y = max_y - 2
        msg_y = max_y - 1

        safe_addstr(
            win,
            fkey1_y,
            1,
            "F1=Help  F2=Save/Sync GDrive  F3=Save&Exit  F4=Spell(aspell)  F5=Find/Replace  F6=Center",
        )
        safe_addstr(
            win,
            fkey2_y,
            1,
            "F8=Reflow Para  F11=View Mode  F12=Exit  ^B=Bold  ^U=Underline  ^R=Highlight  ^Z/^Y=Undo/Redo",
        )
        safe_addstr(win, msg_y, 1, self.status_message[: max_x - 2], curses.A_BOLD)

        # Position hardware cursor inside document
        cursor_screen_y = view_top + (self.cy - self.scroll_y)
        cursor_screen_x = min(gutter_x + paper_w - 2, text_start_x + self.cx)
        try:
            win.move(cursor_screen_y, cursor_screen_x)
        except curses.error:
            pass

    def _render_formatted_mda_line(
        self, win, y: int, start_x: int, line: str, max_width: int
    ):
        """
        Renders inline Markdown/OV400 styling natively using monochrome MDA attributes:
          - # Heading     -> A_BOLD | A_UNDERLINE
          - **bold**      -> A_BOLD
          - __underline__ -> A_UNDERLINE
          - ==highlight== -> A_REVERSE
        Keeps character column positions aligned 1:1 with cursor `cx` so editing in
        formatted mode never shifts cursor offsets!
        """
        stripped = line.lstrip()
        if stripped.startswith("#"):
            safe_addstr(win, y, start_x, line[:max_width], curses.A_BOLD | curses.A_UNDERLINE)
            return

        # Compute per-character attribute map so cursor position always matches 1:1
        n = min(len(line), max_width)
        attrs = [curses.A_NORMAL] * n

        # Highlight **bold**
        for m in re.finditer(r"\*\*(.+?)\*\*", line):
            for i in range(m.start(), min(m.end(), n)):
                attrs[i] |= curses.A_BOLD

        # Highlight __underline__
        for m in re.finditer(r"__(.+?)__", line):
            for i in range(m.start(), min(m.end(), n)):
                attrs[i] |= curses.A_UNDERLINE

        # Highlight ==reverse==
        for m in re.finditer(r"==(.+?)==", line):
            for i in range(m.start(), min(m.end(), n)):
                attrs[i] |= curses.A_REVERSE

        # Render spans of identical attributes for speed
        idx = 0
        while idx < n:
            cur_attr = attrs[idx]
            end_idx = idx + 1
            while end_idx < n and attrs[end_idx] == cur_attr:
                end_idx += 1
            safe_addstr(win, y, start_x + idx, line[idx:end_idx], cur_attr)
            idx = end_idx

    # =========================================================================
    # Word Processor Key Handling & Editing Operations
    # =========================================================================
    def _handle_key(self, stdscr, ch: int, max_y: int, max_x: int) -> bool:
        view_height = max(5, max_y - 6)

        # F1 = Help
        if ch == curses.KEY_F1:
            self.show_help = True
            return True

        # F2 or Ctrl+S (19) = Save & Sync to Google Drive
        if ch in (curses.KEY_F2, 19):
            if self.read_only:
                self.status_message = "Document is opened in Read-Only mode."
                return True
            self._save_document(stdscr, sync_cloud=True)
            return True

        # F3 = Save & Exit
        if ch == curses.KEY_F3:
            if not self.read_only and self.modified:
                self._save_document(stdscr, sync_cloud=self.gdrive.config.get("auto_sync_on_save", True))
            return False

        # F12 or Esc = Exit (prompt if modified)
        if ch in (curses.KEY_F12, 27):
            if not self.read_only and self.modified:
                if not self._confirm_discard(stdscr):
                    return True
            return False

        # F4 = Spell Check (aspell)
        if ch == curses.KEY_F4:
            if self.read_only:
                self.status_message = "Spell check disabled in Read-Only mode."
                return True
            self._run_spellcheck(stdscr)
            return True

        # F5 or Ctrl+F (6) = Find & Replace
        if ch in (curses.KEY_F5, 6):
            self._modal_find_replace(stdscr)
            return True

        # F6 or Ctrl+E (5) = Center Current Line
        if ch in (curses.KEY_F6, 5):
            if not self.read_only:
                self._center_current_line()
            return True

        # F8 or Ctrl+J = Reflow Current Paragraph
        if ch == curses.KEY_F8:
            if not self.read_only:
                self._reflow_current_paragraph()
            return True

        # F11 = Toggle Formatted MDA View vs Raw View
        if ch == curses.KEY_F11:
            self.formatted_view = not self.formatted_view
            mode_lbl = "MDA Formatted View" if self.formatted_view else "Raw Markup View"
            self.status_message = f"Switched to {mode_lbl}."
            return True

        # Undo (Ctrl+Z = 26) / Redo (Ctrl+Y = 25)
        if ch == 26 and not self.read_only:
            self._undo()
            return True
        if ch == 25 and not self.read_only:
            self._redo()
            return True

        # Inline Style Shortcuts: Ctrl+B (2) = Bold, Ctrl+U (21) = Underline, Ctrl+R (18) = Reverse
        if ch == 2 and not self.read_only:
            self._insert_style_wrapper("**", "Bold")
            return True
        if ch == 21 and not self.read_only:
            self._insert_style_wrapper("__", "Underline")
            return True
        if ch == 18 and not self.read_only:
            self._insert_style_wrapper("==", "Highlight")
            return True

        # Cut Line (Ctrl+K = 11) / Paste Line (Ctrl+V = 22)
        if ch == 11 and not self.read_only:
            self._save_undo()
            self.clipboard_line = self.lines[self.cy]
            if len(self.lines) > 1:
                self.lines.pop(self.cy)
                if self.cy >= len(self.lines):
                    self.cy = len(self.lines) - 1
            else:
                self.lines[0] = ""
            self.cx = min(self.cx, len(self.lines[self.cy]))
            self.modified = True
            self.status_message = "Cut line to clipboard (Ctrl+V to paste)."
            return True

        if ch == 22 and not self.read_only:
            if self.clipboard_line:
                self._save_undo()
                self.lines.insert(self.cy + 1, self.clipboard_line)
                self.cy += 1
                self.cx = 0
                self.modified = True
                self.status_message = "Pasted line from clipboard."
            return True

        # Navigation Keys
        if ch == curses.KEY_UP:
            if self.cy > 0:
                self.cy -= 1
                self.cx = min(self.cx, len(self.lines[self.cy]))
            return True

        if ch == curses.KEY_DOWN:
            if self.cy < len(self.lines) - 1:
                self.cy += 1
                self.cx = min(self.cx, len(self.lines[self.cy]))
            return True

        if ch == curses.KEY_LEFT:
            if self.cx > 0:
                self.cx -= 1
            elif self.cy > 0:
                self.cy -= 1
                self.cx = len(self.lines[self.cy])
            return True

        if ch == curses.KEY_RIGHT:
            if self.cx < len(self.lines[self.cy]):
                self.cx += 1
            elif self.cy < len(self.lines) - 1:
                self.cy += 1
                self.cx = 0
            return True

        if ch == curses.KEY_HOME or ch == 1:  # Home or Ctrl+A
            self.cx = 0
            return True

        if ch == curses.KEY_END:
            self.cx = len(self.lines[self.cy])
            return True

        if ch == curses.KEY_PPAGE:
            self.cy = max(0, self.cy - view_height)
            self.cx = min(self.cx, len(self.lines[self.cy]))
            return True

        if ch == curses.KEY_NPAGE:
            self.cy = min(len(self.lines) - 1, self.cy + view_height)
            self.cx = min(self.cx, len(self.lines[self.cy]))
            return True

        if ch == curses.KEY_IC:
            self.insert_mode = not self.insert_mode
            self.status_message = f"Mode changed to {'Insert' if self.insert_mode else 'Overtype'}."
            return True

        if self.read_only:
            return True

        # Editing Keys: Enter, Backspace, Delete, Tab, Printable characters
        if ch in (10, 13, curses.KEY_ENTER):
            self._save_undo()
            cur_line = self.lines[self.cy]
            left_part = cur_line[: self.cx]
            right_part = cur_line[self.cx :]
            self.lines[self.cy] = left_part
            self.lines.insert(self.cy + 1, right_part)
            self.cy += 1
            self.cx = 0
            self.modified = True
            return True

        if ch in (curses.KEY_BACKSPACE, 127, 8):
            self._save_undo()
            cur_line = self.lines[self.cy]
            if self.cx > 0:
                self.lines[self.cy] = cur_line[: self.cx - 1] + cur_line[self.cx :]
                self.cx -= 1
                self.modified = True
            elif self.cy > 0:
                prev_len = len(self.lines[self.cy - 1])
                self.lines[self.cy - 1] += cur_line
                self.lines.pop(self.cy)
                self.cy -= 1
                self.cx = prev_len
                self.modified = True
            return True

        if ch == curses.KEY_DC:
            self._save_undo()
            cur_line = self.lines[self.cy]
            if self.cx < len(cur_line):
                self.lines[self.cy] = cur_line[: self.cx] + cur_line[self.cx + 1 :]
                self.modified = True
            elif self.cy < len(self.lines) - 1:
                self.lines[self.cy] += self.lines.pop(self.cy + 1)
                self.modified = True
            return True

        if ch == 9:  # Tab -> 4 spaces
            self._save_undo()
            cur_line = self.lines[self.cy]
            spaces = "    "
            self.lines[self.cy] = cur_line[: self.cx] + spaces + cur_line[self.cx :]
            self.cx += len(spaces)
            self.modified = True
            return True

        if 32 <= ch <= 126:
            self._save_undo()
            char = chr(ch)
            cur_line = self.lines[self.cy]
            if self.insert_mode or self.cx >= len(cur_line):
                self.lines[self.cy] = cur_line[: self.cx] + char + cur_line[self.cx :]
            else:
                self.lines[self.cy] = cur_line[: self.cx] + char + cur_line[self.cx + 1 :]
            self.cx += 1
            self.modified = True

            # Automatic soft word-wrap when exceeding wrap_width
            if self.auto_wrap and len(self.lines[self.cy]) > self.wrap_width:
                self._auto_wrap_current_line()
            return True

        return True

    def _auto_wrap_current_line(self):
        line = self.lines[self.cy]
        if len(line) <= self.wrap_width:
            return
        # Find last space at or before wrap_width
        split_at = line.rfind(" ", 0, self.wrap_width + 1)
        if split_at <= 0:
            return
        head = line[:split_at]
        tail = line[split_at + 1 :]
        self.lines[self.cy] = head
        self.lines.insert(self.cy + 1, tail)
        if self.cx > split_at:
            self.cy += 1
            self.cx = self.cx - split_at - 1

    def _center_current_line(self):
        self._save_undo()
        stripped = self.lines[self.cy].strip()
        if not stripped:
            return
        pad = max(0, (self.wrap_width - len(stripped)) // 2)
        self.lines[self.cy] = (" " * pad) + stripped
        self.cx = len(self.lines[self.cy])
        self.modified = True
        self.status_message = "Centered line between L and R margins."

    def _reflow_current_paragraph(self):
        if not self.lines[self.cy].strip():
            self.status_message = "Cursor is on a blank line."
            return
        self._save_undo()
        start = self.cy
        while start > 0 and self.lines[start - 1].strip() and not self.lines[start - 1].lstrip().startswith("#"):
            start -= 1
        end = self.cy
        while end < len(self.lines) - 1 and self.lines[end + 1].strip() and not self.lines[end + 1].lstrip().startswith("#"):
            end += 1

        para_text = " ".join(l.strip() for l in self.lines[start : end + 1])
        wrapped = textwrap.wrap(para_text, width=self.wrap_width) or [""]
        self.lines[start : end + 1] = wrapped
        self.cy = min(start + len(wrapped) - 1, len(self.lines) - 1)
        self.cx = min(self.cx, len(self.lines[self.cy]))
        self.modified = True
        self.status_message = f"Paragraph reflowed to {self.wrap_width} columns."

    def _insert_style_wrapper(self, marker: str, style_name: str):
        self._save_undo()
        line = self.lines[self.cy]
        # Wrap word under cursor or insert markers at cursor
        m_len = len(marker)
        if self.cx < len(line) and line[self.cx].isalnum():
            start = self.cx
            while start > 0 and line[start - 1].isalnum():
                start -= 1
            end = self.cx
            while end < len(line) and line[end].isalnum():
                end += 1
            self.lines[self.cy] = line[:start] + marker + line[start:end] + marker + line[end:]
            self.cx = end + m_len * 2
        else:
            self.lines[self.cy] = line[: self.cx] + marker + marker + line[self.cx :]
            self.cx += m_len
        self.modified = True
        self.status_message = f"Inserted {style_name} ({marker}text{marker})."

    def _save_document(self, stdscr, sync_cloud: bool = True):
        self.status_message = "Saving document..."
        stdscr.refresh()
        text = "\n".join(self.lines) + "\n"
        ok, msg = self.gdrive.save_document_text(self.entry, text, sync_cloud=sync_cloud)
        if ok:
            self.modified = False
        self.status_message = msg
        self.exit_status_message = msg

    def _confirm_discard(self, stdscr) -> bool:
        max_y, max_x = stdscr.getmaxyx()
        bw, bh = min(64, max_x - 6), 8
        sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
        draw_5250_box(stdscr, sy, sx, bh, bw, "Unsaved Changes (EDTDOC)")
        safe_addstr(stdscr, sy + 2, sx + 3, f"Document '{self.entry.name}' has unsaved changes.")
        safe_addstr(
            stdscr,
            sy + 4,
            sx + 3,
            "[S]=Save & Exit   [D]=Discard & Exit   [F12/Esc]=Cancel",
            curses.A_BOLD,
        )
        stdscr.refresh()
        while True:
            ch = read_key(stdscr)
            if ch in (ord("s"), ord("S")):
                self._save_document(stdscr, sync_cloud=self.gdrive.config.get("auto_sync_on_save", True))
                return True
            if ch in (ord("d"), ord("D"), ord("y"), ord("Y")):
                return True
            if ch in (27, curses.KEY_F12, ord("n"), ord("N")):
                return False

    # =========================================================================
    # Find & Replace Modal (F5)
    # =========================================================================
    def _modal_find_replace(self, stdscr):
        find_str = self.last_search
        repl_str = ""
        idx = 0
        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw, bh = min(68, max_x - 6), 11
            sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
            draw_5250_box(stdscr, sy, sx, bh, bw, "Find and Replace Text (FNDTXT)")

            safe_addstr(stdscr, sy + 2, sx + 3, "Find Text  . . . . . : ")
            a0 = curses.A_BOLD | curses.A_REVERSE if idx == 0 else curses.A_UNDERLINE
            safe_addstr(stdscr, sy + 2, sx + 26, f"{find_str:<36}"[:36], a0)

            safe_addstr(stdscr, sy + 4, sx + 3, "Replace With (Opt) . : ")
            a1 = curses.A_BOLD | curses.A_REVERSE if idx == 1 else curses.A_UNDERLINE
            safe_addstr(stdscr, sy + 4, sx + 26, f"{repl_str:<36}"[:36], a1)

            safe_addstr(
                stdscr,
                sy + bh - 2,
                sx + 3,
                "Enter=Find Next   F16=Replace All   Tab=Switch   F12=Cancel",
            )
            stdscr.refresh()

            ch = read_key(stdscr)
            if ch in (27, curses.KEY_F12, curses.KEY_F3):
                return
            if ch in (9, curses.KEY_DOWN, curses.KEY_UP, curses.KEY_BTAB):
                idx = 1 - idx
            elif ch == curses.KEY_F16 and find_str:
                self._save_undo()
                count = 0
                for i, l in enumerate(self.lines):
                    if find_str in l:
                        count += l.count(find_str)
                        self.lines[i] = l.replace(find_str, repl_str)
                if count > 0:
                    self.modified = True
                self.status_message = f"Replaced {count} occurrence(s) of '{find_str}'."
                return
            elif ch in (10, 13, curses.KEY_ENTER):
                if not find_str:
                    return
                self.last_search = find_str
                # Search forward from current position
                total = len(self.lines)
                for offset in range(total + 1):
                    r = (self.cy + offset) % total
                    start_c = (self.cx + 1) if offset == 0 else 0
                    pos = self.lines[r].lower().find(find_str.lower(), start_c)
                    if pos != -1:
                        self.cy, self.cx = r, pos
                        if repl_str:
                            self._save_undo()
                            line = self.lines[r]
                            self.lines[r] = line[:pos] + repl_str + line[pos + len(find_str) :]
                            self.modified = True
                            self.status_message = f"Replaced '{find_str}' with '{repl_str}' at Ln {r + 1}."
                        else:
                            self.status_message = f"Found '{find_str}' at Ln {r + 1}, Col {pos + 1}."
                        return
                self.status_message = f"Text '{find_str}' not found."
                return
            elif ch in (curses.KEY_BACKSPACE, 127, 8):
                if idx == 0 and find_str:
                    find_str = find_str[:-1]
                elif idx == 1 and repl_str:
                    repl_str = repl_str[:-1]
            elif 32 <= ch <= 126:
                if idx == 0 and len(find_str) < 36:
                    find_str += chr(ch)
                elif idx == 1 and len(repl_str) < 36:
                    repl_str += chr(ch)

    # =========================================================================
    # Interactive Spell Checker via /usr/bin/aspell (F4)
    # =========================================================================
    def _run_spellcheck(self, stdscr):
        aspell_bin = shutil.which("aspell")
        if not aspell_bin:
            self.status_message = "CPF9801: /usr/bin/aspell is not installed."
            return

        word_re = re.compile(r"\b[A-Za-z']+\b")
        ignored_words = set()
        checked_count = 0
        replaced_count = 0

        for line_idx in range(len(self.lines)):
            col_offset = 0
            while True:
                line = self.lines[line_idx]
                m = word_re.search(line, col_offset)
                if not m:
                    break
                word = m.group(0).strip("'")
                w_start, w_end = m.start(), m.end()
                col_offset = w_end

                if len(word) <= 1 or word.lower() in ignored_words:
                    continue

                checked_count += 1
                suggestions = self._query_aspell_word(aspell_bin, word)
                if suggestions is None:
                    # Word is spelled correctly
                    continue

                # Highlight word in editor
                self.cy = line_idx
                self.cx = w_start
                max_y, max_x = stdscr.getmaxyx()
                stdscr.erase()
                self._draw_editor(stdscr, max_y, max_x)

                action, replacement = self._prompt_spell_modal(
                    stdscr, word, line, suggestions, line_idx + 1
                )
                if action == "cancel":
                    self.status_message = (
                        f"Spell check stopped ({replaced_count} corrected)."
                    )
                    return
                if action == "ignore":
                    ignored_words.add(word.lower())
                elif action == "add":
                    self._add_to_aspell_personal_dict(word)
                    ignored_words.add(word.lower())
                elif action == "replace" and replacement:
                    self._save_undo()
                    cur_l = self.lines[line_idx]
                    # Locate actual word span excluding stripped quotes
                    real_start = cur_l.find(word, w_start)
                    if real_start != -1:
                        self.lines[line_idx] = (
                            cur_l[:real_start]
                            + replacement
                            + cur_l[real_start + len(word) :]
                        )
                        col_offset = real_start + len(replacement)
                        replaced_count += 1
                        self.modified = True

        self.status_message = (
            f"Spell check complete: {checked_count} words checked, {replaced_count} corrected."
        )

    def _query_aspell_word(self, aspell_bin: str, word: str) -> Optional[List[str]]:
        """Returns None if spelled right, or a list of suggestions if misspelled."""
        try:
            res = subprocess.run(
                [aspell_bin, "-a"],
                input=word + "\n",
                capture_output=True,
                text=True,
                timeout=3,
            )
            for out_line in res.stdout.splitlines():
                if out_line.startswith("*") or out_line.startswith("+") or out_line.startswith("-"):
                    return None
                if out_line.startswith("&"):
                    # Format: & original count offset: sugg1, sugg2, ...
                    parts = out_line.split(":", 1)
                    if len(parts) == 2:
                        suggs = [s.strip() for s in parts[1].split(",") if s.strip()]
                        return suggs[:9]
                    return []
                if out_line.startswith("#"):
                    return []
        except Exception:
            return None
        return None

    def _add_to_aspell_personal_dict(self, word: str):
        pws_path = os.path.expanduser("~/.aspell.en.pws")
        try:
            if not os.path.exists(pws_path):
                with open(pws_path, "w", encoding="utf-8") as f:
                    f.write("personal_ws-1.1 en 1\n" + word + "\n")
            else:
                with open(pws_path, "a", encoding="utf-8") as f:
                    f.write(word + "\n")
        except Exception:
            pass

    def _prompt_spell_modal(
        self,
        stdscr,
        word: str,
        context_line: str,
        suggestions: List[str],
        line_num: int,
    ) -> Tuple[str, str]:
        custom_input = ""
        typing_custom = False
        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw, bh = min(72, max_x - 6), 15
            sy, sx = max(2, max_y - bh - 3), (max_x - bw) // 2
            draw_5250_box(stdscr, sy, sx, bh, bw, f"OfficeVision/400 Spell Check (Line {line_num})")

            safe_addstr(stdscr, sy + 2, sx + 3, f"Unrecognized Word : {word}", curses.A_BOLD)
            safe_addstr(stdscr, sy + 3, sx + 3, f"Context           : {context_line.strip()[:bw - 24]}")

            if suggestions:
                safe_addstr(stdscr, sy + 5, sx + 3, "Suggestions (Press 1-9 to Replace):", curses.A_UNDERLINE)
                for idx, s in enumerate(suggestions[:9]):
                    col = sx + 4 + (idx % 3) * 21
                    row = sy + 6 + (idx // 3)
                    safe_addstr(stdscr, row, col, f"{idx + 1}. {s:<16.16}", curses.A_BOLD)
            else:
                safe_addstr(stdscr, sy + 6, sx + 3, "(No dictionary suggestions found)")

            safe_addstr(stdscr, sy + 10, sx + 3, "Custom Replacement: ")
            c_attr = curses.A_BOLD | curses.A_REVERSE if typing_custom else curses.A_UNDERLINE
            safe_addstr(stdscr, sy + 10, sx + 23, f"{custom_input:<30}"[:30], c_attr)

            safe_addstr(
                stdscr,
                sy + bh - 2,
                sx + 3,
                "[1-9]=Replace  [I]=Ignore  [A]=Add to Dict  [R]=Custom  [F12]=Stop",
                curses.A_BOLD,
            )
            stdscr.refresh()

            ch = read_key(stdscr)
            if typing_custom:
                if ch in (10, 13, curses.KEY_ENTER):
                    if custom_input.strip():
                        return "replace", custom_input.strip()
                    typing_custom = False
                elif ch in (27, curses.KEY_F12):
                    typing_custom = False
                elif ch in (curses.KEY_BACKSPACE, 127, 8):
                    custom_input = custom_input[:-1]
                elif 32 <= ch <= 126 and len(custom_input) < 30:
                    custom_input += chr(ch)
                continue

            if ch in (27, curses.KEY_F12, curses.KEY_F3):
                return "cancel", ""
            if ch in (ord("i"), ord("I"), ord(" "), 10, 13):
                return "ignore", ""
            if ch in (ord("a"), ord("A")):
                return "add", ""
            if ch in (ord("r"), ord("R")):
                typing_custom = True
                custom_input = word
                continue
            if ord("1") <= ch <= ord("9"):
                s_idx = ch - ord("1")
                if s_idx < len(suggestions):
                    return "replace", suggestions[s_idx]

    def _draw_help_modal(self, win, max_y: int, max_x: int):
        bw, bh = min(76, max_x - 6), 20
        sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
        draw_5250_box(win, sy, sx, bh, bw, "OfficeVision/400 Word Processor - Help")
        lines = [
            "DisplayWrite/400 & MDA Word Processing Controls:",
            "",
            "  F2 / Ctrl+S  : Save document locally & sync to Google Docs",
            "  F3           : Save & Exit to Document List / Main Menu",
            "  F4           : Spell Check document using /usr/bin/aspell",
            "  F5 / Ctrl+F  : Find and Replace text",
            "  F6 / Ctrl+E  : Center current line between L and R margins",
            "  F8           : Reflow / wrap current paragraph to margin width",
            "  F11          : Toggle Live MDA Rich Formatting vs Raw Markup",
            "  F12 / Esc    : Exit editor (prompts if unsaved changes exist)",
            "",
            "Monochrome MDA Text Styling & Editing:",
            "  Ctrl+B       : Bold (**text**)       Ctrl+U : Underline (__text__)",
            "  Ctrl+R       : Highlight (==text==)  ---    : Page Break line",
            "  Ctrl+Z / ^Y  : Undo / Redo           Ctrl+K / ^V : Cut / Paste line",
            "",
            "Press [Enter], [Space], or [F12] to close Help.",
        ]
        for i, line in enumerate(lines):
            safe_addstr(win, sy + 2 + i, sx + 3, line[: bw - 6])


# =============================================================================
# SCREEN: Configure Google Drive & Google Docs Connection (CFGDRV - F10)
# =============================================================================
class GoogleDriveSetupScreen:
    """
    IBM 5250 Configuration Screen for connecting Google Drive / Google Docs.
    Supports:
      1. Direct Google Drive REST API (OAuth Client ID + Secret or credentials.json)
      2. Rclone remote (e.g. gdrive:)
      3. Local-only mode
    """

    def __init__(
        self,
        gdrive: GoogleDriveManager,
        system_name: str = "AS400",
        launch_external_cb: Optional[Callable] = None,
    ):
        self.gdrive = gdrive
        self.system_name = system_name
        self.launch_external_cb = launch_external_cb
        self.field_idx = 0
        self.status_message = "Configure Google Drive connection settings and press Enter to save."
        self.json_file_path = "~/credentials.json"
        self.manual_auth_code = ""

        self.fields = [
            {"key": "backend", "label": "Sync Backend (auto/api/rclone/local)", "max": 12},
            {"key": "client_id", "label": "Google OAuth Client ID", "max": 48},
            {"key": "client_secret", "label": "Google OAuth Client Secret", "max": 48},
            {"key": "json_path", "label": "Or Load credentials.json Path", "max": 48},
            {"key": "rclone_remote", "label": "Rclone Remote Name (if using rclone)", "max": 24},
            {"key": "page_width", "label": "Default Word Wrap Margin Width", "max": 6},
        ]

    def _get_val(self, key: str) -> str:
        if key == "json_path":
            return self.json_file_path
        return str(self.gdrive.config.get(key, ""))

    def _set_val(self, key: str, val: str):
        if key == "json_path":
            self.json_file_path = val
        elif key == "page_width":
            if val.isdigit():
                self.gdrive.config[key] = max(40, min(84, int(val)))
        else:
            self.gdrive.config[key] = val

    def run(self, stdscr):
        while True:
            max_y, max_x = stdscr.getmaxyx()
            stdscr.erase()

            safe_addstr(stdscr, 0, 1, "CFGDRV", curses.A_BOLD)
            title = "Configure Google Drive, Docs, Sheets & Calendar"
            safe_addstr(stdscr, 0, (max_x - len(title)) // 2, title, curses.A_BOLD)
            safe_addstr(stdscr, 0, max_x - 16, f"System: {self.system_name}")

            safe_addstr(
                stdscr,
                2,
                2,
                f"Current Status: {self.gdrive.get_status_summary()}",
                curses.A_BOLD,
            )

            start_y = 4
            for i, f in enumerate(self.fields):
                y = start_y + i * 2
                safe_addstr(stdscr, y, 3, f"{f['label']:<38}: ")
                val = self._get_val(f["key"])
                disp = f"{val:<{f['max']}}"[: f["max"]]
                attr = (
                    curses.A_BOLD | curses.A_REVERSE
                    if i == self.field_idx
                    else curses.A_UNDERLINE
                )
                safe_addstr(stdscr, y, 43, disp, attr)

            # Instructions Box for both API OAuth and Rclone
            box_y = start_y + len(self.fields) * 2 + 1
            if box_y + 11 < max_y - 3:
                draw_5250_box(
                    stdscr,
                    box_y,
                    2,
                    11,
                    max_x - 4,
                    "Google Drive Setup Options (Works over Local or Telnet MDA)",
                )
                info = [
                    "Option A - Direct Google Drive API (Zero external dependencies):",
                    "  1. Put your Google OAuth 'credentials.json' path above & press [F6] to load it",
                    "     (or type Client ID & Client Secret directly into the fields above).",
                    "  2. Press [F7] for Automatic Browser Sign-In, OR press [F8] for Headless/Telnet",
                    "     Sign-In (displays a URL you can open on any phone/PC and paste the code).",
                    "",
                    "Option B - Rclone Backend (if rclone is installed):",
                    "  * Set Rclone Remote Name (e.g. 'gdrive:') or press [F9] to launch 'rclone config'.",
                ]
                for idx, line in enumerate(info):
                    safe_addstr(stdscr, box_y + 2 + idx, 4, line[: max_x - 8])

            safe_addstr(
                stdscr,
                max_y - 3,
                1,
                "F6=Load JSON File   F7=Browser OAuth   F8=Telnet/Code OAuth   F9=Rclone Config",
            )
            safe_addstr(
                stdscr,
                max_y - 2,
                1,
                "Enter=Save Settings   Tab/Arrows=Move Field   F5=Test Sync   F12=Return",
            )
            safe_addstr(stdscr, max_y - 1, 1, self.status_message[: max_x - 2], curses.A_BOLD)

            stdscr.refresh()
            ch = read_key(stdscr)
            if ch == -1:
                continue

            if ch in (27, curses.KEY_F12, curses.KEY_F3):
                self.gdrive.save_config()
                break

            if ch in (9, curses.KEY_DOWN):
                self.field_idx = (self.field_idx + 1) % len(self.fields)
                continue

            if ch in (curses.KEY_BTAB, curses.KEY_UP):
                self.field_idx = (self.field_idx - 1) % len(self.fields)
                continue

            # F6 = Load credentials.json
            if ch == curses.KEY_F6:
                ok, msg = self.gdrive.load_client_secrets_file(self.json_file_path)
                self.status_message = msg
                continue

            # F7 = Browser Loopback OAuth
            if ch == curses.KEY_F7:
                self.gdrive.save_config()
                if not self.gdrive.config.get("client_id"):
                    self.status_message = "Enter Client ID & Secret (or press F6 to load JSON) first."
                    continue
                self.status_message = "Waiting for browser authorization at http://localhost:8765 ..."
                stdscr.refresh()
                ok, msg = self.gdrive.run_local_oauth_listener()
                self.status_message = msg
                continue

            # F8 = Headless / Telnet OAuth URL + Code Input
            if ch == curses.KEY_F8:
                self.gdrive.save_config()
                if not self.gdrive.config.get("client_id"):
                    self.status_message = "Enter Client ID & Secret (or press F6 to load JSON) first."
                    continue
                self._modal_headless_oauth(stdscr)
                continue

            # F9 = Run rclone config
            if ch == curses.KEY_F9:
                rclone_bin = shutil.which("rclone")
                if not rclone_bin:
                    self.status_message = "rclone is not installed (sudo apt install rclone)."
                elif self.launch_external_cb:
                    self.launch_external_cb(stdscr, [rclone_bin, "config"], "rclone config")
                    self.status_message = f"Returned from rclone. Status: {self.gdrive.get_status_summary()}"
                continue

            # F5 = Test Sync
            if ch == curses.KEY_F5:
                self.gdrive.save_config()
                self.status_message = "Testing Google Drive sync..."
                stdscr.refresh()
                ok, msg = self.gdrive.sync_all_from_gdrive()
                self.status_message = msg
                continue

            if ch in (10, 13, curses.KEY_ENTER):
                self.gdrive.save_config()
                self.status_message = "Configuration saved."
                continue

            cur_f = self.fields[self.field_idx]
            cur_val = self._get_val(cur_f["key"])
            if ch in (curses.KEY_BACKSPACE, 127, 8):
                self._set_val(cur_f["key"], cur_val[:-1])
            elif 32 <= ch <= 126 and len(cur_val) < cur_f["max"]:
                self._set_val(cur_f["key"], cur_val + chr(ch))

    def _modal_headless_oauth(self, stdscr):
        redirect_uri = "http://localhost:8765"
        auth_url = self.gdrive.build_auth_url(redirect_uri=redirect_uri)
        code_input = ""
        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw, bh = min(84, max_x - 4), 16
            sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
            draw_5250_box(stdscr, sy, sx, bh, bw, "Telnet / Headless Google OAuth Sign-In")

            safe_addstr(stdscr, sy + 2, sx + 2, "1. Open this URL in any browser:", curses.A_BOLD)
            wrapped_url = textwrap.wrap(auth_url, width=bw - 6)
            for i, uline in enumerate(wrapped_url[:4]):
                safe_addstr(stdscr, sy + 3 + i, sx + 3, uline)

            safe_addstr(
                stdscr,
                sy + 8,
                sx + 2,
                "2. After signing in, copy the 'code=...' parameter from the redirected",
            )
            safe_addstr(stdscr, sy + 9, sx + 2, "   localhost URL and paste it below:")

            safe_addstr(stdscr, sy + 11, sx + 3, "Auth Code: ")
            safe_addstr(
                stdscr,
                sy + 11,
                sx + 14,
                f"{code_input:<{bw - 18}}"[: bw - 18],
                curses.A_BOLD | curses.A_REVERSE,
            )

            safe_addstr(stdscr, sy + bh - 2, sx + 3, "Enter=Submit Code   F12/Esc=Cancel")
            stdscr.refresh()

            ch = read_key(stdscr)
            if ch in (27, curses.KEY_F12, curses.KEY_F3):
                return
            if ch in (10, 13, curses.KEY_ENTER) and code_input.strip():
                raw = code_input.strip()
                # If user pasted the full localhost redirect URL, extract code=...
                if "code=" in raw:
                    parsed = urllib.parse.urlparse(raw)
                    qs = urllib.parse.parse_qs(parsed.query)
                    if "code" in qs:
                        raw = qs["code"][0]
                ok, msg = self.gdrive.exchange_code_for_token(raw, redirect_uri=redirect_uri)
                self.status_message = msg
                return
            if ch in (curses.KEY_BACKSPACE, 127, 8):
                code_input = code_input[:-1]
            elif 32 <= ch <= 126:
                code_input += chr(ch)
