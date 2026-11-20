#!/usr/bin/env python3
"""
IBM OfficeVision/400 (OV/400) Replica for 90x43 Monochrome MDA Terminal.
Includes:
  - Authentic AS/400 5250 Sign On screen (QDSIGNON)
  - Authentic OfficeVision/400 Main Menu (OFCMAIN) with side-by-side Calendar
  - Direct integration with Alpine (/usr/bin/alpine) when selecting Option 2 (Mail)
"""

import argparse
import calendar
import curses
import datetime
import getpass
import os
import shutil
import socket
import subprocess
import sys
import time
from typing import List, Optional, Tuple

from calapp import CalendarApp
from gdrive import GoogleDriveManager
from lotus123 import Lotus123App
from wordproc import DocumentManagerScreen, read_key


MENU_OPTIONS: List[Tuple[int, str, str]] = [
    (1, "Calendars", "Google Calendar fullscreen month & day schedule (WRKCAL)"),
    (2, "Mail", "Work with electronic mail (Launches Alpine)"),
    (3, "Send message", "Send an immediate message to a user"),
    (4, "Reminders", "Work with personal reminders"),
    (5, "Directories", "Explore folders, local files, and Google Drive"),
    (6, "Documents and folders", "Manage folders, Google Docs, and Drive files"),
    (7, "Antigravity", "Launch Antigravity AI Assistant (agy-vga)"),
    (8, "Plex", "Launch Plex MDA Music Player (pmda)"),
    (9, "Word processing", "Create, revise, and print documents"),
    (10, "Lotus 1-2-3", "Spreadsheet with Google Sheets sync"),
    (90, "Sign off", "End current OfficeVision/400 session"),
]


class OfficeVisionApp:
    def __init__(self, skip_login: bool = False):
        self.screen_state = "MAIN" if skip_login else "LOGIN"
        raw_host = socket.gethostname().split(".")[0].upper()
        self.system_name = (raw_host[:8] if raw_host else "AS400")
        self.subsystem_name = "QINTER"
        self.display_name = "QPADEV0001"

        default_user = getpass.getuser().upper()[:10]
        # Sign-On screen fields: [User, Password, Program/procedure, Menu, Current library]
        self.login_fields = [
            {"label": "User", "value": default_user, "secret": False, "max_len": 10},
            {"label": "Password", "value": "", "secret": True, "max_len": 10},
            {"label": "Program/procedure", "value": "", "secret": False, "max_len": 10},
            {"label": "Menu", "value": "OFCMAIN", "secret": False, "max_len": 10},
            {"label": "Current library", "value": "", "secret": False, "max_len": 10},
        ]
        self.login_field_idx = 1 if default_user else 0
        self.login_cursor_pos = len(self.login_fields[self.login_field_idx]["value"])
        self.login_message = ""

        # Main Menu (OFCMAIN) state
        self.current_user = default_user or "QUSER"
        self.selected_menu_idx = 1  # Default highlight on 2. Mail for convenience, or 0
        self.command_input = ""
        self.command_history: List[str] = []
        self.history_idx = -1
        self.status_message = ""
        self.show_help = False

        now = datetime.date.today()
        self.cal_year = now.year
        self.cal_month = now.month

        self.gdrive = GoogleDriveManager()
        self.calendar_app = CalendarApp(
            self.gdrive,
            system_name=self.system_name,
            current_user=self.current_user,
            launch_external_cb=self._launch_external_program,
        )
        self.doc_manager = DocumentManagerScreen(
            self.gdrive,
            system_name=self.system_name,
            current_user=self.current_user,
            launch_external_cb=self._launch_external_program,
        )
        self.lotus_app = Lotus123App(
            self.gdrive,
            system_name=self.system_name,
            current_user=self.current_user,
            launch_external_cb=self._launch_external_program,
        )

    def run(self):
        # Ensure 90x43 MDA dimensions are respected when MDA=1 or vt100
        if os.environ.get("MDA") == "1" or os.environ.get("TERM") in ("vt100", "wsvt25", "ansi", "dumb", ""):
            os.environ["LINES"] = os.environ.get("LINES", "43")
            os.environ["COLUMNS"] = os.environ.get("COLUMNS", "90")
            try:
                curses.use_env(True)
            except Exception:
                pass
            sys.stdout.write("\x1b[1;43r")
            sys.stdout.flush()

        try:
            curses.wrapper(self._main_loop)
        finally:
            if os.environ.get("MDA") == "1" or os.environ.get("TERM") == "vt100":
                sys.stdout.write("\x1b[r\x1b[1;43r\x1b[H\x1b[2J\x1b[0m\x1b[?25h")
            else:
                sys.stdout.write("\x1b[r\x1b[H\x1b[2J\x1b[0m\x1b[?25h")
            sys.stdout.flush()

    def _safe_addstr(self, win, y: int, x: int, text: str, attr: int = curses.A_NORMAL):
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

    def _main_loop(self, stdscr):
        stdscr.timeout(250)  # 250ms tick so clock updates live
        try:
            curses.use_default_colors()
        except curses.error:
            pass

        while True:
            max_y, max_x = stdscr.getmaxyx()
            stdscr.erase()

            if self.screen_state == "LOGIN":
                self._draw_login_screen(stdscr, max_y, max_x)
            else:
                self._draw_ofcmain_screen(stdscr, max_y, max_x)
                if self.show_help:
                    self._draw_help_modal(stdscr, max_y, max_x)

            stdscr.refresh()

            ch = read_key(stdscr, restore_timeout=250)
            if ch == -1:
                continue

            if self.show_help:
                if ch in (27, ord("q"), ord("Q"), curses.KEY_F3, curses.KEY_F12, 10, 13, ord(" ")):
                    self.show_help = False
                continue

            if self.screen_state == "LOGIN":
                if not self._handle_login_input(ch):
                    break
            else:
                if not self._handle_ofcmain_input(stdscr, ch):
                    break

    # =========================================================================
    # SCREEN 1: IBM AS/400 Sign On (QDSIGNON)
    # =========================================================================
    def _draw_login_screen(self, win, max_y: int, max_x: int):
        try:
            curses.curs_set(1)
        except curses.error:
            pass

        # Title Row (Row 0)
        title = "Sign On"
        self._safe_addstr(win, 0, (max_x - len(title)) // 2, title, curses.A_BOLD)

        # System Metadata Block (Rows 1-3, right-aligned as on 5250 QDSIGNON)
        sys_col = max(46, max_x - 36)
        self._safe_addstr(win, 1, sys_col, f"System  . . . . . :   {self.system_name:<10}")
        self._safe_addstr(win, 2, sys_col, f"Subsystem . . . . :   {self.subsystem_name:<10}")
        self._safe_addstr(win, 3, sys_col, f"Display . . . . . :   {self.display_name:<10}")

        # Center the sign-on form vertically on larger (43-row) or standard (24-row) screens
        form_start_y = 8 if max_y >= 30 else 5
        row_spacing = 2 if max_y >= 28 else 1
        label_col = max(12, (max_x - 54) // 2)
        input_col = label_col + 36

        cursor_y, cursor_x = form_start_y, input_col

        for idx, field in enumerate(self.login_fields):
            y = form_start_y + idx * row_spacing
            label = field["label"]
            # Build 5250 leader dots: "User  . . . . . . . . . . . . . ."
            dots_str = self._format_5250_leader(label, 34)
            self._safe_addstr(win, y, label_col, dots_str)

            val = field["value"]
            max_len = field["max_len"]
            if field["secret"]:
                # Show asterisks or underscores for password field
                disp_val = "*" * len(val)
            else:
                disp_val = val

            padded = disp_val + "_" * max(0, max_len - len(disp_val))
            attr = curses.A_BOLD | curses.A_UNDERLINE if idx == self.login_field_idx else curses.A_UNDERLINE
            self._safe_addstr(win, y, input_col, padded, attr)

            if idx == self.login_field_idx:
                pos = min(len(val), max_len - 1)
                cursor_y = y
                cursor_x = input_col + pos

        # Optional banner / instructions for MDA convenience
        info_y = form_start_y + len(self.login_fields) * row_spacing + 3
        if max_y >= 34:
            box_w = min(66, max_x - 8)
            box_x = (max_x - box_w) // 2
            self._safe_addstr(win, info_y, box_x, "+" + "-" * (box_w - 2) + "+")
            self._safe_addstr(
                win,
                info_y + 1,
                box_x,
                "|" + " IBM OfficeVision/400 Terminal Environment (5250 Emulation) ".center(box_w - 2) + "|",
                curses.A_BOLD,
            )
            self._safe_addstr(
                win,
                info_y + 2,
                box_x,
                "|" + " Press [Enter] to Sign On to OFCMAIN, or [Tab] to edit fields. ".center(box_w - 2) + "|",
            )
            self._safe_addstr(win, info_y + 3, box_x, "+" + "-" * (box_w - 2) + "+")

        # Bottom Message & Copyright lines
        msg_y = max_y - 3
        fkey_y = max_y - 2
        copy_y = max_y - 1

        if self.login_message:
            self._safe_addstr(win, msg_y, 2, self.login_message, curses.A_BOLD)

        self._safe_addstr(win, fkey_y, 2, "F3=Exit   Tab/Arrows=Next Field   Enter=Sign On")
        copyright_text = "(C) COPYRIGHT IBM CORP. 1980, 1994."
        self._safe_addstr(win, copy_y, max(2, max_x - len(copyright_text) - 2), copyright_text)

        try:
            win.move(cursor_y, min(cursor_x, max_x - 2))
        except curses.error:
            pass

    def _format_5250_leader(self, label: str, width: int) -> str:
        res = list(" " * width)
        for i, ch in enumerate(label):
            if i < width:
                res[i] = ch
        start_dot = len(label) + 2
        if start_dot % 2 == 1:
            start_dot += 1
        for i in range(start_dot, width - 1, 2):
            res[i] = "."
        return "".join(res)

    def _handle_login_input(self, ch: int) -> bool:
        field = self.login_fields[self.login_field_idx]

        # F3 or ESC = Exit
        if ch in (curses.KEY_F3, 27):
            return False

        # Enter = Sign On
        if ch in (10, 13, curses.KEY_ENTER):
            user_val = self.login_fields[0]["value"].strip().upper()
            if not user_val:
                self.login_message = "CPF1120: User profile name is required to sign on."
                self.login_field_idx = 0
                return True
            self.current_user = user_val
            self.login_message = ""
            self.status_message = f"Signed on as {self.current_user} on {self.system_name} ({self.display_name})."
            self.screen_state = "MAIN"
            return True

        # Tab / Down = next field
        if ch in (9, curses.KEY_DOWN):
            self.login_field_idx = (self.login_field_idx + 1) % len(self.login_fields)
            return True

        # Shift+Tab / Up = previous field
        if ch in (curses.KEY_BTAB, curses.KEY_UP):
            self.login_field_idx = (self.login_field_idx - 1) % len(self.login_fields)
            return True

        # Backspace
        if ch in (curses.KEY_BACKSPACE, 127, 8):
            if field["value"]:
                field["value"] = field["value"][:-1]
            return True

        # Printable character input
        if 32 <= ch <= 126:
            char = chr(ch)
            if not field["secret"]:
                char = char.upper()
            if len(field["value"]) < field["max_len"]:
                field["value"] += char
            else:
                # Auto-advance if field filled
                self.login_field_idx = (self.login_field_idx + 1) % len(self.login_fields)
            return True

        return True

    # =========================================================================
    # SCREEN 2: OfficeVision/400 Main Menu (OFCMAIN)
    # =========================================================================
    def _draw_ofcmain_screen(self, win, max_y: int, max_x: int):
        try:
            curses.curs_set(1)
        except curses.error:
            pass

        now = datetime.datetime.now()
        date_str = now.strftime("%m/%d/%y")
        time_str = now.strftime("%H:%M:%S")
        day_name = now.strftime("%A")

        # Row 0: OFCMAIN header
        self._safe_addstr(win, 0, 1, "OFCMAIN", curses.A_BOLD)
        title = "OfficeVision/400"
        self._safe_addstr(win, 0, (max_x - len(title)) // 2, title, curses.A_BOLD)
        sys_hdr = f"System: {self.system_name}"
        self._safe_addstr(win, 0, max_x - len(sys_hdr) - 2, sys_hdr)

        # Row 1: User / Date / Time
        user_info = f"User: {self.current_user}"
        self._safe_addstr(win, 1, 1, user_info)
        dt_hdr = f"{day_name}  {date_str}  {time_str}"
        self._safe_addstr(win, 1, max_x - len(dt_hdr) - 2, dt_hdr)

        # Row 3: Prompt
        self._safe_addstr(win, 3, 1, "Select one of the following:")

        # Left Column: Menu Options (1-9, 90)
        menu_start_y = 5
        row_step = 2 if max_y >= 34 else 1

        for idx, (opt_num, opt_label, opt_desc) in enumerate(MENU_OPTIONS):
            # Add extra blank line before 90. Sign off
            y = menu_start_y + idx * row_step
            if opt_num == 90:
                y += 1

            is_selected = idx == self.selected_menu_idx
            line_text = f" {opt_num:>2}. {opt_label:<26}"
            attr = curses.A_REVERSE | curses.A_BOLD if is_selected else curses.A_NORMAL
            self._safe_addstr(win, y, 5, line_text, attr)

        # Right Column: Monthly Calendar + Daily Status Box
        cal_col = max(44, max_x - 36)
        self._draw_calendar_panel(win, 3, cal_col, max_y, max_x, now.date())

        # Option Description / Context Bar on tall (43-row) MDA screens
        cmd_label_y = max_y - 5
        cmd_line_y = max_y - 4
        fkey_line_1_y = max_y - 2
        msg_line_y = max_y - 1

        if max_y >= 30:
            desc_box_y = cmd_label_y - 3
            _, sel_title, sel_desc = MENU_OPTIONS[self.selected_menu_idx]
            sel_num = MENU_OPTIONS[self.selected_menu_idx][0]
            hint = f" Option {sel_num} ({sel_title}): {sel_desc} "
            self._safe_addstr(win, desc_box_y, 1, "-" * (max_x - 2))
            self._safe_addstr(win, desc_box_y + 1, 1, hint[: max_x - 2], curses.A_BOLD)

        # Command Line ("Selection or command ===>")
        self._safe_addstr(win, cmd_label_y, 1, "Selection or command")
        prompt_prefix = " ===> "
        self._safe_addstr(win, cmd_line_y, 1, prompt_prefix, curses.A_BOLD)

        field_width = max(10, max_x - len(prompt_prefix) - 3)
        cmd_display = self.command_input + "_" * max(0, field_width - len(self.command_input))
        self._safe_addstr(win, cmd_line_y, 1 + len(prompt_prefix), cmd_display[:field_width], curses.A_BOLD)

        # Function Keys Bar
        fkeys = "F1=Help   F3=Exit   F9=Retrieve   F12=Cancel   F19=Prev Mo   F20=Next Mo"
        self._safe_addstr(win, fkey_line_1_y, 1, fkeys[: max_x - 2])

        # Message Line / IBM Copyright
        if self.status_message:
            self._safe_addstr(win, msg_line_y, 1, self.status_message[: max_x - 2], curses.A_BOLD)
        else:
            copyright_text = "(C) COPYRIGHT IBM CORP. 1980, 1994."
            self._safe_addstr(
                win,
                msg_line_y,
                max(1, max_x - len(copyright_text) - 2),
                copyright_text,
            )

        # Place hardware cursor on the ===> command line
        cursor_x = 1 + len(prompt_prefix) + min(len(self.command_input), field_width - 1)
        try:
            win.move(cmd_line_y, min(cursor_x, max_x - 2))
        except curses.error:
            pass

    def _draw_calendar_panel(self, win, start_y: int, start_x: int, max_y: int, max_x: int, today: datetime.date):
        cal_width = 30
        if start_x + cal_width >= max_x:
            start_x = max(36, max_x - cal_width - 1)

        # Top border of calendar panel
        self._safe_addstr(win, start_y, start_x, "+" + "-" * (cal_width - 2) + "+")

        month_title = f"{calendar.month_name[self.cal_month]} {self.cal_year}"
        self._safe_addstr(win, start_y + 1, start_x, "|" + month_title.center(cal_width - 2) + "|", curses.A_BOLD)

        dow_header = "  S   M   T   W   T   F   S "
        self._safe_addstr(win, start_y + 2, start_x, "|" + dow_header.center(cal_width - 2) + "|", curses.A_UNDERLINE)

        cal = calendar.Calendar(firstweekday=6)  # Sunday first, standard IBM OV/400 layout
        month_weeks = cal.monthdayscalendar(self.cal_year, self.cal_month)

        row_y = start_y + 3
        for week in month_weeks:
            self._safe_addstr(win, row_y, start_x, "|")
            self._safe_addstr(win, row_y, start_x + cal_width - 1, "|")
            col_x = start_x + 1
            for day in week:
                if day == 0:
                    cell = "    "
                    self._safe_addstr(win, row_y, col_x, cell)
                else:
                    is_today = (
                        day == today.day
                        and self.cal_month == today.month
                        and self.cal_year == today.year
                    )
                    if is_today:
                        cell = f"[{day:>2}]"
                        self._safe_addstr(win, row_y, col_x, cell, curses.A_REVERSE | curses.A_BOLD)
                    else:
                        cell = f" {day:>2} "
                        self._safe_addstr(win, row_y, col_x, cell)
                col_x += 4
            row_y += 1

        # Pad to 6 week rows so calendar box height stays uniform
        while row_y < start_y + 9:
            self._safe_addstr(win, row_y, start_x, "|" + " " * (cal_width - 2) + "|")
            row_y += 1

        self._safe_addstr(win, row_y, start_x, "+" + "-" * (cal_width - 2) + "+")
        row_y += 1

        # Calendar navigation hint
        nav_hint = "[ / ]: Prev/Next Month  t: Today"
        self._safe_addstr(win, row_y, start_x, nav_hint[:cal_width])
        row_y += 2

        # On 43-line MDA screen, show Daily Schedule / OfficeVision Status summary beneath Calendar
        if row_y + 8 < max_y - 7:
            self._safe_addstr(win, row_y, start_x, "+" + "-" * (cal_width - 2) + "+")
            self._safe_addstr(
                win,
                row_y + 1,
                start_x,
                "|" + " OfficeVision/400 Status ".center(cal_width - 2) + "|",
                curses.A_BOLD,
            )
            self._safe_addstr(win, row_y + 2, start_x, "+" + "-" * (cal_width - 2) + "+")

            local_bin = os.path.expanduser("~/.local/bin")
            alpine_avail = "Ready" if shutil.which("alpine") else "Not Found"
            agy_avail = (
                "Ready"
                if (
                    shutil.which("agy-vga")
                    or os.path.exists(os.path.join(local_bin, "agy-vga"))
                    or shutil.which("agy")
                    or os.path.exists(os.path.join(local_bin, "agy"))
                )
                else "Not Found"
            )
            plex_avail = (
                "Ready"
                if (shutil.which("pmda") or os.path.exists(os.path.join(local_bin, "pmda")))
                else "Not Found"
            )
            gdrive_mode = self.gdrive.get_active_backend().upper()
            cal_mode = self.calendar_app.gcal.get_short_backend_label()

            status_lines = [
                f" Cal (1)     : WRKCAL ({cal_mode})",
                f" Mail (2)    : Alpine ({alpine_avail})",
                f" Docs (6/9)  : OV/400 WP ({gdrive_mode})",
                f" 1-2-3 (10)  : Lotus ({gdrive_mode})",
                f" Antigrav (7): agy-vga ({agy_avail})",
                f" Plex (8)    : pmda ({plex_avail})",
                f" Terminal    : {os.environ.get('COLUMNS', '90')}x{os.environ.get('LINES', '43')} MDA",
            ]
            for i, sline in enumerate(status_lines):
                self._safe_addstr(
                    win,
                    row_y + 3 + i,
                    start_x,
                    "|" + f"{sline:<{cal_width - 2}}"[: cal_width - 2] + "|",
                )
            self._safe_addstr(win, row_y + 3 + len(status_lines), start_x, "+" + "-" * (cal_width - 2) + "+")

    def _draw_help_modal(self, win, max_y: int, max_x: int):
        box_w = min(68, max_x - 6)
        box_h = 16
        start_y = max(2, (max_y - box_h) // 2)
        start_x = max(2, (max_x - box_w) // 2)

        help_lines = [
            "OfficeVision/400 - Help",
            "",
            "Navigation & Selection:",
            "  * Type an option number (e.g. 2 for Mail, 90 for Sign off)",
            "    on the ===> command line and press [Enter].",
            "  * Or use [Up] / [Down] arrow keys to highlight an option",
            "    and press [Enter] when the command line is empty.",
            "",
            "Calendar Controls:",
            "  * [F19] or '[' : Previous month",
            "  * [F20] or ']' : Next month",
            "  * 't'          : Jump back to current month",
            "",
            "Press [Enter], [Space], [Esc], or [F12] to close Help.",
        ]

        for r in range(box_h):
            y = start_y + r
            if r == 0 or r == box_h - 1:
                self._safe_addstr(win, y, start_x, "+" + "=" * (box_w - 2) + "+", curses.A_BOLD)
            else:
                txt = help_lines[r - 1] if (r - 1) < len(help_lines) else ""
                attr = curses.A_BOLD | curses.A_REVERSE if r == 1 else curses.A_NORMAL
                padded = f" {txt:<{box_w - 4}} "
                self._safe_addstr(win, y, start_x, "|" + padded[: box_w - 2] + "|", attr)

    def _handle_ofcmain_input(self, stdscr, ch: int) -> bool:
        # F1 = Help
        if ch == curses.KEY_F1 or (ch == ord("?") and not self.command_input):
            self.show_help = True
            return True

        # F3 = Exit application
        if ch == curses.KEY_F3:
            return False

        # F12 or Esc = Clear command input or return to Sign On
        if ch in (curses.KEY_F12, 27):
            if self.command_input:
                self.command_input = ""
                self.status_message = "Command line cleared."
            else:
                self.status_message = "Press F3 or select option 90 to Sign off."
            return True

        # F9 = Retrieve previous command
        if ch == curses.KEY_F9:
            if self.command_history:
                if self.history_idx < 0:
                    self.history_idx = len(self.command_history) - 1
                else:
                    self.history_idx = max(0, self.history_idx - 1)
                self.command_input = self.command_history[self.history_idx]
            return True

        # Calendar Month Navigation: F19 / F7 / [ = Prev Month, F20 / F8 / ] = Next Month
        if ch in (curses.KEY_F7, curses.KEY_F19) or (ch == ord("[") and not self.command_input):
            self._shift_calendar_month(-1)
            return True
        if ch in (curses.KEY_F8, curses.KEY_F20) or (ch == ord("]") and not self.command_input):
            self._shift_calendar_month(1)
            return True
        if ch == ord("t") and not self.command_input:
            today = datetime.date.today()
            self.cal_year, self.cal_month = today.year, today.month
            self.status_message = "Calendar reset to current month."
            return True

        # Up / Down arrow keys move menu selection bar and clear stale command line text
        if ch == curses.KEY_UP:
            self.selected_menu_idx = (self.selected_menu_idx - 1) % len(MENU_OPTIONS)
            self.command_input = ""
            return True
        if ch == curses.KEY_DOWN:
            self.selected_menu_idx = (self.selected_menu_idx + 1) % len(MENU_OPTIONS)
            self.command_input = ""
            return True

        # Backspace on command line
        if ch in (curses.KEY_BACKSPACE, 127, 8):
            if self.command_input:
                self.command_input = self.command_input[:-1]
            return True

        # Enter = Execute command line option or currently highlighted menu item
        if ch in (10, 13, curses.KEY_ENTER):
            cmd = self.command_input.strip()
            self.command_input = ""
            if cmd:
                self.command_history.append(cmd)
                self.history_idx = -1
                return self._execute_menu_command(stdscr, cmd)
            else:
                opt_num = MENU_OPTIONS[self.selected_menu_idx][0]
                return self._execute_menu_command(stdscr, str(opt_num))

        # Printable character input onto ===> command line
        if 32 <= ch <= 126:
            self.command_input += chr(ch)
            # If user typed a matching option number, also move highlight bar to match
            stripped = self.command_input.strip()
            if stripped.isdigit():
                num = int(stripped)
                for i, (opt_num, _, _) in enumerate(MENU_OPTIONS):
                    if opt_num == num:
                        self.selected_menu_idx = i
                        break
            return True

        return True

    def _shift_calendar_month(self, delta: int):
        m = self.cal_month + delta
        y = self.cal_year
        if m < 1:
            m = 12
            y -= 1
        elif m > 12:
            m = 1
            y += 1
        self.cal_month, self.cal_year = m, y

    def _execute_menu_command(self, stdscr, cmd: str) -> bool:
        upper_cmd = cmd.upper()

        # Direct AS/400 commands
        if upper_cmd in ("SIGNOFF", "EXIT", "QUIT", "LOGOFF", "90"):
            # Return to AS/400 Sign On screen
            self.login_fields[1]["value"] = ""  # Clear password
            self.login_field_idx = 0
            self.login_message = f"User {self.current_user} signed off from OfficeVision/400."
            self.screen_state = "LOGIN"
            return True

        if upper_cmd in ("HELP", "?"):
            self.show_help = True
            return True

        if upper_cmd in ("CAL", "CALENDAR", "CALENDARS", "WRKCAL", "GCAL", "1"):
            self._launch_calendar(stdscr)
            return True

        if upper_cmd in ("ALPINE", "MAIL", "2"):
            self._launch_alpine(stdscr)
            return True

        if upper_cmd in ("DOCS", "DOCUMENTS", "DIRECTORIES", "DIR", "FILES", "WRKDOC", "WRKFLR", "GDRIVE", "5", "6"):
            self._launch_documents(stdscr)
            return True

        if upper_cmd in ("ANTIGRAVITY", "AGY", "AGY-VGA", "7"):
            self._launch_antigravity(stdscr)
            return True

        if upper_cmd in ("PLEX", "PMDA", "MUSIC", "8"):
            self._launch_plex(stdscr)
            return True

        if upper_cmd in ("WP", "WORD", "WORDPROC", "EDTDOC", "9"):
            self._launch_wordproc(stdscr)
            return True

        if upper_cmd in ("10", "123", "LOTUS", "LOTUS123", "SHEET", "SHEETS", "GSHEETS", "SPREADSHEET", "CALC"):
            self._launch_lotus123(stdscr)
            return True

        if cmd.isdigit():
            opt_num = int(cmd)
            for idx, (num, label, _) in enumerate(MENU_OPTIONS):
                if num == opt_num:
                    self.selected_menu_idx = idx
                    if opt_num == 1:
                        self._launch_calendar(stdscr)
                    elif opt_num == 2:
                        self._launch_alpine(stdscr)
                    elif opt_num in (5, 6):
                        self._launch_documents(stdscr)
                    elif opt_num == 7:
                        self._launch_antigravity(stdscr)
                    elif opt_num == 8:
                        self._launch_plex(stdscr)
                    elif opt_num == 9:
                        self._launch_wordproc(stdscr)
                    elif opt_num == 10:
                        self._launch_lotus123(stdscr)
                    elif opt_num == 90:
                        return self._execute_menu_command(stdscr, "90")
                    else:
                        self.status_message = (
                            f"CPF9898: Option {opt_num} ({label}) is not yet active in this build."
                        )
                    return True

        self.status_message = f"CPD0030: Selection or command '{cmd}' is not valid."
        return True

    def _launch_external_program(self, stdscr, cmd_args: List[str], label: str):
        # Cleanly suspend curses and prepare the 90x43 MDA terminal
        try:
            curses.def_prog_mode()
        except curses.error:
            pass
        try:
            curses.endwin()
        except curses.error:
            pass

        env = os.environ.copy()
        local_bin = os.path.expanduser("~/.local/bin")
        if local_bin not in env.get("PATH", ""):
            env["PATH"] = f"{local_bin}:{env.get('PATH', '')}"

        if env.get("MDA") == "1" or env.get("TERM") in ("vt100", "wsvt25", "ansi", "dumb", ""):
            env["TERM"] = "vt100"
            env["MDA"] = "1"
            env["LINES"] = "43"
            env["COLUMNS"] = "90"
            sys.stdout.write("\x1b[r\x1b[1;43r\x1b[H\x1b[2J\x1b[0m\x1b[?25h")
        else:
            sys.stdout.write("\x1b[H\x1b[2J\x1b[0m\x1b[?25h")
        sys.stdout.flush()

        try:
            subprocess.run(cmd_args, env=env)
        except Exception as exc:
            self.status_message = f"Error launching {label}: {exc}"
        else:
            self.status_message = f"Returned from {label}."

        # Restore MDA scroll region and resume curses screen
        if env.get("MDA") == "1" or env.get("TERM") == "vt100":
            sys.stdout.write("\x1b[r\x1b[1;43r\x1b[H\x1b[2J\x1b[0m")
            sys.stdout.flush()

        try:
            curses.reset_prog_mode()
        except curses.error:
            pass
        stdscr.clear()
        stdscr.refresh()

    def _launch_alpine(self, stdscr):
        alpine_bin = shutil.which("alpine") or "/usr/bin/alpine"
        if not os.path.exists(alpine_bin):
            self.status_message = "CPF9801: Mail program (/usr/bin/alpine) not found on system."
            return
        self._launch_external_program(stdscr, [alpine_bin], "OfficeVision/400 Mail (Alpine)")

    def _launch_antigravity(self, stdscr):
        local_bin = os.path.expanduser("~/.local/bin")
        agy_vga_bin = shutil.which("agy-vga") or os.path.join(local_bin, "agy-vga")
        agy_bin = shutil.which("agy") or os.path.join(local_bin, "agy")
        if os.path.exists(agy_vga_bin):
            target = agy_vga_bin
        elif os.path.exists(agy_bin):
            target = agy_bin
        else:
            self.status_message = "CPF9801: Antigravity (agy-vga / agy) not found in PATH or ~/.local/bin."
            return
        self._launch_external_program(
            stdscr,
            [target, "--settings", '{"altScreenMode":"always"}'],
            "Antigravity (agy-vga)",
        )

    def _launch_plex(self, stdscr):
        local_bin = os.path.expanduser("~/.local/bin")
        pmda_bin = shutil.which("pmda") or os.path.join(local_bin, "pmda")
        if not os.path.exists(pmda_bin):
            self.status_message = "CPF9801: Plex MDA app (pmda) not found in PATH or ~/.local/bin."
            return
        self._launch_external_program(stdscr, [pmda_bin], "Plex MDA Music Player (pmda)")

    def _launch_documents(self, stdscr):
        self.doc_manager.current_user = self.current_user
        self.doc_manager.run(stdscr, direct_editor=False)
        self.status_message = (
            self.doc_manager.status_message
            or "Returned from Work with Documents in Folders (WRKDOC)."
        )

    def _launch_wordproc(self, stdscr):
        self.doc_manager.current_user = self.current_user
        self.doc_manager.run(stdscr, direct_editor=True)
        self.status_message = (
            self.doc_manager.status_message
            or "Returned from OfficeVision/400 Word Processing."
        )

    def _launch_calendar(self, stdscr):
        self.calendar_app.current_user = self.current_user
        self.calendar_app.run(stdscr)
        self.status_message = (
            self.calendar_app.status_message
            or "Returned from OfficeVision/400 Calendar (WRKCAL)."
        )

    def _launch_lotus123(self, stdscr):
        self.lotus_app.current_user = self.current_user
        self.lotus_app.run(stdscr)
        self.status_message = (
            self.lotus_app.status_message
            or "Returned from Lotus 1-2-3 Spreadsheet."
        )


def main():
    parser = argparse.ArgumentParser(
        description="IBM OfficeVision/400 (OV/400) Terminal Environment for 90x43 MDA",
    )
    parser.add_argument(
        "--skip-login",
        action="store_true",
        help="Skip the AS/400 Sign On screen and go directly to OFCMAIN",
    )
    args = parser.parse_args()

    app = OfficeVisionApp(skip_login=args.skip_login)
    app.run()


if __name__ == "__main__":
    main()
