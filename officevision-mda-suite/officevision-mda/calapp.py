#!/usr/bin/env python3
"""
IBM OfficeVision/400 (OV/400) Calendar Application with Google Calendar Sync
Tailored for the 90x43 Monochrome MDA Terminal (TERM=vt100).

Screens:
  - WRKCAL : Fullscreen Month View (Primary screen) with inline day events
  - DSPDAY : Fullscreen Day Schedule & Event Details (on selecting a day)
  - ADDEVT : Create / Edit Event 5250 Form (timed, all-day, multi-day, recurring)
  - SELCAL : Google Calendar List Selector (toggle visible / default calendars)
  - DSPAGD : Upcoming Agenda & Event Search
  - CFGDRV : Shared OV/400 Google OAuth Setup (F10)
"""

import calendar as pycal
import curses
import datetime
import os
import sys
import textwrap
from typing import Any, Callable, Dict, List, Optional, Tuple

from gcalendar import (
    REPEAT_OPTIONS,
    CalendarEvent,
    GoogleCalendarManager,
    add_minutes_to_hhmm,
    format_hhmm,
    format_time_12h,
    get_local_timezone_name,
    parse_date_str,
    parse_time_str,
)
from gdrive import GoogleDriveManager
from wordproc import GoogleDriveSetupScreen, draw_5250_box, read_key, safe_addstr


class CalendarApp:
    """IBM OfficeVision/400 Fullscreen Calendar TUI tied to Google Calendar."""

    def __init__(
        self,
        gdrive: GoogleDriveManager,
        system_name: str = "AS400",
        current_user: str = "QUSER",
        launch_external_cb: Optional[Callable] = None,
    ):
        self.gdrive = gdrive
        self.gcal = GoogleCalendarManager(gdrive)
        self.system_name = system_name
        self.current_user = current_user
        self.launch_external_cb = launch_external_cb

        today = datetime.date.today()
        self.selected_date: datetime.date = today
        self.view_year: int = today.year
        self.view_month: int = today.month

        self.status_message: str = ""
        self.show_help: bool = False
        self._initial_sync_done: bool = False

    # =========================================================================
    # Public Entrypoint (called from OV/400 Option 1 or standalone)
    # =========================================================================
    def run(self, stdscr):
        # Reset to today's month on fresh entry if desired? Keep user's last date
        # during the same OV/400 session, or default to today on first open.
        if not self._initial_sync_done:
            self._initial_sync_done = True
            if self.gcal.is_connected():
                self._sync_month_with_feedback(
                    stdscr,
                    self.view_year,
                    self.view_month,
                    refresh_calendars=True,
                )
            else:
                self.status_message = (
                    "Local Mode: Press [F10] to connect Google Calendar, "
                    "[Enter] to view day, [N] for new event."
                )

        while True:
            max_y, max_x = stdscr.getmaxyx()
            stdscr.erase()

            self._draw_month_screen(stdscr, max_y, max_x)
            if self.show_help:
                self._draw_help_modal(stdscr, max_y, max_x)

            stdscr.refresh()

            ch = read_key(stdscr, restore_timeout=500)
            if ch == -1:
                continue

            if self.show_help:
                if ch in (27, ord("q"), ord("Q"), curses.KEY_F1, curses.KEY_F3, curses.KEY_F12, 10, 13, ord(" ")):
                    self.show_help = False
                continue

            if not self._handle_month_input(stdscr, ch):
                break

    def _sync_month_with_feedback(
        self,
        stdscr,
        year: int,
        month: int,
        refresh_calendars: bool = False,
        push_pending: bool = True,
    ):
        max_y, max_x = stdscr.getmaxyx()
        self.status_message = f"Syncing {pycal.month_name[month]} {year} with Google Calendar..."
        stdscr.erase()
        self._draw_month_screen(stdscr, max_y, max_x)
        stdscr.refresh()
        ok, msg = self.gcal.sync_month(
            year,
            month,
            refresh_calendars=refresh_calendars,
            push_pending=push_pending,
        )
        self.status_message = msg
        return ok

    def _maybe_autosync_month(self, stdscr, year: int, month: int):
        if self.gcal.is_connected() and (year, month) not in self.gcal.synced_months:
            self._sync_month_with_feedback(
                stdscr, year, month, refresh_calendars=False, push_pending=False
            )

    def _set_selected_date(self, stdscr, new_date: datetime.date):
        old_ym = (self.view_year, self.view_month)
        self.selected_date = new_date
        self.view_year = new_date.year
        self.view_month = new_date.month
        if (self.view_year, self.view_month) != old_ym:
            self._maybe_autosync_month(stdscr, self.view_year, self.view_month)

    def _shift_month(self, stdscr, delta_months: int):
        total = (self.view_year * 12 + (self.view_month - 1)) + delta_months
        new_y, new_m0 = divmod(total, 12)
        new_m = new_m0 + 1
        max_day = pycal.monthrange(new_y, new_m)[1]
        new_d = min(self.selected_date.day, max_day)
        self._set_selected_date(stdscr, datetime.date(new_y, new_m, new_d))

    # =========================================================================
    # SCREEN 1: Fullscreen Month View (WRKCAL)
    # =========================================================================
    def _draw_month_screen(self, win, max_y: int, max_x: int):
        try:
            curses.curs_set(0)
        except curses.error:
            pass

        now = datetime.datetime.now()
        today = now.date()

        # Row 0: OV/400 Header Bar
        safe_addstr(win, 0, 1, "WRKCAL", curses.A_BOLD)
        month_hdr = f"{pycal.month_name[self.view_month].upper()} {self.view_year}"
        title = f"OfficeVision/400 Calendar  -  {month_hdr}"
        tx = max(9, (max_x - len(title)) // 2)
        safe_addstr(win, 0, tx, title, curses.A_BOLD)
        sys_str = f"System: {self.system_name}"
        if max_x - len(sys_str) - 1 > tx + len(title):
            safe_addstr(win, 0, max_x - len(sys_str) - 1, sys_str)

        # Row 1: User, Google Sync Status, Live Clock
        cal_status = self.gcal.get_status_label()
        enabled_cnt = sum(1 for c in self.gcal.calendars if c.enabled)
        cal_tag = f" ({enabled_cnt} cal)" if len(self.gcal.calendars) > 1 else ""
        left_info = f"User: {self.current_user}   Sync: {cal_status}{cal_tag}"
        dt_str = now.strftime("%a %m/%d/%y %H:%M")
        safe_addstr(win, 1, 1, left_info[: max(10, max_x - len(dt_str) - 3)])
        safe_addstr(win, 1, max(1, max_x - len(dt_str) - 1), dt_str)

        # Compute Fullscreen Month Grid geometry
        # Reserve bottom 4 rows for: selected day preview, F-keys row 1, F-keys row 2, status message
        footer_rows = 4 if max_y >= 16 else 3
        grid_top = 2
        grid_bottom = max_y - footer_rows  # exclusive
        avail_h = max(7, grid_bottom - grid_top)

        # 7 columns with shared 1-char vertical borders: total_w = 1 + 7 * (cell_w + 1)
        cell_w = max(8, (max_x - 3) // 7 - 1)
        grid_w = 1 + 7 * (cell_w + 1)
        grid_x = max(0, (max_x - grid_w) // 2)

        # Build month weeks (Sunday first = 6, classic IBM OV/400 layout)
        first_wd = int(self.gcal.config.get("first_weekday", 6))
        cal = pycal.Calendar(firstweekday=first_wd)
        weeks = cal.monthdatescalendar(self.view_year, self.view_month)
        num_weeks = len(weeks)

        # Query events across visible grid date range in one fast pass
        grid_start_date = weeks[0][0]
        grid_end_date = weeks[-1][-1]
        events_by_day = self.gcal.get_events_map_for_range(grid_start_date, grid_end_date)

        # Row 2: Top Border combined with Weekday Names (+----SUN----+----MON----+...)
        dow_names = (
            ["SUN", "MON", "TUE", "WED", "THU", "FRI", "SAT"]
            if first_wd == 6
            else ["MON", "TUE", "WED", "THU", "FRI", "SAT", "SUN"]
        )
        top_border = "+"
        for dname in dow_names:
            pad_total = cell_w - len(dname)
            left_p = pad_total // 2
            right_p = pad_total - left_p
            top_border += ("-" * left_p) + dname + ("-" * right_p) + "+"
        safe_addstr(win, grid_top, grid_x, top_border, curses.A_BOLD)

        # Vertical space per week: (avail_h - 1) divided by num_weeks
        # Each week consumes (cell_h + 1) rows: cell_h inner rows + 1 horizontal divider
        week_h = max(2, (avail_h - 1) // num_weeks)
        cell_h = max(1, week_h - 1)
        horiz_div = "+" + (("-" * cell_w + "+") * 7)

        cur_y = grid_top + 1
        for w_idx, week in enumerate(weeks):
            if cur_y >= grid_bottom:
                break

            # Draw cell content rows
            for r_line in range(cell_h):
                y = cur_y + r_line
                if y >= grid_bottom:
                    break
                # Draw vertical borders
                for col in range(8):
                    bx = grid_x + col * (cell_w + 1)
                    safe_addstr(win, y, bx, "|")

                for col, day_date in enumerate(week):
                    cx = grid_x + col * (cell_w + 1) + 1
                    in_month = day_date.month == self.view_month
                    is_today = day_date == today
                    is_selected = day_date == self.selected_date
                    day_events = events_by_day.get(day_date, [])

                    if r_line == 0:
                        self._draw_cell_header(
                            win,
                            y,
                            cx,
                            cell_w,
                            day_date,
                            len(day_events),
                            in_month,
                            is_today,
                            is_selected,
                        )
                    else:
                        self._draw_cell_event_line(
                            win,
                            y,
                            cx,
                            cell_w,
                            r_line - 1,
                            cell_h - 1,
                            day_events,
                            in_month,
                            is_selected,
                        )

            cur_y += cell_h
            # Horizontal divider below week
            if cur_y < grid_bottom:
                div_attr = curses.A_BOLD if w_idx == num_weeks - 1 else curses.A_NORMAL
                safe_addstr(win, cur_y, grid_x, horiz_div, div_attr)
                cur_y += 1

        # Bottom Area: Selected Day Quick Preview Bar + Function Keys + Status Line
        sel_events = events_by_day.get(self.selected_date, [])
        if footer_rows >= 4:
            preview_y = max_y - 4
            sel_str = self.selected_date.strftime("%a %b %d, %Y")
            if sel_events:
                cnt_lbl = f"({len(sel_events)} event{'s' if len(sel_events) != 1 else ''})"
                snippets = []
                for ev in sel_events[:4]:
                    if ev.all_day:
                        snippets.append(f"*{ev.summary}")
                    else:
                        snippets.append(f"{ev.start_time} {ev.summary}")
                summary_txt = "; ".join(snippets)
                if len(sel_events) > 4:
                    summary_txt += f"; +{len(sel_events) - 4} more"
                bar_txt = f" {sel_str} {cnt_lbl}: {summary_txt}"
            else:
                bar_txt = f" {sel_str} (No events) - Press [Enter] for Day View or [N] to create event"
            padded_bar = f"{bar_txt:<{max_x - 2}}"[: max_x - 2]
            safe_addstr(win, preview_y, 1, padded_bar, curses.A_REVERSE)

        fk1_y = max_y - 3
        fk2_y = max_y - 2
        msg_y = max_y - 1

        if max_y >= 16:
            safe_addstr(
                win,
                fk1_y,
                1,
                "Enter=Day View  N/F6=New Event  Arrows=Move  [/]=Prev/Next Mo  T=Today"[: max_x - 2],
            )
            safe_addstr(
                win,
                fk2_y,
                1,
                "F1=Help  F2=Agenda  F4=Go Date  F5=Sync  F7=Calendars  F10=Google Setup  F3=Exit"[: max_x - 2],
            )
        else:
            safe_addstr(
                win,
                fk2_y,
                1,
                "Enter=Day  N=New  [/]=Mo  T=Today  F5=Sync  F7=Cals  F10=Setup  F3=Exit"[: max_x - 2],
            )

        if self.status_message:
            safe_addstr(win, msg_y, 1, self.status_message[: max_x - 2], curses.A_BOLD)
        else:
            last_sync = f"Last sync: {self.gcal.last_sync_at}" if self.gcal.last_sync_at else "Ready."
            safe_addstr(win, msg_y, 1, last_sync[: max_x - 2])

    def _draw_cell_header(
        self,
        win,
        y: int,
        x: int,
        cell_w: int,
        day_date: datetime.date,
        ev_count: int,
        in_month: bool,
        is_today: bool,
        is_selected: bool,
    ):
        dnum = f"{day_date.day:>2}"
        if is_selected:
            if is_today:
                right = f"*({ev_count})" if ev_count > 0 else "TODAY"
            elif ev_count > 0:
                right = f"({ev_count})"
            elif not in_month:
                right = pycal.month_abbr[day_date.month].upper()
            else:
                right = ""
            mid_space = max(0, cell_w - 2 - len(dnum) - len(right))
            text = f">{dnum}{' ' * mid_space}{right}"
            if len(text) < cell_w:
                text = text + " " * (cell_w - len(text) - 1) + "<"
            else:
                text = text[: cell_w - 1] + "<"
            safe_addstr(win, y, x, text[:cell_w], curses.A_REVERSE | curses.A_BOLD)
            return

        if is_today:
            right = f"({ev_count})" if ev_count > 0 else "TODAY"
            mid_space = max(0, cell_w - 1 - len(dnum) - len(right))
            text = f"*{dnum}{' ' * mid_space}{right}"
            text = f"{text:<{cell_w}}"[:cell_w]
            safe_addstr(win, y, x, text, curses.A_BOLD | curses.A_UNDERLINE)
            return

        if not in_month:
            badge = f"({ev_count})" if ev_count > 0 else (
                pycal.month_abbr[day_date.month] if day_date.day == 1 else ""
            )
            mid_space = max(0, cell_w - 1 - len(dnum) - len(badge))
            text = f" {dnum}{' ' * mid_space}{badge}"
            text = f"{text:<{cell_w}}"[:cell_w]
            safe_addstr(win, y, x, text, curses.A_DIM)
            return

        badge = f"({ev_count})" if ev_count > 0 else ""
        mid_space = max(0, cell_w - 1 - len(dnum) - len(badge))
        text = f" {dnum}{' ' * mid_space}{badge}"
        text = f"{text:<{cell_w}}"[:cell_w]
        attr = curses.A_BOLD if ev_count > 0 else curses.A_NORMAL
        safe_addstr(win, y, x, text, attr)

    def _draw_cell_event_line(
        self,
        win,
        y: int,
        x: int,
        cell_w: int,
        ev_slot: int,
        max_slots: int,
        day_events: List[CalendarEvent],
        in_month: bool,
        is_selected: bool,
    ):
        if not day_events or ev_slot >= len(day_events):
            return

        # If more events than lines, show "+N more" on the final slot
        if len(day_events) > max_slots and ev_slot == max_slots - 1:
            rem = len(day_events) - (max_slots - 1)
            txt = f"+{rem} more..."[:cell_w]
            attr = curses.A_DIM if not in_month else curses.A_BOLD
            safe_addstr(win, y, x, txt, attr)
            return

        ev = day_events[ev_slot]
        label = ev.compact_cell_label(cell_w)
        if not in_month and not is_selected:
            attr = curses.A_DIM
        elif is_selected:
            attr = curses.A_BOLD
        elif ev.all_day:
            attr = curses.A_BOLD
        else:
            attr = curses.A_NORMAL
        safe_addstr(win, y, x, label, attr)

    def _handle_month_input(self, stdscr, ch: int) -> bool:
        # Exit back to OV/400 Main Menu
        if ch in (curses.KEY_F3, curses.KEY_F12, 27, ord("q"), ord("Q")):
            return False

        # Help
        if ch in (curses.KEY_F1, ord("?")):
            self.show_help = True
            return True

        # Day navigation with arrows (crosses month boundaries seamlessly)
        if ch == curses.KEY_LEFT:
            self._set_selected_date(stdscr, self.selected_date - datetime.timedelta(days=1))
            return True
        if ch == curses.KEY_RIGHT:
            self._set_selected_date(stdscr, self.selected_date + datetime.timedelta(days=1))
            return True
        if ch == curses.KEY_UP:
            self._set_selected_date(stdscr, self.selected_date - datetime.timedelta(days=7))
            return True
        if ch == curses.KEY_DOWN:
            self._set_selected_date(stdscr, self.selected_date + datetime.timedelta(days=7))
            return True

        # Home / End = First / Last day of current month
        if ch == curses.KEY_HOME:
            self._set_selected_date(stdscr, datetime.date(self.view_year, self.view_month, 1))
            return True
        if ch == curses.KEY_END:
            last_d = pycal.monthrange(self.view_year, self.view_month)[1]
            self._set_selected_date(stdscr, datetime.date(self.view_year, self.view_month, last_d))
            return True

        # Month & Year navigation
        if ch in (ord("["), curses.KEY_PPAGE, curses.KEY_F19):
            self._shift_month(stdscr, -1)
            return True
        if ch in (ord("]"), curses.KEY_NPAGE, curses.KEY_F20):
            self._shift_month(stdscr, 1)
            return True
        if ch in (ord("{"), ord("<")):
            self._shift_month(stdscr, -12)
            return True
        if ch in (ord("}"), ord(">")):
            self._shift_month(stdscr, 12)
            return True

        # Jump to Today
        if ch in (ord("t"), ord("T")):
            today = datetime.date.today()
            self._set_selected_date(stdscr, today)
            self.status_message = f"Jumped to Today ({today.strftime('%A, %b %d, %Y')})."
            return True

        # Enter / Space / D = Open Day View for selected date
        if ch in (10, 13, curses.KEY_ENTER, ord(" "), ord("d"), ord("D")):
            self._run_day_view(stdscr)
            return True

        # N / F6 = Create New Event directly on selected date
        if ch in (ord("n"), ord("N"), curses.KEY_F6):
            self._modal_event_form(stdscr, default_date=self.selected_date)
            return True

        # F5 / R = Force full two-way sync with Google Calendar
        if ch in (curses.KEY_F5, ord("r"), ord("R")):
            self._sync_month_with_feedback(
                stdscr,
                self.view_year,
                self.view_month,
                refresh_calendars=True,
                push_pending=True,
            )
            return True

        # F7 / C = Select visible & default Google Calendars
        if ch in (curses.KEY_F7, ord("c"), ord("C")):
            self._modal_select_calendars(stdscr)
            return True

        # F2 / A / '/' = Upcoming Agenda & Search
        if ch in (curses.KEY_F2, ord("a"), ord("A"), ord("/")):
            self._run_agenda_view(stdscr, start_search=(ch == ord("/")))
            return True

        # F4 / G = Jump to specific Date
        if ch in (curses.KEY_F4, ord("g"), ord("G")):
            self._modal_goto_date(stdscr)
            return True

        # F10 = Configure Google OAuth / Credentials (CFGDRV)
        if ch == curses.KEY_F10:
            self._open_google_setup(stdscr)
            return True

        return True

    def _open_google_setup(self, stdscr):
        setup = GoogleDriveSetupScreen(
            self.gdrive,
            system_name=self.system_name,
            launch_external_cb=self.launch_external_cb,
        )
        setup.run(stdscr)
        if self.gcal.is_connected():
            self._sync_month_with_feedback(
                stdscr,
                self.view_year,
                self.view_month,
                refresh_calendars=True,
                push_pending=True,
            )
        else:
            self.status_message = f"Returned from Google Setup. Status: {self.gcal.get_status_label()}"

    # =========================================================================
    # SCREEN 2: Day Schedule & Event Details (DSPDAY)
    # =========================================================================
    def _run_day_view(self, stdscr):
        sel_idx = 0
        scroll_top = 0

        while True:
            max_y, max_x = stdscr.getmaxyx()
            events = self.gcal.get_events_for_date(self.selected_date)
            if sel_idx >= len(events):
                sel_idx = max(0, len(events) - 1)

            stdscr.erase()
            self._draw_day_screen(stdscr, max_y, max_x, events, sel_idx, scroll_top)
            stdscr.refresh()

            ch = read_key(stdscr, restore_timeout=500)
            if ch == -1:
                continue

            # Back to Fullscreen Month View
            if ch in (27, curses.KEY_F12, curses.KEY_F3, ord("q"), ord("Q")):
                break

            # Navigate event list
            if ch == curses.KEY_UP:
                if events:
                    sel_idx = max(0, sel_idx - 1)
                    if sel_idx < scroll_top:
                        scroll_top = sel_idx
                continue
            if ch == curses.KEY_DOWN:
                if events:
                    sel_idx = min(len(events) - 1, sel_idx + 1)
                    list_h = max(3, (max_y - 10) // 2)
                    if sel_idx >= scroll_top + list_h:
                        scroll_top = sel_idx - list_h + 1
                continue

            # Prev / Next Day right inside Day View
            if ch in (curses.KEY_LEFT, ord("["), ord("<")):
                self._set_selected_date(stdscr, self.selected_date - datetime.timedelta(days=1))
                sel_idx = 0
                scroll_top = 0
                continue
            if ch in (curses.KEY_RIGHT, ord("]"), ord(">")):
                self._set_selected_date(stdscr, self.selected_date + datetime.timedelta(days=1))
                sel_idx = 0
                scroll_top = 0
                continue

            # Jump to Today
            if ch in (ord("t"), ord("T")):
                self._set_selected_date(stdscr, datetime.date.today())
                sel_idx = 0
                scroll_top = 0
                continue

            # N / F6 / C = Create New Event on this day
            if ch in (ord("n"), ord("N"), ord("c"), ord("C"), curses.KEY_F6):
                created = self._modal_event_form(stdscr, default_date=self.selected_date)
                if created:
                    new_events = self.gcal.get_events_for_date(self.selected_date)
                    for i, e in enumerate(new_events):
                        if e is created:
                            sel_idx = i
                            break
                continue

            # Enter / E / F2 = Edit Selected Event (or Create New if day is empty!)
            if ch in (10, 13, curses.KEY_ENTER, ord("e"), ord("E"), curses.KEY_F2):
                if not events:
                    self._modal_event_form(stdscr, default_date=self.selected_date)
                else:
                    self._modal_event_form(stdscr, edit_event=events[sel_idx])
                continue

            # Delete / D / F4 = Delete Selected Event
            if ch in (curses.KEY_DC, ord("d"), ord("D"), curses.KEY_F4):
                if events:
                    self._modal_delete_event(stdscr, events[sel_idx])
                continue

            # F5 / R = Sync with Google Calendar
            if ch in (curses.KEY_F5, ord("r"), ord("R")):
                self.status_message = "Syncing with Google Calendar..."
                stdscr.erase()
                self._draw_day_screen(stdscr, max_y, max_x, events, sel_idx, scroll_top)
                stdscr.refresh()
                _, msg = self.gcal.sync_month(
                    self.selected_date.year,
                    self.selected_date.month,
                    refresh_calendars=True,
                    push_pending=True,
                )
                self.status_message = msg
                continue

            # F7 = Select Calendars
            if ch == curses.KEY_F7:
                self._modal_select_calendars(stdscr)
                continue

            # F10 = Google OAuth Setup
            if ch == curses.KEY_F10:
                self._open_google_setup(stdscr)
                continue

    def _draw_day_screen(
        self,
        win,
        max_y: int,
        max_x: int,
        events: List[CalendarEvent],
        sel_idx: int,
        scroll_top: int,
    ):
        try:
            curses.curs_set(0)
        except curses.error:
            pass

        now = datetime.datetime.now()
        is_today = self.selected_date == now.date()
        day_title = self.selected_date.strftime("%A, %B %d, %Y")
        if is_today:
            day_title += " [TODAY]"

        # Row 0: Header
        safe_addstr(win, 0, 1, "DSPDAY", curses.A_BOLD)
        tx = max(9, (max_x - len(day_title)) // 2)
        safe_addstr(win, 0, tx, day_title, curses.A_BOLD)
        sys_str = f"System: {self.system_name}"
        if max_x - len(sys_str) - 1 > tx + len(day_title):
            safe_addstr(win, 0, max_x - len(sys_str) - 1, sys_str)

        # Row 1: Meta
        left_meta = (
            f"User: {self.current_user}   Events: {len(events)}   "
            f"Sync: {self.gcal.get_status_label()}"
        )
        dt_str = now.strftime("%m/%d/%y %H:%M")
        safe_addstr(win, 1, 1, left_meta[: max(10, max_x - len(dt_str) - 3)])
        safe_addstr(win, 1, max(1, max_x - len(dt_str) - 1), dt_str)
        safe_addstr(win, 2, 1, "=" * max(1, max_x - 2), curses.A_BOLD)

        # Split body between Event List (top) and Event Details Box (bottom)
        body_top = 3
        body_bottom = max_y - 3
        body_h = max(8, body_bottom - body_top)

        detail_h = min(14, max(8, body_h // 2))
        list_h = max(3, body_h - detail_h - 2)  # -1 col header, -1 spacer

        # Column widths
        time_w = 18
        sync_w = 7
        cal_w = 16 if max_x >= 76 else 10
        title_w = max(16, max_x - 4 - time_w - cal_w - sync_w - 4)

        # Row 3: Column Headers
        hdr = (
            f" {'Time':<{time_w}} {'Event Title':<{title_w}} "
            f"{'Calendar':<{cal_w}} {'Sync':<{sync_w}}"
        )
        safe_addstr(win, body_top, 1, hdr[: max_x - 2], curses.A_BOLD | curses.A_UNDERLINE)

        list_start_y = body_top + 1
        if not events:
            mid_y = list_start_y + max(1, list_h // 2 - 1)
            msg1 = f"No events scheduled for {self.selected_date.strftime('%A, %B %d, %Y')}."
            msg2 = "Press [N] or [F6] (or [Enter]) to create a new event."
            safe_addstr(win, mid_y, max(2, (max_x - len(msg1)) // 2), msg1, curses.A_BOLD)
            safe_addstr(win, mid_y + 1, max(2, (max_x - len(msg2)) // 2), msg2)
        else:
            visible = events[scroll_top : scroll_top + list_h]
            for i, ev in enumerate(visible):
                abs_idx = scroll_top + i
                y = list_start_y + i
                t_lbl = ev.time_range_label()[:time_w]
                s_lbl = ev.summary[:title_w]
                c_lbl = ev.calendar_name[:cal_w]
                if ev.dirty:
                    sync_lbl = "LOCAL*" if not ev.google_id else "MOD*"
                elif ev.google_id:
                    sync_lbl = "GCAL"
                else:
                    sync_lbl = "LOCAL"
                line = f" {t_lbl:<{time_w}} {s_lbl:<{title_w}} {c_lbl:<{cal_w}} {sync_lbl:<{sync_w}}"
                attr = (curses.A_REVERSE | curses.A_BOLD) if abs_idx == sel_idx else curses.A_NORMAL
                safe_addstr(win, y, 1, line[: max_x - 2], attr)

            if len(events) > list_h:
                scroll_hint = f"Showing {scroll_top + 1}-{min(len(events), scroll_top + list_h)} of {len(events)}"
                safe_addstr(win, list_start_y + list_h, max(1, max_x - len(scroll_hint) - 2), scroll_hint, curses.A_DIM)

        # Details Box
        det_y = body_bottom - detail_h
        draw_5250_box(win, det_y, 1, detail_h, max_x - 2, "Event Details")

        if events and 0 <= sel_idx < len(events):
            ev = events[sel_idx]
            inner_w = max(20, max_x - 8)
            row = det_y + 1

            def put_field(lbl: str, val: str, bold_val: bool = False):
                nonlocal row
                if row >= det_y + detail_h - 1:
                    return
                safe_addstr(win, row, 3, f"{lbl:<11}: ", curses.A_BOLD)
                safe_addstr(
                    win,
                    row,
                    16,
                    val[: max(1, inner_w - 13)],
                    curses.A_BOLD if bold_val else curses.A_NORMAL,
                )
                row += 1

            put_field("Title", ev.summary, bold_val=True)
            if ev.all_day:
                if ev.end_date > ev.start_date:
                    when_str = f"All Day  ({ev.start_date} through {ev.end_date})"
                else:
                    when_str = f"All Day  ({ev.start_date})"
            else:
                st12 = format_time_12h(ev.start_time)
                et12 = format_time_12h(ev.end_time)
                if ev.end_date > ev.start_date:
                    when_str = f"{ev.start_date} {ev.start_time} -> {ev.end_date} {ev.end_time} ({get_local_timezone_name()})"
                else:
                    when_str = f"{ev.start_date}  {ev.start_time} - {ev.end_time} ({st12} - {et12})"
            put_field("When", when_str)

            cal_line = ev.calendar_name
            if ev.recurrence:
                cal_line += f"   |   Repeat: {ev.recurrence.title()}"
            elif ev.recurring_event_id:
                cal_line += "   |   (Recurring event instance)"
            put_field("Calendar", cal_line)

            if ev.location:
                put_field("Location", ev.location)

            sync_desc = (
                "Synced with Google Calendar"
                if (ev.google_id and not ev.dirty)
                else ("Modified locally (Press F5 to sync)" if ev.google_id else "Local event (Press F5 to sync to Google)")
            )
            put_field("Sync State", sync_desc)

            if ev.description and row < det_y + detail_h - 1:
                safe_addstr(win, row, 3, "Notes      : ", curses.A_BOLD)
                wrapped = []
                for para in ev.description.splitlines():
                    wrapped.extend(textwrap.wrap(para, width=max(16, inner_w - 13)) or [""])
                for idx_w, wline in enumerate(wrapped):
                    if row >= det_y + detail_h - 1:
                        break
                    safe_addstr(win, row, 16, wline[: max(1, inner_w - 13)])
                    row += 1
        else:
            safe_addstr(
                win,
                det_y + 2,
                4,
                "Select an event above to inspect its full details, or press [N] / [F6] to add one.",
            )

        # Bottom Function Keys & Status
        safe_addstr(
            win,
            max_y - 3,
            1,
            "N/F6=New Event  Enter/E=Edit  D/Del=Delete  Left/Right=Prev/Next Day  T=Today"[: max_x - 2],
        )
        safe_addstr(
            win,
            max_y - 2,
            1,
            "F5=Sync Google  F7=Calendars  F10=Google Setup  F12/Esc=Back to Month View"[: max_x - 2],
        )
        if self.status_message:
            safe_addstr(win, max_y - 1, 1, self.status_message[: max_x - 2], curses.A_BOLD)

    # =========================================================================
    # MODAL: Create / Edit Event Form (ADDEVT / CHGEVT)
    # =========================================================================
    def _modal_event_form(
        self,
        stdscr,
        default_date: Optional[datetime.date] = None,
        edit_event: Optional[CalendarEvent] = None,
    ) -> Optional[CalendarEvent]:
        is_edit = edit_event is not None
        base_date = default_date or self.selected_date

        writable_cals = self.gcal.get_writable_calendars()
        cal_ids = [c.cal_id for c in writable_cals]

        if is_edit and edit_event is not None:
            summary = edit_event.summary
            start_date = edit_event.start_date
            end_date = edit_event.end_date
            all_day = edit_event.all_day
            start_time = edit_event.start_time or "09:00"
            end_time = edit_event.end_time or "10:00"
            location = edit_event.location
            description = edit_event.description
            cal_idx = cal_ids.index(edit_event.calendar_id) if edit_event.calendar_id in cal_ids else 0
            rec_codes = [r[0] for r in REPEAT_OPTIONS]
            rec_idx = rec_codes.index(edit_event.recurrence) if edit_event.recurrence in rec_codes else 0
            end_date_touched = end_date != start_date
            end_time_touched = True
        else:
            summary = ""
            start_date = base_date.isoformat()
            end_date = base_date.isoformat()
            all_day = False
            now = datetime.datetime.now()
            if base_date == now.date() and now.hour < 22:
                sh = now.hour + 1
                start_time = format_hhmm(sh, 0)
                end_time = format_hhmm(sh + 1, 0)
            else:
                start_time = "09:00"
                end_time = "10:00"
            location = ""
            description = ""
            def_cal = self.gcal.get_default_calendar_id()
            cal_idx = cal_ids.index(def_cal) if def_cal in cal_ids else 0
            rec_idx = 0
            end_date_touched = False
            end_time_touched = False

        # Form state
        values: Dict[str, Any] = {
            "summary": summary,
            "start_date": start_date,
            "all_day": all_day,
            "start_time": start_time,
            "end_time": end_time,
            "end_date": end_date,
            "location": location,
            "cal_idx": cal_idx,
            "rec_idx": rec_idx,
            "description": description,
        }

        fields = [
            {"key": "summary", "label": "Event Title", "kind": "text", "max": 120},
            {"key": "start_date", "label": "Start Date (YYYY-MM-DD)", "kind": "text", "max": 16},
            {"key": "all_day", "label": "All-Day Event [Space=Y/N]", "kind": "bool", "max": 10},
            {"key": "start_time", "label": "Start Time (HH:MM or 2pm)", "kind": "time", "max": 12},
            {"key": "end_time", "label": "End Time   (HH:MM or 3pm)", "kind": "time", "max": 12},
            {"key": "end_date", "label": "End Date   (YYYY-MM-DD)", "kind": "text", "max": 16},
            {"key": "location", "label": "Location", "kind": "text", "max": 120},
            {"key": "cal_idx", "label": "Calendar [Left/Right]", "kind": "cal", "max": 40},
            {"key": "rec_idx", "label": "Repeat   [Left/Right]", "kind": "repeat", "max": 20},
            {"key": "description", "label": "Description / Notes", "kind": "text", "max": 400},
        ]

        field_idx = 0
        cursor_pos = len(values["summary"])
        err_msg = ""

        def is_disabled(idx: int) -> bool:
            return bool(values["all_day"]) and fields[idx]["kind"] == "time"

        def move_field(delta: int):
            nonlocal field_idx, cursor_pos
            n = len(fields)
            for _ in range(n):
                field_idx = (field_idx + delta) % n
                if not is_disabled(field_idx):
                    break
            k = fields[field_idx]["key"]
            if fields[field_idx]["kind"] in ("text", "time"):
                cursor_pos = len(str(values[k]))
            else:
                cursor_pos = 0

        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw = min(84, max_x - 2)
            # Compact spacing if terminal height < 28
            sparse = max_y >= 29
            row_step = 2 if sparse else 1
            bh = min(max_y - 2, len(fields) * row_step + 7)
            sy = max(1, (max_y - bh) // 2)
            sx = max(1, (max_x - bw) // 2)

            title = (
                "Edit Calendar Event (CHGEVT)"
                if is_edit
                else "Create New Calendar Event (ADDEVT)"
            )
            draw_5250_box(stdscr, sy, sx, bh, bw, title)

            label_w = 27
            input_x = sx + label_w + 4
            input_w = max(12, bw - label_w - 7)

            cur_screen_y = sy + 2
            cur_screen_x = input_x

            for i, f in enumerate(fields):
                y = sy + 2 + i * row_step
                if y >= sy + bh - 3:
                    break
                disabled = is_disabled(i)
                active = i == field_idx
                lbl_attr = curses.A_DIM if disabled else (curses.A_BOLD if active else curses.A_NORMAL)
                safe_addstr(win := stdscr, y, sx + 2, f"{f['label']:<{label_w}}: ", lbl_attr)

                kind = f["kind"]
                key = f["key"]
                if kind == "bool":
                    disp = "[X] Yes (All Day)" if values["all_day"] else "[ ] No  (Timed Event)"
                    attr = (curses.A_REVERSE | curses.A_BOLD) if active else curses.A_NORMAL
                    safe_addstr(win, y, input_x, f"{disp:<{input_w}}"[:input_w], attr)
                elif kind == "cal":
                    cal_obj = writable_cals[values["cal_idx"]]
                    disp = f"< {cal_obj.summary} >" if len(writable_cals) > 1 else cal_obj.summary
                    attr = (curses.A_REVERSE | curses.A_BOLD) if active else curses.A_NORMAL
                    safe_addstr(win, y, input_x, f"{disp:<{input_w}}"[:input_w], attr)
                elif kind == "repeat":
                    _, r_label = REPEAT_OPTIONS[values["rec_idx"]]
                    disp = f"< {r_label} >"
                    attr = (curses.A_REVERSE | curses.A_BOLD) if active else curses.A_NORMAL
                    safe_addstr(win, y, input_x, f"{disp:<{input_w}}"[:input_w], attr)
                else:
                    raw_val = "(All-Day)" if disabled else str(values[key])
                    # Horizontal scroll for long text fields
                    view_offset = 0
                    if active and cursor_pos >= input_w:
                        view_offset = cursor_pos - input_w + 1
                    visible_txt = raw_val[view_offset : view_offset + input_w]
                    padded = f"{visible_txt:<{input_w}}"[:input_w]
                    if disabled:
                        attr = curses.A_DIM
                    elif active:
                        attr = curses.A_REVERSE | curses.A_BOLD
                    else:
                        attr = curses.A_UNDERLINE
                    safe_addstr(win, y, input_x, padded, attr)
                    if active and not disabled:
                        cur_screen_y = y
                        cur_screen_x = input_x + max(0, min(cursor_pos - view_offset, input_w - 1))

            # Validation / status message inside modal
            if err_msg:
                safe_addstr(stdscr, sy + bh - 3, sx + 2, f"ERROR: {err_msg}"[: bw - 4], curses.A_BOLD)
            else:
                hint = "Tip: Times accept 24h (14:30) or 12h (2:30pm)."
                safe_addstr(stdscr, sy + bh - 3, sx + 2, hint[: bw - 4], curses.A_DIM)

            fkey_txt = "Enter/F2=Save Event   Tab/Arrows=Next Field   Space=Toggle   F12/Esc=Cancel"
            safe_addstr(stdscr, sy + bh - 2, sx + 2, fkey_txt[: bw - 4], curses.A_BOLD)

            cur_kind = fields[field_idx]["kind"]
            try:
                if cur_kind in ("text", "time") and not is_disabled(field_idx):
                    curses.curs_set(1)
                    stdscr.move(cur_screen_y, cur_screen_x)
                else:
                    curses.curs_set(0)
            except curses.error:
                pass

            stdscr.refresh()
            ch = read_key(stdscr, restore_timeout=500)
            if ch == -1:
                continue

            # Cancel
            if ch in (27, curses.KEY_F12, curses.KEY_F3):
                try:
                    curses.curs_set(0)
                except curses.error:
                    pass
                return None

            # Field navigation
            if ch in (9, curses.KEY_DOWN):
                move_field(1)
                continue
            if ch in (curses.KEY_BTAB, curses.KEY_UP):
                move_field(-1)
                continue

            # Save Event (Enter, F2, Ctrl+S)
            if ch in (10, 13, curses.KEY_ENTER, curses.KEY_F2, 19):
                err_msg = ""
                if not str(values["summary"]).strip():
                    err_msg = "Event Title is required."
                    field_idx = 0
                    cursor_pos = len(str(values["summary"]))
                    continue

                target_cal_id = writable_cals[values["cal_idx"]].cal_id
                rec_code = REPEAT_OPTIONS[values["rec_idx"]][0]

                # Show saving banner before network call
                safe_addstr(
                    stdscr,
                    sy + bh - 3,
                    sx + 2,
                    f"{'Saving to Google Calendar...' if self.gcal.is_connected() else 'Saving event...':<{bw - 4}}"[: bw - 4],
                    curses.A_BOLD,
                )
                stdscr.refresh()

                if is_edit and edit_event is not None:
                    ok, msg = self.gcal.update_event(
                        edit_event,
                        summary=values["summary"],
                        start_date=values["start_date"],
                        end_date=values["end_date"],
                        start_time=values["start_time"],
                        end_time=values["end_time"],
                        all_day=bool(values["all_day"]),
                        location=values["location"],
                        description=values["description"],
                        calendar_id=target_cal_id,
                        recurrence=rec_code,
                    )
                    if not ok:
                        err_msg = msg
                        continue
                    saved_ev = edit_event
                else:
                    ok, msg, saved_ev = self.gcal.create_event(
                        summary=values["summary"],
                        start_date=values["start_date"],
                        end_date=values["end_date"],
                        start_time=values["start_time"],
                        end_time=values["end_time"],
                        all_day=bool(values["all_day"]),
                        location=values["location"],
                        description=values["description"],
                        calendar_id=target_cal_id,
                        recurrence=rec_code,
                    )
                    if not ok:
                        err_msg = msg
                        continue

                self.status_message = msg
                # Follow the saved event's date if user changed start_date
                if saved_ev and saved_ev.start_date_obj():
                    sd = saved_ev.start_date_obj()
                    if sd:
                        self.selected_date = sd
                        self.view_year = sd.year
                        self.view_month = sd.month
                try:
                    curses.curs_set(0)
                except curses.error:
                    pass
                return saved_ev

            # Field-specific editing
            cur_f = fields[field_idx]
            key = cur_f["key"]
            kind = cur_f["kind"]

            if kind == "bool":
                if ch in (ord(" "), curses.KEY_LEFT, curses.KEY_RIGHT):
                    values["all_day"] = not bool(values["all_day"])
                elif ch in (ord("y"), ord("Y")):
                    values["all_day"] = True
                elif ch in (ord("n"), ord("N")):
                    values["all_day"] = False
                continue

            if kind == "cal":
                if ch in (ord(" "), curses.KEY_RIGHT):
                    values["cal_idx"] = (values["cal_idx"] + 1) % len(writable_cals)
                elif ch == curses.KEY_LEFT:
                    values["cal_idx"] = (values["cal_idx"] - 1) % len(writable_cals)
                continue

            if kind == "repeat":
                if ch in (ord(" "), curses.KEY_RIGHT):
                    values["rec_idx"] = (values["rec_idx"] + 1) % len(REPEAT_OPTIONS)
                elif ch == curses.KEY_LEFT:
                    values["rec_idx"] = (values["rec_idx"] - 1) % len(REPEAT_OPTIONS)
                continue

            # Text & Time editing with inline cursor
            text = str(values[key])
            if ch == curses.KEY_LEFT:
                cursor_pos = max(0, cursor_pos - 1)
            elif ch == curses.KEY_RIGHT:
                cursor_pos = min(len(text), cursor_pos + 1)
            elif ch == curses.KEY_HOME:
                cursor_pos = 0
            elif ch == curses.KEY_END:
                cursor_pos = len(text)
            elif ch == 21:  # Ctrl+U clear field
                text = ""
                cursor_pos = 0
            elif ch in (curses.KEY_BACKSPACE, 127, 8):
                if cursor_pos > 0:
                    text = text[: cursor_pos - 1] + text[cursor_pos:]
                    cursor_pos -= 1
            elif ch == curses.KEY_DC:
                if cursor_pos < len(text):
                    text = text[:cursor_pos] + text[cursor_pos + 1 :]
            elif 32 <= ch <= 126 and len(text) < cur_f["max"]:
                text = text[:cursor_pos] + chr(ch) + text[cursor_pos:]
                cursor_pos += 1
            else:
                continue

            values[key] = text
            err_msg = ""

            # Smart auto-sync of End Date / End Time if not manually overridden
            if key == "start_date" and not end_date_touched:
                values["end_date"] = text
            elif key == "end_date":
                end_date_touched = True
            elif key == "start_time" and not end_time_touched:
                pt = parse_time_str(text)
                if pt:
                    new_et, _ = add_minutes_to_hhmm(format_hhmm(pt[0], pt[1]), 60)
                    values["end_time"] = new_et
            elif key == "end_time":
                end_time_touched = True

    # =========================================================================
    # MODAL: Delete Event Confirmation
    # =========================================================================
    def _modal_delete_event(self, stdscr, ev: CalendarEvent):
        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw = min(68, max_x - 4)
            bh = 9
            sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
            draw_5250_box(stdscr, sy, sx, bh, bw, "Confirm Delete Event (DLTEVT)")

            safe_addstr(stdscr, sy + 2, sx + 3, f"Delete event: {ev.summary}"[: bw - 6], curses.A_BOLD)
            safe_addstr(
                stdscr,
                sy + 3,
                sx + 3,
                f"Date/Time   : {ev.start_date}  {ev.time_range_label()}"[: bw - 6],
            )
            if ev.google_id and self.gcal.is_connected():
                safe_addstr(stdscr, sy + 4, sx + 3, "This will also remove the event from Google Calendar."[: bw - 6])
            safe_addstr(stdscr, sy + 6, sx + 3, "Press [Y] to Delete, or [N] / [Esc] to Cancel.", curses.A_BOLD)
            stdscr.refresh()

            ch = read_key(stdscr)
            if ch in (ord("y"), ord("Y")):
                _, msg = self.gcal.delete_event(ev)
                self.status_message = msg
                return
            if ch in (ord("n"), ord("N"), 27, curses.KEY_F12, curses.KEY_F3):
                self.status_message = "Delete cancelled."
                return

    # =========================================================================
    # MODAL: Select Visible / Default Google Calendars (SELCAL - F7)
    # =========================================================================
    def _modal_select_calendars(self, stdscr):
        sel_idx = 0
        changed = False
        while True:
            max_y, max_x = stdscr.getmaxyx()
            cals = self.gcal.calendars
            if sel_idx >= len(cals):
                sel_idx = max(0, len(cals) - 1)

            bw = min(78, max_x - 4)
            bh = min(max_y - 4, max(12, len(cals) + 8))
            sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
            draw_5250_box(stdscr, sy, sx, bh, bw, "Select Google Calendars (SELCAL)")

            def_id = self.gcal.get_default_calendar_id()
            list_h = max(3, bh - 6)
            for i, cal in enumerate(cals[:list_h]):
                y = sy + 2 + i
                chk = "[X]" if cal.enabled else "[ ]"
                def_tag = " (DEFAULT)" if cal.cal_id == def_id else ""
                role_tag = "RW" if cal.writable else "RO"
                name_w = max(16, bw - 26)
                line = f" {chk} {cal.summary[:name_w]:<{name_w}}  [{role_tag}]{def_tag}"
                attr = (curses.A_REVERSE | curses.A_BOLD) if i == sel_idx else curses.A_NORMAL
                safe_addstr(stdscr, y, sx + 2, f"{line:<{bw - 4}}"[: bw - 4], attr)

            safe_addstr(
                stdscr,
                sy + bh - 3,
                sx + 2,
                "Space/Enter=Toggle Visible   D=Set Default   F5=Refresh List"[: bw - 4],
            )
            safe_addstr(
                stdscr,
                sy + bh - 2,
                sx + 2,
                "F12/Esc=Done & Apply"[: bw - 4],
                curses.A_BOLD,
            )
            stdscr.refresh()

            ch = read_key(stdscr)
            if ch in (27, curses.KEY_F12, curses.KEY_F3, ord("q"), ord("Q")):
                self.gcal.save_config()
                self.gcal.save_cache()
                if changed and self.gcal.is_connected():
                    self.gcal.synced_months.clear()
                    self._sync_month_with_feedback(
                        stdscr,
                        self.view_year,
                        self.view_month,
                        refresh_calendars=False,
                        push_pending=False,
                    )
                return
            if ch == curses.KEY_UP and cals:
                sel_idx = (sel_idx - 1) % len(cals)
            elif ch == curses.KEY_DOWN and cals:
                sel_idx = (sel_idx + 1) % len(cals)
            elif ch in (ord(" "), 10, 13, curses.KEY_ENTER) and cals:
                cal = cals[sel_idx]
                cal.enabled = not cal.enabled
                if not any(c.enabled for c in cals):
                    cal.enabled = True
                changed = True
            elif ch in (ord("d"), ord("D")) and cals:
                cal = cals[sel_idx]
                if cal.writable:
                    self.gcal.set_default_calendar_id(cal.cal_id)
                    self.status_message = f"Default calendar set to '{cal.summary}'."
                    changed = True
                else:
                    self.status_message = f"'{cal.summary}' is read-only and cannot be default."
            elif ch == curses.KEY_F5:
                _, msg = self.gcal.sync_calendar_list()
                self.status_message = msg
                changed = True

    # =========================================================================
    # MODAL: Go To Date (F4 / G)
    # =========================================================================
    def _modal_goto_date(self, stdscr):
        buf = self.selected_date.isoformat()
        err = ""
        while True:
            max_y, max_x = stdscr.getmaxyx()
            bw = min(62, max_x - 4)
            bh = 9
            sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
            draw_5250_box(stdscr, sy, sx, bh, bw, "Jump to Date (GODATE)")

            safe_addstr(stdscr, sy + 2, sx + 3, "Enter date (YYYY-MM-DD, MM/DD/YYYY, YYYY-MM, or 'today'):"[: bw - 6])
            safe_addstr(stdscr, sy + 4, sx + 3, "Date: ", curses.A_BOLD)
            fw = max(12, bw - 14)
            safe_addstr(stdscr, sy + 4, sx + 9, f"{buf:<{fw}}"[:fw], curses.A_REVERSE | curses.A_BOLD)
            if err:
                safe_addstr(stdscr, sy + 5, sx + 3, err[: bw - 6], curses.A_BOLD)
            safe_addstr(stdscr, sy + 7, sx + 3, "Enter=Jump   F12/Esc=Cancel")
            stdscr.refresh()

            ch = read_key(stdscr)
            if ch in (27, curses.KEY_F12, curses.KEY_F3):
                return
            if ch in (10, 13, curses.KEY_ENTER):
                d = parse_date_str(buf, ref_date=self.selected_date)
                if not d:
                    err = "Could not parse date. Try YYYY-MM-DD or MM/DD/YYYY."
                    continue
                self._set_selected_date(stdscr, d)
                self.status_message = f"Jumped to {d.strftime('%A, %B %d, %Y')}."
                return
            if ch in (curses.KEY_BACKSPACE, 127, 8):
                buf = buf[:-1]
                err = ""
            elif ch == 21:
                buf = ""
                err = ""
            elif 32 <= ch <= 126 and len(buf) < 24:
                buf += chr(ch)
                err = ""

    # =========================================================================
    # SCREEN: Upcoming Agenda & Event Search (DSPAGD - F2 / '/')
    # =========================================================================
    def _run_agenda_view(self, stdscr, start_search: bool = False):
        query = ""
        searching = start_search
        sel_idx = 0
        scroll_top = 0
        start_from = self.selected_date

        while True:
            max_y, max_x = stdscr.getmaxyx()
            items = self.gcal.get_upcoming_events(start_from, days=90, query=query)
            if sel_idx >= len(items):
                sel_idx = max(0, len(items) - 1)

            stdscr.erase()
            safe_addstr(stdscr, 0, 1, "DSPAGD", curses.A_BOLD)
            title = f"Upcoming Agenda (90 Days from {start_from.isoformat()})"
            safe_addstr(stdscr, 0, max(9, (max_x - len(title)) // 2), title, curses.A_BOLD)
            safe_addstr(stdscr, 0, max(1, max_x - len(self.system_name) - 9), f"System: {self.system_name}")

            # Search bar
            q_lbl = "Search Filter (/): "
            safe_addstr(stdscr, 2, 2, q_lbl, curses.A_BOLD)
            q_w = max(16, max_x - len(q_lbl) - 24)
            q_attr = (curses.A_REVERSE | curses.A_BOLD) if searching else curses.A_UNDERLINE
            safe_addstr(stdscr, 2, 2 + len(q_lbl), f"{query:<{q_w}}"[:q_w], q_attr)
            safe_addstr(stdscr, 2, max(2, max_x - 18), f"Matches: {len(items):<5}")

            hdr = f" {'Date':<14} {'Time':<15} {'Event Title':<{max(18, max_x - 50)}} {'Calendar':<14}"
            safe_addstr(stdscr, 4, 1, hdr[: max_x - 2], curses.A_BOLD | curses.A_UNDERLINE)

            list_h = max(4, max_y - 8)
            if not items:
                msg = "No upcoming events found in local cache for this 90-day window."
                safe_addstr(stdscr, 7, max(2, (max_x - len(msg)) // 2), msg, curses.A_BOLD)
            else:
                if sel_idx < scroll_top:
                    scroll_top = sel_idx
                elif sel_idx >= scroll_top + list_h:
                    scroll_top = sel_idx - list_h + 1

                title_w = max(18, max_x - 50)
                for i, (day, ev) in enumerate(items[scroll_top : scroll_top + list_h]):
                    abs_idx = scroll_top + i
                    d_lbl = day.strftime("%a %m/%d/%y")
                    t_lbl = ev.time_range_label()[:15]
                    row_txt = f" {d_lbl:<14} {t_lbl:<15} {ev.summary[:title_w]:<{title_w}} {ev.calendar_name[:14]:<14}"
                    attr = (curses.A_REVERSE | curses.A_BOLD) if abs_idx == sel_idx else curses.A_NORMAL
                    safe_addstr(stdscr, 5 + i, 1, row_txt[: max_x - 2], attr)

            safe_addstr(
                stdscr,
                max_y - 2,
                1,
                "Enter=Open Day  /=Filter  N=New Event  T=From Today  F12/Esc=Back to Month"[: max_x - 2],
            )
            stdscr.refresh()

            ch = read_key(stdscr, restore_timeout=500)
            if ch == -1:
                continue

            if searching:
                if ch in (27, curses.KEY_F12):
                    searching = False
                elif ch in (10, 13, curses.KEY_ENTER):
                    searching = False
                elif ch in (curses.KEY_BACKSPACE, 127, 8):
                    query = query[:-1]
                    sel_idx = 0
                elif ch == 21:
                    query = ""
                    sel_idx = 0
                elif 32 <= ch <= 126 and len(query) < 48:
                    query += chr(ch)
                    sel_idx = 0
                continue

            if ch in (27, curses.KEY_F12, curses.KEY_F3, ord("q"), ord("Q")):
                return
            if ch == ord("/"):
                searching = True
                continue
            if ch == curses.KEY_UP and items:
                sel_idx = max(0, sel_idx - 1)
            elif ch == curses.KEY_DOWN and items:
                sel_idx = min(len(items) - 1, sel_idx + 1)
            elif ch in (ord("t"), ord("T")):
                start_from = datetime.date.today()
                sel_idx = 0
            elif ch in (ord("n"), ord("N"), curses.KEY_F6):
                self._modal_event_form(stdscr, default_date=self.selected_date)
            elif ch in (10, 13, curses.KEY_ENTER) and items:
                jump_day, _ = items[sel_idx]
                self._set_selected_date(stdscr, jump_day)
                self._run_day_view(stdscr)
                return

    # =========================================================================
    # MODAL: Help Screen (F1)
    # =========================================================================
    def _draw_help_modal(self, win, max_y: int, max_x: int):
        bw = min(76, max_x - 4)
        bh = min(max_y - 4, 24)
        sy, sx = (max_y - bh) // 2, (max_x - bw) // 2
        draw_5250_box(win, sy, sx, bh, bw, "OfficeVision/400 Calendar - Help (WRKCAL)")

        lines = [
            "Fullscreen Month View Controls:",
            "  * Arrow Keys        : Move day-by-day across the month grid",
            "  * [Enter] / [Space] : Open Day View (view, edit, delete, or create)",
            "  * [N] or [F6]       : Create New Event on the selected day",
            "  * '[' / ']' / PgUp/Dn: Previous / Next Month  ('{' / '}' = Prev/Next Year)",
            "  * [T]               : Jump back to Today",
            "  * [F4] or [G]       : Jump to any specific date",
            "  * [F2] / [A] / '/'  : Open Upcoming Agenda & Search",
            "",
            "Day View Controls:",
            "  * [N] / [F6]        : Create New Event on this day",
            "  * [Enter] / [E]     : Edit highlighted event (or Create if empty)",
            "  * [D] / [Delete]    : Delete highlighted event (removes from Google)",
            "  * [Left] / [Right]  : Step to Previous / Next Day",
            "",
            "Google Calendar Sync:",
            "  * [F10]             : Configure Google OAuth (shared with OV/400 Drive/1-2-3)",
            "  * [F5] or [R]       : Two-way Sync (push pending changes + pull month)",
            "  * [F7] or [C]       : Select visible & default Google Calendars",
            "",
            "Press [Enter], [Space], [Esc], or [F12] to close Help.",
        ]
        for i, line in enumerate(lines):
            if sy + 2 + i >= sy + bh - 1:
                break
            attr = curses.A_BOLD if line.endswith(":") else curses.A_NORMAL
            safe_addstr(win, sy + 2 + i, sx + 3, line[: bw - 6], attr)


def main():
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
    app = CalendarApp(gdrive)

    def _wrapped(stdscr):
        try:
            curses.use_default_colors()
        except curses.error:
            pass
        app.run(stdscr)

    try:
        curses.wrapper(_wrapped)
    finally:
        if os.environ.get("MDA") == "1" or os.environ.get("TERM") == "vt100":
            sys.stdout.write("\x1b[r\x1b[1;43r\x1b[H\x1b[2J\x1b[0m\x1b[?25h")
        else:
            sys.stdout.write("\x1b[r\x1b[H\x1b[2J\x1b[0m\x1b[?25h")
        sys.stdout.flush()


if __name__ == "__main__":
    main()
