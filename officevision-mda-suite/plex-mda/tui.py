"""Monochrome MDA Terminal User Interface for Plex Music Player."""

import curses
import os
import time
from typing import List, Optional

from db import Album, Artist, MusicDB, Track
from players.base import PlayerStatus
from players.manager import PlayerManager


class MdaTui:
    TABS = ["search", "artists", "albums", "queue"]

    def __init__(self, db: MusicDB, manager: PlayerManager):
        self.db = db
        self.manager = manager

        # Search tab state
        self.tracks: List[Track] = []
        self.selected_idx: int = 0
        self.top_visible_idx: int = 0

        self.search_query: str = ""
        self.search_input_active: bool = False
        self.search_cursor_pos: int = 0

        # Active tab: "search", "artists", "albums", or "queue"
        self.current_tab: str = "artists"

        # Artists tab state (hierarchy: "artists" -> "albums" -> "tracks")
        self.all_artists: List[Artist] = self.db.get_all_artists()
        self.artists_level: str = "artists"
        self.artists_sel_idx: int = 0
        self.artists_top_idx: int = 0
        self.selected_artist: Optional[Artist] = None
        self.artist_albums: List[Album] = []
        self.artist_albums_sel_idx: int = 0
        self.artist_albums_top_idx: int = 0
        self.selected_artist_album: Optional[Album] = None
        self.artist_album_tracks: List[Track] = []
        self.artist_tracks_sel_idx: int = 0
        self.artist_tracks_top_idx: int = 0

        # Albums tab state (hierarchy: "albums" -> "tracks")
        self.all_albums: List[Album] = self.db.get_all_albums()
        self.albums_level: str = "albums"
        self.albums_sel_idx: int = 0
        self.albums_top_idx: int = 0
        self.selected_album: Optional[Album] = None
        self.album_tracks: List[Track] = []
        self.album_tracks_sel_idx: int = 0
        self.album_tracks_top_idx: int = 0

        # Queue tab state
        self.queue_selected_idx: int = 0
        self.queue_top_visible_idx: int = 0

        self.last_visible_lines: int = 25
        self.show_device_dialog: bool = False
        self.device_dialog_sel_idx: int = 0
        self.show_help_dialog: bool = False
        self.status_message: str = "Disconnected. Press [d] to connect to a device first."
        self.status_message_time: float = time.time()

    def set_status(self, msg: str):
        self.status_message = msg
        self.status_message_time = time.time()

    def _on_device_connected(self, ok: bool, name: str, error: str):
        if ok:
            self.set_status(f"Connected to {name}! Press [Space] or [Enter] to play.")
        else:
            self.set_status(f"Connection failed ({name}): {error or 'Device unreachable'}")

    def _on_devices_scanned(self, updated: list, new_names: list, error: str):
        if error:
            self.set_status(f"Network scan error: {error}")
        elif new_names:
            added = ", ".join(new_names)
            self.set_status(f"Scan complete: {len(updated)} devices saved (New: {added})!")
        else:
            self.set_status(f"Scan complete: {len(updated)} devices saved (no new devices found).")

    def run(self):
        import sys
        is_mda = os.environ.get("MDA") == "1" or os.environ.get("TERM") in (
            "vt100",
            "wsvt25",
            "ansi",
            "dumb",
            "",
        )
        if is_mda:
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
            if is_mda:
                sys.stdout.write("\x1b[r\x1b[1;43r\x1b[H\x1b[2J\x1b[0m\x1b[?25h")
            else:
                sys.stdout.write("\x1b[r\x1b[H\x1b[2J\x1b[0m\x1b[?25h")
            sys.stdout.flush()

    def _safe_addstr(self, win, y: int, x: int, text: str, attr=curses.A_NORMAL):
        max_y, max_x = win.getmaxyx()
        if y < 0 or y >= max_y or x < 0 or x >= max_x:
            return
        avail = max_x - x
        if avail <= 0:
            return
        to_print = text[:avail]
        try:
            # If writing to the bottom-right character, addch/addstr can raise curses.error
            if y == max_y - 1 and len(to_print) == avail:
                to_print = to_print[:-1]
            win.addstr(y, x, to_print, attr)
        except curses.error:
            pass

    def _main_loop(self, stdscr):
        try:
            curses.curs_set(0)
        except curses.error:
            pass

        stdscr.timeout(150)  # 150ms poll for smooth progress bar and responsive keys

        if curses.has_colors():
            try:
                curses.use_default_colors()
            except curses.error:
                pass

        while True:
            max_y, max_x = stdscr.getmaxyx()
            stdscr.erase()

            self._draw_header(stdscr, max_x)
            self._draw_now_playing(stdscr, max_x)
            self._draw_search_and_tabs(stdscr, max_x)
            self._draw_list_panel(stdscr, max_y, max_x)
            self._draw_footer(stdscr, max_y, max_x)

            if self.show_device_dialog:
                self._draw_device_dialog(stdscr, max_y, max_x)
            elif self.show_help_dialog:
                self._draw_help_dialog(stdscr, max_y, max_x)

            stdscr.refresh()

            try:
                ch = stdscr.getch()
            except curses.error:
                ch = -1

            if ch == -1:
                continue

            if self.show_device_dialog:
                self._handle_device_input(ch)
            elif self.show_help_dialog:
                self._handle_help_input(ch)
            elif self.search_input_active:
                self._handle_search_input(ch)
            else:
                if not self._handle_normal_input(ch):
                    break

    def _draw_header(self, win, max_x: int):
        dev_name = self.manager.get_current_device_name()
        conn_state = self.manager.connection_state
        status = self.manager.get_status()
        vol_str = f"Vol: {status.volume}%"

        if conn_state == "CONNECTED":
            dev_info = f" Session: [CONNECTED: {dev_name}]  {vol_str} "
        elif conn_state == "CONNECTING":
            dev_info = f" Session: [CONNECTING -> {dev_name}...] "
        elif conn_state == "ERROR":
            dev_info = f" Session: [ERROR: {dev_name}] (Press [d]) "
        else:
            dev_info = " Session: [DISCONNECTED] (Press [d] to connect) "

        top_line = "+" + "-" * (max_x - 2) + "+"
        title = " PLEX MDA MUSIC PLAYER "

        mid_line = f"|{title}"
        padding = max_x - 2 - len(title) - len(dev_info)
        if padding > 0:
            mid_line += " " * padding + dev_info + "|"
        else:
            mid_line = f"|{title} {dev_info.strip()}|"[: max_x - 1] + "|"

        bot_line = "+" + "-" * (max_x - 2) + "+"

        self._safe_addstr(win, 0, 0, top_line)
        self._safe_addstr(win, 1, 0, mid_line, curses.A_BOLD)
        self._safe_addstr(win, 2, 0, bot_line)

    def _draw_now_playing(self, win, max_x: int):
        status = self.manager.get_status()
        state_tag = f"[{status.state}]"
        track_title = status.track_title or "(No track playing)"
        artist = status.artist or ("" if not status.track_title else "Unknown Artist")
        album = status.album or ("" if not status.track_title else "Unknown Album")

        bar_width = min(50, max(15, max_x - 32))
        progress_bar = status.progress_bar(bar_width)
        time_str = status.time_display()

        line1 = f"| NOW PLAYING: {state_tag}  {track_title}"
        line1 = line1.ljust(max_x - 1) + "|"

        line2 = f"| Artist: {artist}   Album: {album}" if (artist or album) else "| Artist: -                      Album: -"
        line2 = line2.ljust(max_x - 1) + "|"

        line3 = f"| Progress: {progress_bar} {time_str}"
        line3 = line3.ljust(max_x - 1) + "|"

        sep = "+" + "-" * (max_x - 2) + "+"

        self._safe_addstr(win, 3, 0, line1, curses.A_BOLD)
        self._safe_addstr(win, 4, 0, line2)
        self._safe_addstr(win, 5, 0, line3)
        self._safe_addstr(win, 6, 0, sep)

    def _get_current_view_mode(self) -> str:
        """Return 'artists', 'albums', or 'tracks' for the currently displayed list."""
        if self.current_tab == "artists":
            return self.artists_level
        elif self.current_tab == "albums":
            return self.albums_level
        return "tracks"

    def _get_current_items(self) -> list:
        if self.current_tab == "search":
            return self.tracks
        elif self.current_tab == "queue":
            return self.manager.get_queue()
        elif self.current_tab == "artists":
            if self.artists_level == "artists":
                return self.all_artists
            elif self.artists_level == "albums":
                return self.artist_albums
            else:
                return self.artist_album_tracks
        elif self.current_tab == "albums":
            if self.albums_level == "albums":
                return self.all_albums
            else:
                return self.album_tracks
        return []

    def _get_selection_state(self):
        if self.current_tab == "search":
            return self.selected_idx, self.top_visible_idx
        elif self.current_tab == "queue":
            return self.queue_selected_idx, self.queue_top_visible_idx
        elif self.current_tab == "artists":
            if self.artists_level == "artists":
                return self.artists_sel_idx, self.artists_top_idx
            elif self.artists_level == "albums":
                return self.artist_albums_sel_idx, self.artist_albums_top_idx
            else:
                return self.artist_tracks_sel_idx, self.artist_tracks_top_idx
        elif self.current_tab == "albums":
            if self.albums_level == "albums":
                return self.albums_sel_idx, self.albums_top_idx
            else:
                return self.album_tracks_sel_idx, self.album_tracks_top_idx
        return 0, 0

    def _set_selection_state(self, sel_idx: int, top_idx: int):
        if self.current_tab == "search":
            self.selected_idx = sel_idx
            self.top_visible_idx = top_idx
        elif self.current_tab == "queue":
            self.queue_selected_idx = sel_idx
            self.queue_top_visible_idx = top_idx
        elif self.current_tab == "artists":
            if self.artists_level == "artists":
                self.artists_sel_idx = sel_idx
                self.artists_top_idx = top_idx
            elif self.artists_level == "albums":
                self.artist_albums_sel_idx = sel_idx
                self.artist_albums_top_idx = top_idx
            else:
                self.artist_tracks_sel_idx = sel_idx
                self.artist_tracks_top_idx = top_idx
        elif self.current_tab == "albums":
            if self.albums_level == "albums":
                self.albums_sel_idx = sel_idx
                self.albums_top_idx = top_idx
            else:
                self.album_tracks_sel_idx = sel_idx
                self.album_tracks_top_idx = top_idx

    def _draw_search_and_tabs(self, win, max_x: int):
        tabs = [
            ("search", "[ 1: Search ]"),
            ("artists", f"[ 2: Artists ({len(self.all_artists)}) ]"),
            ("albums", f"[ 3: Albums ({len(self.all_albums)}) ]"),
            ("queue", f"[ 4: Queue ({len(self.manager.get_queue())}) ]"),
        ]

        # Draw tabs line
        self._safe_addstr(win, 7, 0, "| View: ")
        cur_x = 8
        for tab_id, label in tabs:
            attr = (curses.A_REVERSE | curses.A_BOLD) if self.current_tab == tab_id else curses.A_NORMAL
            self._safe_addstr(win, 7, cur_x, label, attr)
            cur_x += len(label) + 2
        if cur_x < max_x - 1:
            self._safe_addstr(win, 7, cur_x, " " * (max_x - 1 - cur_x))
        self._safe_addstr(win, 7, max_x - 1, "|")

        # Line 8: Search input (when active or on search tab) OR Breadcrumb / Context bar
        prompt_attr = curses.A_NORMAL
        if self.search_input_active:
            s_text = f"Search: [ {self.search_query}_ ] (Enter to search, Esc to cancel)"
            prompt_attr = curses.A_BOLD
        elif self.current_tab == "search":
            s_text = f"Search: [ {self.search_query} ] ([/] Search  [i] Play Next  [a] Queue)"
        elif self.current_tab == "artists":
            if self.artists_level == "artists":
                s_text = (
                    f"Browse Artists (A-Z): {len(self.all_artists)} artists  "
                    f"([Enter] Albums  [[/]] Jump A-Z  [i] Play Next  [A] Queue All)"
                )
            elif self.artists_level == "albums" and self.selected_artist:
                s_text = (
                    f"Artist: {self.selected_artist.name} ({len(self.artist_albums)} albums A-Z)  "
                    f"([Enter] Tracks  [i] Play Next  [a] Queue  [Bksp] Back)"
                )
                prompt_attr = curses.A_BOLD
            elif self.artists_level == "tracks" and self.selected_artist and self.selected_artist_album:
                s_text = (
                    f"Tracks: {self.selected_artist.name} > {self.selected_artist_album.title} "
                    f"({len(self.artist_album_tracks)} tracks)  ([Enter] Play  [i] Next  [a] Queue  [Bksp] Back)"
                )
                prompt_attr = curses.A_BOLD
            else:
                s_text = "Browse Artists (A-Z)"
        elif self.current_tab == "albums":
            if self.albums_level == "albums":
                s_text = (
                    f"Browse Albums (A-Z): {len(self.all_albums)} albums  "
                    f"([Enter] Tracks  [[/]] Jump A-Z  [i] Play Next  [A] Queue Album)"
                )
            elif self.albums_level == "tracks" and self.selected_album:
                s_text = (
                    f"Album: {self.selected_album.title} - {self.selected_album.artist} "
                    f"({len(self.album_tracks)} tracks)  ([Enter] Play  [i] Play Next  [a] Queue  [Bksp] Back)"
                )
                prompt_attr = curses.A_BOLD
            else:
                s_text = "Browse Albums (A-Z)"
        else:
            q_len = len(self.manager.get_queue())
            s_text = f"Playback Queue: {q_len} tracks  ([Enter] Play Selected  [i] Play Next  [c] Clear Queue)"

        s_line = f"| {s_text[:max_x - 4]}"
        s_line = s_line.ljust(max_x - 1) + "|"
        self._safe_addstr(win, 8, 0, s_line, prompt_attr)

        sep = "+" + "-" * (max_x - 2) + "+"
        self._safe_addstr(win, 9, 0, sep)

    def _draw_list_panel(self, win, max_y: int, max_x: int):
        header_y = 10
        list_start_y = 11
        list_end_y = max_y - 3
        visible_lines = list_end_y - list_start_y

        if visible_lines <= 0:
            return
        self.last_visible_lines = visible_lines

        row_width = max_x - 5
        view_mode = self._get_current_view_mode()

        if view_mode == "artists":
            ar_len = max(20, row_width - 32)
            hdr = f"  {'ARTIST NAME (A-Z)':<{ar_len}} | {'ALBUMS':>13} | {'TRACKS':>13}"
        elif view_mode == "albums":
            avail = max(30, row_width - 20)
            al_len = int(avail * 0.56)
            ar_len = avail - al_len
            hdr = f"  {'ALBUM TITLE (A-Z)':<{al_len}} | {'ARTIST':<{ar_len}} | {'YEAR':>4} | {'TRACKS':>7}"
        else:
            avail_taa = max(30, row_width - 4 - 6 - 6)
            t_len = int(avail_taa * 0.40)
            a_len = int(avail_taa * 0.28)
            al_len = avail_taa - t_len - a_len
            hdr = f"  {'#':<4}{'TRACK TITLE':<{t_len}} | {'ARTIST':<{a_len}} | {'ALBUM':<{al_len}} {'DUR':>5}"

        hdr_line = "|" + hdr[:max_x - 2].ljust(max_x - 2) + "|"
        self._safe_addstr(win, header_y, 0, hdr_line, curses.A_UNDERLINE)

        items = self._get_current_items()
        cur_idx, top_idx = self._get_selection_state()

        if items:
            cur_idx = max(0, min(len(items) - 1, cur_idx))
        else:
            cur_idx = 0

        # Adjust scroll bounds
        if cur_idx < top_idx:
            top_idx = cur_idx
        elif cur_idx >= top_idx + visible_lines:
            top_idx = cur_idx - visible_lines + 1

        self._set_selection_state(cur_idx, top_idx)

        active_q_idx = self.manager.current_queue_index

        for i in range(visible_lines):
            item_idx = top_idx + i
            y = list_start_y + i
            if item_idx < len(items):
                item = items[item_idx]
                is_selected = (item_idx == cur_idx)
                is_currently_playing = (self.current_tab == "queue" and item_idx == active_q_idx)

                prefix = "> " if is_selected else ("* " if is_currently_playing else "  ")
                item_line = prefix + item.display_line(row_width)
                full_line = "|" + item_line.ljust(max_x - 2) + "|"

                attr = curses.A_REVERSE if is_selected else curses.A_NORMAL
                if is_currently_playing and not is_selected:
                    attr = curses.A_BOLD

                self._safe_addstr(win, y, 0, full_line, attr)
            else:
                empty_line = "|" + " " * (max_x - 2) + "|"
                self._safe_addstr(win, y, 0, empty_line)

        sep = "+" + "-" * (max_x - 2) + "+"
        self._safe_addstr(win, list_end_y, 0, sep)

    def _draw_footer(self, win, max_y: int, max_x: int):
        help_line = "| [1-4]View [Ent]Play [i]PlayNext [a]Queue [Bksp]Back [d]Dev [Spc]Pause [n/p]Skip [q]Quit"
        help_line = help_line[:max_x - 1].ljust(max_x - 1) + "|"

        # Bottom status message
        if time.time() - self.status_message_time > 10:
            if self.manager.is_connected:
                self.status_message = f"Connected to {self.manager.get_current_device_name()}. Ready."
            elif self.manager.connection_state == "CONNECTING":
                self.status_message = f"Connecting to {self.manager.get_current_device_name()}..."
            else:
                self.status_message = "Disconnected. Press [d] to connect to a device."
        msg_line = f"| MSG: {self.status_message}"
        msg_line = msg_line[:max_x - 1].ljust(max_x - 1) + "|"

        self._safe_addstr(win, max_y - 2, 0, help_line, curses.A_BOLD)
        self._safe_addstr(win, max_y - 1, 0, msg_line)

    def _open_device_dialog(self):
        devices = self.manager.get_devices()
        cur_id = self.manager.current_device_id
        idx = next((i for i, d in enumerate(devices) if d["id"] == cur_id), 0)
        self.device_dialog_sel_idx = idx
        self.show_device_dialog = True

    def _draw_device_dialog(self, win, max_y: int, max_x: int):
        devices = self.manager.get_devices()
        dialog_width = min(68, max_x - 4)
        list_rows = max(1, len(devices))
        dialog_height = list_rows + 9
        start_y = max(2, (max_y - dialog_height) // 2)
        start_x = max(2, (max_x - dialog_width) // 2)

        cur_dev_id = self.manager.current_device_id
        conn_state = self.manager.connection_state
        if devices:
            self.device_dialog_sel_idx = max(0, min(len(devices) - 1, self.device_dialog_sel_idx))

        # Box
        self._safe_addstr(win, start_y, start_x, "+" + "-" * (dialog_width - 2) + "+")
        title = " CONNECT / DISCONNECT PLAYBACK SESSION "
        self._safe_addstr(win, start_y + 1, start_x, "|" + title.center(dialog_width - 2) + "|", curses.A_BOLD)
        self._safe_addstr(win, start_y + 2, start_x, "+" + "-" * (dialog_width - 2) + "+")

        if not devices:
            empty_msg = "  (No saved devices yet - press [r] to scan your LAN)"
            self._safe_addstr(
                win,
                start_y + 3,
                start_x,
                "|" + empty_msg[: dialog_width - 2].ljust(dialog_width - 2) + "|",
            )
        else:
            for i, dev in enumerate(devices):
                num = i + 1
                num_label = f"{num}" if num <= 9 else "-"
                is_cur_dev = (dev["id"] == cur_dev_id)
                is_cursor = (i == self.device_dialog_sel_idx)
                prefix = ">" if is_cursor else " "

                if is_cur_dev and conn_state == "CONNECTED":
                    tag = " [CONNECTED]"
                elif is_cur_dev and conn_state == "CONNECTING":
                    tag = " [CONNECTING]"
                elif is_cur_dev:
                    tag = " [DEFAULT]"
                else:
                    tag = ""

                if is_cursor:
                    attr = curses.A_REVERSE | curses.A_BOLD
                elif is_cur_dev and conn_state in ["CONNECTED", "CONNECTING"]:
                    attr = curses.A_BOLD
                else:
                    attr = curses.A_NORMAL

                alias = dev.get("alias") or dev.get("model") or dev.get("host", "")
                line = f" {prefix}[{num_label}] {dev['name']:<22} ({alias}){tag}"
                line_str = "|" + line[: dialog_width - 2].ljust(dialog_width - 2) + "|"
                self._safe_addstr(win, start_y + 3 + i, start_x, line_str, attr)

        sep_y = start_y + 3 + list_rows
        self._safe_addstr(win, sep_y, start_x, "|" + (" " * (dialog_width - 2)) + "|")

        if self.manager.is_scanning:
            scan_line = "  [r] Scanning network for devices... please wait"
            scan_attr = curses.A_REVERSE | curses.A_BOLD
        else:
            scan_line = "  [r] Scan network for new devices (updates saved list)"
            scan_attr = curses.A_BOLD
        self._safe_addstr(
            win,
            sep_y + 1,
            start_x,
            "|" + scan_line[: dialog_width - 2].ljust(dialog_width - 2) + "|",
            scan_attr,
        )

        disc_line = "  [0] Disconnect active session (close player session)"
        self._safe_addstr(
            win,
            sep_y + 2,
            start_x,
            "|" + disc_line[: dialog_width - 2].ljust(dialog_width - 2) + "|",
            curses.A_BOLD if conn_state in ["CONNECTED", "CONNECTING"] else curses.A_NORMAL,
        )

        inst_y = sep_y + 3
        max_num = min(9, len(devices))
        if max_num > 0:
            inst = f"[1-{max_num}/Ent] Connect  [r] Scan Network  [0] Disconnect  [Esc] Close"
        else:
            inst = "[r] Scan Network for Cast/HEOS Devices   [Esc] Close"
        self._safe_addstr(win, inst_y, start_x, "|" + inst[: dialog_width - 2].center(dialog_width - 2) + "|")
        self._safe_addstr(win, inst_y + 1, start_x, "+" + "-" * (dialog_width - 2) + "+")

    def _draw_help_dialog(self, win, max_y: int, max_x: int):
        dialog_width = min(68, max_x - 4)
        dialog_height = 21
        start_y = max(2, (max_y - dialog_height) // 2)
        start_x = max(2, (max_x - dialog_width) // 2)

        lines = [
            "KEYBOARD SHORTCUTS REFERENCE",
            "----------------------------------------------------",
            " [1]-[4] / [Tab]  Switch Views: Search, Artists, Albums, Queue",
            " [Enter] / [->]   Select Artist/Album, or Play Track",
            " [Bksp]/[Esc]/[b] Go Back up a level (Tracks -> Albums -> Artists)",
            " [[] / []]        Jump to Prev / Next letter (A-Z) in lists",
            " [PgUp] / [PgDn]  Fast scroll up / down one page",
            " [Home] / [End]   Jump to top [g] / bottom [G] of list",
            " [/]              Search Plex music library",
            " [Space]          Toggle Play / Pause (or select when idle)",
            " [i] / [I]        Play Next in queue (track / album or artist)",
            " [a] / [A]        Add track / entire album (or artist) to queue",
            " [s] / [x]        Stop playback      [c] Clear queue",
            " [n] / [p]        Next / Previous track in queue",
            " [+] / [-]        Volume Up / Down   [>] / [<] Seek +/-10s",
            " [d] / [D]        Device Menu ([r] inside menu scans network)",
            " [q]              Disconnect & Quit to terminal",
        ]

        self._safe_addstr(win, start_y, start_x, "+" + "-" * (dialog_width - 2) + "+")
        for i, text in enumerate(lines):
            attr = curses.A_BOLD if i == 0 else curses.A_NORMAL
            self._safe_addstr(win, start_y + 1 + i, start_x, "|" + text.center(dialog_width - 2) + "|", attr)
        self._safe_addstr(win, start_y + 1 + len(lines), start_x, "|" + "Press any key to close".center(dialog_width - 2) + "|")
        self._safe_addstr(win, start_y + 2 + len(lines), start_x, "+" + "-" * (dialog_width - 2) + "+")

    def _connect_device_from_dialog(self, target: dict):
        self.show_device_dialog = False
        self.set_status(f"Connecting to {target['name']} (opening session, please wait)...")
        self.manager.connect_device(
            target["id"],
            blocking=False,
            on_complete=self._on_device_connected,
        )

    def _handle_device_input(self, ch: int):
        devices = self.manager.get_devices()
        if ch in [27, ord("q"), ord("d")]:  # Escape or q or d
            self.show_device_dialog = False
            return

        if ch in [ord("r"), ord("R"), ord("s"), ord("S")]:
            if not self.manager.is_scanning:
                self.set_status("Scanning local network for Chromecast & HEOS devices...")
                self.manager.scan_devices(
                    blocking=False,
                    on_complete=self._on_devices_scanned,
                )
            return

        if ch in [ord("0"), ord("x"), ord("X")]:
            dev_name = self.manager.get_current_device_name()
            self.manager.disconnect_device()
            self.set_status(f"Disconnected session from {dev_name}.")
            self.show_device_dialog = False
            return

        if ch in [curses.KEY_UP, ord("k")]:
            if devices:
                self.device_dialog_sel_idx = max(0, self.device_dialog_sel_idx - 1)
            return

        if ch in [curses.KEY_DOWN, ord("j")]:
            if devices:
                self.device_dialog_sel_idx = min(len(devices) - 1, self.device_dialog_sel_idx + 1)
            return

        if ch in [10, 13]:  # Enter connects to highlighted device
            if 0 <= self.device_dialog_sel_idx < len(devices):
                self._connect_device_from_dialog(devices[self.device_dialog_sel_idx])
            return

        if ord("1") <= ch <= ord("9"):
            idx = ch - ord("1")
            if 0 <= idx < len(devices):
                self.device_dialog_sel_idx = idx
                self._connect_device_from_dialog(devices[idx])

    def _handle_help_input(self, ch: int):
        self.show_help_dialog = False

    def _handle_search_input(self, ch: int):
        if ch in [10, 13]:  # Enter
            self.search_input_active = False
            if self.search_query.strip():
                self.set_status(f"Searching for '{self.search_query}'...")
                self.tracks = self.db.search_tracks(self.search_query, limit=100)
                self.selected_idx = 0
                self.top_visible_idx = 0
                self.current_tab = "search"
                self.set_status(f"Found {len(self.tracks)} tracks.")
        elif ch in [27]:  # Escape
            self.search_input_active = False
            self.set_status("Search cancelled.")
        elif ch in [curses.KEY_BACKSPACE, 127, 8]:
            if len(self.search_query) > 0:
                self.search_query = self.search_query[:-1]
        elif 32 <= ch <= 126:
            self.search_query += chr(ch)

    def _require_connected(self) -> bool:
        conn = self.manager.connection_state
        if conn == "CONNECTED":
            return True
        if conn == "CONNECTING":
            self.set_status(f"Still connecting to {self.manager.get_current_device_name()}... please wait.")
            return False
        self.set_status("Connect to a device first! Select a device or press [r] to scan.")
        self._open_device_dialog()
        return False

    def _navigate_back(self) -> bool:
        """Step back up one level when browsing Artists or Albums. Returns True if stepped back."""
        if self.current_tab == "artists":
            if self.artists_level == "tracks":
                self.artists_level = "albums"
                if self.selected_artist:
                    self.set_status(f"Back to {self.selected_artist.name} albums ({len(self.artist_albums)} albums).")
                return True
            elif self.artists_level == "albums":
                self.artists_level = "artists"
                self.set_status(f"Back to all Artists ({len(self.all_artists)} artists A-Z).")
                return True
        elif self.current_tab == "albums":
            if self.albums_level == "tracks":
                self.albums_level = "albums"
                self.set_status(f"Back to all Albums ({len(self.all_albums)} albums A-Z).")
                return True
        return False

    def _jump_letter(self, direction: int):
        """Jump to the next (+1) or previous (-1) initial letter in the current list."""
        items = self._get_current_items()
        if not items:
            return
        cur_idx, top_idx = self._get_selection_state()

        def first_char(item) -> str:
            text = getattr(item, "name", None) or getattr(item, "title", "")
            cleaned = self.db._alpha_sort_key(text)
            return cleaned[0].upper() if cleaned else "#"

        cur_letter = first_char(items[cur_idx])
        n = len(items)
        if direction > 0:
            for idx in range(cur_idx + 1, n):
                if first_char(items[idx]) != cur_letter:
                    self._set_selection_state(idx, top_idx)
                    self.set_status(f"Jumped to '{first_char(items[idx])}'")
                    return
        else:
            # Find start of previous letter group
            prev_idx = cur_idx
            while prev_idx > 0 and first_char(items[prev_idx - 1]) == cur_letter:
                prev_idx -= 1
            if prev_idx < cur_idx:
                self._set_selection_state(prev_idx, top_idx)
                self.set_status(f"Start of '{cur_letter}'")
                return
            if prev_idx > 0:
                target_letter = first_char(items[prev_idx - 1])
                idx = prev_idx - 1
                while idx > 0 and first_char(items[idx - 1]) == target_letter:
                    idx -= 1
                self._set_selection_state(idx, top_idx)
                self.set_status(f"Jumped to '{target_letter}'")

    def _activate_selected_item(self):
        """Drill down if viewing Artists or Albums; play track if viewing Tracks."""
        view_mode = self._get_current_view_mode()
        items = self._get_current_items()
        if not items:
            return
        cur_idx, _ = self._get_selection_state()
        if not (0 <= cur_idx < len(items)):
            return

        if self.current_tab == "artists" and view_mode == "artists":
            artist: Artist = items[cur_idx]
            self.selected_artist = artist
            self.artist_albums = self.db.get_artist_albums(artist.id)
            self.artist_albums_sel_idx = 0
            self.artist_albums_top_idx = 0
            self.artists_level = "albums"
            self.set_status(f"Viewing {len(self.artist_albums)} albums by {artist.name} (Press [Bksp] to go back)")
            return

        if self.current_tab == "artists" and view_mode == "albums":
            album: Album = items[cur_idx]
            self.selected_artist_album = album
            self.artist_album_tracks = self.db.get_album_tracks(album.id)
            self.artist_tracks_sel_idx = 0
            self.artist_tracks_top_idx = 0
            self.artists_level = "tracks"
            self.set_status(f"Viewing {len(self.artist_album_tracks)} tracks in '{album.title}' (Press [Enter] to play, [Bksp] back)")
            return

        if self.current_tab == "albums" and view_mode == "albums":
            album: Album = items[cur_idx]
            self.selected_album = album
            self.album_tracks = self.db.get_album_tracks(album.id)
            self.album_tracks_sel_idx = 0
            self.album_tracks_top_idx = 0
            self.albums_level = "tracks"
            self.set_status(f"Viewing {len(self.album_tracks)} tracks in '{album.title}' (Press [Enter] to play, [Bksp] back)")
            return

        # Otherwise we are on a track list (search, artist album tracks, album tracks, or queue)
        self._play_selected_track()

    def _play_selected_track(self):
        if not self._require_connected():
            return
        if self.current_tab == "queue":
            q = self.manager.get_queue()
            if 0 <= self.queue_selected_idx < len(q):
                self.manager.current_queue_index = self.queue_selected_idx
                self.manager._trigger_current_play()
                self.set_status(f"Playing queue item {self.queue_selected_idx + 1}")
            return

        items = self._get_current_items()
        cur_idx, _ = self._get_selection_state()
        if not items or not (0 <= cur_idx < len(items)):
            return

        track: Track = items[cur_idx]
        # If viewing an album's tracks directly, queue from selected track to end of album
        if self.current_tab in ["artists", "albums"] and self._get_current_view_mode() == "tracks":
            queue_tracks = items[cur_idx:]
            self.manager.play_album_now(queue_tracks)
            self.set_status(f"Playing '{track.title}' + {len(queue_tracks) - 1} album tracks")
        else:
            album_tracks = self.db.get_album_tracks_for_track(track.id)
            if album_tracks:
                t_idx = next((i for i, t in enumerate(album_tracks) if t.id == track.id), 0)
                queue_tracks = album_tracks[t_idx:]
                self.manager.play_album_now(queue_tracks)
                self.set_status(f"Playing '{track.title}' + {len(queue_tracks) - 1} album tracks")
            else:
                self.manager.play_track_now(track, replace_queue=True)
                self.set_status(f"Playing '{track.title}'")

    def _queue_selected_item(self, whole_collection: bool = False, play_next: bool = False):
        """Add selected track, album, or artist to the playback queue (at end or next)."""
        view_mode = self._get_current_view_mode()
        items = self._get_current_items()
        cur_idx, top_idx = self._get_selection_state()
        if not items or not (0 <= cur_idx < len(items)):
            return

        if self.current_tab == "queue":
            if play_next:
                item = items[cur_idx]
                if cur_idx == self.manager.current_queue_index:
                    self.set_status(f"'{item.title}' is currently playing.")
                else:
                    new_idx = self.manager.move_queue_item_next(cur_idx)
                    self._set_selection_state(new_idx, top_idx)
                    self.set_status(f"Playing next: '{item.title}'")
            return

        item = items[cur_idx]
        if view_mode == "artists":
            artist_tracks = self.db.get_artist_tracks(item.id)
            if artist_tracks:
                if play_next:
                    self.manager.add_album_to_queue_next(artist_tracks)
                    self.set_status(f"Playing next: all {len(artist_tracks)} tracks by {item.name}")
                else:
                    self.manager.add_album_to_queue(artist_tracks)
                    self.set_status(f"Queued all {len(artist_tracks)} tracks by {item.name}")
        elif view_mode == "albums":
            al_tracks = self.db.get_album_tracks(item.id)
            if al_tracks:
                if play_next:
                    self.manager.add_album_to_queue_next(al_tracks)
                    self.set_status(f"Playing next: album '{item.title}' ({len(al_tracks)} tracks)")
                else:
                    self.manager.add_album_to_queue(al_tracks)
                    self.set_status(f"Queued album '{item.title}' ({len(al_tracks)} tracks)")
        else:
            # Track view
            if whole_collection:
                if self.current_tab in ["artists", "albums"]:
                    col_tracks = items
                else:
                    col_tracks = self.db.get_album_tracks_for_track(item.id)
                if play_next:
                    self.manager.add_album_to_queue_next(col_tracks)
                    self.set_status(f"Playing next: album ({len(col_tracks)} tracks)")
                else:
                    self.manager.add_album_to_queue(col_tracks)
                    self.set_status(f"Queued album ({len(col_tracks)} tracks)")
            else:
                if play_next:
                    self.manager.add_to_queue_next(item)
                    self.set_status(f"Playing next: '{item.title}'")
                else:
                    self.manager.add_to_queue(item)
                    self.set_status(f"Queued: '{item.title}'")

    def _handle_normal_input(self, ch: int) -> bool:
        if ch in [ord("q"), ord("Q")]:
            return False  # Exit TUI

        if ch == ord("?"):
            self.show_help_dialog = True
            return True

        if ch == ord("d"):
            self._open_device_dialog()
            return True

        if ch == ord("D"):
            if self.manager.connection_state in ["CONNECTED", "CONNECTING"]:
                dev_name = self.manager.get_current_device_name()
                self.manager.disconnect_device()
                self.set_status(f"Disconnected session from {dev_name}.")
            else:
                self._open_device_dialog()
            return True

        if ch == ord("/"):
            self.search_input_active = True
            self.search_query = ""
            return True

        if ch == 9:  # Tab cycles through Search -> Artists -> Albums -> Queue
            idx = self.TABS.index(self.current_tab) if self.current_tab in self.TABS else 0
            self.current_tab = self.TABS[(idx + 1) % len(self.TABS)]
            return True

        if ch == ord("1"):
            self.current_tab = "search"
            return True

        if ch == ord("2"):
            if self.current_tab == "artists" and self.artists_level != "artists":
                self.artists_level = "artists"
                self.set_status(f"Browsing all {len(self.all_artists)} Artists (A-Z).")
            else:
                self.current_tab = "artists"
            return True

        if ch == ord("3"):
            if self.current_tab == "albums" and self.albums_level != "albums":
                self.albums_level = "albums"
                self.set_status(f"Browsing all {len(self.all_albums)} Albums (A-Z).")
            else:
                self.current_tab = "albums"
            return True

        if ch == ord("4"):
            self.current_tab = "queue"
            return True

        # Back / Up-level navigation for Artists & Albums hierarchy
        if ch in [curses.KEY_BACKSPACE, 127, 8, 27, curses.KEY_LEFT, ord("b"), ord("B"), ord("h")]:
            self._navigate_back()
            return True

        # Right arrow drills down into Artist or Album
        if ch in [curses.KEY_RIGHT, ord("l")]:
            if self._get_current_view_mode() in ["artists", "albums"]:
                self._activate_selected_item()
            return True

        # Alphabetical letter jump
        if ch == ord("]"):
            self._jump_letter(1)
            return True
        if ch == ord("["):
            self._jump_letter(-1)
            return True

        # List Navigation
        items = self._get_current_items()
        cur_idx, top_idx = self._get_selection_state()

        if ch in [curses.KEY_UP, ord("k")]:
            if items:
                self._set_selection_state(max(0, cur_idx - 1), top_idx)
        elif ch in [curses.KEY_DOWN, ord("j")]:
            if items:
                self._set_selection_state(min(len(items) - 1, cur_idx + 1), top_idx)
        elif ch in [curses.KEY_PPAGE, 2]:  # Page Up / Ctrl+B
            step = max(10, self.last_visible_lines - 1)
            if items:
                self._set_selection_state(max(0, cur_idx - step), max(0, top_idx - step))
        elif ch in [curses.KEY_NPAGE, 6]:  # Page Down / Ctrl+F
            step = max(10, self.last_visible_lines - 1)
            if items:
                self._set_selection_state(min(len(items) - 1, cur_idx + step), top_idx)
        elif ch in [curses.KEY_HOME, ord("g")]:
            if items:
                self._set_selection_state(0, 0)
        elif ch in [curses.KEY_END, ord("G")]:
            if items:
                self._set_selection_state(len(items) - 1, top_idx)

        # Selection / Playback controls
        elif ch in [10, 13]:  # Enter
            self._activate_selected_item()

        elif ch == ord(" "):  # Space
            if self._get_current_view_mode() in ["artists", "albums"]:
                st = self.manager.get_status()
                if st.state in ["IDLE", "STOPPED"] or not self.manager.get_queue():
                    self._activate_selected_item()
                else:
                    self.manager.toggle_play_pause()
                    st = self.manager.get_status()
                    self.set_status(f"Playback: [{st.state}]")
            else:
                if not self._require_connected():
                    return True
                st = self.manager.get_status()
                if st.state in ["IDLE", "STOPPED"] or not self.manager.get_queue():
                    self._play_selected_track()
                else:
                    self.manager.toggle_play_pause()
                    st = self.manager.get_status()
                    self.set_status(f"Playback: [{st.state}]")

        elif ch in [ord("s"), ord("x")]:  # Stop
            self.manager.stop()
            self.set_status("Playback stopped.")

        elif ch in [ord("n"), ord("N")]:  # Next
            if self.manager.next_track():
                st = self.manager.get_status()
                self.set_status(f"Next track: {st.track_title}")
            else:
                self.set_status("End of queue.")

        elif ch in [ord("p"), ord("P")]:  # Prev
            if self.manager.prev_track():
                st = self.manager.get_status()
                self.set_status(f"Previous track: {st.track_title}")

        elif ch in [ord("+"), ord("=")]:  # Volume up
            self.manager.volume_up(5)
            st = self.manager.get_status()
            self.set_status(f"Volume: {st.volume}%")

        elif ch in [ord("-"), ord("_")]:  # Volume down
            self.manager.volume_down(5)
            st = self.manager.get_status()
            self.set_status(f"Volume: {st.volume}%")

        elif ch in [ord(">"), ord(".")]:  # Seek forward 10s
            self.manager.seek_relative(10)
            self.set_status("Seek +10s")

        elif ch in [ord("<"), ord(",")]:  # Seek backward 10s
            self.manager.seek_relative(-10)
            self.set_status("Seek -10s")

        elif ch == ord("i"):  # Play selected item next in queue
            self._queue_selected_item(whole_collection=False, play_next=True)

        elif ch == ord("I"):  # Play entire album / artist next in queue
            self._queue_selected_item(whole_collection=True, play_next=True)

        elif ch == ord("a"):  # Add selected item to queue
            self._queue_selected_item(whole_collection=False, play_next=False)

        elif ch == ord("A"):  # Add entire album / artist to queue
            self._queue_selected_item(whole_collection=True, play_next=False)

        elif ch == ord("c"):  # Clear queue
            self.manager.clear_queue()
            self.set_status("Queue cleared.")

        return True
