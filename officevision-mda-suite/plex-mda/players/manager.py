"""Unified Player Manager: handles active device, queues, and auto-advancing."""

import threading
import time
from typing import List, Optional

from config import discover_devices, load_devices, load_settings, save_settings
from db import MusicDB, Track
from players.base import BasePlayer, PlayerStatus
from players.cast import CastPlayer
from players.heos import HeosPlayer
from streamer import AudioStreamer


class PlayerManager:
    def __init__(self, db: MusicDB, streamer: AudioStreamer, auto_connect: bool = False):
        self.db = db
        self.streamer = streamer
        self.settings = load_settings()
        self.devices = load_devices()
        self.is_scanning: bool = False

        self.current_device_id = self.settings.get("selected_device_id", "")
        if not self.current_device_id and self.devices:
            self.current_device_id = self.devices[0].get("id", "")
        self.active_player: Optional[BasePlayer] = None

        self.queue: List[Track] = []
        self.current_queue_index: int = -1

        self._lock = threading.Lock()
        self._running = True
        self._worker_thread = threading.Thread(target=self._monitor_loop, daemon=True)
        self._worker_thread.start()

        if auto_connect:
            self.connect_device(self.current_device_id, blocking=False)

    @property
    def connection_state(self) -> str:
        with self._lock:
            if not self.active_player:
                return "DISCONNECTED"
            return self.active_player.connection_state

    @property
    def is_connected(self) -> bool:
        return self.connection_state == "CONNECTED"

    def get_devices(self) -> List[dict]:
        with self._lock:
            return list(self.devices)

    def scan_devices(self, blocking: bool = False, on_complete=None):
        """Manually scan the network for new or updated devices and persist the list."""
        with self._lock:
            if self.is_scanning:
                return
            self.is_scanning = True
            snapshot = list(self.devices)

        def _do_scan():
            err = ""
            updated = snapshot
            new_names: List[str] = []
            try:
                updated, new_names = discover_devices(snapshot)
                with self._lock:
                    self.devices = updated
            except Exception as e:
                err = str(e)
            finally:
                with self._lock:
                    self.is_scanning = False
                if on_complete:
                    try:
                        on_complete(updated, new_names, err)
                    except Exception:
                        pass
            return updated, new_names

        if blocking:
            return _do_scan()
        else:
            threading.Thread(target=_do_scan, daemon=True).start()
            return None

    def connect_device(self, device_id: str, blocking: bool = False, on_complete=None) -> bool:
        dev = next((d for d in self.devices if d["id"] == device_id), None)
        if not dev:
            return False

        with self._lock:
            old_player = self.active_player
            self.current_device_id = device_id
            self.settings["selected_device_id"] = device_id
            save_settings(self.settings)

            if dev["type"] == "cast":
                new_player = CastPlayer(dev)
            elif dev["type"] == "heos":
                new_player = HeosPlayer(dev)
            else:
                self.active_player = None
                return False

            new_player.connection_state = "CONNECTING"
            self.active_player = new_player

        if old_player:
            try:
                old_player.stop()
                old_player.disconnect()
            except Exception:
                pass

        def _do_connect():
            ok = new_player.connect()
            if on_complete:
                try:
                    on_complete(ok, dev["name"], new_player.last_error)
                except Exception:
                    pass
            return ok

        if blocking:
            return _do_connect()
        else:
            threading.Thread(target=_do_connect, daemon=True).start()
            return True

    def select_device(self, device_id: str, blocking: bool = False, on_complete=None) -> bool:
        return self.connect_device(device_id, blocking=blocking, on_complete=on_complete)

    def disconnect_device(self) -> bool:
        with self._lock:
            player = self.active_player
            self.active_player = None
        if player:
            try:
                player.stop()
                player.disconnect()
            except Exception:
                pass
            return True
        return False

    def get_current_device_name(self) -> str:
        dev = next((d for d in self.devices if d["id"] == self.current_device_id), None)
        return dev["name"] if dev else "None"

    def play_track_now(self, track: Track, replace_queue: bool = False):
        """Play a track immediately. If replace_queue is True, resets queue to this track."""
        with self._lock:
            if replace_queue or not self.queue:
                self.queue = [track]
                self.current_queue_index = 0
            else:
                # Insert right after current
                insert_pos = max(0, self.current_queue_index + 1)
                self.queue.insert(insert_pos, track)
                self.current_queue_index = insert_pos

        self._trigger_current_play()

    def play_album_now(self, tracks: List[Track]):
        if not tracks:
            return
        with self._lock:
            self.queue = list(tracks)
            self.current_queue_index = 0
        self._trigger_current_play()

    def add_to_queue(self, track: Track):
        with self._lock:
            self.queue.append(track)
            if self.current_queue_index == -1:
                self.current_queue_index = 0
                if self.active_player and self.active_player.is_connected:
                    threading.Thread(target=self._trigger_current_play, daemon=True).start()

    def add_to_queue_next(self, track: Track):
        """Insert a track into the queue immediately after the currently playing track."""
        with self._lock:
            if not self.queue or self.current_queue_index < 0:
                self.queue.insert(0, track)
                self.current_queue_index = 0
                if self.active_player and self.active_player.is_connected:
                    threading.Thread(target=self._trigger_current_play, daemon=True).start()
            else:
                insert_pos = min(len(self.queue), self.current_queue_index + 1)
                self.queue.insert(insert_pos, track)

    def add_album_to_queue(self, tracks: List[Track]):
        if not tracks:
            return
        with self._lock:
            was_empty = len(self.queue) == 0
            self.queue.extend(tracks)
            if was_empty:
                self.current_queue_index = 0
                if self.active_player and self.active_player.is_connected:
                    threading.Thread(target=self._trigger_current_play, daemon=True).start()

    def add_album_to_queue_next(self, tracks: List[Track]):
        """Insert a list of tracks into the queue immediately after the currently playing track."""
        if not tracks:
            return
        with self._lock:
            if not self.queue or self.current_queue_index < 0:
                self.queue[0:0] = list(tracks)
                self.current_queue_index = 0
                if self.active_player and self.active_player.is_connected:
                    threading.Thread(target=self._trigger_current_play, daemon=True).start()
            else:
                insert_pos = min(len(self.queue), self.current_queue_index + 1)
                self.queue[insert_pos:insert_pos] = list(tracks)

    def move_queue_item_next(self, index: int) -> int:
        """Move an existing item in the queue so it plays next right after the current track."""
        with self._lock:
            if not (0 <= index < len(self.queue)):
                return index
            if self.current_queue_index < 0:
                track = self.queue.pop(index)
                self.queue.insert(0, track)
                self.current_queue_index = 0
                if self.active_player and self.active_player.is_connected:
                    threading.Thread(target=self._trigger_current_play, daemon=True).start()
                return 0
            if index == self.current_queue_index or index == self.current_queue_index + 1:
                return index
            track = self.queue.pop(index)
            if index < self.current_queue_index:
                self.current_queue_index -= 1
            insert_pos = min(len(self.queue), self.current_queue_index + 1)
            self.queue.insert(insert_pos, track)
            return insert_pos

    def _trigger_current_play(self):
        self._user_stopped = False
        with self._lock:
            if self.current_queue_index < 0 or self.current_queue_index >= len(self.queue):
                return
            track = self.queue[self.current_queue_index]
            player = self.active_player

        if player and player.is_connected and track:
            stream_url = self.streamer.get_stream_url(track.id)
            # Run in background to avoid blocking caller
            threading.Thread(
                target=player.play_track, args=(track, stream_url), daemon=True
            ).start()

    def next_track(self) -> bool:
        with self._lock:
            if self.current_queue_index + 1 < len(self.queue):
                self.current_queue_index += 1
            else:
                return False
        self._trigger_current_play()
        return True

    def prev_track(self) -> bool:
        with self._lock:
            if self.current_queue_index > 0:
                self.current_queue_index -= 1
            else:
                return False
        self._trigger_current_play()
        return True

    def toggle_play_pause(self) -> bool:
        if not self.active_player:
            return False
        status = self.active_player.get_status()
        if status.state == "PLAYING":
            return self.active_player.pause()
        elif status.state == "PAUSED":
            return self.active_player.resume()
        elif status.state in ["STOPPED", "IDLE"]:
            if self.queue and self.current_queue_index >= 0:
                self._trigger_current_play()
                return True
            elif status.track_title:
                return self.active_player.resume()
        return False

    def stop(self) -> bool:
        self._user_stopped = True
        if self.active_player:
            return self.active_player.stop()
        return False

    def seek_relative(self, delta_seconds: float) -> bool:
        if not self.active_player:
            return False
        status = self.active_player.get_status()
        new_pos = max(0.0, status.current_time + delta_seconds)
        return self.active_player.seek(new_pos)

    def volume_up(self, step: int = 5) -> bool:
        if self.active_player:
            return self.active_player.volume_up(step)
        return False

    def volume_down(self, step: int = 5) -> bool:
        if self.active_player:
            return self.active_player.volume_down(step)
        return False

    def set_volume(self, level: int) -> bool:
        if self.active_player:
            return self.active_player.set_volume(level)
        return False

    def get_status(self) -> PlayerStatus:
        player = self.active_player
        if player:
            if player.connection_state == "CONNECTING":
                return PlayerStatus(state="CONNECTING", device_name=player.name)
            if player.connection_state == "ERROR":
                return PlayerStatus(state="CONN_ERROR", device_name=player.name)
            st = player.get_status()
            # If player doesn't know track name yet, supplement from queue
            if not st.track_title and 0 <= self.current_queue_index < len(self.queue):
                cur_t = self.queue[self.current_queue_index]
                st.track_title = cur_t.title
                st.artist = cur_t.artist
                st.album = cur_t.album
                if st.duration <= 0:
                    st.duration = cur_t.duration_sec
            return st
        return PlayerStatus(state="DISCONNECTED", device_name="Disconnected")

    def get_queue(self) -> List[Track]:
        with self._lock:
            return list(self.queue)

    def clear_queue(self):
        with self._lock:
            self.queue.clear()
            self.current_queue_index = -1
        self.stop()

    def _monitor_loop(self):
        """Check for track completion and auto-advance queue."""
        last_state = "IDLE"
        while self._running:
            time.sleep(0.5)
            if not self.active_player:
                continue

            status = self.active_player.get_status()
            # If state went from PLAYING to IDLE or duration reached
            if last_state == "PLAYING" and status.state in ["IDLE", "STOPPED"]:
                # Song finished, auto-advance if not stopped by user and actually completed
                if getattr(self, "_user_stopped", False):
                    self._user_stopped = False
                elif status.duration > 0 and (
                    status.current_time >= status.duration - 2.0
                    or getattr(self.active_player, "_idle_reason", None) == "FINISHED"
                ):
                    if self.current_queue_index + 1 < len(self.queue):
                        self.next_track()

            last_state = status.state

    def shutdown(self):
        self._running = False
        if self.active_player:
            self.active_player.disconnect()
