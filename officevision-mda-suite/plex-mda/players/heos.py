"""HEOS Player implementation for Denon / Marantz HEOS receivers."""

import json
import socket
import threading
import time
from typing import Optional

from players.base import BasePlayer, PlayerStatus


class HeosPlayer(BasePlayer):
    def __init__(self, device_info: dict):
        super().__init__(device_info)
        self.pid = device_info.get("pid", 0)
        self.sock: Optional[socket.socket] = None
        self._status = PlayerStatus(device_name=self.name)
        self._current_track = None
        self._last_update_time = time.time()
        self._lock = threading.Lock()

    def connect(self) -> bool:
        with self._lock:
            self.connection_state = "CONNECTING"
            self.last_error = ""
            try:
                if self.sock:
                    try:
                        self.sock.close()
                    except Exception:
                        pass
                self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
                self.sock.settimeout(3.0)
                self.sock.connect((self.host, self.port))
                # Check player id
                self._send_cmd("player/get_players")
                resp = self._read_resp()
                if resp and "payload" in resp:
                    for p in resp["payload"]:
                        if p.get("ip") == self.host or self.name.lower() in p.get("name", "").lower():
                            self.pid = p.get("pid", self.pid)
                            self.name = p.get("name", self.name)
                            self._status.device_name = self.name
                            break
                self.connection_state = "CONNECTED"
                return True
            except Exception as e:
                self.sock = None
                self.connection_state = "ERROR"
                self.last_error = str(e)
                return False

    def _ensure_connected(self) -> bool:
        if self.sock:
            return True
        return self.connect()

    def _send_cmd(self, command: str) -> bool:
        if not self.sock:
            return False
        try:
            cmd_bytes = f"heos://{command}\r\n".encode("utf-8")
            self.sock.sendall(cmd_bytes)
            return True
        except Exception:
            self.sock = None
            return False

    def _read_resp(self) -> Optional[dict]:
        if not self.sock:
            return None
        try:
            data = self.sock.recv(4096)
            if not data:
                return None
            lines = data.decode("utf-8", errors="replace").split("\r\n")
            for line in lines:
                if line.strip():
                    try:
                        return json.loads(line)
                    except json.JSONDecodeError:
                        pass
        except Exception:
            pass
        return None

    def play_track(self, track, stream_url: str) -> bool:
        if not self._ensure_connected():
            return False

        with self._lock:
            try:
                # Tell HEOS to play stream
                cmd = f"browse/play_stream?pid={self.pid}&url={stream_url}"
                self._send_cmd(cmd)
                resp = self._read_resp()

                self._current_track = track
                self._status.state = "PLAYING"
                self._status.track_title = track.title
                self._status.artist = track.artist
                self._status.album = track.album
                self._status.duration = track.duration_sec
                self._status.current_time = 0.0
                self._last_update_time = time.time()
                return True
            except Exception:
                return False

    def pause(self) -> bool:
        if not self._ensure_connected():
            return False
        with self._lock:
            try:
                self._send_cmd(f"player/set_play_state?pid={self.pid}&state=pause")
                self._read_resp()
                self._status.state = "PAUSED"
                return True
            except Exception:
                return False

    def resume(self) -> bool:
        if not self._ensure_connected():
            return False
        with self._lock:
            try:
                self._send_cmd(f"player/set_play_state?pid={self.pid}&state=play")
                self._read_resp()
                self._status.state = "PLAYING"
                self._last_update_time = time.time()
                return True
            except Exception:
                return False

    def stop(self) -> bool:
        if not self._ensure_connected():
            return False
        with self._lock:
            try:
                self._send_cmd(f"player/set_play_state?pid={self.pid}&state=stop")
                self._read_resp()
                self._status.state = "STOPPED"
                self._status.current_time = 0.0
                return True
            except Exception:
                return False

    def seek(self, position_sec: float) -> bool:
        # HEOS direct streams don't always support arbitrary byte seek command over CLI,
        # but time estimation handles it
        with self._lock:
            self._status.current_time = max(0.0, min(self._status.duration, position_sec))
            self._last_update_time = time.time()
            return True

    def set_volume(self, level: int) -> bool:
        if not self._ensure_connected():
            return False
        with self._lock:
            try:
                level = max(0, min(100, level))
                self._send_cmd(f"player/set_volume?pid={self.pid}&level={level}")
                self._read_resp()
                self._status.volume = level
                return True
            except Exception:
                return False

    def get_status(self) -> PlayerStatus:
        with self._lock:
            if self._status.state == "PLAYING" and self._status.duration > 0:
                elapsed = time.time() - self._last_update_time
                est_time = min(self._status.duration, self._status.current_time + elapsed)
                if est_time >= self._status.duration and self._status.duration > 0:
                    self._status.state = "IDLE"
                    self._status.current_time = self._status.duration
                else:
                    self._status.current_time = est_time
                self._last_update_time = time.time()

            return self._status

    def disconnect(self):
        with self._lock:
            self.connection_state = "DISCONNECTED"
            self._status.state = "IDLE"
            self._status.current_time = 0.0
            if self.sock:
                try:
                    self.sock.close()
                except Exception:
                    pass
                self.sock = None
