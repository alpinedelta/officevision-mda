"""Base player interface and data classes."""

from dataclasses import dataclass
from typing import Optional


@dataclass
class PlayerStatus:
    state: str = "IDLE"  # "PLAYING", "PAUSED", "BUFFERING", "IDLE", "STOPPED", "ERROR"
    current_time: float = 0.0
    duration: float = 0.0
    volume: int = 50
    muted: bool = False
    track_title: str = ""
    artist: str = ""
    album: str = ""
    device_name: str = ""

    @property
    def progress_percent(self) -> float:
        if self.duration <= 0:
            return 0.0
        return min(100.0, max(0.0, (self.current_time / self.duration) * 100.0))

    def time_display(self) -> str:
        cur_min = int(self.current_time) // 60
        cur_sec = int(self.current_time) % 60
        dur_min = int(self.duration) // 60
        dur_sec = int(self.duration) % 60
        return f"{cur_min:02d}:{cur_sec:02d} / {dur_min:02d}:{dur_sec:02d}"

    def progress_bar(self, width: int = 30) -> str:
        """Render ASCII progress bar, e.g. [========>-------------------]"""
        if width < 5:
            width = 5
        bar_len = width - 2  # subtract brackets
        if self.duration <= 0:
            return "[" + "-" * bar_len + "]"
        filled = int((self.current_time / self.duration) * bar_len)
        filled = min(bar_len, max(0, filled))
        if filled == 0:
            bar = ">" + "-" * (bar_len - 1)
        elif filled >= bar_len:
            bar = "=" * bar_len
        else:
            bar = "=" * (filled - 1) + ">" + "-" * (bar_len - filled)
        return f"[{bar}]"


class BasePlayer:
    def __init__(self, device_info: dict):
        self.device_info = device_info
        self.device_id = device_info.get("id", "")
        self.name = device_info.get("name", "Unknown Device")
        self.host = device_info.get("host", "")
        self.port = device_info.get("port", 0)
        self.connection_state: str = "DISCONNECTED"  # "DISCONNECTED", "CONNECTING", "CONNECTED", "ERROR"
        self.last_error: str = ""

    @property
    def is_connected(self) -> bool:
        return self.connection_state == "CONNECTED"

    def connect(self) -> bool:
        raise NotImplementedError

    def disconnect(self):
        pass

    def play_track(self, track, stream_url: str) -> bool:
        raise NotImplementedError

    def pause(self) -> bool:
        raise NotImplementedError

    def resume(self) -> bool:
        raise NotImplementedError

    def stop(self) -> bool:
        raise NotImplementedError

    def seek(self, position_sec: float) -> bool:
        raise NotImplementedError

    def set_volume(self, level: int) -> bool:
        raise NotImplementedError

    def volume_up(self, step: int = 5) -> bool:
        status = self.get_status()
        new_vol = min(100, status.volume + step)
        return self.set_volume(new_vol)

    def volume_down(self, step: int = 5) -> bool:
        status = self.get_status()
        new_vol = max(0, status.volume - step)
        return self.set_volume(new_vol)

    def get_status(self) -> PlayerStatus:
        raise NotImplementedError
