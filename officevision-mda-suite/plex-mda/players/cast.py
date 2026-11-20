"""Chromecast Player implementation (Chromecast Audio, Displays, Streamers, TVs)."""

import threading
import time
from typing import Optional
from uuid import UUID

import pychromecast
from pychromecast.controllers.media import MediaStatusListener, METADATA_TYPE_MUSICTRACK

from players.base import BasePlayer, PlayerStatus


class CastStatusCallback(MediaStatusListener):
    def __init__(self, player):
        self.player = player

    def new_media_status(self, status):
        self.player._on_media_status(status)

    def load_media_failed(self, queue_item_id, error_code):
        self.player._on_load_media_failed(queue_item_id, error_code)


class CastPlayer(BasePlayer):
    def __init__(self, device_info: dict):
        super().__init__(device_info)
        self.cc: Optional[pychromecast.Chromecast] = None
        self._browser = None
        self._status = PlayerStatus(device_name=self.name)
        self._current_track = None
        self._last_update_time = time.time()
        self._play_requested_at = 0.0
        self._lock = threading.Lock()
        self._is_connecting = False
        self._connect_event = threading.Event()

    def connect(self) -> bool:
        with self._lock:
            if (
                self.connection_state == "CONNECTED"
                and self.cc
                and getattr(self.cc, "socket_client", None)
                and self.cc.socket_client.is_connected
            ):
                return True
            if self._is_connecting:
                should_connect = False
            else:
                self._is_connecting = True
                self.connection_state = "CONNECTING"
                self.last_error = ""
                self._connect_event.clear()
                should_connect = True

        if not should_connect:
            self._connect_event.wait(timeout=12)
            with self._lock:
                return self.connection_state == "CONNECTED"

        try:
            target = None
            # 1. Fast direct connection using cached host/port (no 5s mDNS wait)
            if self.host:
                try:
                    raw_uuid = self.device_info.get("uuid")
                    uuid_val = UUID(raw_uuid) if raw_uuid else UUID(int=0)
                    port_val = int(self.port or 8009)
                    cast_type = self.device_info.get("cast_type") or (
                        "group" if port_val != 8009 else "cast"
                    )
                    service = pychromecast.HostServiceInfo(self.host, port_val)
                    cast_info = pychromecast.CastInfo(
                        services={service},
                        uuid=uuid_val,
                        model_name=self.device_info.get("model") or self.device_info.get("alias") or "Chromecast",
                        friendly_name=self.name,
                        host=self.host,
                        port=port_val,
                        cast_type=cast_type,
                        manufacturer=self.device_info.get("manufacturer", ""),
                    )
                    direct_cc = pychromecast.Chromecast(
                        cast_info=cast_info,
                        timeout=3.5,
                        tries=1,
                    )
                    direct_cc.wait(timeout=4.5)
                    target = direct_cc
                except Exception:
                    target = None

            # 2. Fallback to mDNS discovery if direct IP connection failed (e.g. DHCP IP changed)
            if not target:
                casts, browser = pychromecast.get_chromecasts(
                    known_hosts=[self.host] if self.host else None,
                    timeout=4,
                )
                self._browser = browser

                for c in casts:
                    if self.host and c.cast_info.host == self.host:
                        target = c
                        break
                if not target:
                    for c in casts:
                        if self.name.lower() in c.name.lower() or c.name.lower() in self.name.lower():
                            target = c
                            break
                if target:
                    target.wait(timeout=8)

            if not target:
                with self._lock:
                    self.connection_state = "ERROR"
                    self.last_error = f"Device {self.name} ({self.host}) not found"
                return False

            self.cc = target
            if target.cast_info and target.cast_info.friendly_name:
                self.name = target.cast_info.friendly_name
            elif getattr(target, "name", None):
                self.name = target.name
            self._status.device_name = self.name
            if self.cc.status:
                self._status.volume = int((self.cc.status.volume_level or 0.5) * 100)
                self._status.muted = bool(self.cc.status.volume_muted)

            # Register media status listener
            self.cc.media_controller.register_status_listener(CastStatusCallback(self))

            # Launch Default Media Receiver session so the device wakes up and opens a session immediately
            try:
                if self.cc.app_id != pychromecast.APP_MEDIA_RECEIVER:
                    self.cc.start_app(pychromecast.APP_MEDIA_RECEIVER)
            except Exception:
                pass

            with self._lock:
                self.connection_state = "CONNECTED"
                self._status.state = "IDLE"
            return True
        except Exception as e:
            self.cc = None
            with self._lock:
                self.connection_state = "ERROR"
                self.last_error = str(e)
            return False
        finally:
            with self._lock:
                self._is_connecting = False
                self._connect_event.set()

    def _ensure_connected(self) -> bool:
        if (
            self.connection_state == "CONNECTED"
            and self.cc
            and getattr(self.cc, "socket_client", None)
            and self.cc.socket_client.is_connected
        ):
            return True
        return self.connect()

    def _on_load_media_failed(self, queue_item_id, error_code):
        with self._lock:
            self._status.state = "IDLE"
            self.last_error = f"Media load failed (code {error_code})"

    def _on_media_status(self, ms):
        with self._lock:
            state = ms.player_state or "IDLE"
            idle_reason = getattr(ms, "idle_reason", None)
            since_play = time.time() - self._play_requested_at

            if state != "UNKNOWN":
                # Ignore transient IDLE (or previous track INTERRUPTED/FINISHED) while newly requested track is loading
                if (
                    self._status.state == "BUFFERING"
                    and state == "IDLE"
                    and since_play < 5.0
                    and idle_reason != "ERROR"
                ):
                    pass
                else:
                    self._status.state = state

            if ms.current_time is not None and (state != "IDLE" or since_play >= 5.0):
                self._status.current_time = float(ms.current_time)
            if ms.duration is not None and ms.duration > 0:
                self._status.duration = float(ms.duration)
            self._last_update_time = time.time()
            self._idle_reason = idle_reason

            if ms.title:
                self._status.track_title = ms.title
            if ms.artist:
                self._status.artist = ms.artist
            if ms.album_name:
                self._status.album = ms.album_name

    def play_track(self, track, stream_url: str) -> bool:
        if not self._ensure_connected():
            return False

        try:
            with self._lock:
                self._current_track = track
                self._play_requested_at = time.time()
                self._status.state = "BUFFERING"
                self._status.track_title = track.title
                self._status.artist = track.artist
                self._status.album = track.album
                self._status.duration = track.duration_sec
                self._status.current_time = 0.0
                self._last_update_time = time.time()
                self._idle_reason = None
                if self.cc and self.cc.status:
                    self._status.volume = int((self.cc.status.volume_level or 0.5) * 100)
                    self._status.muted = bool(self.cc.status.volume_muted)

            mime = "audio/flac" if track.format == "FLAC" else "audio/mpeg"
            mc = self.cc.media_controller
            metadata = {
                "metadataType": METADATA_TYPE_MUSICTRACK,
                "title": track.title,
                "artist": track.artist,
                "albumName": track.album,
            }
            if track.track_num:
                metadata["trackNumber"] = track.track_num

            mc.play_media(
                url=stream_url,
                content_type=mime,
                title=track.title,
                stream_type="BUFFERED",
                metadata=metadata,
                autoplay=True,
            )
            mc.block_until_active(timeout=8)
            return True
        except Exception as e:
            with self._lock:
                self._status.state = "IDLE"
                self.last_error = str(e)
            return False

    def pause(self) -> bool:
        if not self.cc:
            return False
        try:
            self.cc.media_controller.pause()
            self._status.state = "PAUSED"
            return True
        except Exception:
            return False

    def resume(self) -> bool:
        if not self.cc:
            return False
        try:
            self.cc.media_controller.play()
            self._status.state = "PLAYING"
            self._last_update_time = time.time()
            return True
        except Exception:
            return False

    def stop(self) -> bool:
        if not self.cc:
            return False
        try:
            self.cc.media_controller.stop()
            self._status.state = "STOPPED"
            self._status.current_time = 0.0
            return True
        except Exception:
            return False

    def seek(self, position_sec: float) -> bool:
        if not self.cc:
            return False
        try:
            position_sec = max(0.0, min(self._status.duration, position_sec))
            self.cc.media_controller.seek(position_sec)
            self._status.current_time = position_sec
            self._last_update_time = time.time()
            return True
        except Exception:
            return False

    def set_volume(self, level: int) -> bool:
        if not self.cc:
            return False
        try:
            level = max(0, min(100, level))
            self.cc.set_volume(level / 100.0)
            self._status.volume = level
            return True
        except Exception:
            return False

    def get_status(self) -> PlayerStatus:
        with self._lock:
            # Estimate elapsed time if currently playing
            if self._status.state == "PLAYING" and self._status.duration > 0:
                elapsed = time.time() - self._last_update_time
                est_time = min(self._status.duration, self._status.current_time + elapsed)
                # Check for track end
                if est_time >= self._status.duration and self._status.duration > 0:
                    self._status.state = "IDLE"
                    self._status.current_time = self._status.duration
                else:
                    self._status.current_time = est_time
                self._last_update_time = time.time()

            if self.cc and self.cc.status:
                self._status.volume = int((self.cc.status.volume_level or 0.5) * 100)
                self._status.muted = bool(self.cc.status.volume_muted)

            return self._status

    def disconnect(self):
        with self._lock:
            self.connection_state = "DISCONNECTED"
            self._status.state = "IDLE"
            self._status.current_time = 0.0
            cc = self.cc
            self.cc = None
            browser = self._browser
            self._browser = None
        try:
            if cc:
                try:
                    if cc.app_id == pychromecast.APP_MEDIA_RECEIVER:
                        cc.quit_app()
                except Exception:
                    pass
                cc.disconnect()
            if browser:
                browser.stop_discovery()
        except Exception:
            pass
