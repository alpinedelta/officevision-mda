"""Local HTTP Streaming Server with byte-range (206) support for LAN devices."""

import mimetypes
import os
import re
import socket
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

from config import get_local_ip, get_stream_port
from db import MusicDB

RANGE_RE = re.compile(r"bytes=(\d+)-(\d*)")


class StreamRequestHandler(BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"

    def log_message(self, format, *args):
        # Suppress standard logging to prevent terminal corruption
        pass

    def do_HEAD(self):
        self._handle_request(send_body=False)

    def do_GET(self):
        self._handle_request(send_body=True)

    def _handle_request(self, send_body: bool = True):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == "/health":
            self.send_response(200)
            self.send_header("Content-Type", "text/plain")
            self.send_header("Content-Length", "2")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            if send_body:
                self.wfile.write(b"OK")
            return

        if path.startswith("/stream/"):
            try:
                track_id = int(path.split("/stream/")[1].split("?")[0])
            except ValueError:
                self.send_error(400, "Invalid track ID")
                return

            db = getattr(self.server, "db", None)
            if not db:
                db = MusicDB()
            track = db.get_track(track_id)
            if not track or not track.file_path or not os.path.exists(track.file_path):
                self.send_error(404, "Track or audio file not found")
                return

            self._serve_file(track.file_path, send_body=send_body)
            return

        self.send_error(404, "Not Found")

    def _serve_file(self, filepath: str, send_body: bool = True):
        try:
            file_size = os.path.getsize(filepath)
            content_type, _ = mimetypes.guess_type(filepath)
            if not content_type:
                ext = os.path.splitext(filepath)[1].lower()
                if ext == ".flac":
                    content_type = "audio/flac"
                elif ext == ".mp3":
                    content_type = "audio/mpeg"
                else:
                    content_type = "application/octet-stream"

            range_header = self.headers.get("Range")
            start = 0
            end = file_size - 1
            status_code = 200

            if range_header:
                match = RANGE_RE.match(range_header)
                if match:
                    status_code = 206
                    start_str, end_str = match.groups()
                    start = int(start_str)
                    if end_str:
                        end = int(end_str)
                    end = min(end, file_size - 1)

            length = end - start + 1

            self.send_response(status_code)
            self.send_header("Content-Type", content_type)
            self.send_header("Accept-Ranges", "bytes")
            self.send_header("Content-Length", str(length))
            self.send_header("Access-Control-Allow-Origin", "*")
            if status_code == 206:
                self.send_header(
                    "Content-Range", f"bytes {start}-{end}/{file_size}"
                )
            self.send_header("Connection", "keep-alive")
            self.end_headers()

            if not send_body:
                return

            # Stream requested byte chunk
            with open(filepath, "rb") as f:
                f.seek(start)
                remaining = length
                chunk_size = 64 * 1024
                while remaining > 0:
                    read_len = min(remaining, chunk_size)
                    buf = f.read(read_len)
                    if not buf:
                        break
                    self.wfile.write(buf)
                    remaining -= len(buf)

        except (BrokenPipeError, ConnectionResetError):
            # Normal occurrence when player seeks or skips track
            pass
        except Exception:
            pass


class AudioStreamer:
    def __init__(self, db: MusicDB, port: int = None):
        self.db = db
        self.port = port if port is not None else get_stream_port()
        self.host_ip = get_local_ip()
        self.server: ThreadingHTTPServer = None
        self._thread: threading.Thread = None
        self._start_server()

    def _start_server(self):
        curr_port = self.port
        for attempt in range(10):
            try:
                server = ThreadingHTTPServer(("0.0.0.0", curr_port), StreamRequestHandler)
                server.db = self.db
                self.server = server
                self.port = curr_port
                self._thread = threading.Thread(target=server.serve_forever, daemon=True)
                self._thread.start()
                return
            except OSError:
                curr_port += 1

        raise RuntimeError(f"Could not bind streaming server to port {self.port}-8899")

    def get_stream_url(self, track_id: int) -> str:
        return f"http://{self.host_ip}:{self.port}/stream/{track_id}"

    def stop(self):
        if self.server:
            self.server.shutdown()
            self.server.server_close()
