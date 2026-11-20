#!/usr/bin/env python3
"""
Google Drive & Google Docs Integration for IBM OfficeVision/400 (ov400).
Supports:
  1. Direct Google Drive REST API v3 (zero external dependencies using urllib)
     - Native Google Docs (application/vnd.google-apps.document) import/export
       using Markdown or Plain Text so formatting (bold, underline, headings)
       round-trips cleanly between OV/400 and Google Docs.
     - Standard text/markdown files on Google Drive.
     - OAuth 2.0 flow supporting both local browser loopback and copy-paste
       authorization URL for headless / 90x43 MDA Telnet sessions.
  2. Rclone backend support (if rclone is installed and has a Drive remote).
  3. Local Document Library & Offline Sync Cache in:
     ~/.local/share/officevision-mda/documents/
"""

import datetime
import http.server
import json
import os
import shutil
import socket
import subprocess
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import uuid
from typing import Any, Dict, List, Optional, Tuple


APP_DATA_DIR = os.environ.get("OV400_DATA_DIR") or os.path.expanduser(
    "~/.local/share/officevision-mda"
)
DOCS_DIR = os.path.join(APP_DATA_DIR, "documents")
CONFIG_PATH = os.path.join(APP_DATA_DIR, "gdrive_config.json")
TOKEN_PATH = os.path.join(APP_DATA_DIR, "gdrive_token.json")
INDEX_PATH = os.path.join(APP_DATA_DIR, "doc_index.json")

GOOGLE_AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_URL = "https://oauth2.googleapis.com/token"
DRIVE_API_BASE = "https://www.googleapis.com/drive/v3"
DRIVE_UPLOAD_BASE = "https://www.googleapis.com/upload/drive/v3"
CALENDAR_SCOPE = "https://www.googleapis.com/auth/calendar"
DRIVE_SCOPE = (
    "https://www.googleapis.com/auth/drive "
    "https://www.googleapis.com/auth/spreadsheets "
    "https://www.googleapis.com/auth/calendar"
)

GDOC_MIME = "application/vnd.google-apps.document"
GSHEET_MIME = "application/vnd.google-apps.spreadsheet"
GFOLDER_MIME = "application/vnd.google-apps.folder"


class DocumentEntry:
    """Represents an OfficeVision/400 Document or Folder (local + Google Drive)."""

    def __init__(
        self,
        doc_id: str,
        name: str,
        folder: str = "/My Drive",
        description: str = "",
        is_folder: bool = False,
        gdrive_id: Optional[str] = None,
        mime_type: str = GDOC_MIME,
        local_path: str = "",
        updated_at: str = "",
        synced: bool = False,
        dirty: bool = False,
        size: int = 0,
    ):
        self.doc_id = doc_id
        self.name = name
        self.folder = folder
        self.description = description
        self.is_folder = is_folder
        self.gdrive_id = gdrive_id
        self.mime_type = mime_type
        self.local_path = local_path
        self.updated_at = updated_at or datetime.datetime.now().strftime("%m/%d/%y %H:%M")
        self.synced = synced
        self.dirty = dirty
        self.size = size

    def to_dict(self) -> Dict[str, Any]:
        return {
            "doc_id": self.doc_id,
            "name": self.name,
            "folder": self.folder,
            "description": self.description,
            "is_folder": self.is_folder,
            "gdrive_id": self.gdrive_id,
            "mime_type": self.mime_type,
            "local_path": self.local_path,
            "updated_at": self.updated_at,
            "synced": self.synced,
            "dirty": self.dirty,
            "size": self.size,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "DocumentEntry":
        return cls(
            doc_id=data.get("doc_id", str(uuid.uuid4())[:8].upper()),
            name=data.get("name", "UNTITLED"),
            folder=data.get("folder", "/My Drive"),
            description=data.get("description", ""),
            is_folder=bool(data.get("is_folder", False)),
            gdrive_id=data.get("gdrive_id"),
            mime_type=data.get("mime_type", GDOC_MIME),
            local_path=data.get("local_path", ""),
            updated_at=data.get("updated_at", ""),
            synced=bool(data.get("synced", False)),
            dirty=bool(data.get("dirty", False)),
            size=int(data.get("size", 0)),
        )


class GoogleDriveManager:
    """Manages local OV/400 documents and synchronizes with Google Drive / Docs."""

    def __init__(self):
        os.makedirs(DOCS_DIR, exist_ok=True)
        self.config: Dict[str, Any] = self._load_json(
            CONFIG_PATH,
            {
                "backend": "auto",  # "api", "rclone", "local", or "auto"
                "client_id": "",
                "client_secret": "",
                "rclone_remote": "gdrive:",
                "root_folder_id": "root",
                "root_folder_name": "/My Drive",
                "export_format": "markdown",  # "markdown" or "txt"
                "auto_sync_on_save": True,
                "page_width": 72,
            },
        )
        self.token: Dict[str, Any] = self._load_json(TOKEN_PATH, {})
        self.entries: List[DocumentEntry] = []
        self._load_index()
        self._ensure_welcome_doc()

    def _load_json(self, path: str, default: Any) -> Any:
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                return default
        return default

    def _save_json(self, path: str, data: Any):
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def save_config(self):
        self._save_json(CONFIG_PATH, self.config)

    def _load_index(self):
        raw = self._load_json(INDEX_PATH, [])
        self.entries = [DocumentEntry.from_dict(item) for item in raw]
        # Verify local files still exist and scan for any unindexed files in DOCS_DIR
        indexed_paths = {e.local_path for e in self.entries if e.local_path}
        for fname in sorted(os.listdir(DOCS_DIR)):
            if fname.startswith("."):
                continue
            fpath = os.path.join(DOCS_DIR, fname)
            if os.path.isfile(fpath) and fpath not in indexed_paths:
                stat = os.stat(fpath)
                dt = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%m/%d/%y %H:%M")
                base_name = os.path.splitext(fname)[0]
                self.entries.append(
                    DocumentEntry(
                        doc_id=str(uuid.uuid4())[:8].upper(),
                        name=base_name,
                        folder="/My Drive",
                        description="Local OV/400 Document",
                        local_path=fpath,
                        updated_at=dt,
                        synced=False,
                        dirty=True,
                        size=stat.st_size,
                    )
                )
        self.save_index()

    def save_index(self):
        self._save_json(INDEX_PATH, [e.to_dict() for e in self.entries])

    def _ensure_welcome_doc(self):
        if any(not e.is_folder for e in self.entries):
            return
        welcome_path = os.path.join(DOCS_DIR, "WELCOME_OV400.md")
        welcome_text = (
            "# IBM OfficeVision/400 Word Processing\n"
            "\n"
            "Welcome to the **OfficeVision/400 Document & Word Processing System** "
            "designed for the **90x43 Monochrome MDA Terminal**.\n"
            "\n"
            "## Key Features\n"
            "\n"
            "1. **DisplayWrite/400 Style Editor**: Includes a live 5250 margin ruler, "
            "automatic soft word-wrapping at column 72, paragraph reflow (`F8`), "
            "line centering (`F6`), and multi-level Undo/Redo (`Ctrl+Z` / `Ctrl+Y`).\n"
            "2. **Monochrome MDA Formatting**: Use **bold** (`Ctrl+B` or `**text**`), "
            "__underline__ (`Ctrl+U` or `__text__`), and ==reverse highlight== "
            "(`Ctrl+R` or `==text==`). Toggle between live formatted MDA view and "
            "raw markup with `F11`.\n"
            "3. **Integrated Spell Checker**: Press `F4` at any time to run `aspell` "
            "across your document with interactive replacement suggestions.\n"
            "4. **Google Drive & Google Docs Sync**: Press `F2` to save locally and "
            "push directly to **Google Docs** on your Google Drive. Use `F10` in the "
            "Document Manager (Option 6) to configure your Google Drive connection.\n"
            "\n"
            "---\n"
            "\n"
            "### Quick Keyboard Reference\n"
            "\n"
            "* `F1`         : Display Help Screen\n"
            "* `F2` / `^S`  : Save Document & Sync to Google Drive\n"
            "* `F3`         : Save & Exit to Document List / Menu\n"
            "* `F4`         : Spell Check Document (`aspell`)\n"
            "* `F5` / `^F`  : Find & Replace Text\n"
            "* `F6` / `^E`  : Center Current Line between Margins\n"
            "* `F8` / `^J`  : Reflow / Justify Current Paragraph\n"
            "* `F11`        : Toggle MDA Formatted View vs Raw Markup\n"
            "* `F12`        : Cancel / Exit Editor\n"
        )
        try:
            with open(welcome_path, "w", encoding="utf-8") as f:
                f.write(welcome_text)
            self.entries.append(
                DocumentEntry(
                    doc_id="OV400DOC",
                    name="WELCOME_OV400",
                    folder="/My Drive",
                    description="OfficeVision/400 Word Processor Guide",
                    local_path=welcome_path,
                    updated_at=datetime.datetime.now().strftime("%m/%d/%y %H:%M"),
                    synced=False,
                    dirty=False,
                    size=len(welcome_text.encode("utf-8")),
                )
            )
            self.save_index()
        except Exception:
            pass

    # =========================================================================
    # Connection & Backend Status
    # =========================================================================
    def get_active_backend(self) -> str:
        """Returns 'api', 'rclone', or 'local'."""
        pref = self.config.get("backend", "auto")
        if pref == "local":
            return "local"
        if pref == "api":
            return "api" if self.is_api_authenticated() else "local"
        if pref == "rclone":
            return "rclone" if self.is_rclone_available() else "local"

        # Auto-detect
        if self.is_api_authenticated():
            return "api"
        if self.is_rclone_available():
            return "rclone"
        return "local"

    def is_api_authenticated(self) -> bool:
        return bool(self.token.get("access_token") or self.token.get("refresh_token"))

    def has_calendar_scope(self) -> bool:
        if not self.is_api_authenticated():
            return False
        scope_str = str(self.token.get("scope", ""))
        if not scope_str:
            # Older tokens without recorded scope field: allow API call to verify
            return True
        return "calendar" in scope_str.lower()

    def is_rclone_available(self) -> bool:
        rclone_bin = shutil.which("rclone")
        if not rclone_bin:
            return False
        remote = self.config.get("rclone_remote", "gdrive:")
        if not remote.endswith(":"):
            remote += ":"
        try:
            res = subprocess.run(
                [rclone_bin, "listremotes"],
                capture_output=True,
                text=True,
                timeout=3,
            )
            remotes = [r.strip() for r in res.stdout.splitlines() if r.strip()]
            return remote in remotes
        except Exception:
            return False

    def get_status_summary(self) -> str:
        backend = self.get_active_backend()
        if backend == "api":
            return "Connected (Google API)"
        if backend == "rclone":
            return f"Connected (rclone {self.config.get('rclone_remote', 'gdrive:')})"
        return "Local Mode (F10=GDrive Setup)"

    # =========================================================================
    # OAuth 2.0 Flow for Direct Google Drive API
    # =========================================================================
    def load_client_secrets_file(self, filepath: str) -> Tuple[bool, str]:
        """Loads client_id and client_secret from a downloaded Google credentials.json."""
        expanded = os.path.expanduser(filepath.strip())
        if not os.path.exists(expanded):
            return False, f"File not found: {expanded}"
        try:
            with open(expanded, "r", encoding="utf-8") as f:
                data = json.load(f)
            block = data.get("installed") or data.get("web") or data
            cid = block.get("client_id", "").strip()
            csec = block.get("client_secret", "").strip()
            if not cid or not csec:
                return False, "JSON file does not contain client_id and client_secret."
            self.config["client_id"] = cid
            self.config["client_secret"] = csec
            self.save_config()
            return True, "Loaded Google OAuth Client ID & Secret from file."
        except Exception as exc:
            return False, f"Failed to parse JSON: {exc}"

    def build_auth_url(self, redirect_uri: str = "http://localhost:8765") -> str:
        params = {
            "client_id": self.config.get("client_id", ""),
            "redirect_uri": redirect_uri,
            "response_type": "code",
            "scope": DRIVE_SCOPE,
            "access_type": "offline",
            "prompt": "consent",
        }
        return f"{GOOGLE_AUTH_URL}?{urllib.parse.urlencode(params)}"

    def exchange_code_for_token(
        self, code: str, redirect_uri: str = "http://localhost:8765"
    ) -> Tuple[bool, str]:
        cid = self.config.get("client_id", "").strip()
        csec = self.config.get("client_secret", "").strip()
        if not cid or not csec:
            return False, "Client ID and Client Secret are required first."

        payload = urllib.parse.urlencode(
            {
                "code": code.strip(),
                "client_id": cid,
                "client_secret": csec,
                "redirect_uri": redirect_uri,
                "grant_type": "authorization_code",
            }
        ).encode("utf-8")

        req = urllib.request.Request(
            GOOGLE_TOKEN_URL,
            data=payload,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                token_data = json.loads(resp.read().decode("utf-8"))
            token_data["obtained_at"] = int(time.time())
            self.token = token_data
            self._save_json(TOKEN_PATH, self.token)
            self.config["backend"] = "api"
            self.save_config()
            return True, "Successfully authenticated with Google Drive API!"
        except urllib.error.HTTPError as e:
            err_body = e.read().decode("utf-8", errors="replace")
            return False, f"OAuth error ({e.code}): {err_body[:120]}"
        except Exception as exc:
            return False, f"Network/OAuth error: {exc}"

    def run_local_oauth_listener(self, port: int = 8765, timeout_sec: int = 120) -> Tuple[bool, str]:
        """
        Starts a temporary localhost HTTP listener on 127.0.0.1:port to capture
        the OAuth redirect code automatically if a browser is available.
        """
        captured: Dict[str, Optional[str]] = {"code": None, "error": None}
        redirect_uri = f"http://localhost:{port}"

        class OAuthHandler(http.server.BaseHTTPRequestHandler):
            def do_GET(self):
                parsed = urllib.parse.urlparse(self.path)
                qs = urllib.parse.parse_qs(parsed.query)
                if "code" in qs:
                    captured["code"] = qs["code"][0]
                    body = (
                        b"<html><body style='font-family:monospace;background:#000;color:#0f0;padding:2em;'>"
                        b"<h2>IBM OfficeVision/400 - Google Account Connected</h2>"
                        b"<p>Drive, Sheets &amp; Calendar authentication complete. You may return to your 90x43 MDA terminal.</p>"
                        b"</body></html>"
                    )
                else:
                    captured["error"] = qs.get("error", ["Unknown error"])[0]
                    body = b"<html><body><h3>Authentication Failed or Cancelled.</h3></body></html>"
                self.send_response(200)
                self.send_header("Content-Type", "text/html; charset=utf-8")
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

            def log_message(self, format, *args):
                pass

        try:
            server = http.server.HTTPServer(("127.0.0.1", port), OAuthHandler)
            server.timeout = timeout_sec
        except Exception as exc:
            return False, f"Could not bind localhost:{port}: {exc}"

        auth_url = self.build_auth_url(redirect_uri=redirect_uri)
        try:
            # Try launching xdg-open if DISPLAY is available
            if os.environ.get("DISPLAY"):
                subprocess.Popen(
                    ["xdg-open", auth_url],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
        except Exception:
            pass

        server.handle_request()
        server.server_close()

        if captured["code"]:
            return self.exchange_code_for_token(captured["code"], redirect_uri=redirect_uri)
        return False, f"OAuth cancelled or timed out ({captured['error'] or 'timeout'})."

    def get_access_token(self) -> Optional[str]:
        """Public accessor for a valid OAuth 2.0 access token (auto-refreshes)."""
        return self._get_valid_access_token()

    def _get_valid_access_token(self) -> Optional[str]:
        if not self.token:
            return None
        now = int(time.time())
        obtained = int(self.token.get("obtained_at", 0))
        expires_in = int(self.token.get("expires_in", 3600))
        if self.token.get("access_token") and (now - obtained) < (expires_in - 120):
            return self.token["access_token"]

        refresh_tok = self.token.get("refresh_token")
        cid = self.config.get("client_id", "").strip()
        csec = self.config.get("client_secret", "").strip()
        if not (refresh_tok and cid and csec):
            return self.token.get("access_token")

        payload = urllib.parse.urlencode(
            {
                "client_id": cid,
                "client_secret": csec,
                "refresh_token": refresh_tok,
                "grant_type": "refresh_token",
            }
        ).encode("utf-8")
        req = urllib.request.Request(
            GOOGLE_TOKEN_URL,
            data=payload,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=10) as resp:
                new_data = json.loads(resp.read().decode("utf-8"))
            self.token["access_token"] = new_data["access_token"]
            self.token["expires_in"] = new_data.get("expires_in", 3600)
            self.token["obtained_at"] = int(time.time())
            self._save_json(TOKEN_PATH, self.token)
            return self.token["access_token"]
        except Exception:
            return self.token.get("access_token")

    # =========================================================================
    # Document CRUD & Cloud Sync Operations
    # =========================================================================
    def list_documents(self, folder: Optional[str] = None) -> List[DocumentEntry]:
        items = list(self.entries)
        if folder:
            items = [e for e in items if e.folder == folder]
        # Sort folders first, then alphabetically by name
        items.sort(key=lambda e: (0 if e.is_folder else 1, e.name.upper()))
        return items

    def get_folders(self) -> List[str]:
        folders = {"/My Drive"}
        for e in self.entries:
            if e.folder:
                folders.add(e.folder)
            if e.is_folder:
                full = f"{e.folder.rstrip('/')}/{e.name}"
                folders.add(full)
        return sorted(folders)

    def create_folder(self, folder_name: str, parent_folder: str = "/My Drive") -> Tuple[bool, str]:
        clean = folder_name.strip().replace("/", "_")
        if not clean:
            return False, "Folder name cannot be empty."
        full_path = f"{parent_folder.rstrip('/')}/{clean}"
        for e in self.entries:
            if e.is_folder and e.name.upper() == clean.upper() and e.folder == parent_folder:
                return False, f"Folder '{clean}' already exists."

        entry = DocumentEntry(
            doc_id=str(uuid.uuid4())[:8].upper(),
            name=clean,
            folder=parent_folder,
            description="Folder",
            is_folder=True,
            mime_type=GFOLDER_MIME,
            updated_at=datetime.datetime.now().strftime("%m/%d/%y %H:%M"),
            synced=False,
        )
        self.entries.append(entry)
        self.save_index()
        return True, f"Created folder '{full_path}'."

    def create_document(
        self,
        name: str,
        description: str = "",
        folder: str = "/My Drive",
        initial_text: str = "",
    ) -> DocumentEntry:
        clean_name = name.strip()
        if not clean_name:
            clean_name = f"DOC_{datetime.datetime.now().strftime('%H%M%S')}"
        safe_fname = "".join(c if c.isalnum() or c in ("-", "_", ".") else "_" for c in clean_name)
        if not os.path.splitext(safe_fname)[1]:
            safe_fname += ".md"
        local_path = os.path.join(DOCS_DIR, safe_fname)
        counter = 1
        while os.path.exists(local_path):
            stem, ext = os.path.splitext(safe_fname)
            local_path = os.path.join(DOCS_DIR, f"{stem}_{counter}{ext}")
            counter += 1

        with open(local_path, "w", encoding="utf-8") as f:
            f.write(initial_text)

        entry = DocumentEntry(
            doc_id=str(uuid.uuid4())[:8].upper(),
            name=clean_name,
            folder=folder,
            description=description or "OfficeVision/400 Document",
            is_folder=False,
            mime_type=GDOC_MIME,
            local_path=local_path,
            updated_at=datetime.datetime.now().strftime("%m/%d/%y %H:%M"),
            synced=False,
            dirty=True,
            size=len(initial_text.encode("utf-8")),
        )
        self.entries.append(entry)
        self.save_index()
        return entry

    def read_document_text(self, entry: DocumentEntry) -> str:
        if entry.local_path and os.path.exists(entry.local_path):
            try:
                with open(entry.local_path, "r", encoding="utf-8", errors="replace") as f:
                    return f.read()
            except Exception:
                return ""
        # If not cached locally yet, pull from Google Drive
        ok, msg = self.pull_document(entry)
        if ok and entry.local_path and os.path.exists(entry.local_path):
            try:
                with open(entry.local_path, "r", encoding="utf-8", errors="replace") as f:
                    return f.read()
            except Exception:
                return ""
        return ""

    def save_document_text(
        self, entry: DocumentEntry, text: str, sync_cloud: bool = True
    ) -> Tuple[bool, str]:
        if not entry.local_path:
            safe_fname = "".join(
                c if c.isalnum() or c in ("-", "_", ".") else "_" for c in entry.name
            )
            if not os.path.splitext(safe_fname)[1]:
                safe_fname += ".md"
            entry.local_path = os.path.join(DOCS_DIR, safe_fname)

        try:
            with open(entry.local_path, "w", encoding="utf-8") as f:
                f.write(text)
            entry.size = len(text.encode("utf-8"))
            entry.updated_at = datetime.datetime.now().strftime("%m/%d/%y %H:%M")
            entry.dirty = True
            self.save_index()
        except Exception as exc:
            return False, f"Local save error: {exc}"

        if sync_cloud and self.get_active_backend() != "local":
            ok, cloud_msg = self.push_document(entry)
            if ok:
                return True, f"Saved locally & synced to Google Drive ({entry.name})."
            return True, f"Saved locally (Cloud sync warning: {cloud_msg})"

        return True, f"Document '{entry.name}' saved locally ({entry.size} bytes)."

    def copy_document(self, entry: DocumentEntry, new_name: str) -> Tuple[bool, str]:
        if entry.is_folder:
            return False, "Cannot copy folder directly."
        text = self.read_document_text(entry)
        new_entry = self.create_document(
            name=new_name,
            description=f"Copy of {entry.name}",
            folder=entry.folder,
            initial_text=text,
        )
        return True, f"Copied '{entry.name}' to '{new_entry.name}'."

    def rename_document(self, entry: DocumentEntry, new_name: str, new_desc: str = "") -> Tuple[bool, str]:
        clean = new_name.strip()
        if not clean:
            return False, "New document name cannot be empty."
        old_name = entry.name
        entry.name = clean
        if new_desc.strip():
            entry.description = new_desc.strip()
        entry.updated_at = datetime.datetime.now().strftime("%m/%d/%y %H:%M")
        entry.dirty = True
        self.save_index()
        return True, f"Renamed '{old_name}' to '{clean}'."

    def delete_document(self, entry: DocumentEntry, delete_remote: bool = False) -> Tuple[bool, str]:
        if entry.local_path and os.path.exists(entry.local_path):
            try:
                os.remove(entry.local_path)
            except Exception:
                pass

        if delete_remote and entry.gdrive_id and self.get_active_backend() == "api":
            token = self._get_valid_access_token()
            if token:
                req = urllib.request.Request(
                    f"{DRIVE_API_BASE}/files/{entry.gdrive_id}",
                    headers={"Authorization": f"Bearer {token}"},
                    method="DELETE",
                )
                try:
                    urllib.request.urlopen(req, timeout=10)
                except Exception:
                    pass

        self.entries = [e for e in self.entries if e.doc_id != entry.doc_id]
        self.save_index()
        return True, f"Deleted '{entry.name}'."

    # =========================================================================
    # Google Drive Push / Pull / Full Directory Sync
    # =========================================================================
    def sync_all_from_gdrive(self) -> Tuple[bool, str]:
        """Fetches file list from Google Drive (via REST API or rclone) and updates index."""
        backend = self.get_active_backend()
        if backend == "api":
            return self._sync_list_via_api()
        if backend == "rclone":
            return self._sync_list_via_rclone()
        return False, "Google Drive not configured yet. Press F10 to set up Google Drive."

    def _sync_list_via_api(self) -> Tuple[bool, str]:
        token = self._get_valid_access_token()
        if not token:
            return False, "Google Drive OAuth token expired or missing. Press F10."

        # Query Google Docs, Markdown, and Plain Text files in Drive
        q = (
            "trashed = false and ("
            f"mimeType = '{GDOC_MIME}' or "
            "mimeType = 'text/plain' or "
            "mimeType = 'text/markdown'"
            ")"
        )
        params = urllib.parse.urlencode(
            {
                "q": q,
                "pageSize": 100,
                "fields": "files(id, name, mimeType, modifiedTime, size, description)",
                "orderBy": "modifiedTime desc",
            }
        )
        req = urllib.request.Request(
            f"{DRIVE_API_BASE}/files?{params}",
            headers={"Authorization": f"Bearer {token}"},
        )
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read().decode("utf-8"))
        except Exception as exc:
            return False, f"Google Drive API error: {exc}"

        files = data.get("files", [])
        by_gdrive_id = {e.gdrive_id: e for e in self.entries if e.gdrive_id}
        added, updated = 0, 0

        for f in files:
            gid = f["id"]
            name = f.get("name", "Untitled")
            mime = f.get("mimeType", GDOC_MIME)
            mtime = f.get("modifiedTime", "")
            disp_time = mtime[:16].replace("T", " ") if mtime else ""
            size = int(f.get("size", 0))
            desc = f.get("description") or ("Google Doc" if mime == GDOC_MIME else "Drive Text File")

            if gid in by_gdrive_id:
                existing = by_gdrive_id[gid]
                existing.name = name
                existing.mime_type = mime
                if disp_time:
                    existing.updated_at = disp_time
                existing.synced = True
                updated += 1
            else:
                safe_fname = "".join(
                    c if c.isalnum() or c in ("-", "_", ".") else "_" for c in name
                )
                if not os.path.splitext(safe_fname)[1]:
                    safe_fname += ".md"
                local_path = os.path.join(DOCS_DIR, f"gd_{gid[:6]}_{safe_fname}")
                self.entries.append(
                    DocumentEntry(
                        doc_id=gid[:8].upper(),
                        name=name,
                        folder="/My Drive",
                        description=desc,
                        is_folder=False,
                        gdrive_id=gid,
                        mime_type=mime,
                        local_path=local_path,
                        updated_at=disp_time,
                        synced=True,
                        dirty=False,
                        size=size,
                    )
                )
                added += 1

        self.save_index()
        return True, f"Google Drive synced: {added} new, {updated} updated documents."

    def _sync_list_via_rclone(self) -> Tuple[bool, str]:
        rclone_bin = shutil.which("rclone")
        if not rclone_bin:
            return False, "rclone binary not found."
        remote = self.config.get("rclone_remote", "gdrive:")
        if not remote.endswith(":"):
            remote += ":"

        try:
            res = subprocess.run(
                [
                    rclone_bin,
                    "lsjson",
                    remote,
                    "--max-depth",
                    "1",
                    "--drive-export-formats",
                    "md,txt",
                ],
                capture_output=True,
                text=True,
                timeout=20,
            )
            if res.returncode != 0:
                return False, f"rclone error: {res.stderr.strip()[:80]}"
            items = json.loads(res.stdout)
        except Exception as exc:
            return False, f"rclone sync failed: {exc}"

        by_name = {e.name: e for e in self.entries}
        added, updated = 0, 0
        for item in items:
            if item.get("IsDir"):
                continue
            name = item.get("Name", "")
            ext = os.path.splitext(name)[1].lower()
            if ext not in (".md", ".txt", ".doc", ".docx", ".odt", ""):
                continue
            stem = os.path.splitext(name)[0] or name
            gid = item.get("ID", stem[:8].upper())
            mtime = item.get("ModTime", "")[:16].replace("T", " ")
            size = int(item.get("Size", 0))
            if size < 0:
                size = 0

            if stem in by_name:
                ent = by_name[stem]
                ent.gdrive_id = gid
                ent.synced = True
                if mtime:
                    ent.updated_at = mtime
                updated += 1
            else:
                local_path = os.path.join(DOCS_DIR, f"{stem}.md")
                self.entries.append(
                    DocumentEntry(
                        doc_id=gid[:8].upper(),
                        name=stem,
                        folder="/My Drive",
                        description=f"Google Drive ({name})",
                        is_folder=False,
                        gdrive_id=name,  # store remote filename for rclone
                        mime_type=item.get("MimeType", GDOC_MIME),
                        local_path=local_path,
                        updated_at=mtime,
                        synced=True,
                        dirty=False,
                        size=size,
                    )
                )
                added += 1

        self.save_index()
        return True, f"rclone synced from {remote}: {added} new, {updated} updated."

    def pull_document(self, entry: DocumentEntry) -> Tuple[bool, str]:
        """Downloads/exports the latest content of a document from Google Drive."""
        backend = self.get_active_backend()
        if backend == "api" and entry.gdrive_id:
            token = self._get_valid_access_token()
            if not token:
                return False, "Missing Google OAuth token."
            if entry.mime_type == GDOC_MIME:
                # Export native Google Doc as Markdown (or fallback to text/plain)
                export_mime = (
                    "text/markdown"
                    if self.config.get("export_format") == "markdown"
                    else "text/plain"
                )
                url = (
                    f"{DRIVE_API_BASE}/files/{entry.gdrive_id}/export?"
                    f"{urllib.parse.urlencode({'mimeType': export_mime})}"
                )
            else:
                url = f"{DRIVE_API_BASE}/files/{entry.gdrive_id}?alt=media"

            req = urllib.request.Request(url, headers={"Authorization": f"Bearer {token}"})
            try:
                with urllib.request.urlopen(req, timeout=15) as resp:
                    content = resp.read().decode("utf-8", errors="replace")
            except urllib.error.HTTPError:
                # Fallback to text/plain if text/markdown export not supported for this doc
                url_txt = (
                    f"{DRIVE_API_BASE}/files/{entry.gdrive_id}/export?"
                    f"{urllib.parse.urlencode({'mimeType': 'text/plain'})}"
                )
                req2 = urllib.request.Request(
                    url_txt, headers={"Authorization": f"Bearer {token}"}
                )
                try:
                    with urllib.request.urlopen(req2, timeout=15) as resp2:
                        content = resp2.read().decode("utf-8", errors="replace")
                except Exception as exc:
                    return False, f"Failed to export Google Doc: {exc}"
            except Exception as exc:
                return False, f"Download failed: {exc}"

            if not entry.local_path:
                entry.local_path = os.path.join(DOCS_DIR, f"gd_{entry.doc_id}.md")
            with open(entry.local_path, "w", encoding="utf-8") as f:
                f.write(content)
            entry.size = len(content.encode("utf-8"))
            entry.synced = True
            entry.dirty = False
            self.save_index()
            return True, f"Pulled '{entry.name}' from Google Docs."

        if backend == "rclone":
            rclone_bin = shutil.which("rclone")
            remote = self.config.get("rclone_remote", "gdrive:")
            if not remote.endswith(":"):
                remote += ":"
            remote_name = entry.gdrive_id or f"{entry.name}.md"
            if not entry.local_path:
                entry.local_path = os.path.join(DOCS_DIR, f"{entry.name}.md")
            try:
                res = subprocess.run(
                    [
                        rclone_bin,
                        "copyto",
                        f"{remote}{remote_name}",
                        entry.local_path,
                        "--drive-export-formats",
                        "md,txt",
                    ],
                    capture_output=True,
                    text=True,
                    timeout=20,
                )
                if res.returncode == 0 and os.path.exists(entry.local_path):
                    entry.size = os.path.getsize(entry.local_path)
                    entry.synced = True
                    entry.dirty = False
                    self.save_index()
                    return True, f"Pulled '{entry.name}' from {remote}."
                return False, f"rclone pull error: {res.stderr.strip()[:80]}"
            except Exception as exc:
                return False, f"rclone error: {exc}"

        return False, "No active Google Drive connection."

    def push_document(self, entry: DocumentEntry) -> Tuple[bool, str]:
        """Uploads/converts the local document to Google Docs / Google Drive."""
        if not entry.local_path or not os.path.exists(entry.local_path):
            return False, "Local file missing."

        with open(entry.local_path, "r", encoding="utf-8", errors="replace") as f:
            content_bytes = f.read().encode("utf-8")

        backend = self.get_active_backend()
        if backend == "api":
            token = self._get_valid_access_token()
            if not token:
                return False, "Missing Google OAuth token."

            boundary = f"ov400_boundary_{uuid.uuid4().hex}"
            metadata: Dict[str, Any] = {
                "name": entry.name,
                "mimeType": GDOC_MIME,
                "description": entry.description or "Created by IBM OfficeVision/400",
            }
            content_mime = (
                "text/markdown"
                if self.config.get("export_format") == "markdown"
                else "text/plain"
            )

            body_parts = [
                f"--{boundary}\r\n"
                "Content-Type: application/json; charset=UTF-8\r\n\r\n"
                f"{json.dumps(metadata)}\r\n"
                f"--{boundary}\r\n"
                f"Content-Type: {content_mime}; charset=UTF-8\r\n\r\n"
            ]
            body = body_parts[0].encode("utf-8") + content_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")

            if entry.gdrive_id:
                url = f"{DRIVE_UPLOAD_BASE}/files/{entry.gdrive_id}?uploadType=multipart"
                method = "PATCH"
            else:
                url = f"{DRIVE_UPLOAD_BASE}/files?uploadType=multipart"
                method = "POST"

            req = urllib.request.Request(
                url,
                data=body,
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": f"multipart/related; boundary={boundary}",
                },
                method=method,
            )
            try:
                with urllib.request.urlopen(req, timeout=20) as resp:
                    resp_data = json.loads(resp.read().decode("utf-8"))
                entry.gdrive_id = resp_data.get("id", entry.gdrive_id)
                entry.synced = True
                entry.dirty = False
                self.save_index()
                return True, f"Synced '{entry.name}' to Google Docs."
            except Exception as exc:
                return False, f"Drive upload failed: {exc}"

        if backend == "rclone":
            rclone_bin = shutil.which("rclone")
            remote = self.config.get("rclone_remote", "gdrive:")
            if not remote.endswith(":"):
                remote += ":"
            remote_target = f"{remote}{entry.name}.md"
            try:
                res = subprocess.run(
                    [
                        rclone_bin,
                        "copyto",
                        entry.local_path,
                        remote_target,
                        "--drive-import-formats",
                        "md,txt",
                    ],
                    capture_output=True,
                    text=True,
                    timeout=20,
                )
                if res.returncode == 0:
                    entry.gdrive_id = f"{entry.name}.md"
                    entry.synced = True
                    entry.dirty = False
                    self.save_index()
                    return True, f"Synced '{entry.name}' to {remote} as Google Doc."
                return False, f"rclone upload error: {res.stderr.strip()[:80]}"
            except Exception as exc:
                return False, f"rclone upload error: {exc}"

        return False, "Google Drive not connected."
