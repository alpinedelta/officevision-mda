#!/usr/bin/env python3
"""
Google Sheets & Local Spreadsheet Sync Manager for Lotus 1-2-3 on IBM OV/400.

Zero external dependencies (pure Python 3 stdlib + shared gdrive.py OAuth/rclone):
  1. Direct Google Sheets API v4 (multi-tab, live formula round-trip with
     valueRenderOption=FORMULA and valueInputOption=USER_ENTERED)
  2. Automatic Google Drive API v3 .xlsx fallback (if Google Sheets API v4
     isn't explicitly enabled in the user's Cloud project — still preserves
     multiple tabs and live formulas!)
  3. Rclone backend support (exports/imports .xlsx with tabs & formulas)
  4. Local .123.json library in ~/.local/share/officevision-mda/spreadsheets/
"""

import datetime
import json
import os
import re
import shutil
import subprocess
import tempfile
import urllib.error
import urllib.parse
import urllib.request
import uuid
from typing import Any, Dict, List, Optional, Tuple

from calc_engine import (
    Workbook,
    Worksheet,
    build_sample_workbook,
    export_workbook_to_xlsx_bytes,
    import_csv_to_worksheet,
    import_xlsx_bytes_to_workbook,
    to_gsheets_formula,
)
from gdrive import (
    APP_DATA_DIR,
    DRIVE_API_BASE,
    DRIVE_UPLOAD_BASE,
    GSHEET_MIME,
    GoogleDriveManager,
)

SHEETS_DIR = os.path.join(APP_DATA_DIR, "spreadsheets")
SHEET_INDEX_PATH = os.path.join(APP_DATA_DIR, "sheet_index.json")
SHEETS_API_BASE = "https://sheets.googleapis.com/v4/spreadsheets"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def extract_spreadsheet_id(url_or_id: str) -> str:
    """Extracts a Google Spreadsheet ID from a full URL or raw ID string."""
    s = url_or_id.strip()
    if not s:
        return ""
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", s)
    if m:
        return m.group(1)
    # Strip query params or slashes if user pasted partial URL
    return s.split("?")[0].split("#")[0].strip("/")


class SheetEntry:
    """Metadata index entry for a Lotus 1-2-3 / Google Sheets workbook."""

    def __init__(
        self,
        sheet_id: str,
        name: str,
        local_path: str = "",
        gsheet_id: Optional[str] = None,
        updated_at: str = "",
        synced: bool = False,
        dirty: bool = False,
        tab_count: int = 1,
        description: str = "",
    ):
        self.sheet_id = sheet_id
        self.name = name
        self.local_path = local_path
        self.gsheet_id = gsheet_id
        self.updated_at = updated_at or datetime.datetime.now().strftime("%m/%d/%y %H:%M")
        self.synced = synced
        self.dirty = dirty
        self.tab_count = tab_count
        self.description = description

    def to_dict(self) -> Dict[str, Any]:
        return {
            "sheet_id": self.sheet_id,
            "name": self.name,
            "local_path": self.local_path,
            "gsheet_id": self.gsheet_id,
            "updated_at": self.updated_at,
            "synced": self.synced,
            "dirty": self.dirty,
            "tab_count": self.tab_count,
            "description": self.description,
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "SheetEntry":
        return cls(
            sheet_id=d.get("sheet_id", str(uuid.uuid4())[:8].upper()),
            name=d.get("name", "UNTITLED"),
            local_path=d.get("local_path", ""),
            gsheet_id=d.get("gsheet_id"),
            updated_at=d.get("updated_at", ""),
            synced=bool(d.get("synced", False)),
            dirty=bool(d.get("dirty", False)),
            tab_count=int(d.get("tab_count", 1)),
            description=d.get("description", ""),
        )


class GoogleSheetsManager:
    """Manages local Lotus 1-2-3 workbooks and syncs with Google Sheets."""

    def __init__(self, gdrive: GoogleDriveManager):
        self.gdrive = gdrive
        os.makedirs(SHEETS_DIR, exist_ok=True)
        self.entries: List[SheetEntry] = []
        self._load_index()
        self._ensure_sample_workbook()

    def _load_index(self):
        raw: List[Dict[str, Any]] = []
        if os.path.exists(SHEET_INDEX_PATH):
            try:
                with open(SHEET_INDEX_PATH, "r", encoding="utf-8") as f:
                    raw = json.load(f)
            except Exception:
                raw = []
        self.entries = [SheetEntry.from_dict(item) for item in raw]

        # Discover any unindexed .123.json files in SHEETS_DIR
        known = {e.local_path for e in self.entries if e.local_path}
        for fname in sorted(os.listdir(SHEETS_DIR)):
            if not fname.endswith(".123.json"):
                continue
            fpath = os.path.join(SHEETS_DIR, fname)
            if os.path.isfile(fpath) and fpath not in known:
                stat = os.stat(fpath)
                dt = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%m/%d/%y %H:%M")
                base = fname[:-9]
                self.entries.append(
                    SheetEntry(
                        sheet_id=str(uuid.uuid4())[:8].upper(),
                        name=base,
                        local_path=fpath,
                        updated_at=dt,
                        synced=False,
                        dirty=False,
                        description="Local Lotus 1-2-3 Workbook",
                    )
                )
        self.save_index()

    def save_index(self):
        try:
            with open(SHEET_INDEX_PATH, "w", encoding="utf-8") as f:
                json.dump([e.to_dict() for e in self.entries], f, indent=2)
        except Exception:
            pass

    def _ensure_sample_workbook(self):
        if self.entries:
            return
        wb = build_sample_workbook()
        fpath = os.path.join(SHEETS_DIR, "Q4_BUDGET.123.json")
        try:
            with open(fpath, "w", encoding="utf-8") as f:
                json.dump(wb.to_dict(), f, indent=2)
            self.entries.append(
                SheetEntry(
                    sheet_id="LOTUS123",
                    name="Q4_BUDGET",
                    local_path=fpath,
                    updated_at=datetime.datetime.now().strftime("%m/%d/%y %H:%M"),
                    synced=False,
                    dirty=False,
                    tab_count=len(wb.sheets),
                    description="Sample Lotus 1-2-3 Budget Workbook",
                )
            )
            self.save_index()
        except Exception:
            pass

    def get_backend(self) -> str:
        return self.gdrive.get_active_backend()

    def get_status_summary(self) -> str:
        b = self.get_backend()
        if b == "api":
            return "Google Sheets API"
        if b == "rclone":
            return f"rclone ({self.gdrive.config.get('rclone_remote', 'gdrive:')})"
        return "Local Mode"

    def list_entries(self) -> List[SheetEntry]:
        return sorted(self.entries, key=lambda e: (0 if e.gsheet_id else 1, e.name.upper()))

    def _make_local_path(self, name: str) -> str:
        safe = "".join(c if c.isalnum() or c in ("-", "_") else "_" for c in name.strip())
        if not safe:
            safe = f"SHEET_{datetime.datetime.now().strftime('%H%M%S')}"
        path = os.path.join(SHEETS_DIR, f"{safe}.123.json")
        counter = 1
        while os.path.exists(path):
            path = os.path.join(SHEETS_DIR, f"{safe}_{counter}.123.json")
            counter += 1
        return path

    def create_workbook_entry(
        self,
        name: str,
        wb: Optional[Workbook] = None,
        gsheet_id: Optional[str] = None,
        description: str = "",
    ) -> Tuple[SheetEntry, Workbook]:
        clean = name.strip() or f"SHEET_{datetime.datetime.now().strftime('%H%M%S')}"
        if wb is None:
            wb = Workbook(title=clean)
        else:
            wb.title = clean
        wb.gsheet_id = gsheet_id
        local_path = self._make_local_path(clean)
        entry = SheetEntry(
            sheet_id=str(uuid.uuid4())[:8].upper(),
            name=clean,
            local_path=local_path,
            gsheet_id=gsheet_id,
            updated_at=datetime.datetime.now().strftime("%m/%d/%y %H:%M"),
            synced=bool(gsheet_id),
            dirty=True,
            tab_count=len(wb.sheets),
            description=description or "Lotus 1-2-3 Workbook",
        )
        self.entries.append(entry)
        self.save_workbook(entry, wb, sync_cloud=False)
        return entry, wb

    def load_workbook(self, entry: SheetEntry) -> Tuple[Workbook, str]:
        if entry.local_path and os.path.exists(entry.local_path):
            try:
                with open(entry.local_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                wb = Workbook.from_dict(data)
                wb.title = entry.name
                wb.gsheet_id = entry.gsheet_id or wb.gsheet_id
                wb.dirty = False
                return wb, f"Loaded '{entry.name}' ({len(wb.sheets)} tab(s))."
            except Exception as exc:
                return Workbook(title=entry.name), f"Load warning ({exc}); opened blank sheet."

        # Not cached locally yet — try pulling from Google Sheets
        if entry.gsheet_id and self.get_backend() != "local":
            ok, msg, wb = self.pull_spreadsheet(entry)
            if ok and wb is not None:
                return wb, msg

        wb = Workbook(title=entry.name)
        wb.gsheet_id = entry.gsheet_id
        return wb, f"Opened new workbook '{entry.name}'."

    def save_workbook(
        self, entry: SheetEntry, wb: Workbook, sync_cloud: bool = True
    ) -> Tuple[bool, str]:
        if not entry.local_path:
            entry.local_path = self._make_local_path(entry.name)
        wb.title = entry.name
        if wb.gsheet_id:
            entry.gsheet_id = wb.gsheet_id
        entry.tab_count = len(wb.sheets)
        entry.updated_at = datetime.datetime.now().strftime("%m/%d/%y %H:%M")

        try:
            with open(entry.local_path, "w", encoding="utf-8") as f:
                json.dump(wb.to_dict(), f, indent=2)
            wb.dirty = False
            entry.dirty = True
            self.save_index()
        except Exception as exc:
            return False, f"Local save failed: {exc}"

        auto_sync = bool(self.gdrive.config.get("auto_sync_on_save", True))
        if sync_cloud and auto_sync and self.get_backend() != "local" and entry.gsheet_id:
            ok, smsg = self.push_spreadsheet(entry, wb)
            if ok:
                return True, f"Saved & synced '{entry.name}' to Google Sheets."
            return True, f"Saved locally (Sync note: {smsg})"

        return True, f"Saved '{entry.name}' locally ({len(wb.sheets)} tab(s))."

    def delete_entry(self, entry: SheetEntry) -> Tuple[bool, str]:
        if entry.local_path and os.path.exists(entry.local_path):
            try:
                os.remove(entry.local_path)
            except Exception:
                pass
        self.entries = [e for e in self.entries if e.sheet_id != entry.sheet_id]
        self.save_index()
        return True, f"Removed local workbook '{entry.name}'."

    # =========================================================================
    # Google Sheets Remote Listing
    # =========================================================================
    def sync_spreadsheet_list(self) -> Tuple[bool, str]:
        """Lists Google Sheets from Google Drive and merges them into local index."""
        backend = self.get_backend()
        if backend == "api":
            return self._sync_list_via_api()
        if backend == "rclone":
            return self._sync_list_via_rclone()
        return False, "Google Drive/Sheets not configured. Use /Sheets Config (or F10 -> Setup)."

    def _sync_list_via_api(self) -> Tuple[bool, str]:
        token = self.gdrive.get_access_token()
        if not token:
            return False, "Google OAuth token missing/expired. Open Config to sign in."

        q = f"trashed = false and mimeType = '{GSHEET_MIME}'"
        params = urllib.parse.urlencode(
            {
                "q": q,
                "pageSize": 100,
                "fields": "files(id, name, modifiedTime, description)",
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
        except urllib.error.HTTPError as e:
            err = e.read().decode("utf-8", errors="replace")[:120]
            return False, f"Google Drive API error ({e.code}): {err}"
        except Exception as exc:
            return False, f"Network error listing Google Sheets: {exc}"

        files = data.get("files", [])
        by_gid = {e.gsheet_id: e for e in self.entries if e.gsheet_id}
        added, updated = 0, 0
        for f in files:
            gid = f["id"]
            name = f.get("name", "Untitled Sheet")
            mtime = f.get("modifiedTime", "")
            disp_time = mtime[:16].replace("T", " ") if mtime else ""
            desc = f.get("description") or "Google Sheet"

            if gid in by_gid:
                ent = by_gid[gid]
                ent.name = name
                if disp_time:
                    ent.updated_at = disp_time
                ent.synced = True
                updated += 1
            else:
                self.entries.append(
                    SheetEntry(
                        sheet_id=gid[:8].upper(),
                        name=name,
                        local_path="",
                        gsheet_id=gid,
                        updated_at=disp_time,
                        synced=True,
                        dirty=False,
                        description=desc,
                    )
                )
                added += 1

        self.save_index()
        return True, f"Google Sheets refreshed: {added} new, {updated} existing ({len(files)} total)."

    def _sync_list_via_rclone(self) -> Tuple[bool, str]:
        rclone_bin = shutil.which("rclone")
        if not rclone_bin:
            return False, "rclone binary not found."
        remote = self.gdrive.config.get("rclone_remote", "gdrive:")
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
                    "xlsx,csv",
                ],
                capture_output=True,
                text=True,
                timeout=20,
            )
            if res.returncode != 0:
                return False, f"rclone error: {res.stderr.strip()[:80]}"
            items = json.loads(res.stdout)
        except Exception as exc:
            return False, f"rclone failed: {exc}"

        by_name = {e.name: e for e in self.entries}
        added, updated = 0, 0
        for item in items:
            if item.get("IsDir"):
                continue
            fname = item.get("Name", "")
            ext = os.path.splitext(fname)[1].lower()
            if ext not in (".xlsx", ".csv", ".ods"):
                continue
            stem = os.path.splitext(fname)[0] or fname
            mtime = item.get("ModTime", "")[:16].replace("T", " ")
            if stem in by_name:
                ent = by_name[stem]
                ent.gsheet_id = fname
                ent.synced = True
                if mtime:
                    ent.updated_at = mtime
                updated += 1
            else:
                self.entries.append(
                    SheetEntry(
                        sheet_id=str(uuid.uuid4())[:8].upper(),
                        name=stem,
                        local_path="",
                        gsheet_id=fname,
                        updated_at=mtime,
                        synced=True,
                        dirty=False,
                        description=f"Google Sheet ({fname})",
                    )
                )
                added += 1

        self.save_index()
        return True, f"rclone found {added} new, {updated} existing sheets on {remote}."

    # =========================================================================
    # Pull from Google Sheets -> Local Lotus 1-2-3 Workbook
    # =========================================================================
    def pull_spreadsheet(self, entry: SheetEntry) -> Tuple[bool, str, Optional[Workbook]]:
        if not entry.gsheet_id:
            return False, "Workbook is not linked to a Google Sheet yet. Use Push or Link.", None

        backend = self.get_backend()
        if backend == "api":
            return self._pull_via_api(entry)
        if backend == "rclone":
            return self._pull_via_rclone(entry)
        return False, "No active Google connection. Configure via /Sheets Config.", None

    def _api_json_request(
        self, url: str, token: str, method: str = "GET", body: Optional[Dict[str, Any]] = None
    ) -> Tuple[int, Any]:
        data_bytes = json.dumps(body).encode("utf-8") if body is not None else None
        headers = {"Authorization": f"Bearer {token}"}
        if data_bytes is not None:
            headers["Content-Type"] = "application/json; charset=UTF-8"
        req = urllib.request.Request(url, data=data_bytes, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                raw = resp.read().decode("utf-8", errors="replace")
                return resp.status, (json.loads(raw) if raw else {})
        except urllib.error.HTTPError as e:
            raw = e.read().decode("utf-8", errors="replace")
            try:
                return e.code, json.loads(raw)
            except Exception:
                return e.code, {"error": {"message": raw[:180]}}

    def _pull_via_api(self, entry: SheetEntry) -> Tuple[bool, str, Optional[Workbook]]:
        token = self.gdrive.get_access_token()
        if not token:
            return False, "Missing Google OAuth token.", None

        gid = entry.gsheet_id
        # 1. Try Google Sheets API v4 (preserves exact formulas & all tabs)
        meta_url = (
            f"{SHEETS_API_BASE}/{urllib.parse.quote(gid)}?"
            f"{urllib.parse.urlencode({'fields': 'properties.title,sheets(properties(sheetId,title))'})}"
        )
        try:
            status, meta = self._api_json_request(meta_url, token)
        except Exception as exc:
            return False, f"Network error pulling Google Sheet: {exc}", None

        if status == 200 and isinstance(meta, dict):
            title = meta.get("properties", {}).get("title", entry.name)
            sheet_props = [s.get("properties", {}) for s in meta.get("sheets", [])]
            tab_names = [sp.get("title", f"Sheet{i + 1}") for i, sp in enumerate(sheet_props)]
            if not tab_names:
                tab_names = ["Sheet1"]

            # Batch get all sheet tabs with FORMULA render option
            query_parts = [("valueRenderOption", "FORMULA")]
            for tname in tab_names:
                escaped_t = tname.replace("'", "''")
                query_parts.append(("ranges", f"'{escaped_t}'!A1:ZZ2000"))
            batch_url = (
                f"{SHEETS_API_BASE}/{urllib.parse.quote(gid)}/values:batchGet?"
                f"{urllib.parse.urlencode(query_parts)}"
            )
            try:
                bstatus, bdata = self._api_json_request(batch_url, token)
            except Exception as exc:
                return False, f"Failed fetching sheet values: {exc}", None

            if bstatus == 200 and isinstance(bdata, dict):
                wb = Workbook(title=title)
                wb.gsheet_id = gid
                wb.sheets = []
                value_ranges = bdata.get("valueRanges", [])
                for idx, tname in enumerate(tab_names):
                    ws = Worksheet(tname)
                    vr = value_ranges[idx] if idx < len(value_ranges) else {}
                    rows_data = vr.get("values", [])
                    for r, row_vals in enumerate(rows_data):
                        for c, val in enumerate(row_vals):
                            if val is None or val == "":
                                continue
                            ws.set_raw(r, c, str(val))
                    # Auto-fit initial column widths nicely
                    _, max_c = ws.get_bounds()
                    for c in range(min(max_c + 1, 26)):
                        ws.auto_fit_col(c, min_w=9, max_w=24)
                    wb.sheets.append(ws)
                wb.recalculate()
                wb.dirty = False

                entry.name = title
                if not entry.local_path:
                    entry.local_path = self._make_local_path(title)
                entry.tab_count = len(wb.sheets)
                entry.synced = True
                entry.dirty = False
                entry.updated_at = datetime.datetime.now().strftime("%m/%d/%y %H:%M")
                with open(entry.local_path, "w", encoding="utf-8") as f:
                    json.dump(wb.to_dict(), f, indent=2)
                self.save_index()
                return True, f"Pulled '{title}' via Google Sheets API ({len(wb.sheets)} tab(s)).", wb

        # 2. Fallback: Drive API v3 XLSX export (works even if Sheets API v4 is disabled!)
        if status in (403, 404):
            return self._pull_via_drive_export(entry, token)

        err_msg = ""
        if isinstance(meta, dict):
            err_msg = meta.get("error", {}).get("message", str(meta))[:120]
        return False, f"Sheets API error ({status}): {err_msg}", None

    def _pull_via_drive_export(
        self, entry: SheetEntry, token: str
    ) -> Tuple[bool, str, Optional[Workbook]]:
        gid = entry.gsheet_id or ""
        export_url = (
            f"{DRIVE_API_BASE}/files/{urllib.parse.quote(gid)}/export?"
            f"{urllib.parse.urlencode({'mimeType': XLSX_MIME})}"
        )
        req = urllib.request.Request(export_url, headers={"Authorization": f"Bearer {token}"})
        try:
            with urllib.request.urlopen(req, timeout=25) as resp:
                xlsx_bytes = resp.read()
            wb = import_xlsx_bytes_to_workbook(xlsx_bytes, title=entry.name)
            wb.gsheet_id = gid
        except Exception:
            # Final fallback: CSV export of primary tab
            csv_url = (
                f"{DRIVE_API_BASE}/files/{urllib.parse.quote(gid)}/export?"
                f"{urllib.parse.urlencode({'mimeType': 'text/csv'})}"
            )
            req2 = urllib.request.Request(csv_url, headers={"Authorization": f"Bearer {token}"})
            try:
                with urllib.request.urlopen(req2, timeout=20) as resp2:
                    csv_text = resp2.read().decode("utf-8", errors="replace")
                wb = Workbook(title=entry.name)
                wb.gsheet_id = gid
                wb.sheets = [import_csv_to_worksheet(csv_text, "Sheet1")]
                wb.recalculate()
            except Exception as exc2:
                return False, f"Drive export failed: {exc2}", None

        for ws in wb.sheets:
            _, max_c = ws.get_bounds()
            for c in range(min(max_c + 1, 26)):
                ws.auto_fit_col(c, min_w=9, max_w=24)

        wb.dirty = False
        if not entry.local_path:
            entry.local_path = self._make_local_path(entry.name)
        entry.tab_count = len(wb.sheets)
        entry.synced = True
        entry.dirty = False
        entry.updated_at = datetime.datetime.now().strftime("%m/%d/%y %H:%M")
        with open(entry.local_path, "w", encoding="utf-8") as f:
            json.dump(wb.to_dict(), f, indent=2)
        self.save_index()
        return True, f"Pulled '{entry.name}' via Drive XLSX export ({len(wb.sheets)} tab(s)).", wb

    def _pull_via_rclone(self, entry: SheetEntry) -> Tuple[bool, str, Optional[Workbook]]:
        rclone_bin = shutil.which("rclone")
        if not rclone_bin:
            return False, "rclone binary not found.", None
        remote = self.gdrive.config.get("rclone_remote", "gdrive:")
        if not remote.endswith(":"):
            remote += ":"
        remote_name = entry.gsheet_id or f"{entry.name}.xlsx"

        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", prefix="ov400_sheet_")
        os.close(tmp_fd)
        try:
            res = subprocess.run(
                [
                    rclone_bin,
                    "copyto",
                    f"{remote}{remote_name}",
                    tmp_path,
                    "--drive-export-formats",
                    "xlsx",
                ],
                capture_output=True,
                text=True,
                timeout=25,
            )
            if res.returncode != 0:
                return False, f"rclone pull error: {res.stderr.strip()[:90]}", None
            with open(tmp_path, "rb") as f:
                raw_bytes = f.read()
            if remote_name.lower().endswith(".csv"):
                wb = Workbook(title=entry.name)
                wb.sheets = [import_csv_to_worksheet(raw_bytes.decode("utf-8", errors="replace"), "Sheet1")]
                wb.recalculate()
            else:
                wb = import_xlsx_bytes_to_workbook(raw_bytes, title=entry.name)
            wb.gsheet_id = remote_name
            if not entry.local_path:
                entry.local_path = self._make_local_path(entry.name)
            entry.tab_count = len(wb.sheets)
            entry.synced = True
            entry.dirty = False
            with open(entry.local_path, "w", encoding="utf-8") as f:
                json.dump(wb.to_dict(), f, indent=2)
            self.save_index()
            return True, f"Pulled '{entry.name}' from {remote} ({len(wb.sheets)} tab(s)).", wb
        except Exception as exc:
            return False, f"rclone pull failed: {exc}", None
        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

    # =========================================================================
    # Push Local Lotus 1-2-3 Workbook -> Google Sheets
    # =========================================================================
    def push_spreadsheet(self, entry: SheetEntry, wb: Workbook) -> Tuple[bool, str]:
        backend = self.get_backend()
        if backend == "api":
            return self._push_via_api(entry, wb)
        if backend == "rclone":
            return self._push_via_rclone(entry, wb)
        return False, "No active Google connection. Configure via /Sheets Config."

    def _sheet_to_value_grid(self, ws: Worksheet) -> List[List[Any]]:
        max_r, max_c = ws.get_bounds()
        if max_r < 0:
            return [[""]]
        grid: List[List[Any]] = []
        for r in range(max_r + 1):
            row_vals: List[Any] = []
            for c in range(max_c + 1):
                cell = ws.get_cell(r, c)
                if not cell or cell.raw == "":
                    row_vals.append("")
                else:
                    row_vals.append(to_gsheets_formula(cell.raw))
            grid.append(row_vals)
        return grid

    def _push_via_api(self, entry: SheetEntry, wb: Workbook) -> Tuple[bool, str]:
        token = self.gdrive.get_access_token()
        if not token:
            return False, "Missing Google OAuth token."

        gid = entry.gsheet_id or wb.gsheet_id

        # 1. Create spreadsheet via Sheets API v4 if not linked yet
        if not gid:
            create_body = {
                "properties": {"title": entry.name},
                "sheets": [{"properties": {"title": ws.name}} for ws in wb.sheets],
            }
            try:
                cstatus, cdata = self._api_json_request(
                    SHEETS_API_BASE, token, method="POST", body=create_body
                )
            except Exception as exc:
                return False, f"Network error creating Google Sheet: {exc}"

            if cstatus == 200 and isinstance(cdata, dict) and cdata.get("spreadsheetId"):
                gid = cdata["spreadsheetId"]
                entry.gsheet_id = gid
                wb.gsheet_id = gid
            elif cstatus in (403, 404):
                # Fallback to Drive multipart XLSX upload
                return self._push_via_drive_xlsx(entry, wb, token)
            else:
                msg = cdata.get("error", {}).get("message", str(cdata))[:110] if isinstance(cdata, dict) else ""
                return False, f"Sheets create error ({cstatus}): {msg}"
        else:
            # Ensure all local sheet tabs exist remotely
            meta_url = (
                f"{SHEETS_API_BASE}/{urllib.parse.quote(gid)}?"
                f"{urllib.parse.urlencode({'fields': 'sheets(properties(sheetId,title))'})}"
            )
            try:
                mstatus, mdata = self._api_json_request(meta_url, token)
            except Exception as exc:
                return False, f"Network error checking Google Sheet tabs: {exc}"

            if mstatus in (403, 404):
                return self._push_via_drive_xlsx(entry, wb, token)
            if mstatus != 200:
                msg = mdata.get("error", {}).get("message", "")[:110] if isinstance(mdata, dict) else ""
                return False, f"Sheets API error ({mstatus}): {msg}"

            remote_tabs = {
                s.get("properties", {}).get("title", "")
                for s in mdata.get("sheets", [])
            }
            add_requests = [
                {"addSheet": {"properties": {"title": ws.name}}}
                for ws in wb.sheets
                if ws.name not in remote_tabs
            ]
            if add_requests:
                self._api_json_request(
                    f"{SHEETS_API_BASE}/{urllib.parse.quote(gid)}:batchUpdate",
                    token,
                    method="POST",
                    body={"requests": add_requests},
                )

        # 2. Clear existing values on each tab, then push fresh formulas/values
        clear_ranges = [f"'{ws.name.replace(chr(39), chr(39)*2)}'!A1:ZZ2000" for ws in wb.sheets]
        try:
            self._api_json_request(
                f"{SHEETS_API_BASE}/{urllib.parse.quote(gid)}/values:batchClear",
                token,
                method="POST",
                body={"ranges": clear_ranges},
            )
        except Exception:
            pass

        data_payload = []
        total_cells = 0
        for ws in wb.sheets:
            grid = self._sheet_to_value_grid(ws)
            total_cells += sum(1 for row in grid for v in row if v != "")
            escaped_t = ws.name.replace("'", "''")
            data_payload.append(
                {
                    "range": f"'{escaped_t}'!A1",
                    "majorDimension": "ROWS",
                    "values": grid,
                }
            )

        update_body = {
            "valueInputOption": "USER_ENTERED",
            "data": data_payload,
        }
        try:
            ustatus, udata = self._api_json_request(
                f"{SHEETS_API_BASE}/{urllib.parse.quote(gid)}/values:batchUpdate",
                token,
                method="POST",
                body=update_body,
            )
        except Exception as exc:
            return False, f"Network error updating Google Sheet: {exc}"

        if ustatus == 200:
            entry.gsheet_id = gid
            wb.gsheet_id = gid
            entry.synced = True
            entry.dirty = False
            entry.updated_at = datetime.datetime.now().strftime("%m/%d/%y %H:%M")
            self.save_index()
            return (
                True,
                f"Pushed '{entry.name}' to Google Sheets ({len(wb.sheets)} tab(s), {total_cells} cells).",
            )

        if ustatus in (403, 404):
            return self._push_via_drive_xlsx(entry, wb, token)

        msg = udata.get("error", {}).get("message", str(udata))[:110] if isinstance(udata, dict) else ""
        return False, f"Sheets push error ({ustatus}): {msg}"

    def _push_via_drive_xlsx(
        self, entry: SheetEntry, wb: Workbook, token: str
    ) -> Tuple[bool, str]:
        """Uploads a generated .xlsx and lets Google Drive convert it to a native Google Sheet."""
        xlsx_bytes = export_workbook_to_xlsx_bytes(wb)
        boundary = f"ov400_123_{uuid.uuid4().hex}"
        metadata: Dict[str, Any] = {
            "name": entry.name,
            "mimeType": GSHEET_MIME,
            "description": entry.description or "Synced by Lotus 1-2-3 for IBM OV/400",
        }
        header_part = (
            f"--{boundary}\r\n"
            "Content-Type: application/json; charset=UTF-8\r\n\r\n"
            f"{json.dumps(metadata)}\r\n"
            f"--{boundary}\r\n"
            f"Content-Type: {XLSX_MIME}\r\n\r\n"
        ).encode("utf-8")
        body = header_part + xlsx_bytes + f"\r\n--{boundary}--\r\n".encode("utf-8")

        gid = entry.gsheet_id or wb.gsheet_id
        if gid:
            url = f"{DRIVE_UPLOAD_BASE}/files/{urllib.parse.quote(gid)}?uploadType=multipart"
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
            with urllib.request.urlopen(req, timeout=25) as resp:
                resp_data = json.loads(resp.read().decode("utf-8"))
            new_gid = resp_data.get("id", gid)
            entry.gsheet_id = new_gid
            wb.gsheet_id = new_gid
            entry.synced = True
            entry.dirty = False
            entry.updated_at = datetime.datetime.now().strftime("%m/%d/%y %H:%M")
            self.save_index()
            return True, f"Pushed '{entry.name}' to Google Sheets via Drive XLSX ({len(wb.sheets)} tab(s))."
        except Exception as exc:
            return False, f"Drive XLSX push failed: {exc}"

    def _push_via_rclone(self, entry: SheetEntry, wb: Workbook) -> Tuple[bool, str]:
        rclone_bin = shutil.which("rclone")
        if not rclone_bin:
            return False, "rclone binary not found."
        remote = self.gdrive.config.get("rclone_remote", "gdrive:")
        if not remote.endswith(":"):
            remote += ":"
        remote_target = f"{remote}{entry.name}.xlsx"

        tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx", prefix="ov400_push_")
        os.close(tmp_fd)
        try:
            with open(tmp_path, "wb") as f:
                f.write(export_workbook_to_xlsx_bytes(wb))
            res = subprocess.run(
                [
                    rclone_bin,
                    "copyto",
                    tmp_path,
                    remote_target,
                    "--drive-import-formats",
                    "xlsx",
                ],
                capture_output=True,
                text=True,
                timeout=25,
            )
            if res.returncode == 0:
                entry.gsheet_id = f"{entry.name}.xlsx"
                wb.gsheet_id = entry.gsheet_id
                entry.synced = True
                entry.dirty = False
                self.save_index()
                return True, f"Pushed '{entry.name}' to {remote} as Google Sheet."
            return False, f"rclone push error: {res.stderr.strip()[:90]}"
        except Exception as exc:
            return False, f"rclone push failed: {exc}"
        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
