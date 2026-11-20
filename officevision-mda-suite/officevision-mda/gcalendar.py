#!/usr/bin/env python3
"""
Google Calendar REST API v3 Client & Offline Cache Manager
for IBM OfficeVision/400 (OV/400) on 90x43 Monochrome MDA Terminal.

Zero external dependencies (pure Python 3 stdlib + shared gdrive.py OAuth):
  - Full two-way Google Calendar API v3 sync:
      * List user calendars (primary + secondary/subscribed calendars)
      * Fetch events for month/range (with recurring event expansion)
      * Create, update, and delete events (timed, all-day, multi-day, recurring)
  - Local JSON cache & offline queue in:
      ~/.local/share/officevision-mda/calendar_events.json
      ~/.local/share/officevision-mda/calendar_config.json
"""

import calendar as pycal
import datetime
import json
import os
import re
import urllib.error
import urllib.parse
import urllib.request
import uuid
from typing import Any, Dict, List, Optional, Set, Tuple

from gdrive import APP_DATA_DIR, GoogleDriveManager

CALENDAR_API_BASE = "https://www.googleapis.com/calendar/v3"
CAL_EVENTS_PATH = os.path.join(APP_DATA_DIR, "calendar_events.json")
CAL_CONFIG_PATH = os.path.join(APP_DATA_DIR, "calendar_config.json")

REPEAT_OPTIONS: List[Tuple[str, str]] = [
    ("", "None"),
    ("DAILY", "Daily"),
    ("WEEKLY", "Weekly"),
    ("MONTHLY", "Monthly"),
    ("YEARLY", "Yearly"),
]


# =============================================================================
# Timezone, Date, and Time Parsing Helpers
# =============================================================================
def get_local_tzinfo() -> datetime.tzinfo:
    return datetime.datetime.now().astimezone().tzinfo or datetime.timezone.utc


def get_local_timezone_name() -> str:
    tz_env = os.environ.get("TZ", "").strip()
    if tz_env and "/" in tz_env:
        return tz_env.lstrip(":")
    if os.path.exists("/etc/timezone"):
        try:
            with open("/etc/timezone", "r", encoding="utf-8") as f:
                name = f.read().strip()
                if name:
                    return name
        except Exception:
            pass
    if os.path.exists("/etc/localtime"):
        try:
            real = os.path.realpath("/etc/localtime")
            marker = "zoneinfo/"
            idx = real.find(marker)
            if idx != -1:
                return real[idx + len(marker) :]
        except Exception:
            pass
    return "UTC"


def parse_date_str(
    raw: str, ref_date: Optional[datetime.date] = None
) -> Optional[datetime.date]:
    """Parses common date formats into datetime.date."""
    s = raw.strip()
    if not s:
        return None
    base = ref_date or datetime.date.today()
    low = s.lower()
    if low in ("today", "t"):
        return datetime.date.today()
    if low in ("tomorrow", "tmr"):
        return datetime.date.today() + datetime.timedelta(days=1)
    if low in ("yesterday", "yest"):
        return datetime.date.today() - datetime.timedelta(days=1)

    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%m-%d-%Y", "%m/%d/%y", "%m-%d-%y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            pass

    # MM/DD or MM-DD -> assume reference year
    m = re.fullmatch(r"(\d{1,2})[/-](\d{1,2})", s)
    if m:
        try:
            return datetime.date(base.year, int(m.group(1)), int(m.group(2)))
        except ValueError:
            return None

    # YYYY-MM or MM/YYYY -> 1st of that month
    m2 = re.fullmatch(r"(\d{4})[/-](\d{1,2})", s)
    if m2:
        try:
            return datetime.date(int(m2.group(1)), int(m2.group(2)), 1)
        except ValueError:
            return None
    m3 = re.fullmatch(r"(\d{1,2})[/-](\d{4})", s)
    if m3:
        try:
            return datetime.date(int(m3.group(2)), int(m3.group(1)), 1)
        except ValueError:
            return None

    return None


def parse_time_str(raw: str) -> Optional[Tuple[int, int]]:
    """
    Parses 24-hour or 12-hour time strings into (hour, minute).
    Accepts: '9:00', '09:30', '14:15', '2:30pm', '9am', '9:30 AM', '1430', '0900', '9'.
    """
    s = raw.strip().lower().replace(" ", "")
    if not s:
        return None

    ampm = ""
    if s.endswith("am") or s.endswith("pm"):
        ampm = s[-2:]
        s = s[:-2]
    elif s.endswith("a") or s.endswith("p"):
        ampm = s[-1] + "m"
        s = s[:-1]

    if not s:
        return None

    if ":" in s:
        parts = s.split(":")
        if len(parts) != 2 or not parts[0].isdigit() or not parts[1].isdigit():
            return None
        hh, mm = int(parts[0]), int(parts[1])
    elif s.isdigit():
        if len(s) in (1, 2):
            hh, mm = int(s), 0
        elif len(s) == 3:
            hh, mm = int(s[0]), int(s[1:])
        elif len(s) == 4:
            hh, mm = int(s[:2]), int(s[2:])
        else:
            return None
    else:
        return None

    if mm < 0 or mm > 59:
        return None

    if ampm:
        if hh < 1 or hh > 12:
            return None
        if ampm == "am":
            hh = 0 if hh == 12 else hh
        else:
            hh = 12 if hh == 12 else hh + 12
    else:
        if hh < 0 or hh > 23:
            return None

    return (hh, mm)


def format_hhmm(hh: int, mm: int) -> str:
    return f"{hh:02d}:{mm:02d}"


def format_time_12h(hhmm: str, compact: bool = False) -> str:
    parsed = parse_time_str(hhmm)
    if not parsed:
        return hhmm
    hh, mm = parsed
    suffix = "a" if hh < 12 else "p"
    h12 = hh % 12
    if h12 == 0:
        h12 = 12
    if compact:
        if mm == 0:
            return f"{h12}{suffix}"
        return f"{h12}:{mm:02d}{suffix}"
    return f"{h12}:{mm:02d}{suffix.upper()}M"


def add_minutes_to_hhmm(hhmm: str, minutes: int = 60) -> Tuple[str, int]:
    """Adds minutes to HH:MM, returning (new_hhmm, day_overflow)."""
    parsed = parse_time_str(hhmm)
    if not parsed:
        return ("10:00", 0)
    total = parsed[0] * 60 + parsed[1] + minutes
    day_delta, rem = divmod(total, 24 * 60)
    return (format_hhmm(rem // 60, rem % 60), day_delta)


def _parse_iso_datetime_local(iso_str: str) -> Optional[datetime.datetime]:
    if not iso_str:
        return None
    s = iso_str.strip()
    if s.endswith("Z"):
        s = s[:-1] + "+00:00"
    try:
        dt = datetime.datetime.fromisoformat(s)
        if dt.tzinfo is not None:
            dt = dt.astimezone(get_local_tzinfo())
        else:
            dt = dt.replace(tzinfo=get_local_tzinfo())
        return dt
    except ValueError:
        return None


def _month_grid_bounds(year: int, month: int) -> Tuple[datetime.date, datetime.date]:
    """Returns padded (start_date, end_date) covering all visible weeks of a month."""
    first_day = datetime.date(year, month, 1)
    last_day = datetime.date(year, month, pycal.monthrange(year, month)[1])
    return (
        first_day - datetime.timedelta(days=7),
        last_day + datetime.timedelta(days=7),
    )


# =============================================================================
# Data Models
# =============================================================================
class CalendarInfo:
    """Represents a Google Calendar entry from calendarList."""

    def __init__(
        self,
        cal_id: str,
        summary: str,
        primary: bool = False,
        access_role: str = "owner",
        enabled: bool = True,
        time_zone: str = "",
    ):
        self.cal_id = cal_id
        self.summary = summary or cal_id
        self.primary = primary
        self.access_role = access_role
        self.enabled = enabled
        self.time_zone = time_zone or get_local_timezone_name()

    @property
    def writable(self) -> bool:
        return self.access_role in ("owner", "writer")

    def to_dict(self) -> Dict[str, Any]:
        return {
            "cal_id": self.cal_id,
            "summary": self.summary,
            "primary": self.primary,
            "access_role": self.access_role,
            "enabled": self.enabled,
            "time_zone": self.time_zone,
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "CalendarInfo":
        return cls(
            cal_id=d.get("cal_id", "primary"),
            summary=d.get("summary", "Primary Calendar"),
            primary=bool(d.get("primary", False)),
            access_role=d.get("access_role", "owner"),
            enabled=bool(d.get("enabled", True)),
            time_zone=d.get("time_zone", ""),
        )


class CalendarEvent:
    """Represents a calendar event (Google Calendar API v3 + local cache)."""

    def __init__(
        self,
        event_id: str,
        summary: str,
        start_date: str,
        end_date: str = "",
        start_time: str = "",
        end_time: str = "",
        all_day: bool = False,
        location: str = "",
        description: str = "",
        calendar_id: str = "primary",
        calendar_name: str = "Primary",
        google_id: Optional[str] = None,
        recurrence: str = "",
        recurring_event_id: Optional[str] = None,
        dirty: bool = False,
        deleted: bool = False,
        updated_at: str = "",
    ):
        self.event_id = event_id
        self.summary = summary or "(No Title)"
        self.start_date = start_date
        self.end_date = end_date or start_date
        self.start_time = "" if all_day else start_time
        self.end_time = "" if all_day else end_time
        self.all_day = all_day
        self.location = location
        self.description = description
        self.calendar_id = calendar_id or "primary"
        self.calendar_name = calendar_name or "Primary"
        self.google_id = google_id
        self.recurrence = recurrence
        self.recurring_event_id = recurring_event_id
        self.dirty = dirty
        self.deleted = deleted
        self.updated_at = updated_at or datetime.datetime.now().strftime(
            "%Y-%m-%d %H:%M"
        )

    def start_date_obj(self) -> Optional[datetime.date]:
        return parse_date_str(self.start_date)

    def end_date_obj(self) -> Optional[datetime.date]:
        return parse_date_str(self.end_date) or self.start_date_obj()

    def occurs_on(self, day: datetime.date) -> bool:
        if self.deleted:
            return False
        s_d = self.start_date_obj()
        e_d = self.end_date_obj()
        if not s_d or not e_d:
            return False
        if not self.all_day and e_d > s_d and self.end_time == "00:00":
            # Timed event ending exactly at midnight does not belong to next day
            effective_end = e_d - datetime.timedelta(days=1)
        else:
            effective_end = e_d
        return s_d <= day <= effective_end

    def sort_key(self) -> Tuple[int, str, str]:
        return (0 if self.all_day else 1, self.start_time or "00:00", self.summary.lower())

    def time_range_label(self) -> str:
        if self.all_day:
            if self.end_date and self.end_date > self.start_date:
                return f"ALL DAY ({self.start_date[5:]}..{self.end_date[5:]})"
            return "ALL DAY"
        st = self.start_time or "00:00"
        et = self.end_time or st
        if self.end_date and self.end_date > self.start_date:
            return f"{st} - {self.end_date[5:]} {et}"
        return f"{st} - {et}"

    def compact_cell_label(self, width: int) -> str:
        """Formats event cleanly for narrow month grid cells (e.g. 11 chars on 90x43 MDA)."""
        if self.all_day:
            prefix = "*"
        elif width >= 15 and self.start_time:
            prefix = f"{self.start_time} "
        elif self.start_time:
            prefix = f"{format_time_12h(self.start_time, compact=True)} "
        else:
            prefix = "- "
        avail = max(1, width - len(prefix))
        return (prefix + self.summary)[:width] if avail > 0 else self.summary[:width]

    def to_dict(self) -> Dict[str, Any]:
        return {
            "event_id": self.event_id,
            "summary": self.summary,
            "start_date": self.start_date,
            "end_date": self.end_date,
            "start_time": self.start_time,
            "end_time": self.end_time,
            "all_day": self.all_day,
            "location": self.location,
            "description": self.description,
            "calendar_id": self.calendar_id,
            "calendar_name": self.calendar_name,
            "google_id": self.google_id,
            "recurrence": self.recurrence,
            "recurring_event_id": self.recurring_event_id,
            "dirty": self.dirty,
            "deleted": self.deleted,
            "updated_at": self.updated_at,
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "CalendarEvent":
        return cls(
            event_id=d.get("event_id", str(uuid.uuid4())[:10].upper()),
            summary=d.get("summary", "(No Title)"),
            start_date=d.get("start_date", datetime.date.today().isoformat()),
            end_date=d.get("end_date", d.get("start_date", datetime.date.today().isoformat())),
            start_time=d.get("start_time", ""),
            end_time=d.get("end_time", ""),
            all_day=bool(d.get("all_day", False)),
            location=d.get("location", ""),
            description=d.get("description", ""),
            calendar_id=d.get("calendar_id", "primary"),
            calendar_name=d.get("calendar_name", "Primary"),
            google_id=d.get("google_id"),
            recurrence=d.get("recurrence", ""),
            recurring_event_id=d.get("recurring_event_id"),
            dirty=bool(d.get("dirty", False)),
            deleted=bool(d.get("deleted", False)),
            updated_at=d.get("updated_at", ""),
        )

    @classmethod
    def from_google_item(
        cls, item: Dict[str, Any], calendar_id: str, calendar_name: str
    ) -> Optional["CalendarEvent"]:
        if item.get("status") == "cancelled":
            return None
        gid = item.get("id", "")
        summary = item.get("summary") or "(No Title)"
        start_obj = item.get("start") or {}
        end_obj = item.get("end") or {}

        if "date" in start_obj:
            all_day = True
            start_date = start_obj["date"][:10]
            raw_end = (end_obj.get("date") or start_date)[:10]
            # Google Calendar all-day end.date is EXCLUSIVE; convert to inclusive end_date
            s_d = parse_date_str(start_date)
            e_d = parse_date_str(raw_end)
            if s_d and e_d and e_d > s_d:
                end_date = (e_d - datetime.timedelta(days=1)).isoformat()
            else:
                end_date = start_date
            start_time = ""
            end_time = ""
        elif "dateTime" in start_obj:
            all_day = False
            s_dt = _parse_iso_datetime_local(start_obj["dateTime"])
            e_dt = _parse_iso_datetime_local(end_obj.get("dateTime", "")) or s_dt
            if not s_dt:
                return None
            if not e_dt:
                e_dt = s_dt + datetime.timedelta(hours=1)
            start_date = s_dt.date().isoformat()
            end_date = e_dt.date().isoformat()
            start_time = s_dt.strftime("%H:%M")
            end_time = e_dt.strftime("%H:%M")
        else:
            return None

        recurrence = ""
        rec_list = item.get("recurrence") or []
        if rec_list:
            joined = " ".join(rec_list).upper()
            for freq in ("DAILY", "WEEKLY", "MONTHLY", "YEARLY"):
                if f"FREQ={freq}" in joined:
                    recurrence = freq
                    break
            if not recurrence:
                recurrence = rec_list[0][:32]

        return cls(
            event_id=f"GC_{gid[:12]}" if gid else str(uuid.uuid4())[:10].upper(),
            summary=summary,
            start_date=start_date,
            end_date=end_date,
            start_time=start_time,
            end_time=end_time,
            all_day=all_day,
            location=item.get("location", "") or "",
            description=item.get("description", "") or "",
            calendar_id=calendar_id,
            calendar_name=calendar_name,
            google_id=gid or None,
            recurrence=recurrence,
            recurring_event_id=item.get("recurringEventId"),
            dirty=False,
            deleted=False,
            updated_at=(item.get("updated") or "")[:16].replace("T", " "),
        )

    def to_google_payload(self) -> Dict[str, Any]:
        payload: Dict[str, Any] = {
            "summary": self.summary,
            "location": self.location,
            "description": self.description,
        }
        tz_name = get_local_timezone_name()
        if self.all_day:
            s_d = self.start_date_obj() or datetime.date.today()
            e_d = self.end_date_obj() or s_d
            if e_d < s_d:
                e_d = s_d
            # Google Calendar API requires exclusive end date for all-day events
            excl_end = e_d + datetime.timedelta(days=1)
            payload["start"] = {"date": s_d.isoformat()}
            payload["end"] = {"date": excl_end.isoformat()}
        else:
            s_d = self.start_date_obj() or datetime.date.today()
            e_d = self.end_date_obj() or s_d
            st = parse_time_str(self.start_time) or (9, 0)
            et = parse_time_str(self.end_time) or ((st[0] + 1) % 24, st[1])
            tz = get_local_tzinfo()
            s_dt = datetime.datetime(s_d.year, s_d.month, s_d.day, st[0], st[1], tzinfo=tz)
            e_dt = datetime.datetime(e_d.year, e_d.month, e_d.day, et[0], et[1], tzinfo=tz)
            if e_dt <= s_dt:
                e_dt = s_dt + datetime.timedelta(hours=1)
            payload["start"] = {"dateTime": s_dt.isoformat(), "timeZone": tz_name}
            payload["end"] = {"dateTime": e_dt.isoformat(), "timeZone": tz_name}

        if self.recurrence in ("DAILY", "WEEKLY", "MONTHLY", "YEARLY"):
            payload["recurrence"] = [f"RRULE:FREQ={self.recurrence}"]
        return payload


# =============================================================================
# Google Calendar Manager
# =============================================================================
class GoogleCalendarManager:
    """Manages Google Calendar API v3 synchronization and local offline cache."""

    def __init__(self, gdrive: GoogleDriveManager):
        self.gdrive = gdrive
        os.makedirs(APP_DATA_DIR, exist_ok=True)
        self.config: Dict[str, Any] = self._load_json(
            CAL_CONFIG_PATH,
            {
                "default_calendar_id": "primary",
                "enabled_calendars": ["primary"],
                "use_24h": False,
                "first_weekday": 6,  # 6=Sunday (standard IBM OV/400)
            },
        )
        self.calendars: List[CalendarInfo] = []
        self.events: List[CalendarEvent] = []
        self.synced_months: Set[Tuple[int, int]] = set()
        self.last_sync_at: str = ""
        self._load_cache()

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
        self.config["enabled_calendars"] = [
            c.cal_id for c in self.calendars if c.enabled
        ] or ["primary"]
        self._save_json(CAL_CONFIG_PATH, self.config)

    def _load_cache(self):
        raw = self._load_json(CAL_EVENTS_PATH, {})
        if isinstance(raw, dict):
            self.calendars = [
                CalendarInfo.from_dict(c) for c in raw.get("calendars", [])
            ]
            self.events = [
                CalendarEvent.from_dict(e) for e in raw.get("events", [])
            ]
            self.last_sync_at = raw.get("last_sync_at", "")
        if not self.calendars:
            self.calendars = [
                CalendarInfo(
                    cal_id="primary",
                    summary="Primary Calendar",
                    primary=True,
                    access_role="owner",
                    enabled=True,
                )
            ]

    def save_cache(self):
        data = {
            "calendars": [c.to_dict() for c in self.calendars],
            "events": [e.to_dict() for e in self.events],
            "last_sync_at": self.last_sync_at,
        }
        self._save_json(CAL_EVENTS_PATH, data)

    # =========================================================================
    # Connection State
    # =========================================================================
    def is_connected(self) -> bool:
        pref = self.gdrive.config.get("backend", "auto")
        if pref == "local":
            return False
        return self.gdrive.is_api_authenticated()

    def get_status_label(self) -> str:
        if not self.is_connected():
            return "Local Mode (F10=Google Setup)"
        if not self.gdrive.has_calendar_scope():
            return "Re-Auth Needed (F10 -> F7/F8)"
        pending = self.pending_change_count()
        if pending > 0:
            return f"Google Calendar ({pending} unsynced)"
        return "Google Calendar Connected"

    def get_short_backend_label(self) -> str:
        if not self.is_connected():
            return "LOCAL"
        if not self.gdrive.has_calendar_scope():
            return "REAUTH"
        return "GCAL"

    def pending_change_count(self) -> int:
        return sum(1 for e in self.events if e.dirty or e.deleted)

    def get_calendar_name(self, cal_id: str) -> str:
        for c in self.calendars:
            if c.cal_id == cal_id:
                return c.summary
        return "Primary" if cal_id == "primary" else cal_id

    def get_writable_calendars(self) -> List[CalendarInfo]:
        w = [c for c in self.calendars if c.writable]
        return w or [self.calendars[0]]

    def get_default_calendar_id(self) -> str:
        pref = self.config.get("default_calendar_id", "primary")
        writable = self.get_writable_calendars()
        for c in writable:
            if c.cal_id == pref:
                return c.cal_id
        for c in writable:
            if c.primary:
                return c.cal_id
        return writable[0].cal_id

    def set_default_calendar_id(self, cal_id: str):
        self.config["default_calendar_id"] = cal_id
        for c in self.calendars:
            if c.cal_id == cal_id:
                c.enabled = True
        self.save_config()
        self.save_cache()

    # =========================================================================
    # Low-Level Google Calendar REST API v3 HTTP Client
    # =========================================================================
    def _api_request(
        self,
        method: str,
        path: str,
        params: Optional[Dict[str, str]] = None,
        body: Optional[Dict[str, Any]] = None,
        timeout: int = 15,
    ) -> Tuple[bool, Any, str]:
        token = self.gdrive.get_access_token()
        if not token:
            return False, None, "Not signed into Google OAuth. Press F10 to connect."

        url = f"{CALENDAR_API_BASE}{path}"
        if params:
            url += "?" + urllib.parse.urlencode(params)

        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        }
        data_bytes: Optional[bytes] = None
        if body is not None:
            data_bytes = json.dumps(body).encode("utf-8")
            headers["Content-Type"] = "application/json; charset=utf-8"

        req = urllib.request.Request(url, data=data_bytes, headers=headers, method=method)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8", errors="replace").strip()
                if not raw:
                    return True, {}, ""
                return True, json.loads(raw), ""
        except urllib.error.HTTPError as e:
            err_raw = e.read().decode("utf-8", errors="replace")
            reason = ""
            msg = f"HTTP {e.code}"
            try:
                err_json = json.loads(err_raw)
                err_obj = err_json.get("error", {})
                if isinstance(err_obj, dict):
                    msg = err_obj.get("message", msg)
                    errs = err_obj.get("errors") or []
                    if errs and isinstance(errs[0], dict):
                        reason = errs[0].get("reason", "")
            except Exception:
                msg = err_raw[:120] or msg

            if e.code == 401:
                return (
                    False,
                    None,
                    "Google OAuth token expired/invalid. Press F10 and re-authorize.",
                )
            if e.code == 403:
                low = (reason + " " + msg).lower()
                if "insufficient" in low or "scope" in low:
                    return (
                        False,
                        None,
                        "Google token lacks Calendar scope. Press F10 -> F7/F8 to re-authorize.",
                    )
                if "accessnotconfigured" in low or "not been used" in low or "disabled" in low:
                    return (
                        False,
                        None,
                        "Enable 'Google Calendar API' in your Google Cloud Console project.",
                    )
                return False, None, f"Google Calendar 403: {msg[:90]}"
            if e.code == 410 and method == "DELETE":
                return True, {}, ""
            return False, None, f"Google Calendar API error ({e.code}): {msg[:90]}"
        except urllib.error.URLError as e:
            return False, None, f"Network error reaching Google Calendar: {e.reason}"
        except Exception as exc:
            return False, None, f"Google Calendar error: {exc}"

    # =========================================================================
    # Sync Operations
    # =========================================================================
    def sync_calendar_list(self) -> Tuple[bool, str]:
        if not self.is_connected():
            return False, "Not connected to Google API. Press F10 for Google Setup."

        ok, data, err = self._api_request(
            "GET",
            "/users/me/calendarList",
            params={"minAccessRole": "reader", "maxResults": "100"},
        )
        if not ok:
            return False, err

        prev_enabled = {c.cal_id: c.enabled for c in self.calendars}
        cfg_enabled = set(self.config.get("enabled_calendars", []))
        items = data.get("items") or []
        new_cals: List[CalendarInfo] = []

        for item in items:
            cid = item.get("id", "")
            if not cid:
                continue
            is_primary = bool(item.get("primary", False))
            summary = item.get("summaryOverride") or item.get("summary") or cid
            role = item.get("accessRole", "reader")
            tz = item.get("timeZone", "")

            if cid in prev_enabled:
                enabled = prev_enabled[cid]
            elif is_primary and "primary" in prev_enabled:
                enabled = prev_enabled["primary"]
            elif cfg_enabled and cfg_enabled != {"primary"}:
                enabled = cid in cfg_enabled or is_primary
            else:
                enabled = bool(item.get("selected", is_primary))

            new_cals.append(
                CalendarInfo(
                    cal_id=cid,
                    summary=summary,
                    primary=is_primary,
                    access_role=role,
                    enabled=enabled,
                    time_zone=tz,
                )
            )

        if new_cals:
            # Sort primary first, then writable, then alphabetically
            new_cals.sort(
                key=lambda c: (0 if c.primary else 1, 0 if c.writable else 1, c.summary.lower())
            )
            if not any(c.enabled for c in new_cals):
                new_cals[0].enabled = True
            self.calendars = new_cals
            # Migrate any local "primary" events to the real primary calendar ID
            primary_cal = next((c for c in self.calendars if c.primary), None)
            if primary_cal:
                if self.config.get("default_calendar_id") == "primary":
                    self.config["default_calendar_id"] = primary_cal.cal_id
                for ev in self.events:
                    if ev.calendar_id == "primary":
                        ev.calendar_id = primary_cal.cal_id
                        ev.calendar_name = primary_cal.summary
            self.save_config()
            self.save_cache()

        return True, f"Loaded {len(self.calendars)} Google Calendar(s)."

    def push_pending_changes(self) -> Tuple[int, int, str]:
        """Pushes locally queued creates/edits/deletes to Google Calendar."""
        if not self.is_connected():
            return 0, 0, ""

        pushed = 0
        errors = 0
        last_err = ""
        remaining_events: List[CalendarEvent] = []

        for ev in list(self.events):
            if ev.deleted:
                if ev.google_id:
                    cal_q = urllib.parse.quote(ev.calendar_id, safe="")
                    ev_q = urllib.parse.quote(ev.google_id, safe="")
                    ok, _, err = self._api_request(
                        "DELETE", f"/calendars/{cal_q}/events/{ev_q}"
                    )
                    if ok:
                        pushed += 1
                        continue  # drop from cache
                    errors += 1
                    last_err = err
                    remaining_events.append(ev)
                else:
                    continue  # local-only deleted event
            elif ev.dirty:
                cal_q = urllib.parse.quote(ev.calendar_id, safe="")
                payload = ev.to_google_payload()
                if ev.google_id:
                    ev_q = urllib.parse.quote(ev.google_id, safe="")
                    ok, data, err = self._api_request(
                        "PUT", f"/calendars/{cal_q}/events/{ev_q}", body=payload
                    )
                else:
                    ok, data, err = self._api_request(
                        "POST", f"/calendars/{cal_q}/events", body=payload
                    )
                if ok:
                    if isinstance(data, dict) and data.get("id"):
                        ev.google_id = data["id"]
                    ev.dirty = False
                    pushed += 1
                else:
                    errors += 1
                    last_err = err
                remaining_events.append(ev)
            else:
                remaining_events.append(ev)

        self.events = remaining_events
        if pushed > 0:
            self.save_cache()
        return pushed, errors, last_err

    def sync_month(
        self,
        year: int,
        month: int,
        refresh_calendars: bool = False,
        push_pending: bool = True,
    ) -> Tuple[bool, str]:
        if not self.is_connected():
            return False, "Local mode (Press F10 to connect Google Calendar)."

        if refresh_calendars or (
            len(self.calendars) == 1 and self.calendars[0].cal_id == "primary"
        ):
            ok_c, msg_c = self.sync_calendar_list()
            if not ok_c:
                return False, msg_c

        pushed = 0
        if push_pending and self.pending_change_count() > 0:
            pushed, push_errs, push_err_msg = self.push_pending_changes()
            if push_errs and not pushed:
                return False, push_err_msg

        win_start, win_end = _month_grid_bounds(year, month)
        tz = get_local_tzinfo()
        t_min = datetime.datetime(
            win_start.year, win_start.month, win_start.day, 0, 0, 0, tzinfo=tz
        ).isoformat()
        t_max = datetime.datetime(
            win_end.year, win_end.month, win_end.day, 23, 59, 59, tzinfo=tz
        ).isoformat()

        enabled_cals = [c for c in self.calendars if c.enabled]
        if not enabled_cals:
            enabled_cals = self.calendars[:1]

        fetched_events: List[CalendarEvent] = []
        fetched_cal_ids: Set[str] = set()
        first_error = ""

        for cal in enabled_cals:
            cal_q = urllib.parse.quote(cal.cal_id, safe="")
            page_token = ""
            cal_ok = True
            for _ in range(5):  # up to 1250 events per month per calendar
                params: Dict[str, str] = {
                    "timeMin": t_min,
                    "timeMax": t_max,
                    "singleEvents": "true",
                    "orderBy": "startTime",
                    "maxResults": "250",
                }
                if page_token:
                    params["pageToken"] = page_token
                ok, data, err = self._api_request(
                    "GET", f"/calendars/{cal_q}/events", params=params
                )
                if not ok:
                    cal_ok = False
                    if not first_error:
                        first_error = err
                    break
                for item in data.get("items") or []:
                    ev = CalendarEvent.from_google_item(item, cal.cal_id, cal.summary)
                    if ev:
                        fetched_events.append(ev)
                page_token = data.get("nextPageToken", "")
                if not page_token:
                    break
            if cal_ok:
                fetched_cal_ids.add(cal.cal_id)

        if not fetched_cal_ids and first_error:
            return False, first_error

        # Merge fetched events into cache: keep events outside this window or still dirty
        kept: List[CalendarEvent] = []
        dirty_gids = {e.google_id for e in self.events if (e.dirty or e.deleted) and e.google_id}

        for ev in self.events:
            if ev.dirty or ev.deleted:
                kept.append(ev)
                continue
            if ev.calendar_id not in fetched_cal_ids:
                kept.append(ev)
                continue
            s_d = ev.start_date_obj()
            e_d = ev.end_date_obj()
            if s_d and e_d and not (e_d < win_start or s_d > win_end):
                # Inside refreshed window -> replace with fresh data from Google
                continue
            kept.append(ev)

        seen_keys: Set[Tuple[str, str]] = {
            (e.calendar_id, e.google_id or e.event_id) for e in kept
        }
        for fev in fetched_events:
            if fev.google_id and fev.google_id in dirty_gids:
                continue
            key = (fev.calendar_id, fev.google_id or fev.event_id)
            if key not in seen_keys:
                kept.append(fev)
                seen_keys.add(key)

        self.events = kept
        self.synced_months.add((year, month))
        self.last_sync_at = datetime.datetime.now().strftime("%m/%d %H:%M")
        self.save_cache()

        month_label = f"{pycal.month_abbr[month]} {year}"
        extra = f" (pushed {pushed} pending)" if pushed else ""
        if first_error:
            return True, f"Synced {month_label}: {len(fetched_events)} events{extra} (partial warn)."
        return True, f"Synced {month_label}: {len(fetched_events)} event(s) from Google Calendar{extra}."

    # =========================================================================
    # Event CRUD Operations (Immediate Google Sync + Local Fallback)
    # =========================================================================
    def create_event(
        self,
        summary: str,
        start_date: str,
        end_date: str = "",
        start_time: str = "",
        end_time: str = "",
        all_day: bool = False,
        location: str = "",
        description: str = "",
        calendar_id: Optional[str] = None,
        recurrence: str = "",
    ) -> Tuple[bool, str, Optional[CalendarEvent]]:
        clean_title = summary.strip()
        if not clean_title:
            return False, "Event title is required.", None

        s_d = parse_date_str(start_date)
        if not s_d:
            return False, f"Invalid start date: '{start_date}'. Use YYYY-MM-DD.", None
        e_d = parse_date_str(end_date, ref_date=s_d) if end_date.strip() else s_d
        if not e_d:
            return False, f"Invalid end date: '{end_date}'. Use YYYY-MM-DD.", None
        if e_d < s_d:
            return False, "End date cannot be earlier than start date.", None

        norm_start_t = ""
        norm_end_t = ""
        if not all_day:
            pt_s = parse_time_str(start_time)
            if not pt_s:
                return False, f"Invalid start time: '{start_time}'. Use HH:MM or 9:00am.", None
            norm_start_t = format_hhmm(pt_s[0], pt_s[1])
            if end_time.strip():
                pt_e = parse_time_str(end_time)
                if not pt_e:
                    return False, f"Invalid end time: '{end_time}'. Use HH:MM or 10:00am.", None
                norm_end_t = format_hhmm(pt_e[0], pt_e[1])
            else:
                norm_end_t, day_ov = add_minutes_to_hhmm(norm_start_t, 60)
                if day_ov and e_d == s_d:
                    e_d = s_d + datetime.timedelta(days=day_ov)
            if e_d == s_d and norm_end_t <= norm_start_t:
                return False, "End time must be after start time on the same day.", None

        target_cal_id = calendar_id or self.get_default_calendar_id()
        cal_name = self.get_calendar_name(target_cal_id)

        ev = CalendarEvent(
            event_id=f"EV_{str(uuid.uuid4())[:8].upper()}",
            summary=clean_title,
            start_date=s_d.isoformat(),
            end_date=e_d.isoformat(),
            start_time=norm_start_t,
            end_time=norm_end_t,
            all_day=all_day,
            location=location.strip(),
            description=description.strip(),
            calendar_id=target_cal_id,
            calendar_name=cal_name,
            recurrence=recurrence,
            dirty=True,
        )

        # Ensure target calendar is visible
        for c in self.calendars:
            if c.cal_id == target_cal_id:
                c.enabled = True

        self.events.append(ev)

        if self.is_connected():
            cal_q = urllib.parse.quote(ev.calendar_id, safe="")
            ok, data, err = self._api_request(
                "POST", f"/calendars/{cal_q}/events", body=ev.to_google_payload()
            )
            if ok:
                if isinstance(data, dict) and data.get("id"):
                    ev.google_id = data["id"]
                ev.dirty = False
                self.save_cache()
                return True, f"Created & synced '{ev.summary}' to Google Calendar.", ev
            self.save_cache()
            return True, f"Saved '{ev.summary}' locally (Sync warning: {err})", ev

        self.save_cache()
        return True, f"Created '{ev.summary}' locally (Press F10/F5 to sync to Google).", ev

    def update_event(
        self,
        ev: CalendarEvent,
        summary: str,
        start_date: str,
        end_date: str = "",
        start_time: str = "",
        end_time: str = "",
        all_day: bool = False,
        location: str = "",
        description: str = "",
        calendar_id: Optional[str] = None,
        recurrence: str = "",
    ) -> Tuple[bool, str]:
        clean_title = summary.strip()
        if not clean_title:
            return False, "Event title is required."

        s_d = parse_date_str(start_date)
        if not s_d:
            return False, f"Invalid start date: '{start_date}'. Use YYYY-MM-DD."
        e_d = parse_date_str(end_date, ref_date=s_d) if end_date.strip() else s_d
        if not e_d:
            return False, f"Invalid end date: '{end_date}'. Use YYYY-MM-DD."
        if e_d < s_d:
            return False, "End date cannot be earlier than start date."

        norm_start_t = ""
        norm_end_t = ""
        if not all_day:
            pt_s = parse_time_str(start_time)
            if not pt_s:
                return False, f"Invalid start time: '{start_time}'."
            norm_start_t = format_hhmm(pt_s[0], pt_s[1])
            if end_time.strip():
                pt_e = parse_time_str(end_time)
                if not pt_e:
                    return False, f"Invalid end time: '{end_time}'."
                norm_end_t = format_hhmm(pt_e[0], pt_e[1])
            else:
                norm_end_t, day_ov = add_minutes_to_hhmm(norm_start_t, 60)
                if day_ov and e_d == s_d:
                    e_d = s_d + datetime.timedelta(days=day_ov)
            if e_d == s_d and norm_end_t <= norm_start_t:
                return False, "End time must be after start time on the same day."

        old_cal_id = ev.calendar_id
        new_cal_id = calendar_id or ev.calendar_id

        ev.summary = clean_title
        ev.start_date = s_d.isoformat()
        ev.end_date = e_d.isoformat()
        ev.start_time = norm_start_t
        ev.end_time = norm_end_t
        ev.all_day = all_day
        ev.location = location.strip()
        ev.description = description.strip()
        ev.recurrence = recurrence
        ev.calendar_id = new_cal_id
        ev.calendar_name = self.get_calendar_name(new_cal_id)
        ev.updated_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        ev.dirty = True

        if self.is_connected():
            # If user moved an existing synced event to a different calendar, create on new & delete from old
            if ev.google_id and new_cal_id != old_cal_id:
                new_q = urllib.parse.quote(new_cal_id, safe="")
                ok, data, err = self._api_request(
                    "POST", f"/calendars/{new_q}/events", body=ev.to_google_payload()
                )
                if ok:
                    old_q = urllib.parse.quote(old_cal_id, safe="")
                    ev_q = urllib.parse.quote(ev.google_id, safe="")
                    self._api_request("DELETE", f"/calendars/{old_q}/events/{ev_q}")
                    if isinstance(data, dict) and data.get("id"):
                        ev.google_id = data["id"]
                    ev.dirty = False
                    self.save_cache()
                    return True, f"Moved & updated '{ev.summary}' on Google Calendar."
                self.save_cache()
                return True, f"Updated locally (Sync warning: {err})"

            cal_q = urllib.parse.quote(ev.calendar_id, safe="")
            if ev.google_id:
                ev_q = urllib.parse.quote(ev.google_id, safe="")
                ok, data, err = self._api_request(
                    "PUT", f"/calendars/{cal_q}/events/{ev_q}", body=ev.to_google_payload()
                )
            else:
                ok, data, err = self._api_request(
                    "POST", f"/calendars/{cal_q}/events", body=ev.to_google_payload()
                )
            if ok:
                if isinstance(data, dict) and data.get("id"):
                    ev.google_id = data["id"]
                ev.dirty = False
                self.save_cache()
                return True, f"Updated '{ev.summary}' on Google Calendar."
            self.save_cache()
            return True, f"Updated locally (Sync warning: {err})"

        self.save_cache()
        return True, f"Updated '{ev.summary}' locally."

    def delete_event(self, ev: CalendarEvent) -> Tuple[bool, str]:
        title = ev.summary
        if self.is_connected() and ev.google_id:
            cal_q = urllib.parse.quote(ev.calendar_id, safe="")
            ev_q = urllib.parse.quote(ev.google_id, safe="")
            ok, _, err = self._api_request(
                "DELETE", f"/calendars/{cal_q}/events/{ev_q}"
            )
            if ok:
                self.events = [e for e in self.events if e is not ev]
                self.save_cache()
                return True, f"Deleted '{title}' from Google Calendar."
            ev.deleted = True
            self.save_cache()
            return True, f"Removed locally (Google delete queued: {err})"

        if ev.google_id:
            ev.deleted = True
            self.save_cache()
            return True, f"Deleted '{title}' locally (will delete on Google at next sync)."

        self.events = [e for e in self.events if e is not ev]
        self.save_cache()
        return True, f"Deleted '{title}'."

    # =========================================================================
    # Query Helpers for Views
    # =========================================================================
    def _enabled_calendar_ids(self) -> Set[str]:
        enabled = {c.cal_id for c in self.calendars if c.enabled}
        return enabled or {c.cal_id for c in self.calendars} or {"primary"}

    def get_events_for_date(self, day: datetime.date) -> List[CalendarEvent]:
        enabled = self._enabled_calendar_ids()
        matched = [
            e
            for e in self.events
            if not e.deleted and e.calendar_id in enabled and e.occurs_on(day)
        ]
        matched.sort(key=lambda e: e.sort_key())
        return matched

    def get_events_map_for_range(
        self, start_day: datetime.date, end_day: datetime.date
    ) -> Dict[datetime.date, List[CalendarEvent]]:
        enabled = self._enabled_calendar_ids()
        by_day: Dict[datetime.date, List[CalendarEvent]] = {}
        for ev in self.events:
            if ev.deleted or ev.calendar_id not in enabled:
                continue
            s_d = ev.start_date_obj()
            e_d = ev.end_date_obj()
            if not s_d or not e_d:
                continue
            if not ev.all_day and e_d > s_d and ev.end_time == "00:00":
                eff_end = e_d - datetime.timedelta(days=1)
            else:
                eff_end = e_d
            if eff_end < start_day or s_d > end_day:
                continue
            cur = max(start_day, s_d)
            stop = min(end_day, eff_end)
            while cur <= stop:
                by_day.setdefault(cur, []).append(ev)
                cur += datetime.timedelta(days=1)

        for d_list in by_day.values():
            d_list.sort(key=lambda e: e.sort_key())
        return by_day

    def get_upcoming_events(
        self, start_date: datetime.date, days: int = 60, query: str = ""
    ) -> List[Tuple[datetime.date, CalendarEvent]]:
        end_date = start_date + datetime.timedelta(days=days)
        by_day = self.get_events_map_for_range(start_date, end_date)
        q = query.strip().lower()
        results: List[Tuple[datetime.date, CalendarEvent]] = []
        for day in sorted(by_day.keys()):
            for ev in by_day[day]:
                if q:
                    hay = f"{ev.summary} {ev.location} {ev.description} {ev.calendar_name}".lower()
                    if q not in hay:
                        continue
                results.append((day, ev))
        return results
