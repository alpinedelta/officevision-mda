"""Plex SQLite Database reader for fast local music search."""

import os
import sqlite3
from dataclasses import dataclass
from typing import List, Optional

from config import get_db_path

@dataclass
class Artist:
    id: int
    name: str
    album_count: int
    track_count: int

    def display_line(self, max_width: int = 85) -> str:
        """Format a single line summary for an artist."""
        ar_len = max(20, max_width - 32)
        name = (self.name[:ar_len - 2] + "..") if len(self.name) > ar_len else self.name
        al_label = "album " if self.album_count == 1 else "albums"
        tr_label = "track " if self.track_count == 1 else "tracks"
        fmt = f"{name:<{ar_len}} | {self.album_count:>6} {al_label} | {self.track_count:>6} {tr_label}"
        return fmt[:max_width]


@dataclass
class Album:
    id: int
    title: str
    artist_id: Optional[int]
    artist: str
    year: Optional[int]
    track_count: int

    def display_line(self, max_width: int = 85) -> str:
        """Format a single line summary for an album."""
        year_str = f"{self.year:4d}" if self.year else "   -"
        tr_str = f"{self.track_count:3d} trk"
        avail = max(30, max_width - 20)
        al_len = int(avail * 0.56)
        ar_len = avail - al_len
        title = (self.title[:al_len - 2] + "..") if len(self.title) > al_len else self.title
        artist = (self.artist[:ar_len - 2] + "..") if len(self.artist) > ar_len else self.artist
        fmt = f"{title:<{al_len}} | {artist:<{ar_len}} | {year_str} | {tr_str}"
        return fmt[:max_width]


@dataclass
class Track:
    id: int
    title: str
    artist: str
    album: str
    year: Optional[int]
    track_num: Optional[int]
    duration_ms: int
    file_path: str
    format: str

    @property
    def duration_sec(self) -> float:
        return self.duration_ms / 1000.0

    @property
    def duration_str(self) -> str:
        total_sec = int(self.duration_sec)
        minutes = total_sec // 60
        seconds = total_sec % 60
        return f"{minutes:02d}:{seconds:02d}"

    def display_line(self, max_width: int = 85) -> str:
        """Format a single line summary fitting within terminal width."""
        num_str = f"{self.track_num:2d}. " if self.track_num else "    "
        dur_str = f" {self.duration_str}"
        avail = max(30, max_width - len(num_str) - len(dur_str) - 6)
        # Allocate roughly 40% for title, 28% for artist, 32% for album
        t_len = int(avail * 0.40)
        a_len = int(avail * 0.28)
        al_len = avail - t_len - a_len
        title = (self.title[:t_len - 2] + "..") if len(self.title) > t_len else self.title
        artist = (self.artist[:a_len - 2] + "..") if len(self.artist) > a_len else self.artist
        album = (self.album[:al_len - 2] + "..") if len(self.album) > al_len else self.album
        fmt = f"{num_str}{title:<{t_len}} | {artist:<{a_len}} | {album:<{al_len}}{dur_str}"
        return fmt[:max_width]


class MusicDB:
    def __init__(self, db_path: Optional[str] = None):
        self.db_path = os.path.expanduser(db_path) if db_path else get_db_path()
        if not os.path.exists(self.db_path):
            raise FileNotFoundError(
                f"Plex database not found at '{self.db_path}'. "
                "Set PLEX_DB_PATH=/path/to/com.plexapp.plugins.library.db or pass --db."
            )
        if not os.access(self.db_path, os.R_OK):
            raise PermissionError(
                f"Cannot read Plex database at '{self.db_path}'. "
                "Ensure your user has read permissions (e.g. add your user to the 'plex' group)."
            )

    def _get_connection(self):
        return sqlite3.connect(f"file:{self.db_path}?mode=ro", uri=True)

    def _format_from_path(self, path: str) -> str:
        if not path:
            return "AUDIO"
        ext = os.path.splitext(path)[1].lower()
        if ext == ".flac":
            return "FLAC"
        elif ext == ".mp3":
            return "MP3"
        elif ext in [".m4a", ".aac"]:
            return "AAC"
        elif ext == ".wav":
            return "WAV"
        elif ext == ".ogg":
            return "OGG"
        return ext.strip(".").upper()

    def search_tracks(self, query: str, limit: int = 100) -> List[Track]:
        """Search tracks by title, album, or artist."""
        if not query or not query.strip():
            return []

        terms = query.strip().split()
        sql = """
        SELECT m.id,
               COALESCE(m.title, 'Unknown Title') as title,
               COALESCE(parent.title, 'Unknown Album') as album,
               COALESCE(grandparent.title, 'Unknown Artist') as artist,
               m.year,
               m."index" as track_num,
               COALESCE(mi.duration, p.duration, 0) as dur,
               p.file
        FROM metadata_items m
        LEFT JOIN metadata_items parent ON parent.id = m.parent_id
        LEFT JOIN metadata_items grandparent ON grandparent.id = parent.parent_id
        LEFT JOIN media_items mi ON mi.metadata_item_id = m.id
        LEFT JOIN media_parts p ON p.media_item_id = mi.id
        WHERE m.metadata_type = 10
        """
        params = []
        for term in terms:
            sql += " AND (m.title LIKE ? OR parent.title LIKE ? OR grandparent.title LIKE ?)"
            pattern = f"%{term}%"
            params.extend([pattern, pattern, pattern])

        sql += " ORDER BY grandparent.title, parent.title, m.\"index\" LIMIT ?"
        params.append(limit)

        tracks = []
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(sql, params)
            for row in cur.fetchall():
                t_id, title, album, artist, year, track_num, dur, fpath = row
                tracks.append(
                    Track(
                        id=t_id,
                        title=title,
                        album=album,
                        artist=artist,
                        year=year,
                        track_num=track_num,
                        duration_ms=dur or 0,
                        file_path=fpath or "",
                        format=self._format_from_path(fpath),
                    )
                )
        return tracks

    def get_track(self, track_id: int) -> Optional[Track]:
        """Fetch a single track by its Plex metadata ID."""
        sql = """
        SELECT m.id,
               COALESCE(m.title, 'Unknown Title') as title,
               COALESCE(parent.title, 'Unknown Album') as album,
               COALESCE(grandparent.title, 'Unknown Artist') as artist,
               m.year,
               m."index" as track_num,
               COALESCE(mi.duration, p.duration, 0) as dur,
               p.file
        FROM metadata_items m
        LEFT JOIN metadata_items parent ON parent.id = m.parent_id
        LEFT JOIN metadata_items grandparent ON grandparent.id = parent.parent_id
        LEFT JOIN media_items mi ON mi.metadata_item_id = m.id
        LEFT JOIN media_parts p ON p.media_item_id = mi.id
        WHERE m.id = ? AND m.metadata_type = 10
        LIMIT 1
        """
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(sql, (track_id,))
            row = cur.fetchone()
            if not row:
                return None
            t_id, title, album, artist, year, track_num, dur, fpath = row
            return Track(
                id=t_id,
                title=title,
                album=album,
                artist=artist,
                year=year,
                track_num=track_num,
                duration_ms=dur or 0,
                file_path=fpath or "",
                format=self._format_from_path(fpath),
            )

    def get_album_tracks_for_track(self, track_id: int) -> List[Track]:
        """Get all tracks belonging to the same album as the given track."""
        sql = """
        SELECT m2.id,
               COALESCE(m2.title, 'Unknown Title') as title,
               COALESCE(parent.title, 'Unknown Album') as album,
               COALESCE(grandparent.title, 'Unknown Artist') as artist,
               m2.year,
               m2."index" as track_num,
               COALESCE(mi.duration, p.duration, 0) as dur,
               p.file
        FROM metadata_items m1
        JOIN metadata_items m2 ON m2.parent_id = m1.parent_id
        LEFT JOIN metadata_items parent ON parent.id = m2.parent_id
        LEFT JOIN metadata_items grandparent ON grandparent.id = parent.parent_id
        LEFT JOIN media_items mi ON mi.metadata_item_id = m2.id
        LEFT JOIN media_parts p ON p.media_item_id = mi.id
        WHERE m1.id = ? AND m2.metadata_type = 10
        ORDER BY COALESCE(m2.absolute_index, 1) ASC, COALESCE(m2."index", 0) ASC, m2.id ASC
        """
        tracks = []
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(sql, (track_id,))
            for row in cur.fetchall():
                t_id, title, album, artist, year, track_num, dur, fpath = row
                tracks.append(
                    Track(
                        id=t_id,
                        title=title,
                        album=album,
                        artist=artist,
                        year=year,
                        track_num=track_num,
                        duration_ms=dur or 0,
                        file_path=fpath or "",
                        format=self._format_from_path(fpath),
                    )
                )
        return tracks

    @staticmethod
    def _alpha_sort_key(text: str) -> str:
        """Normalize leading quotes/whitespace for clean A-Z alphabetical sorting."""
        return (text or "").lstrip(' "\'“”‘’').lower()

    def get_all_artists(self) -> List[Artist]:
        """Get all artists in the library sorted alphabetically A-Z."""
        sql = """
        SELECT ar.id,
               COALESCE(ar.title, 'Unknown Artist') as artist_name,
               COUNT(DISTINCT al.id) as album_count,
               COUNT(t.id) as track_count
        FROM metadata_items ar
        JOIN metadata_items al ON al.parent_id = ar.id AND al.metadata_type = 9
        JOIN metadata_items t ON t.parent_id = al.id AND t.metadata_type = 10
        WHERE ar.metadata_type = 8
        GROUP BY ar.id
        """
        artists = []
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(sql)
            for row in cur.fetchall():
                ar_id, name, al_count, tr_count = row
                artists.append(
                    Artist(
                        id=ar_id,
                        name=name,
                        album_count=al_count or 0,
                        track_count=tr_count or 0,
                    )
                )
        artists.sort(key=lambda a: (self._alpha_sort_key(a.name), a.id))
        return artists

    def get_artist_albums(self, artist_id: int) -> List[Album]:
        """Get all albums for a specific artist sorted alphabetically A-Z."""
        sql = """
        SELECT al.id,
               COALESCE(al.title, 'Unknown Album') as album_title,
               ar.id as artist_id,
               COALESCE(ar.title, 'Unknown Artist') as artist_name,
               al.year,
               COUNT(t.id) as track_count
        FROM metadata_items al
        LEFT JOIN metadata_items ar ON ar.id = al.parent_id
        JOIN metadata_items t ON t.parent_id = al.id AND t.metadata_type = 10
        WHERE al.parent_id = ? AND al.metadata_type = 9
        GROUP BY al.id
        """
        albums = []
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(sql, (artist_id,))
            for row in cur.fetchall():
                al_id, title, ar_id, ar_name, year, tr_count = row
                albums.append(
                    Album(
                        id=al_id,
                        title=title,
                        artist_id=ar_id,
                        artist=ar_name,
                        year=year,
                        track_count=tr_count or 0,
                    )
                )
        albums.sort(key=lambda al: (self._alpha_sort_key(al.title), al.id))
        return albums

    def get_all_albums(self) -> List[Album]:
        """Get all albums in the library sorted alphabetically A-Z."""
        sql = """
        SELECT al.id,
               COALESCE(al.title, 'Unknown Album') as album_title,
               ar.id as artist_id,
               COALESCE(ar.title, 'Unknown Artist') as artist_name,
               al.year,
               COUNT(t.id) as track_count
        FROM metadata_items al
        LEFT JOIN metadata_items ar ON ar.id = al.parent_id
        JOIN metadata_items t ON t.parent_id = al.id AND t.metadata_type = 10
        WHERE al.metadata_type = 9
        GROUP BY al.id
        """
        albums = []
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(sql)
            for row in cur.fetchall():
                al_id, title, ar_id, ar_name, year, tr_count = row
                albums.append(
                    Album(
                        id=al_id,
                        title=title,
                        artist_id=ar_id,
                        artist=ar_name,
                        year=year,
                        track_count=tr_count or 0,
                    )
                )
        albums.sort(key=lambda al: (self._alpha_sort_key(al.title), self._alpha_sort_key(al.artist), al.id))
        return albums

    def get_album_tracks(self, album_id: int) -> List[Track]:
        """Get all tracks inside a specific album, ordered by disc and track number."""
        sql = """
        SELECT m.id,
               COALESCE(m.title, 'Unknown Title') as title,
               COALESCE(parent.title, 'Unknown Album') as album,
               COALESCE(grandparent.title, 'Unknown Artist') as artist,
               COALESCE(m.year, parent.year) as year,
               m."index" as track_num,
               COALESCE(mi.duration, p.duration, 0) as dur,
               p.file
        FROM metadata_items m
        LEFT JOIN metadata_items parent ON parent.id = m.parent_id
        LEFT JOIN metadata_items grandparent ON grandparent.id = parent.parent_id
        LEFT JOIN media_items mi ON mi.metadata_item_id = m.id
        LEFT JOIN media_parts p ON p.media_item_id = mi.id
        WHERE m.parent_id = ? AND m.metadata_type = 10
        ORDER BY COALESCE(m.absolute_index, 1) ASC, COALESCE(m."index", 0) ASC, m.id ASC
        """
        tracks = []
        with self._get_connection() as conn:
            cur = conn.cursor()
            cur.execute(sql, (album_id,))
            for row in cur.fetchall():
                t_id, title, album, artist, year, track_num, dur, fpath = row
                tracks.append(
                    Track(
                        id=t_id,
                        title=title,
                        album=album,
                        artist=artist,
                        year=year,
                        track_num=track_num,
                        duration_ms=dur or 0,
                        file_path=fpath or "",
                        format=self._format_from_path(fpath),
                    )
                )
        return tracks

    def get_artist_tracks(self, artist_id: int) -> List[Track]:
        """Get all tracks by an artist across all their albums (albums A-Z, then track order)."""
        albums = self.get_artist_albums(artist_id)
        all_tracks = []
        for al in albums:
            all_tracks.extend(self.get_album_tracks(al.id))
        return all_tracks
