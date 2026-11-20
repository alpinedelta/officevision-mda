#!/usr/bin/env python3
"""Main entry point for Plex MDA Music Player & Controller."""

import argparse
import sys
import time

from config import discover_devices, load_devices
from db import MusicDB
from streamer import AudioStreamer


def _print_devices(devices, header="Available Playback Devices (Cached):"):
    print(f"\n{header}")
    print("-" * 88)
    print(f" {'#':<3} {'DEVICE ID':<20} {'NAME':<24} {'MODEL / ALIAS':<22} {'HOST':<15}")
    print("-" * 88)
    for i, d in enumerate(devices, 1):
        alias = d.get("alias") or d.get("model") or d.get("type", "")
        host = f"{d.get('host', '')}:{d.get('port', '')}"
        print(f" [{i}] {d['id']:<20} {d['name'][:23]:<24} {alias[:21]:<22} {host:<15}")
    print("-" * 88)


def main():
    parser = argparse.ArgumentParser(
        description="Plex MDA Terminal Music Controller",
        epilog="Run with no arguments to launch the full-screen MDA interactive TUI.",
    )
    parser.add_argument(
        "--db",
        dest="db_path",
        default=None,
        help="Path to Plex SQLite database (com.plexapp.plugins.library.db). Overrides PLEX_DB_PATH.",
    )
    subparsers = parser.add_subparsers(dest="command", help="Command to run (or leave empty for TUI)")

    # CLI subcommands
    dev_p = subparsers.add_parser("devices", help="List saved playback devices (or scan with --scan)")
    dev_p.add_argument("-s", "--scan", action="store_true", help="Scan network for new Chromecast/HEOS devices and save")

    subparsers.add_parser("scan", help="Scan network for new Chromecast/HEOS devices and update saved list")

    search_p = subparsers.add_parser("search", help="Search Plex music library")
    search_p.add_argument("query", help="Search term (artist, album, or track)")
    search_p.add_argument("-l", "--limit", type=int, default=20, help="Max results")

    artists_p = subparsers.add_parser("artists", help="Browse all artists (A-Z) or albums for an artist ID")
    artists_p.add_argument("artist_id", type=int, nargs="?", help="Optional artist ID to list their albums")

    albums_p = subparsers.add_parser("albums", help="Browse all albums (A-Z) or tracks for an album ID")
    albums_p.add_argument("album_id", type=int, nargs="?", help="Optional album ID to list its tracks")

    play_p = subparsers.add_parser("play", help="Play a track by ID")
    play_p.add_argument("track_id", type=int, help="Plex track ID")
    play_p.add_argument(
        "-d",
        "--device",
        default=None,
        help="Target device ID (defaults to last selected or first saved device)",
    )

    subparsers.add_parser("status", help="Get now playing status")
    subparsers.add_parser("pause", help="Pause playback")
    subparsers.add_parser("resume", help="Resume playback")
    subparsers.add_parser("stop", help="Stop playback")
    subparsers.add_parser("next", help="Skip to next track")
    subparsers.add_parser("prev", help="Go to previous track")

    vol_p = subparsers.add_parser("volume", help="Set volume level (0-100)")
    vol_p.add_argument("level", type=int, help="Volume percentage (0-100)")

    args = parser.parse_args()

    if args.command in ["devices", "scan"]:
        do_scan = (args.command == "scan") or getattr(args, "scan", False)
        if do_scan:
            print("Scanning local network for Chromecast & HEOS playback devices...")
            devices, new_names = discover_devices()
            if new_names:
                print(f"Discovered {len(new_names)} new device(s): {', '.join(new_names)}")
            else:
                print("Scan complete (no new devices found; existing devices updated).")
            _print_devices(devices, header=f"Saved Playback Devices ({len(devices)} total):")
        else:
            devices = load_devices()
            if not devices:
                print("\nNo saved playback devices yet. Run 'pmda scan' to discover Chromecast/HEOS devices on your LAN.")
            else:
                _print_devices(devices, header=f"Saved Playback Devices ({len(devices)} cached — run 'pmda scan' to rescan):")
        return

    # Initialize core services
    try:
        db = MusicDB(db_path=args.db_path)
    except Exception as e:
        print(f"Error accessing Plex database: {e}", file=sys.stderr)
        sys.exit(1)

    if args.command == "search":
        tracks = db.search_tracks(args.query, limit=args.limit)
        print(f"\nFound {len(tracks)} tracks for '{args.query}':")
        print("-" * 88)
        print(f" {'ID':<6} {'#':<3} {'TRACK TITLE':<27} {'ARTIST':<18} {'ALBUM':<22} {'TIME':>5}")
        print("-" * 88)
        for t in tracks:
            num = f"{t.track_num:2d}" if t.track_num else "  "
            title = (t.title[:25] + "..") if len(t.title) > 27 else t.title
            artist = (t.artist[:16] + "..") if len(t.artist) > 18 else t.artist
            album = (t.album[:20] + "..") if len(t.album) > 22 else t.album
            print(f" [{t.id:<5}] {num}  {title:<27} {artist:<18} {album:<22} {t.duration_str:>5}")
        print("-" * 88)
        return

    if args.command == "artists":
        if args.artist_id is not None:
            albums = db.get_artist_albums(args.artist_id)
            print(f"\nFound {len(albums)} albums for Artist ID {args.artist_id}:")
            print("-" * 88)
            print(f" {'ALBUM ID':<8} {'ALBUM TITLE (A-Z)':<42} {'ARTIST':<22} {'YEAR':<6} {'TRACKS':>6}")
            print("-" * 88)
            for al in albums:
                title = (al.title[:40] + "..") if len(al.title) > 42 else al.title
                artist = (al.artist[:20] + "..") if len(al.artist) > 22 else al.artist
                yr = str(al.year) if al.year else "-"
                print(f" [{al.id:<6}] {title:<42} {artist:<22} {yr:<6} {al.track_count:>6}")
            print("-" * 88)
        else:
            artists = db.get_all_artists()
            print(f"\nAll Artists ({len(artists)} A-Z):")
            print("-" * 88)
            print(f" {'ARTIST ID':<9} {'ARTIST NAME (A-Z)':<50} {'ALBUMS':>12} {'TRACKS':>12}")
            print("-" * 88)
            for ar in artists:
                name = (ar.name[:48] + "..") if len(ar.name) > 50 else ar.name
                print(f" [{ar.id:<7}] {name:<50} {ar.album_count:>12} {ar.track_count:>12}")
            print("-" * 88)
        return

    if args.command == "albums":
        if args.album_id is not None:
            tracks = db.get_album_tracks(args.album_id)
            print(f"\nFound {len(tracks)} tracks in Album ID {args.album_id}:")
            print("-" * 88)
            print(f" {'ID':<6} {'#':<3} {'TRACK TITLE':<32} {'ARTIST':<20} {'ALBUM':<17} {'TIME':>5}")
            print("-" * 88)
            for t in tracks:
                num = f"{t.track_num:2d}" if t.track_num else "  "
                title = (t.title[:30] + "..") if len(t.title) > 32 else t.title
                artist = (t.artist[:18] + "..") if len(t.artist) > 20 else t.artist
                album = (t.album[:15] + "..") if len(t.album) > 17 else t.album
                print(f" [{t.id:<5}] {num}  {title:<32} {artist:<20} {album:<17} {t.duration_str:>5}")
            print("-" * 88)
        else:
            albums = db.get_all_albums()
            print(f"\nAll Albums ({len(albums)} A-Z):")
            print("-" * 88)
            print(f" {'ALBUM ID':<8} {'ALBUM TITLE (A-Z)':<42} {'ARTIST':<22} {'YEAR':<6} {'TRACKS':>6}")
            print("-" * 88)
            for al in albums:
                title = (al.title[:40] + "..") if len(al.title) > 42 else al.title
                artist = (al.artist[:20] + "..") if len(al.artist) > 22 else al.artist
                yr = str(al.year) if al.year else "-"
                print(f" [{al.id:<6}] {title:<42} {artist:<22} {yr:<6} {al.track_count:>6}")
            print("-" * 88)
        return

    # For commands that control playback or launch TUI, start streamer & manager
    try:
        from players.manager import PlayerManager
        from tui import MdaTui
    except ModuleNotFoundError as e:
        print(
            f"Missing dependency ({e}). Install Plex MDA dependencies with:\n"
            "  python3 -m venv ~/.local/share/plex-mda-venv\n"
            "  ~/.local/share/plex-mda-venv/bin/pip install -r ~/.local/share/plex-mda/requirements.txt",
            file=sys.stderr,
        )
        sys.exit(1)

    streamer = AudioStreamer(db)
    manager = PlayerManager(db, streamer)

    try:
        if args.command == "play":
            track = db.get_track(args.track_id)
            if not track:
                print(f"Track ID {args.track_id} not found.", file=sys.stderr)
                sys.exit(1)
            dev_id = args.device or manager.current_device_id
            if not dev_id and manager.devices:
                dev_id = manager.devices[0]["id"]
            if not dev_id:
                print("No playback device configured. Run 'pmda scan' first to discover devices.", file=sys.stderr)
                sys.exit(1)
            print(f"Connecting to {dev_id}...")
            if not manager.connect_device(dev_id, blocking=True):
                print(f"Failed to connect to {dev_id}.", file=sys.stderr)
                sys.exit(1)
            print(f"Connected! Playing '{track.title}' by {track.artist} on {manager.get_current_device_name()}... (Press Ctrl+C to stop)")
            manager.play_track_now(track, replace_queue=True)
            try:
                while True:
                    time.sleep(1)
                    st = manager.get_status()
                    sys.stdout.write(f"\rStatus: [{st.state}] {st.progress_bar(20)} {st.time_display()} Vol: {st.volume}%  ")
                    sys.stdout.flush()
                    if st.state in ["IDLE", "STOPPED"] and st.current_time > 0:
                        break
            except KeyboardInterrupt:
                print("\nPlayback stopped.")
            print()
            return

        if args.command == "status":
            st = manager.get_status()
            print(f"\nDevice : {manager.get_current_device_name()}")
            print(f"State  : [{st.state}]")
            print(f"Track  : {st.track_title or '(None)'}")
            print(f"Artist : {st.artist}")
            print(f"Album  : {st.album}")
            print(f"Time   : {st.progress_bar(30)} {st.time_display()}")
            print(f"Volume : {st.volume}% {'(Muted)' if st.muted else ''}\n")
            return

        if args.command == "pause":
            manager.active_player.pause() if manager.active_player else None
            print("Playback paused.")
            return

        if args.command == "resume":
            manager.active_player.resume() if manager.active_player else None
            print("Playback resumed.")
            return

        if args.command == "stop":
            manager.stop()
            print("Playback stopped.")
            return

        if args.command == "next":
            manager.next_track()
            print("Skipped to next track.")
            return

        if args.command == "prev":
            manager.prev_track()
            print("Previous track.")
            return

        if args.command == "volume":
            manager.set_volume(args.level)
            print(f"Volume set to {args.level}%.")
            return

        # No command passed -> Launch full-screen Interactive MDA TUI
        tui = MdaTui(db, manager)
        tui.run()

    finally:
        manager.shutdown()
        streamer.stop()


if __name__ == "__main__":
    main()
