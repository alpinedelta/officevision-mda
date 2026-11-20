# Plex MDA Terminal Music Controller (`pmda`)

A lightweight, high-performance terminal music player and network playback controller designed specifically for Monochrome Display Adapter (MDA) and retro/VT100 terminals (80x25 / 90x43), as well as standard modern Linux/macOS terminals.

## Features

* **Full-Screen Monochrome TUI**: Pure ASCII borders and text bars, high-contrast reverse-video navigation, zero color bleeding, no Unicode glitches.
* **Instant Plex Search**: Fast SQLite search across your entire Plex music library (tracks, artists, and albums) directly from `com.plexapp.plugins.library.db`.
* **LAN Streaming**: Integrated HTTP audio streamer with byte-range (`206 Partial Content`) support. Streams lossless FLAC, MP3, AAC/M4A, WAV, and OGG directly from your local media paths with zero transcoding lag.
* **Network Target Control**: Discovers and casts directly to Google Cast / Chromecast / Chromecast Audio speakers, displays, streamers, and Denon/Marantz HEOS receivers on your local network.
* **Live Now Playing Display**: Real-time progress bar, track title, artist, album, elapsed/total time, volume percentage, and playback state.
* **Auto-Advancing Queue**: Playing a track automatically queues the remainder of the album and advances seamlessly when a song finishes.

---

## Quick Start

1. Scan your local network for Chromecast / HEOS speakers:
   ```bash
   pmda scan
   ```
2. Launch the interactive MDA TUI:
   ```bash
   pmda
   ```
3. Press `d` inside the TUI to connect to a discovered playback device (or press `r` inside the Device dialog to scan for new devices).

---

## Keyboard Shortcuts in TUI

| Key | Action |
| :--- | :--- |
| **`1` - `4`** / **`Tab`** | Switch Views: **`1` Search**, **`2` Artists (A-Z)**, **`3` Albums (A-Z)**, **`4` Queue** |
| **`Enter`** / **`Right`** | **Select** Artist (views their albums), **Select** Album (views its tracks), or **Play** Track |
| **`Bksp`** / **`Esc`** / **`b`** | **Go Back** up a level (from Album Tracks -> Artist's Albums -> All Artists, or Album Tracks -> All Albums) |
| **`[`** / **`]`** | **Jump** to Previous / Next starting letter (A-Z) in Artist or Album lists |
| **`PgUp`** / **`PgDn`** | Scroll up / down one full page (`Home`/`g` and `End`/`G` jump to top/bottom) |
| **`/`** | **Search** library (type search query, press **Enter** to search, **Esc** to cancel) |
| **`Space`** | **Toggle Play / Pause** (or select/play when idle) |
| **`s`** or **`x`** | **Stop** playback |
| **`n`** | **Next** track in queue |
| **`p`** | **Previous** track in queue |
| **`+`** / **`-`** | **Volume Up / Down** (in steps of 5%) |
| **`>`** / **`<`** | **Seek Forward / Backward** 10 seconds |
| **`i`** | **Play Next**: put highlighted track (or album/artist) next in queue ahead of other songs |
| **`I`** | **Play Album Next**: put entire album (or all tracks by artist) next in queue |
| **`a`** | **Add highlighted track** (or album) to end of queue |
| **`A`** | **Add entire album** (or all tracks by artist) to end of queue |
| **`c`** | **Clear** current playback queue |
| **`d`** | Open **Device Selection** menu (`1`-`9` / `Enter` to connect, **`r`** to scan network for new devices, `0` to disconnect) |
| **`?`** | Open **Help** overlay |
| **`q`** | **Quit** to terminal |

---

## Command Line Mode (Scriptable)

```bash
# Scan the network for Chromecast / Chromecast Audio / HEOS devices and save them
pmda scan

# List cached playback devices (instant)
pmda devices

# Search library
pmda search "Miles Davis"

# Browse artists and albums
pmda artists
pmda albums

# Play a track by ID on a discovered device ID
pmda play 1234 --device my_speaker

# Specify a custom Plex database path
pmda --db "/path/to/com.plexapp.plugins.library.db" search "Bach"

# View status and control playback
pmda status
pmda pause
pmda resume
pmda next
pmda stop
pmda volume 65
```
