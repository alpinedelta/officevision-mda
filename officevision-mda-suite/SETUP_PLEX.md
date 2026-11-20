# Setting Up Plex MDA Music Controller (`pmda`)

`pmda` reads your local Plex Media Server SQLite database directly (read-only) for instant search and A-Z browsing, and runs a lightweight local HTTP server with byte-range (`206`) streaming to cast lossless FLAC/MP3/AAC/WAV/OGG audio to Chromecast and Denon/Marantz HEOS devices on your LAN.

---

## 1. Prerequisites

1. **Plex Media Server** installed on the same machine (or with its SQLite database and media paths accessible locally at the exact paths stored in the Plex DB).
2. **Python dependencies** installed (handled automatically by `./install.sh`):
   - `PyChromecast>=14.0.0`
   - `zeroconf>=0.130.0`

If you skipped venv creation during install, set it up manually:
```bash
python3 -m venv ~/.local/share/plex-mda-venv
~/.local/share/plex-mda-venv/bin/pip install -r ~/.local/share/plex-mda/requirements.txt
```

---

## 2. Plex Database Path & Permissions

By default, `pmda` automatically checks standard Plex Media Server database locations:
- **Linux**: `/var/lib/plexmediaserver/Library/Application Support/Plex Media Server/Plug-in Support/Databases/com.plexapp.plugins.library.db`
- **macOS**: `~/Library/Application Support/Plex Media Server/Plug-in Support/Databases/com.plexapp.plugins.library.db`

### Granting Read Permissions on Linux

On Linux, `/var/lib/plexmediaserver` is typically owned by the `plex` user/group. To let your user read the database (read-only) without `sudo`:

```bash
# Add your user to the plex group (log out and back in for group changes to apply)
sudo usermod -aG plex "$USER"
```

If directory permissions still block reading, verify group read/execute access on the database path:
```bash
ls -la "/var/lib/plexmediaserver/Library/Application Support/Plex Media Server/Plug-in Support/Databases/com.plexapp.plugins.library.db"
```

### Using a Custom Database Path

If your Plex database lives elsewhere (e.g., Docker volume mount, Snap, Flatpak, or NAS):

- **Environment variable** (add to `~/.bashrc`):
  ```bash
  export PLEX_DB_PATH="/path/to/com.plexapp.plugins.library.db"
  ```
- **Or in `~/.local/share/plex-mda/settings.json`**:
  ```json
  {
    "plex_db_path": "/path/to/com.plexapp.plugins.library.db",
    "stream_port": 8888
  }
  ```
- **Or per-command flag**:
  ```bash
  pmda --db "/path/to/com.plexapp.plugins.library.db"
  ```

> **Note:** Your user must also have read access to the actual audio files at the file paths recorded inside the Plex database.

---

## 3. Firewall Configuration (LAN Streaming)

When playing a track, `pmda` starts a local HTTP audio server on **TCP port `8888`** (or next available port up to `8899`) so your Chromecast / HEOS speaker can stream the audio file directly over your LAN.

If you run a firewall (such as `ufw` on Ubuntu/Mint), allow LAN access to port `8888`:

```bash
# Replace 192.168.1.0/24 with your own LAN subnet
sudo ufw allow from 192.168.1.0/24 to any port 8888:8899 proto tcp
```

You can change the default streaming port via `PMDA_STREAM_PORT=8888` or `"stream_port"` in `settings.json`.

---

## 4. Discovering Your Playback Devices

Run a LAN scan to detect Google Cast / Chromecast / Chromecast Audio speakers, smart displays, TVs, and Denon/Marantz HEOS receivers:

```bash
pmda scan
```

Or launch `pmda`, press **`d`** to open the Device dialog, and press **`r`** to scan.

Discovered devices are saved locally in `~/.local/share/plex-mda/devices.json` so future startups connect instantly without waiting for mDNS.

### Manually Adding a Device (Optional)

If mDNS / discovery is blocked on your network, you can manually create `~/.local/share/plex-mda/devices.json`:

```json
[
  {
    "id": "living_room_speaker",
    "name": "Living Room Speaker",
    "alias": "Chromecast Audio",
    "type": "cast",
    "cast_type": "audio",
    "host": "192.168.1.50",
    "port": 8009
  },
  {
    "id": "av_receiver",
    "name": "Denon Receiver",
    "alias": "HEOS AVR",
    "type": "heos",
    "host": "192.168.1.60",
    "port": 1255
  }
]
```

---

## 5. Troubleshooting

- **"Plex database not found"**: Check `PLEX_DB_PATH` or pass `--db /path/to/com.plexapp.plugins.library.db`.
- **"Cannot read Plex database"**: Ensure your user is in the `plex` group (and that you logged out/in after `usermod -aG plex`).
- **Device connects, track shows `BUFFERING` then `IDLE`**: Check that your firewall allows incoming TCP on port `8888` from your speaker's IP, and that your user can read the audio files referenced by Plex.
- **No devices found on scan**: Ensure your computer and speakers are on the same LAN/VLAN and not isolated by AP client isolation.
