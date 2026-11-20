# IBM OfficeVision/400 MDA Terminal Suite

A retro **IBM AS/400 OfficeVision/400 (OV/400)** monochrome terminal productivity & media suite built in Python (`curses`), tailored for **90x43 Monochrome Display Adapter (MDA) / VT100 / Telnet** terminals while working smoothly in any modern Linux or macOS terminal emulator.

## Included Applications

1. **OfficeVision/400 Shell (`ov400`)**
   - Authentic IBM AS/400 `QDSIGNON` Sign-On screen and `OFCMAIN` Main Menu with live clock and monthly calendar sidebar.
   - Unified launcher for all built-in sub-applications, plus optional external integrations (Alpine Mail, Antigravity CLI, Plex MDA).
   - **Zero external Python dependencies** (100% Python 3 standard library).

2. **Word Processor & Document Manager (OV/400 Options 5, 6, and 9)**
   - **WRKDOC**: Work with Documents in Folders with local library and optional Google Drive / Google Docs sync.
   - **EDTDOC**: DisplayWrite/400-style fullscreen word processor featuring a live margin ruler, soft word-wrap, paragraph reflow (`F8`), line centering (`F6`), bold/underline/reverse MDA markup (`Ctrl+B`, `Ctrl+U`, `Ctrl+R`, toggle live preview with `F11`), multi-level Undo/Redo (`Ctrl+Z` / `Ctrl+Y`), Find & Replace (`F5`), and interactive `aspell` spell checking (`F4`).

3. **Lotus 1-2-3 Spreadsheet (OV/400 Option 10 or standalone `lotus123`)**
   - Authentic 3-line Control Panel with reverse-video mode indicator (`READY`, `VALUE`, `LABEL`, `EDIT`, `MENU`, `GRAPH`, etc.) and classic two-line `/` Slash Command Menu.
   - Multi-sheet Release 3 tabs, dual Lotus 1-2-3 (`@SUM(A1..B5)`, `+B2*C2`, `' " ^ \`) and Google Sheets (`=SUM(A1:B5)`) formula engine.
   - Monochrome ASCII bar/line charts (`/Graph`), CSV and multi-sheet `.xlsx` import/export, print to OV/400 Word Processor, and two-way Google Sheets sync.

4. **Calendar (`WRKCAL` — OV/400 Option 1 or standalone `ov400-cal`)**
   - Fullscreen Month Grid (`WRKCAL`), Day Schedule (`DSPDAY`), Event Creation/Editing (`ADDEVT` supporting timed, all-day, multi-day, and recurring events), Calendar Selector (`SELCAL`), and Upcoming Agenda Search (`DSPAGD`).
   - Works completely offline in Local Mode or with full two-way Google Calendar API v3 synchronization.

5. **Plex MDA Music Controller (OV/400 Option 8 or standalone `pmda`)**
   - Monochrome TUI and scriptable CLI that searches your local Plex Media Server SQLite music library (`com.plexapp.plugins.library.db`) and streams lossless FLAC/MP3/AAC/OGG/WAV directly over LAN HTTP with byte-range (`206`) seeking.
   - Automatic discovery and control of Google Cast / Chromecast / Chromecast Audio speakers, displays, TVs, and Denon/Marantz HEOS receivers.

---

## Requirements

### Core Suite (`ov400`, Word Processor, Lotus 1-2-3, Calendar)
- **OS**: Linux or macOS (or Windows via WSL2)
- **Python**: Python 3.9 or newer (with standard `curses` support)
- **No pip packages required** for OfficeVision/400, Word Processor, Lotus 1-2-3, or Calendar.

### Optional Integrations
- **`aspell`** (plus English dictionary `aspell-en`) — Interactive spell checker (`F4`) in Word Processor
- **`alpine`** — Mail client launched by OfficeVision Option 2 (`Mail`)
- **`rclone`** — Optional alternative backend for Google Drive / Sheets sync
- **Plex MDA (`pmda`)** — Requires `python3-venv` (to install `PyChromecast` and `zeroconf`) and read access to a local Plex Media Server database + media files.

On Debian / Ubuntu / Linux Mint, you can install the optional system tools with:
```bash
sudo apt update
sudo apt install python3 python3-venv aspell aspell-en alpine rclone
```

On macOS (via Homebrew):
```bash
brew install python aspell alpine rclone
```

---

## Quick Installation

1. Extract the package (if using the `.tar.gz` or `.zip` archive) and enter the directory:
   ```bash
   tar -xzf officevision-mda-suite.tar.gz
   cd officevision-mda-suite
   ```
2. Run the installer:
   ```bash
   ./install.sh
   ```
   - Install **without** Plex MDA: `./install.sh --no-plex`
   - Install to a custom prefix: `./install.sh --prefix /opt/ov400`
   - Skip creating the Plex virtualenv (e.g., offline install): `./install.sh --skip-venv`

3. Ensure `~/.local/bin` is in your `PATH` (add to `~/.bashrc` or `~/.zshrc` if not already present):
   ```bash
   export PATH="$HOME/.local/bin:$PATH"
   ```
   Then reload your shell (`source ~/.bashrc` or open a new terminal).

4. Launch OfficeVision/400:
   ```bash
   ov400
   ```
   *(Tip: pass `ov400 --skip-login` to jump straight to the `OFCMAIN` menu!)*

---

## Launching Applications

| Command | Application |
| :--- | :--- |
| `ov400` | Full OfficeVision/400 environment (Sign-On screen -> Main Menu) |
| `ov400 --skip-login` | OfficeVision/400 Main Menu directly |
| `lotus123` | Lotus 1-2-3 Spreadsheet standalone |
| `ov400-cal` | OfficeVision/400 Calendar standalone |
| `pmda` | Plex MDA Music Player TUI (or CLI subcommands like `pmda scan`, `pmda search`) |

On first launch:
- **Word Processor** automatically creates a starter guide (`WELCOME_OV400.md`).
- **Lotus 1-2-3** automatically creates a multi-tab sample budget workbook (`Q4_BUDGET`).
- **Calendar** starts immediately in Local Mode — press `F10` anytime to connect Google.
- **Plex MDA** starts with an empty device list — run `pmda scan` or press `d` then `r` inside the TUI to discover your speakers/receivers.

---

## Cloud & Media Configuration Guides

- **Google Drive, Docs, Sheets & Calendar Sync**: See [SETUP_GOOGLE.md](SETUP_GOOGLE.md)
- **Plex Music Library & Chromecast/HEOS Streaming**: See [SETUP_PLEX.md](SETUP_PLEX.md)

---

## Keyboard Reference

### OfficeVision/400 Main Menu (`ov400`)
- **Type an option number** (`1`-`10`, `90`) or command (`CAL`, `WP`, `DOCS`, `123`, `LOTUS`, `PLEX`, `MAIL`, `SIGNOFF`) and press `Enter`, or use `Up`/`Down` arrows + `Enter`.
- `[` / `]` (or `F19`/`F20`): Previous / Next month on sidebar calendar (`t` = Today)
- `F1` or `?`: Help | `F9`: Retrieve previous command | `F12`: Clear command | `F3`: Exit

### Word Processor (`EDTDOC` & `WRKDOC`)
- `F1`: Help | `F2` / `Ctrl+S`: Save (& sync to Google Docs) | `F3`: Save & Exit
- `F4`: Spell check (`aspell`) | `F5` / `Ctrl+F`: Find & Replace | `F6` / `Ctrl+E`: Center line
- `F8` / `Ctrl+J`: Reflow paragraph | `F10` (in `WRKDOC`): Google Drive Setup (`CFGDRV`)
- `F11`: Toggle MDA formatted preview vs raw Markdown | `F12`: Cancel / Back
- `Ctrl+B` / `Ctrl+U` / `Ctrl+R`: Bold / Underline / Reverse highlight | `Ctrl+Z` / `Ctrl+Y`: Undo / Redo

### Lotus 1-2-3 (`lotus123`)
- `/`: Open classic two-line Lotus command menu (`W` Worksheet, `R` Range, `C` Copy, `M` Move, `F` File, `P` Print, `G` Graph, `D` Data, `S` Sheets, `Q` Quit)
- `F1`: Help | `F2`: Edit cell | `F4`: Workbook / Google Sheets Browser | `F9`: Recalculate | `F10`: Google Sheets Sync Menu
- `Ctrl+PgUp` / `Ctrl+PgDn` (or `[` / `]`): Switch worksheet tabs

### Calendar (`ov400-cal` / `WRKCAL`)
- Arrows: Move date | `Enter`: Open Day Schedule (`DSPDAY`) | `N`: New event | `E`: Edit | `D`: Delete
- `PgUp` / `PgDn` (or `[` / `]`): Previous / Next month | `T`: Today | `G`: Go to date
- `A` or `/`: Upcoming Agenda & Search | `C`: Select visible Google Calendars | `F5`: Sync now | `F10`: Google OAuth Setup | `F3`/`F12`: Exit

### Plex MDA (`pmda`)
- `1`-`4` / `Tab`: Search, Artists (A-Z), Albums (A-Z), Queue | `/`: Search | `[` / `]`: Jump letter A-Z
- `Enter`: Drill down / Play | `Bksp` / `Esc` / `b`: Back | `Space`: Play/Pause | `s`/`x`: Stop
- `n` / `p`: Next / Previous track | `+` / `-`: Volume | `<` / `>`: Seek +/-10s
- `i` / `I`: Play track / album next | `a` / `A`: Append to queue | `c`: Clear queue
- `d`: Device dialog (`r` scans LAN, `1`-`9` connects, `0` disconnects) | `?`: Help | `q`: Quit

---

## Optional: Retro 90x43 MDA / VT100 / Telnet Terminal Setup

When `TERM` is `vt100`, `wsvt25`, `ansi`, `dumb`, or when `MDA=1` is exported, all launchers automatically configure a **90x43** monochrome layout with clean VT100 scroll-region resets.

If you connect over a retro serial terminal, Telnet, or ESP32/FabGL terminal, you can optionally add this snippet to your `~/.bashrc`:

```bash
# Optional: 90x43 Monochrome MDA / VT100 Telnet session configuration
if [[ "$TERM" == "wsvt25" || "$TERM" == "vt100" || "$TERM" == "ansi" || "$MDA" == "1" ]]; then
    export TERM=vt100
    export MDA=1
    stty rows 43 cols 90 2>/dev/null
    export LINES=43
    export COLUMNS=90
    printf '\0337\033[?6l\033[1;43r\0338'
    alias ls='ls --color=never'
    alias grep='grep --color=never'
fi
```

---

## Data Locations & Uninstalling

All user data is stored locally under your home directory (never shared with the repository):
- `~/.local/share/officevision-mda/` — Documents, spreadsheets, calendar cache, and Google OAuth credentials/tokens
- `~/.local/share/plex-mda/` — Saved playback devices (`devices.json`) and preferences (`settings.json`)

To uninstall:
```bash
./uninstall.sh          # Removes launchers & code, keeps your documents/configs
./uninstall.sh --purge  # Removes everything including documents, configs, and venv
```
