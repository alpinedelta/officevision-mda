# Setting Up Google Drive, Docs, Sheets & Calendar Sync

OfficeVision/400 shares a single Google OAuth 2.0 configuration (`CFGDRV`) across the **Word Processor (Google Docs)**, **Lotus 1-2-3 (Google Sheets)**, and **Calendar (Google Calendar)**.

You can run all three applications completely offline in **Local Mode**, or connect them to your Google account using either **Direct Google API (recommended)** or **rclone** (Drive & Sheets only).

---

## Option A: Direct Google API (OAuth 2.0 — Recommended)

Works for **Google Drive / Docs**, **Google Sheets**, and **Google Calendar** with zero extra Python packages.

### Step 1: Create a Google Cloud Project & Enable APIs

1. Open the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project (e.g., `OfficeVision-Personal`).
3. Go to **APIs & Services -> Library** and enable:
   - **Google Drive API**
   - **Google Sheets API**
   - **Google Calendar API**

### Step 2: Configure the OAuth Consent Screen

1. Go to **APIs & Services -> OAuth consent screen** (or **Google Auth Platform -> Branding / Audience**).
2. Choose **External** (unless you have a Google Workspace org and prefer Internal).
3. Fill in the required App Name (e.g., `OfficeVision400`) and your email address.
4. Under **Test users**, add your own Google account email address (required while the app is in "Testing" mode).
5. Save your changes.

### Step 3: Create Desktop OAuth Credentials

1. Go to **APIs & Services -> Credentials -> Create Credentials -> OAuth client ID**.
2. Choose Application type: **Desktop app**.
3. Give it a name (e.g., `OV400 Terminal`) and click **Create**.
4. Click **Download JSON** and save the file to your computer (for example, `~/credentials.json`).

> **Security Note:** Never share your `credentials.json`, `gdrive_config.json`, or `gdrive_token.json` files online.

### Step 4: Connect Inside OfficeVision/400

1. Launch `ov400`, open **Documents** (Option `6`), **Calendar** (Option `1`), or **Lotus 1-2-3** (Option `10`), and press **`F10`** to open the **Google Setup (`CFGDRV`)** screen.
2. Load your credentials:
   - Enter the path to your downloaded JSON file (e.g., `~/credentials.json`) in the credentials file field and press **`F6`**, **OR**
   - Manually enter your **Client ID** and **Client Secret**.
3. Authorize your account:
   - **On a desktop Linux/macOS machine**: Press **`F7`** (Browser OAuth). A temporary listener starts on `http://localhost:8765` and opens your browser. Sign in and grant Drive, Sheets, and Calendar permissions.
   - **Over SSH, Telnet, or a headless/retro terminal**: Press **`F8`** (Telnet/Code OAuth). Copy the displayed Google authorization URL into a browser on any phone or PC, complete sign-in, and copy the redirected `http://localhost:8765/?code=...` URL (or just the `code` value) back into the `Auth Code` prompt.
4. Press **`F5`** to test sync, then **`F12`** to return!

Your credentials and OAuth tokens are stored locally at:
- `~/.local/share/officevision-mda/gdrive_config.json`
- `~/.local/share/officevision-mda/gdrive_token.json`

---

## Option B: `rclone` Backend (Word Processor & Lotus 1-2-3)

If you already use `rclone` with Google Drive and only need Drive/Docs + Sheets sync (note: Google Calendar requires Option A):

1. Install `rclone` (`sudo apt install rclone` or `brew install rclone`).
2. Configure a remote named `gdrive`:
   ```bash
   rclone config
   ```
   (Or press **`F9`** inside the `CFGDRV` screen in OfficeVision/400.)
3. Set the Rclone Remote field in `CFGDRV` to `gdrive:` (or your chosen remote name) and press `Enter` to save.

---

## Troubleshooting

- **"Google token lacks Calendar scope"**: Press `F10` -> `F7` (or `F8`) to re-authorize with all three scopes (Drive, Sheets, Calendar).
- **"Access blocked: app has not completed Google verification"**: Make sure your Google account email is listed under **Test users** in Google Cloud Console -> OAuth consent screen.
- **"API has not been used in project / disabled"**: Verify that Google Drive API, Google Sheets API, and Google Calendar API are all enabled in your Google Cloud project.
- **Changing data directory**: You can override `~/.local/share/officevision-mda` by exporting `OV400_DATA_DIR=/your/custom/path`.
