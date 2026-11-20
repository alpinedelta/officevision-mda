#!/usr/bin/env bash
set -euo pipefail

# OfficeVision/400 MDA Suite Installer
# Installs OfficeVision/400 (Shell, Word Processor, Lotus 1-2-3, Calendar) and Plex MDA (pmda).

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PREFIX="${HOME}/.local"
INSTALL_PLEX=1
SKIP_VENV=0

usage() {
    cat <<EOF
Usage: ./install.sh [OPTIONS]

Options:
  --prefix DIR     Install prefix (default: ~/.local)
  --no-plex        Install only OfficeVision/400 (skip Plex MDA music player)
  --skip-venv      Copy Plex MDA files without creating Python virtualenv / pip install
  -h, --help       Show this help message
EOF
}

while [[ $# -gt 0 ]]; do
    case "$1" in
        --prefix)
            PREFIX="$2"
            shift 2
            ;;
        --no-plex)
            INSTALL_PLEX=0
            shift
            ;;
        --skip-venv)
            SKIP_VENV=1
            shift
            ;;
        -h|--help)
            usage
            exit 0
            ;;
        *)
            echo "Unknown option: $1" >&2
            usage >&2
            exit 1
            ;;
    esac
done

BIN_DIR="${PREFIX}/bin"
SHARE_DIR="${PREFIX}/share"
OV_DIR="${SHARE_DIR}/officevision-mda"
PMDA_DIR="${SHARE_DIR}/plex-mda"
PMDA_VENV="${SHARE_DIR}/plex-mda-venv"

echo "================================================================"
echo "  IBM OfficeVision/400 MDA Terminal Suite - Installer"
echo "================================================================"
echo "  Install prefix : ${PREFIX}"
echo "  Launchers      : ${BIN_DIR}"
echo "  OfficeVision   : ${OV_DIR}"
if [[ "${INSTALL_PLEX}" -eq 1 ]]; then
    echo "  Plex MDA       : ${PMDA_DIR}"
    echo "  Plex Venv      : ${PMDA_VENV}"
fi
echo ""

# 1. Check Python 3
if ! command -v python3 >/dev/null 2>&1; then
    echo "ERROR: python3 is required but was not found in PATH." >&2
    exit 1
fi

PY_VER=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
PY_OK=$(python3 -c 'import sys; print(1 if sys.version_info >= (3, 9) else 0)')
if [[ "${PY_OK}" != "1" ]]; then
    echo "WARNING: Python 3.9+ is recommended (detected Python ${PY_VER})."
fi

# Verify curses module is available
if ! python3 -c "import curses" 2>/dev/null; then
    echo "ERROR: Python's 'curses' module is not available in this Python installation." >&2
    exit 1
fi

# 2. Create directories
mkdir -p "${BIN_DIR}" "${OV_DIR}"

# 3. Install OfficeVision/400 code (preserves existing user documents, spreadsheets, configs, and tokens)
echo "[1/3] Installing OfficeVision/400 (Shell, Word Processor, Lotus 1-2-3, Calendar)..."
cp "${SCRIPT_DIR}/officevision-mda/"*.py "${OV_DIR}/"
chmod +x "${OV_DIR}/main.py" "${OV_DIR}/lotus123.py" "${OV_DIR}/calapp.py"

for launcher in ov400 lotus123 ov400-cal; do
    sed "s|\$HOME/.local/share/officevision-mda|${OV_DIR}|g" "${SCRIPT_DIR}/bin/${launcher}" > "${BIN_DIR}/${launcher}"
    chmod +x "${BIN_DIR}/${launcher}"
done
echo "      Installed launchers: ov400, lotus123, ov400-cal"

# 4. Install Plex MDA (optional)
if [[ "${INSTALL_PLEX}" -eq 1 ]]; then
    echo "[2/3] Installing Plex MDA Music Controller (pmda)..."
    mkdir -p "${PMDA_DIR}/players"
    cp "${SCRIPT_DIR}/plex-mda/"*.py "${PMDA_DIR}/"
    cp "${SCRIPT_DIR}/plex-mda/players/"*.py "${PMDA_DIR}/players/"
    cp "${SCRIPT_DIR}/plex-mda/requirements.txt" "${PMDA_DIR}/"
    cp "${SCRIPT_DIR}/plex-mda/README.md" "${PMDA_DIR}/"
    chmod +x "${PMDA_DIR}/main.py"

    sed -e "s|\$HOME/.local/share/plex-mda-venv|${PMDA_VENV}|g" \
        -e "s|\$HOME/.local/share/plex-mda|${PMDA_DIR}|g" \
        "${SCRIPT_DIR}/bin/pmda" > "${BIN_DIR}/pmda"
    chmod +x "${BIN_DIR}/pmda"

    if [[ "${SKIP_VENV}" -eq 0 ]]; then
        echo "      Setting up Python virtual environment for Plex MDA..."
        if python3 -m venv "${PMDA_VENV}" 2>/dev/null; then
            echo "      Installing PyChromecast & zeroconf..."
            if "${PMDA_VENV}/bin/pip" install --quiet --upgrade pip && \
               "${PMDA_VENV}/bin/pip" install --quiet -r "${PMDA_DIR}/requirements.txt"; then
                echo "      Plex MDA dependencies installed successfully."
            else
                echo "      WARNING: pip install failed (no internet or missing build tools)."
                echo "      Retry later with: ${PMDA_VENV}/bin/pip install -r ${PMDA_DIR}/requirements.txt"
            fi
        else
            echo "      WARNING: Could not create virtualenv (on Debian/Ubuntu, run: sudo apt install python3-venv)."
            echo "      Retry later with:"
            echo "        python3 -m venv ${PMDA_VENV}"
            echo "        ${PMDA_VENV}/bin/pip install -r ${PMDA_DIR}/requirements.txt"
        fi
    else
        echo "      Skipped virtualenv setup (--skip-venv)."
    fi
else
    echo "[2/3] Skipping Plex MDA (--no-plex)."
fi

# 5. Check optional tools & PATH
echo "[3/3] Checking environment and optional integrations..."
echo ""
check_opt() {
    local cmd="$1"
    local desc="$2"
    if command -v "${cmd}" >/dev/null 2>&1; then
        printf "  [OK]       %-10s - %s\n" "${cmd}" "${desc}"
    else
        printf "  [OPTIONAL] %-10s - %s (not installed)\n" "${cmd}" "${desc}"
    fi
}

check_opt "aspell" "Spell checker for Word Processor (F4)"
check_opt "alpine" "Mail client for OV/400 Option 2"
check_opt "rclone" "Alternative Google Drive/Sheets sync backend"

echo ""
if [[ ":${PATH}:" != *":${BIN_DIR}:"* ]]; then
    echo "NOTE: ${BIN_DIR} is not currently in your PATH."
    echo "Add the following line to your ~/.bashrc or ~/.zshrc, then restart your shell:"
    echo ""
    echo "    export PATH=\"${BIN_DIR}:\$PATH\""
    echo ""
fi

echo "================================================================"
echo "  Installation complete!"
echo "================================================================"
echo "  Launch commands:"
echo "    ov400       - IBM OfficeVision/400 Main Environment"
echo "    lotus123    - Lotus 1-2-3 Spreadsheet (standalone)"
echo "    ov400-cal   - OfficeVision/400 Calendar (standalone)"
if [[ "${INSTALL_PLEX}" -eq 1 ]]; then
    echo "    pmda        - Plex MDA Music Player (run 'pmda scan' first)"
fi
echo ""
echo "  See README.md, SETUP_GOOGLE.md, and SETUP_PLEX.md for configuration."
