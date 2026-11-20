#!/usr/bin/env bash
set -euo pipefail

# OfficeVision/400 MDA Suite Uninstaller
# By default removes launchers and program code while preserving user documents,
# spreadsheets, calendar events, and configs. Pass --purge to delete all data too.

PREFIX="${HOME}/.local"
PURGE=0

usage() {
    cat <<EOF
Usage: ./uninstall.sh [OPTIONS]

Options:
  --prefix DIR     Install prefix used during install (default: ~/.local)
  --purge          Also delete user data (documents, spreadsheets, configs, tokens, venv)
  -h, --help       Show this help message
EOF
}

while [[ $# -gt 0 ]]; do
    case "$1" in
        --prefix)
            PREFIX="$2"
            shift 2
            ;;
        --purge)
            PURGE=1
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

echo "Removing launchers from ${BIN_DIR}..."
rm -f "${BIN_DIR}/ov400" "${BIN_DIR}/lotus123" "${BIN_DIR}/ov400-cal" "${BIN_DIR}/pmda"

if [[ "${PURGE}" -eq 1 ]]; then
    echo "Purging application directories, virtualenv, and user data..."
    rm -rf "${OV_DIR}" "${PMDA_DIR}" "${PMDA_VENV}"
    echo "Full uninstall complete."
else
    echo "Removing program code (preserving user documents, spreadsheets, and configs)..."
    rm -f "${OV_DIR}/"*.py
    rm -rf "${OV_DIR}/__pycache__"
    rm -f "${PMDA_DIR}/"*.py "${PMDA_DIR}/requirements.txt" "${PMDA_DIR}/README.md"
    rm -rf "${PMDA_DIR}/players" "${PMDA_DIR}/__pycache__"
    echo "Uninstall complete. User data (if any) was preserved in:"
    echo "  ${OV_DIR}"
    echo "  ${PMDA_DIR}"
    echo "Re-run with --purge to remove user data and ${PMDA_VENV}."
fi
