#!/usr/bin/env python3
r"""
Lotus 1-2-3 & Google Sheets Compatible Spreadsheet Engine for IBM OfficeVision/400.

Pure Python 3 standard library (zero external dependencies):
  - Multi-sheet Workbook / Worksheet / Cell model with undo/redo snapshots
  - Recursive-descent formula parser supporting BOTH Lotus 1-2-3 syntax
    (@SUM(A1..B5), +B2*C2, label prefixes ' " ^ \) AND Google Sheets syntax
    (=SUM(A1:B5), Sheet2!A1, $A$1 absolute references)
  - Bidirectional Lotus <-> Google Sheets formula conversion
  - 55+ built-in mathematical, statistical, logical, lookup, text, financial, date functions
  - Relative/absolute formula shifting for Copy, Move, Insert/Delete Row & Column
  - Pure-Python CSV, multi-sheet XLSX (with formulas!), and ASCII report import/export
"""

import copy
import csv
import datetime
import io
import json
import math
import random
import re
import statistics
import xml.etree.ElementTree as ET
import zipfile
from typing import Any, Dict, List, Optional, Set, Tuple, Union


# =============================================================================
# Column / Cell Reference Helpers
# =============================================================================

MAX_COLS = 256   # A .. IV (classic Lotus 1-2-3 Release 2/3 width!)
MAX_ROWS = 8192  # Classic Lotus 1-2-3 Release 2.2 row limit (8192 rows)

CELL_REF_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
SHEET_CELL_RE = re.compile(
    r"^(?:(?:'([^']+)'|([A-Za-z0-9_ ]+))!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)$"
)
LOTUS_RANGE_RE = re.compile(
    r"(\$?[A-Za-z]{1,3}\$?\d+)\.\.(\$?[A-Za-z]{1,3}\$?\d+)"
)

LOTUS_TO_GSHEETS_FUNCS = {
    "AVG": "AVERAGE",
    "STD": "STDEV",
    "LENGTH": "LEN",
    "REPEAT": "REPT",
    "STRING": "TEXT",
    "ISNA": "ISNA",
    "ISERR": "ISERR",
    "ISNUMBER": "ISNUMBER",
    "ISSTRING": "ISTEXT",
}


class FormulaError(Exception):
    """Represents a spreadsheet formula evaluation error (#DIV/0!, #REF!, etc.)."""

    def __init__(self, code: str = "ERR"):
        super().__init__(code)
        self.code = code


def col_to_idx(col_str: str) -> int:
    """Converts column letters ('A'->0, 'Z'->25, 'AA'->26) to 0-based index."""
    idx = 0
    for ch in col_str.upper():
        if not ("A" <= ch <= "Z"):
            raise ValueError(f"Invalid column: {col_str}")
        idx = idx * 26 + (ord(ch) - ord("A") + 1)
    return idx - 1


def idx_to_col(idx: int) -> str:
    """Converts 0-based column index (0->'A', 25->'Z', 26->'AA') to letters."""
    if idx < 0:
        return "A"
    res = []
    n = idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        res.append(chr(ord("A") + rem))
    return "".join(reversed(res))


def format_cell_ref(row: int, col: int, abs_row: bool = False, abs_col: bool = False) -> str:
    c_prefix = "$" if abs_col else ""
    r_prefix = "$" if abs_row else ""
    return f"{c_prefix}{idx_to_col(col)}{r_prefix}{row + 1}"


def parse_cell_ref(ref: str) -> Optional[Tuple[int, int, bool, bool]]:
    """Returns (row_0based, col_0based, abs_row, abs_col) or None."""
    m = CELL_REF_RE.match(ref.strip())
    if not m:
        return None
    abs_c, col_s, abs_r, row_s = m.groups()
    col = col_to_idx(col_s)
    row = int(row_s) - 1
    if row < 0 or col < 0 or row >= MAX_ROWS or col >= MAX_COLS:
        return None
    return (row, col, bool(abs_r), bool(abs_c))


def parse_range_str(range_str: str) -> Optional[Tuple[int, int, int, int]]:
    """
    Parses 'A1..C5', 'A1:C5', or single cell 'B3' into (r1, c1, r2, c2)
    normalized so r1 <= r2 and c1 <= c2.
    """
    s = range_str.strip()
    if ".." in s:
        parts = s.split("..", 1)
    elif ":" in s:
        parts = s.split(":", 1)
    else:
        p = parse_cell_ref(s)
        if not p:
            return None
        return (p[0], p[1], p[0], p[1])

    p1 = parse_cell_ref(parts[0])
    p2 = parse_cell_ref(parts[1])
    if not p1 or not p2:
        return None
    return (min(p1[0], p2[0]), min(p1[1], p2[1]), max(p1[0], p2[0]), max(p1[1], p2[1]))


def is_formula_raw(raw: str) -> bool:
    """Determines whether raw cell input should be treated as a formula."""
    if not raw:
        return False
    s = raw.strip()
    if not s:
        return False
    if s[0] in ("'", '"', "^", "\\"):
        return False
    if s[0] in ("=", "@"):
        return True
    if s[0] in ("+", "-", "("):
        # Plain signed numeric literal like -42.5 or +100 is a number, not a formula
        try:
            float(s.replace(",", ""))
            return False
        except ValueError:
            return True
    return False


def try_parse_number(raw: str) -> Optional[float]:
    """Attempts to parse a plain numeric literal (including $, %, commas)."""
    s = raw.strip()
    if not s:
        return None
    if s[0] in ("'", '"', "^", "\\", "=", "@"):
        return None
    # Handle currency / percent literals entered directly by user
    neg = False
    clean = s
    if clean.startswith("(") and clean.endswith(")"):
        neg = True
        clean = clean[1:-1].strip()
    is_pct = False
    if clean.endswith("%"):
        is_pct = True
        clean = clean[:-1].strip()
    if clean.startswith("$"):
        clean = clean[1:].strip()
    clean = clean.replace(",", "")
    try:
        val = float(clean)
        if math.isnan(val) or math.isinf(val):
            return None
        if neg:
            val = -val
        if is_pct:
            val /= 100.0
        return val
    except ValueError:
        return None


def _split_outside_quotes(expr: str) -> List[Tuple[bool, str]]:
    """Splits formula into (is_string_literal, segment) chunks."""
    out: List[Tuple[bool, str]] = []
    buf: List[str] = []
    in_str = False
    i = 0
    while i < len(expr):
        ch = expr[i]
        if ch == '"':
            if in_str:
                buf.append(ch)
                if i + 1 < len(expr) and expr[i + 1] == '"':
                    buf.append('"')
                    i += 2
                    continue
                out.append((True, "".join(buf)))
                buf = []
                in_str = False
            else:
                if buf:
                    out.append((False, "".join(buf)))
                    buf = []
                buf.append(ch)
                in_str = True
        else:
            buf.append(ch)
        i += 1
    if buf:
        out.append((in_str, "".join(buf)))
    return out


def to_gsheets_formula(raw: str) -> str:
    """
    Converts Lotus 1-2-3 formula syntax into Google Sheets compatible syntax:
      - @SUM(A1..A5)  -> =SUM(A1:A5)
      - @AVG(B2..B10) -> =AVERAGE(B2:B10)
      - +A1+B1        -> =A1+B1
      - (A1+B1)*2     -> =(A1+B1)*2
    Leaves labels (stripping Lotus alignment prefix) and numbers clean for Sheets.
    """
    if not raw:
        return ""
    if raw[0] in ("'", '"', "^"):
        return raw[1:]
    if raw[0] == "\\":
        # Repeating fill character -> expand to a readable divider string for Sheets
        ch = raw[1:] or "-"
        return (ch * 12)[:12]
    if not is_formula_raw(raw):
        return raw

    body = raw.strip()
    if body.startswith("="):
        body = body[1:]
    elif body.startswith("+"):
        body = body[1:]
    elif body.startswith("@"):
        pass  # @FUNC handled below

    chunks = _split_outside_quotes(body)
    converted_parts: List[str] = []
    for is_str, seg in chunks:
        if is_str:
            converted_parts.append(seg)
            continue
        # Convert @FUNC( -> FUNC( with Lotus name mapping
        def _repl_at_func(m):
            fname = m.group(1).upper()
            mapped = LOTUS_TO_GSHEETS_FUNCS.get(fname, fname)
            return f"{mapped}("

        seg = re.sub(r"@([A-Za-z_][A-Za-z0-9_]*)\s*\(", _repl_at_func, seg)
        # Convert standalone @PI, @TODAY, @NOW, @TRUE, @FALSE
        seg = re.sub(
            r"@(PI|TODAY|NOW|TRUE|FALSE|RAND)\b",
            lambda m: f"{m.group(1).upper()}()",
            seg,
            flags=re.IGNORECASE,
        )
        # Convert A1..B5 -> A1:B5
        seg = LOTUS_RANGE_RE.sub(r"\1:\2", seg)
        # Convert #AND# #OR# #NOT#? Keep standard operators; != -> <>
        seg = seg.replace("!=", "<>")
        converted_parts.append(seg)

    return "=" + "".join(converted_parts)


def shift_formula(
    raw: str,
    drow: int,
    dcol: int,
    min_row: int = 0,
    min_col: int = 0,
    respect_abs: bool = True,
) -> str:
    """
    Adjusts cell references inside a formula by (drow, dcol).
    Used for Copy, Move, and Insert/Delete Row or Column.
    """
    if not is_formula_raw(raw):
        return raw

    chunks = _split_outside_quotes(raw)
    out: List[str] = []
    # Match cell references not immediately followed by '(' (which would be a function name)
    ref_pattern = re.compile(r"(?<![A-Za-z0-9_])(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?!\s*\()")

    for is_str, seg in chunks:
        if is_str:
            out.append(seg)
            continue

        def _shift_match(m):
            abs_c, col_s, abs_r, row_s = m.groups()
            try:
                c = col_to_idx(col_s)
                r = int(row_s) - 1
            except ValueError:
                return m.group(0)
            if c >= MAX_COLS or r >= MAX_ROWS or r < 0:
                return m.group(0)

            new_r = r
            new_c = c
            if r >= min_row and (not respect_abs or not abs_r):
                new_r = r + drow
            if c >= min_col and (not respect_abs or not abs_c):
                new_c = c + dcol

            if new_r < 0 or new_c < 0 or new_r >= MAX_ROWS or new_c >= MAX_COLS:
                return "#REF!"
            return f"{abs_c}{idx_to_col(new_c)}{abs_r}{new_r + 1}"

        out.append(ref_pattern.sub(_shift_match, seg))
    return "".join(out)


# =============================================================================
# Data Model: Cell, Worksheet, Workbook
# =============================================================================

class Cell:
    """Represents a single cell in a Lotus 1-2-3 / Google Sheets worksheet."""

    __slots__ = ("raw", "fmt", "value", "error")

    def __init__(self, raw: str = "", fmt: str = ""):
        self.raw: str = raw
        self.fmt: str = fmt  # "", "G", "F0".."F4", "C0".."C2", "P0".."P2", ",0"..",2", "S2", "D"
        self.value: Any = None
        self.error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        d: Dict[str, Any] = {"raw": self.raw}
        if self.fmt:
            d["fmt"] = self.fmt
        return d

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "Cell":
        return cls(raw=str(d.get("raw", "")), fmt=str(d.get("fmt", "")))


class Worksheet:
    """Represents a single tab/sheet inside a Workbook."""

    def __init__(self, name: str = "Sheet1"):
        self.name: str = name
        self.cells: Dict[Tuple[int, int], Cell] = {}
        self.col_widths: Dict[int, int] = {}
        self.default_col_width: int = 11
        self.global_format: str = "G"
        self.freeze_rows: int = 0
        self.freeze_cols: int = 0

    def get_col_width(self, col: int) -> int:
        return self.col_widths.get(col, self.default_col_width)

    def set_col_width(self, col: int, width: int):
        self.col_widths[col] = max(3, min(48, width))

    def get_cell(self, row: int, col: int) -> Optional[Cell]:
        return self.cells.get((row, col))

    def get_raw(self, row: int, col: int) -> str:
        c = self.cells.get((row, col))
        return c.raw if c else ""

    def set_raw(self, row: int, col: int, raw: str, fmt: Optional[str] = None):
        if raw == "" and not fmt:
            existing = self.cells.get((row, col))
            if existing and existing.fmt:
                existing.raw = ""
                existing.value = None
                existing.error = None
            else:
                self.cells.pop((row, col), None)
            return
        existing = self.cells.get((row, col))
        if existing:
            existing.raw = raw
            if fmt is not None:
                existing.fmt = fmt
        else:
            # Auto-detect currency/percent format when user types $123 or 15%
            auto_fmt = fmt or ""
            if not auto_fmt and not is_formula_raw(raw):
                s = raw.strip()
                if s.startswith("$") or (s.startswith("($") and s.endswith(")")):
                    auto_fmt = "C2"
                elif s.endswith("%") and try_parse_number(s) is not None:
                    auto_fmt = "P1"
            self.cells[(row, col)] = Cell(raw=raw, fmt=auto_fmt)

    def set_format(self, row: int, col: int, fmt: str):
        cell = self.cells.get((row, col))
        if cell:
            cell.fmt = fmt
            if not cell.raw and not cell.fmt:
                self.cells.pop((row, col), None)
        elif fmt:
            self.cells[(row, col)] = Cell(raw="", fmt=fmt)

    def clear_cell(self, row: int, col: int):
        self.cells.pop((row, col), None)

    def get_bounds(self) -> Tuple[int, int]:
        """Returns (max_row_used, max_col_used) or (-1, -1) if empty."""
        if not self.cells:
            return (-1, -1)
        max_r = max(r for r, _ in self.cells.keys())
        max_c = max(c for _, c in self.cells.keys())
        return (max_r, max_c)

    def insert_row(self, at_row: int):
        new_cells: Dict[Tuple[int, int], Cell] = {}
        for (r, c), cell in self.cells.items():
            nr = r + 1 if r >= at_row else r
            cell.raw = shift_formula(cell.raw, 1, 0, min_row=at_row, respect_abs=False)
            if nr < MAX_ROWS:
                new_cells[(nr, c)] = cell
        self.cells = new_cells

    def delete_row(self, at_row: int):
        new_cells: Dict[Tuple[int, int], Cell] = {}
        for (r, c), cell in self.cells.items():
            if r == at_row:
                continue
            nr = r - 1 if r > at_row else r
            cell.raw = shift_formula(cell.raw, -1, 0, min_row=at_row + 1, respect_abs=False)
            new_cells[(nr, c)] = cell
        self.cells = new_cells

    def insert_col(self, at_col: int):
        new_cells: Dict[Tuple[int, int], Cell] = {}
        for (r, c), cell in self.cells.items():
            nc = c + 1 if c >= at_col else c
            cell.raw = shift_formula(cell.raw, 0, 1, min_col=at_col, respect_abs=False)
            if nc < MAX_COLS:
                new_cells[(r, nc)] = cell
        self.cells = new_cells
        new_widths: Dict[int, int] = {}
        for c, w in self.col_widths.items():
            new_widths[c + 1 if c >= at_col else c] = w
        self.col_widths = new_widths

    def delete_col(self, at_col: int):
        new_cells: Dict[Tuple[int, int], Cell] = {}
        for (r, c), cell in self.cells.items():
            if c == at_col:
                continue
            nc = c - 1 if c > at_col else c
            cell.raw = shift_formula(cell.raw, 0, -1, min_col=at_col + 1, respect_abs=False)
            new_cells[(r, nc)] = cell
        self.cells = new_cells
        new_widths: Dict[int, int] = {}
        for c, w in self.col_widths.items():
            if c == at_col:
                continue
            new_widths[c - 1 if c > at_col else c] = w
        self.col_widths = new_widths

    def auto_fit_col(self, col: int, min_w: int = 6, max_w: int = 36) -> int:
        best = min_w
        for (r, c), cell in self.cells.items():
            if c == col:
                disp = format_cell_display(cell, self.global_format, max_w)
                best = max(best, min(max_w, len(disp) + 2))
        self.col_widths[col] = best
        return best

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "default_col_width": self.default_col_width,
            "global_format": self.global_format,
            "freeze_rows": self.freeze_rows,
            "freeze_cols": self.freeze_cols,
            "col_widths": {str(k): v for k, v in self.col_widths.items()},
            "cells": {
                format_cell_ref(r, c): cell.to_dict()
                for (r, c), cell in sorted(self.cells.items())
                if cell.raw != "" or cell.fmt != ""
            },
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "Worksheet":
        ws = cls(name=str(d.get("name", "Sheet1")))
        ws.default_col_width = int(d.get("default_col_width", 11))
        ws.global_format = str(d.get("global_format", "G"))
        ws.freeze_rows = int(d.get("freeze_rows", 0))
        ws.freeze_cols = int(d.get("freeze_cols", 0))
        for k, v in d.get("col_widths", {}).items():
            try:
                ws.col_widths[int(k)] = int(v)
            except Exception:
                pass
        for ref, cdata in d.get("cells", {}).items():
            parsed = parse_cell_ref(ref)
            if parsed:
                ws.cells[(parsed[0], parsed[1])] = Cell.from_dict(cdata)
        return ws


class Workbook:
    """Multi-sheet Lotus 1-2-3 / Google Sheets Workbook with full recalculation."""

    def __init__(self, title: str = "UNTITLED"):
        self.title: str = title
        self.sheets: List[Worksheet] = [Worksheet("Sheet1")]
        self.active_sheet_idx: int = 0
        self.named_ranges: Dict[str, Tuple[str, str]] = {}  # UPPER_NAME -> (sheet_name, range_str)
        self.auto_recalc: bool = True
        self.has_circular_ref: bool = False
        self.dirty: bool = False
        self.gsheet_id: Optional[str] = None
        self.undo_stack: List[Dict[str, Any]] = []
        self.redo_stack: List[Dict[str, Any]] = []

    @property
    def active_sheet(self) -> Worksheet:
        if not self.sheets:
            self.sheets.append(Worksheet("Sheet1"))
        self.active_sheet_idx = max(0, min(self.active_sheet_idx, len(self.sheets) - 1))
        return self.sheets[self.active_sheet_idx]

    def get_sheet_by_name(self, name: str) -> Optional[Worksheet]:
        target = name.strip().upper()
        for ws in self.sheets:
            if ws.name.upper() == target:
                return ws
        # Also support Lotus Release 3 single-letter sheet references: A:, B:, C:
        if len(target) == 1 and "A" <= target <= "Z":
            idx = ord(target) - ord("A")
            if 0 <= idx < len(self.sheets):
                return self.sheets[idx]
        return None

    def add_sheet(self, name: Optional[str] = None) -> Worksheet:
        if not name:
            n = len(self.sheets) + 1
            while any(s.name.upper() == f"SHEET{n}" for s in self.sheets):
                n += 1
            name = f"Sheet{n}"
        ws = Worksheet(name)
        self.sheets.append(ws)
        self.active_sheet_idx = len(self.sheets) - 1
        self.dirty = True
        return ws

    def delete_active_sheet(self) -> Tuple[bool, str]:
        if len(self.sheets) <= 1:
            return False, "Cannot delete the only worksheet in the workbook."
        removed = self.sheets.pop(self.active_sheet_idx)
        self.active_sheet_idx = max(0, self.active_sheet_idx - 1)
        self.dirty = True
        self.recalculate()
        return True, f"Deleted sheet '{removed.name}'."

    def snapshot_state(self) -> Dict[str, Any]:
        return {
            "active_sheet_idx": self.active_sheet_idx,
            "named_ranges": copy.deepcopy(self.named_ranges),
            "sheets": [s.to_dict() for s in self.sheets],
        }

    def push_undo(self):
        self.undo_stack.append(self.snapshot_state())
        if len(self.undo_stack) > 40:
            self.undo_stack.pop(0)
        self.redo_stack.clear()

    def undo(self) -> bool:
        if not self.undo_stack:
            return False
        self.redo_stack.append(self.snapshot_state())
        snap = self.undo_stack.pop()
        self._restore_snapshot(snap)
        return True

    def redo(self) -> bool:
        if not self.redo_stack:
            return False
        self.undo_stack.append(self.snapshot_state())
        snap = self.redo_stack.pop()
        self._restore_snapshot(snap)
        return True

    def _restore_snapshot(self, snap: Dict[str, Any]):
        self.sheets = [Worksheet.from_dict(s) for s in snap.get("sheets", [])]
        if not self.sheets:
            self.sheets = [Worksheet("Sheet1")]
        self.active_sheet_idx = max(0, min(snap.get("active_sheet_idx", 0), len(self.sheets) - 1))
        self.named_ranges = copy.deepcopy(snap.get("named_ranges", {}))
        self.dirty = True
        self.recalculate()

    def to_dict(self) -> Dict[str, Any]:
        return {
            "version": 1,
            "title": self.title,
            "gsheet_id": self.gsheet_id,
            "active_sheet_idx": self.active_sheet_idx,
            "auto_recalc": self.auto_recalc,
            "named_ranges": self.named_ranges,
            "sheets": [s.to_dict() for s in self.sheets],
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> "Workbook":
        wb = cls(title=str(d.get("title", "UNTITLED")))
        wb.gsheet_id = d.get("gsheet_id")
        wb.auto_recalc = bool(d.get("auto_recalc", True))
        wb.named_ranges = {
            k.upper(): (v[0], v[1])
            for k, v in d.get("named_ranges", {}).items()
            if isinstance(v, (list, tuple)) and len(v) == 2
        }
        sheets_raw = d.get("sheets", [])
        if sheets_raw:
            wb.sheets = [Worksheet.from_dict(s) for s in sheets_raw]
        wb.active_sheet_idx = max(0, min(int(d.get("active_sheet_idx", 0)), len(wb.sheets) - 1))
        wb.recalculate()
        return wb

    def recalculate(self):
        """Recalculates all cells across all sheets with circular-ref detection."""
        self.has_circular_ref = False
        cache: Dict[Tuple[str, int, int], Any] = {}
        evaluator = FormulaEvaluator(self, cache)

        for ws in self.sheets:
            for (r, c), cell in list(ws.cells.items()):
                try:
                    val = evaluator.eval_cell(ws, r, c, set())
                    cell.value = val
                    cell.error = None
                except FormulaError as ferr:
                    cell.value = ferr.code
                    cell.error = ferr.code
                except Exception:
                    cell.value = "ERR"
                    cell.error = "ERR"


# =============================================================================
# Recursive-Descent Formula Lexer & Evaluator
# =============================================================================

TOKEN_SPEC = re.compile(
    r"""
    (?P<SPACE>\s+)
    | (?P<STRING>"(?:""|[^"])*")
    | (?P<NUMBER>(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)
    | (?P<QUOTED_SHEET>'[^']+'!)
    | (?P<OP><=|>=|<>|!=|=|<|>|\+|-|\*|/|\^|&|%)
    | (?P<DOTDOT>\.\.)
    | (?P<COLON>:)
    | (?P<BANG>!)
    | (?P<COMMA>,)
    | (?P<LPAREN>\()
    | (?P<RPAREN>\))
    | (?P<IDENT>@?[A-Za-z_$][A-Za-z0-9_$.]*)
    """,
    re.VERBOSE,
)


def _tokenize_formula(expr: str) -> List[Tuple[str, str]]:
    tokens: List[Tuple[str, str]] = []
    pos = 0
    n = len(expr)
    while pos < n:
        m = TOKEN_SPEC.match(expr, pos)
        if not m:
            raise FormulaError("#NAME?")
        kind = m.lastgroup
        text = m.group()
        pos = m.end()
        if kind == "SPACE":
            continue
        if kind == "IDENT" and ".." in text:
            # Split A1..B5 that matched IDENT because '.' is allowed in IDENT
            parts = text.split("..")
            for idx, p in enumerate(parts):
                if p:
                    tokens.append(("IDENT", p))
                if idx < len(parts) - 1:
                    tokens.append(("DOTDOT", ".."))
            continue
        tokens.append((kind, text))
    return tokens


def _to_number(val: Any) -> float:
    if val is None or val == "":
        return 0.0
    if isinstance(val, bool):
        return 1.0 if val else 0.0
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        num = try_parse_number(val)
        if num is not None:
            return num
    raise FormulaError("#VALUE!")


def _to_bool(val: Any) -> bool:
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val != 0
    if isinstance(val, str):
        u = val.strip().upper()
        if u in ("TRUE", "1"):
            return True
        if u in ("FALSE", "0", ""):
            return False
    return bool(val)


def _to_str(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, bool):
        return "TRUE" if val else "FALSE"
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return f"{val:.10g}"
    return str(val)


def _flatten_values(args: List[Any]) -> List[Any]:
    flat: List[Any] = []
    for a in args:
        if isinstance(a, list):
            for item in a:
                if isinstance(item, list):
                    flat.extend(item)
                else:
                    flat.append(item)
        else:
            flat.append(a)
    return flat


def _flatten_numbers(args: List[Any]) -> List[float]:
    nums: List[float] = []
    for item in _flatten_values(args):
        if isinstance(item, bool):
            continue
        if isinstance(item, (int, float)):
            nums.append(float(item))
    return nums


class FormulaEvaluator:
    """Evaluates cells & formulas inside a Workbook."""

    def __init__(self, workbook: Workbook, cache: Dict[Tuple[str, int, int], Any]):
        self.wb = workbook
        self.cache = cache

    def eval_cell(
        self, ws: Worksheet, row: int, col: int, stack: Set[Tuple[str, int, int]]
    ) -> Any:
        key = (ws.name.upper(), row, col)
        if key in self.cache:
            cached = self.cache[key]
            if isinstance(cached, FormulaError):
                raise cached
            return cached

        if key in stack:
            self.wb.has_circular_ref = True
            raise FormulaError("CIRC")

        cell = ws.get_cell(row, col)
        if cell is None or cell.raw == "":
            return None

        raw = cell.raw
        # Lotus 1-2-3 label prefixes
        if raw[0] in ("'", '"', "^"):
            val: Any = raw[1:]
            self.cache[key] = val
            return val
        if raw[0] == "\\":
            self.cache[key] = raw
            return raw

        # Check plain numeric literal
        if not is_formula_raw(raw):
            u = raw.strip().upper()
            if u == "TRUE":
                self.cache[key] = True
                return True
            if u == "FALSE":
                self.cache[key] = False
                return False
            num = try_parse_number(raw)
            if num is not None:
                self.cache[key] = num
                return num
            self.cache[key] = raw
            return raw

        # Formula evaluation
        expr = raw.strip()
        if expr.startswith("="):
            expr = expr[1:]

        stack.add(key)
        try:
            tokens = _tokenize_formula(expr)
            if not tokens:
                val = None
            else:
                parser = _ExprParser(self, ws, tokens, stack)
                val = parser.parse()
                if parser.pos < len(tokens):
                    raise FormulaError("ERR")
            self.cache[key] = val
            return val
        except FormulaError as ferr:
            self.cache[key] = ferr
            raise
        finally:
            stack.discard(key)

    def eval_range_matrix(
        self,
        ws: Worksheet,
        r1: int,
        c1: int,
        r2: int,
        c2: int,
        stack: Set[Tuple[str, int, int]],
    ) -> List[List[Any]]:
        r_lo, r_hi = min(r1, r2), max(r1, r2)
        c_lo, c_hi = min(c1, c2), max(c1, c2)
        # Bound huge open-ended ranges to active sheet bounds so performance stays snappy
        max_r, max_c = ws.get_bounds()
        r_hi = min(r_hi, max(r_lo, max_r, r_lo))
        c_hi = min(c_hi, max(c_lo, max_c, c_lo))
        # Cap total cells per single range op for safety
        if (r_hi - r_lo + 1) * (c_hi - c_lo + 1) > 25000:
            r_hi = min(r_hi, r_lo + 1000)
            c_hi = min(c_hi, c_lo + 25)

        matrix: List[List[Any]] = []
        for r in range(r_lo, r_hi + 1):
            row_vals: List[Any] = []
            for c in range(c_lo, c_hi + 1):
                row_vals.append(self.eval_cell(ws, r, c, stack))
            matrix.append(row_vals)
        return matrix


class _ExprParser:
    """Recursive-descent expression parser for spreadsheet formulas."""

    def __init__(
        self,
        evaluator: FormulaEvaluator,
        ws: Worksheet,
        tokens: List[Tuple[str, str]],
        stack: Set[Tuple[str, int, int]],
    ):
        self.ev = evaluator
        self.ws = ws
        self.tokens = tokens
        self.stack = stack
        self.pos = 0

    def peek(self, offset: int = 0) -> Optional[Tuple[str, str]]:
        idx = self.pos + offset
        if 0 <= idx < len(self.tokens):
            return self.tokens[idx]
        return None

    def consume(self) -> Tuple[str, str]:
        tok = self.tokens[self.pos]
        self.pos += 1
        return tok

    def parse(self) -> Any:
        return self._parse_comparison()

    def _parse_comparison(self) -> Any:
        left = self._parse_concat()
        while True:
            tok = self.peek()
            if tok and tok[0] == "OP" and tok[1] in ("=", "<>", "!=", "<", "<=", ">", ">="):
                op = self.consume()[1]
                right = self._parse_concat()
                left = self._apply_compare(op, left, right)
            else:
                break
        return left

    def _apply_compare(self, op: str, a: Any, b: Any) -> bool:
        # Normalize None for comparisons
        if a is None and isinstance(b, (int, float)):
            a = 0.0
        elif b is None and isinstance(a, (int, float)):
            b = 0.0
        elif a is None and isinstance(b, str):
            a = ""
        elif b is None and isinstance(a, str):
            b = ""

        if isinstance(a, str) and isinstance(b, str):
            aa, bb = a.lower(), b.lower()
        elif isinstance(a, (int, float, bool)) and isinstance(b, (int, float, bool)):
            aa, bb = float(a), float(b)
        else:
            aa, bb = _to_str(a).lower(), _to_str(b).lower()

        if op == "=":
            return aa == bb
        if op in ("<>", "!="):
            return aa != bb
        if op == "<":
            return aa < bb
        if op == "<=":
            return aa <= bb
        if op == ">":
            return aa > bb
        if op == ">=":
            return aa >= bb
        return False

    def _parse_concat(self) -> Any:
        left = self._parse_add()
        while True:
            tok = self.peek()
            if tok and tok[0] == "OP" and tok[1] == "&":
                self.consume()
                right = self._parse_add()
                left = _to_str(left) + _to_str(right)
            else:
                break
        return left

    def _parse_add(self) -> Any:
        left = self._parse_mul()
        while True:
            tok = self.peek()
            if tok and tok[0] == "OP" and tok[1] in ("+", "-"):
                op = self.consume()[1]
                right = self._parse_mul()
                a, b = _to_number(left), _to_number(right)
                left = a + b if op == "+" else a - b
            else:
                break
        return left

    def _parse_mul(self) -> Any:
        left = self._parse_power()
        while True:
            tok = self.peek()
            if tok and tok[0] == "OP" and tok[1] in ("*", "/"):
                op = self.consume()[1]
                right = self._parse_power()
                a, b = _to_number(left), _to_number(right)
                if op == "*":
                    left = a * b
                else:
                    if b == 0.0:
                        raise FormulaError("#DIV/0!")
                    left = a / b
            else:
                break
        return left

    def _parse_power(self) -> Any:
        left = self._parse_unary()
        tok = self.peek()
        if tok and tok[0] == "OP" and tok[1] == "^":
            self.consume()
            right = self._parse_power()
            a, b = _to_number(left), _to_number(right)
            try:
                res = a ** b
                if isinstance(res, complex) or math.isnan(res) or math.isinf(res):
                    raise FormulaError("#NUM!")
                return float(res)
            except (OverflowError, ZeroDivisionError):
                raise FormulaError("#NUM!")
        return left

    def _parse_unary(self) -> Any:
        tok = self.peek()
        if tok and tok[0] == "OP" and tok[1] in ("+", "-"):
            op = self.consume()[1]
            val = self._parse_unary()
            num = _to_number(val)
            return -num if op == "-" else num
        val = self._parse_postfix()
        return val

    def _parse_postfix(self) -> Any:
        val = self._parse_primary()
        while True:
            tok = self.peek()
            if tok and tok[0] == "OP" and tok[1] == "%":
                self.consume()
                val = _to_number(val) / 100.0
            else:
                break
        return val

    def _parse_primary(self) -> Any:
        tok = self.peek()
        if tok is None:
            raise FormulaError("ERR")

        kind, text = tok

        if kind == "NUMBER":
            self.consume()
            return float(text)

        if kind == "STRING":
            self.consume()
            return text[1:-1].replace('""', '"')

        if kind == "LPAREN":
            self.consume()
            val = self.parse()
            nxt = self.peek()
            if not nxt or nxt[0] != "RPAREN":
                raise FormulaError("ERR")
            self.consume()
            return val

        # Quoted sheet reference: 'My Sheet'!A1 or 'My Sheet'!A1:B5
        if kind == "QUOTED_SHEET":
            self.consume()
            sheet_name = text[1:-2]
            target_ws = self.ev.wb.get_sheet_by_name(sheet_name)
            if target_ws is None:
                raise FormulaError("#REF!")
            return self._parse_ref_or_range(target_ws)

        if kind == "IDENT":
            # Check for unquoted sheet reference: Sheet2!A1
            nxt = self.peek(1)
            if nxt and nxt[0] == "BANG":
                self.consume()  # sheet name
                self.consume()  # '!'
                target_ws = self.ev.wb.get_sheet_by_name(text)
                if target_ws is None:
                    raise FormulaError("#REF!")
                return self._parse_ref_or_range(target_ws)

            # Check for Lotus Release 3 sheet letter prefix: B:A1 (only if single letter + ':' + cell)
            if (
                len(text) == 1
                and "A" <= text.upper() <= "Z"
                and nxt
                and nxt[0] == "COLON"
                and self.peek(2)
                and self.peek(2)[0] == "IDENT"
                and parse_cell_ref(self.peek(2)[1]) is not None
            ):
                # Distinguish from cell ref; single letter alone is not a cell ref!
                self.consume()  # sheet letter
                self.consume()  # ':'
                target_ws = self.ev.wb.get_sheet_by_name(text)
                if target_ws is None:
                    raise FormulaError("#REF!")
                return self._parse_ref_or_range(target_ws)

            # Function call: SUM(...) or @SUM(...)
            if nxt and nxt[0] == "LPAREN":
                self.consume()  # func name
                return self._parse_func_call(text)

            # Boolean & constant keywords (including Lotus @PI, @TODAY, @NOW, @TRUE, @FALSE)
            clean_u = text.lstrip("@").upper()
            if clean_u == "TRUE":
                self.consume()
                return True
            if clean_u == "FALSE":
                self.consume()
                return False
            if text.startswith("@") and clean_u in ("PI", "TODAY", "NOW", "RAND"):
                self.consume()
                return self._eval_func(clean_u, [])

            # Named range lookup
            if clean_u in self.ev.wb.named_ranges:
                self.consume()
                sname, rstr = self.ev.wb.named_ranges[clean_u]
                target_ws = self.ev.wb.get_sheet_by_name(sname) or self.ws
                rng = parse_range_str(rstr)
                if not rng:
                    raise FormulaError("#REF!")
                r1, c1, r2, c2 = rng
                if r1 == r2 and c1 == c2:
                    return self.ev.eval_cell(target_ws, r1, c1, self.stack)
                return self.ev.eval_range_matrix(target_ws, r1, c1, r2, c2, self.stack)

            # Cell reference or cell range (A1 or A1:B5 or A1..B5)
            if parse_cell_ref(text) is not None:
                return self._parse_ref_or_range(self.ws)

            raise FormulaError("#NAME?")

        raise FormulaError("ERR")

    def _parse_ref_or_range(self, target_ws: Worksheet) -> Any:
        tok1 = self.peek()
        if not tok1 or tok1[0] != "IDENT":
            raise FormulaError("#REF!")
        p1 = parse_cell_ref(tok1[1])
        if p1 is None:
            raise FormulaError("#REF!")
        self.consume()

        nxt = self.peek()
        if nxt and nxt[0] in ("COLON", "DOTDOT"):
            self.consume()
            tok2 = self.peek()
            if not tok2 or tok2[0] != "IDENT":
                raise FormulaError("#REF!")
            p2 = parse_cell_ref(tok2[1])
            if p2 is None:
                raise FormulaError("#REF!")
            self.consume()
            return self.ev.eval_range_matrix(target_ws, p1[0], p1[1], p2[0], p2[1], self.stack)

        return self.ev.eval_cell(target_ws, p1[0], p1[1], self.stack)

    def _parse_func_call(self, raw_name: str) -> Any:
        fname = raw_name.lstrip("@").upper()
        self.consume()  # '('

        # Special short-circuit / error-catching functions: IF, IFERROR, ISERR, ISNA
        if fname == "IF":
            cond = self.parse()
            args_rest: List[Any] = []
            # Parse remaining branches lazily? Evaluate eagerly safely or skip unused branch
            if self.peek() and self.peek()[0] == "COMMA":
                self.consume()
                if _to_bool(cond):
                    res = self.parse()
                    if self.peek() and self.peek()[0] == "COMMA":
                        self.consume()
                        self._skip_arg()
                else:
                    self._skip_arg()
                    if self.peek() and self.peek()[0] == "COMMA":
                        self.consume()
                        res = self.parse()
                    else:
                        res = False
            else:
                res = _to_bool(cond)
            if self.peek() and self.peek()[0] == "RPAREN":
                self.consume()
            return res

        if fname in ("IFERROR", "ISERR", "ISERROR", "ISNA"):
            err_caught: Optional[str] = None
            val: Any = None
            try:
                val = self.parse()
            except FormulaError as fe:
                err_caught = fe.code
            if fname == "IFERROR":
                fallback: Any = ""
                if self.peek() and self.peek()[0] == "COMMA":
                    self.consume()
                    if err_caught is not None:
                        fallback = self.parse()
                    else:
                        self._skip_arg()
                if self.peek() and self.peek()[0] == "RPAREN":
                    self.consume()
                return fallback if err_caught is not None else val
            if self.peek() and self.peek()[0] == "RPAREN":
                self.consume()
            if fname == "ISNA":
                return err_caught == "#N/A"
            return err_caught is not None

        args: List[Any] = []
        if self.peek() and self.peek()[0] != "RPAREN":
            while True:
                args.append(self.parse())
                if self.peek() and self.peek()[0] == "COMMA":
                    self.consume()
                else:
                    break
        if not self.peek() or self.peek()[0] != "RPAREN":
            raise FormulaError("ERR")
        self.consume()
        return self._eval_func(fname, args)

    def _skip_arg(self):
        depth = 0
        while self.pos < len(self.tokens):
            kind, _ = self.tokens[self.pos]
            if kind == "LPAREN":
                depth += 1
            elif kind == "RPAREN":
                if depth == 0:
                    return
                depth -= 1
            elif kind == "COMMA" and depth == 0:
                return
            self.pos += 1

    def _eval_func(self, fname: str, args: List[Any]) -> Any:
        # ---- Aggregation & Math ----
        if fname == "SUM":
            return sum(_flatten_numbers(args))
        if fname in ("AVG", "AVERAGE"):
            nums = _flatten_numbers(args)
            if not nums:
                raise FormulaError("#DIV/0!")
            return sum(nums) / len(nums)
        if fname == "MIN":
            nums = _flatten_numbers(args)
            return min(nums) if nums else 0.0
        if fname == "MAX":
            nums = _flatten_numbers(args)
            return max(nums) if nums else 0.0
        if fname == "MEDIAN":
            nums = _flatten_numbers(args)
            if not nums:
                raise FormulaError("#NUM!")
            return float(statistics.median(nums))
        if fname == "COUNT":
            return float(len(_flatten_numbers(args)))
        if fname == "COUNTA":
            return float(sum(1 for v in _flatten_values(args) if v is not None and v != ""))
        if fname == "COUNTBLANK":
            return float(sum(1 for v in _flatten_values(args) if v is None or v == ""))
        if fname == "PRODUCT":
            nums = _flatten_numbers(args)
            if not nums:
                return 0.0
            res = 1.0
            for x in nums:
                res *= x
            return res
        if fname in ("STDEV", "STD", "STDEV.S"):
            nums = _flatten_numbers(args)
            if len(nums) < 2:
                raise FormulaError("#DIV/0!")
            return float(statistics.stdev(nums))
        if fname in ("STDEVP", "STDEV.P"):
            nums = _flatten_numbers(args)
            if not nums:
                raise FormulaError("#DIV/0!")
            return float(statistics.pstdev(nums))
        if fname in ("VAR", "VAR.S"):
            nums = _flatten_numbers(args)
            if len(nums) < 2:
                raise FormulaError("#DIV/0!")
            return float(statistics.variance(nums))
        if fname in ("VARP", "VAR.P"):
            nums = _flatten_numbers(args)
            if not nums:
                raise FormulaError("#DIV/0!")
            return float(statistics.pvariance(nums))
        if fname == "COUNTIF":
            if len(args) < 2:
                raise FormulaError("#VALUE!")
            vals = _flatten_values([args[0]])
            return float(sum(1 for v in vals if self._matches_criteria(v, args[1])))
        if fname == "SUMIF":
            if len(args) < 2:
                raise FormulaError("#VALUE!")
            crit_vals = _flatten_values([args[0]])
            sum_vals = _flatten_values([args[2]]) if len(args) >= 3 else crit_vals
            total = 0.0
            for i, cv in enumerate(crit_vals):
                if self._matches_criteria(cv, args[1]) and i < len(sum_vals):
                    sv = sum_vals[i]
                    if isinstance(sv, (int, float)) and not isinstance(sv, bool):
                        total += float(sv)
            return total
        if fname == "ABS":
            return abs(_to_number(args[0]))
        if fname == "INT":
            return float(math.floor(_to_number(args[0])))
        if fname == "ROUND":
            val = _to_number(args[0])
            digits = int(_to_number(args[1])) if len(args) > 1 else 0
            return float(round(val, digits))
        if fname == "ROUNDUP":
            val = _to_number(args[0])
            digits = int(_to_number(args[1])) if len(args) > 1 else 0
            factor = 10 ** digits
            return float(math.ceil(val * factor) / factor if val >= 0 else math.floor(val * factor) / factor)
        if fname == "ROUNDDOWN":
            val = _to_number(args[0])
            digits = int(_to_number(args[1])) if len(args) > 1 else 0
            factor = 10 ** digits
            return float(math.floor(val * factor) / factor if val >= 0 else math.ceil(val * factor) / factor)
        if fname == "MOD":
            a, b = _to_number(args[0]), _to_number(args[1])
            if b == 0:
                raise FormulaError("#DIV/0!")
            return float(a % b)
        if fname == "SQRT":
            v = _to_number(args[0])
            if v < 0:
                raise FormulaError("#NUM!")
            return math.sqrt(v)
        if fname == "POWER":
            return float(_to_number(args[0]) ** _to_number(args[1]))
        if fname == "EXP":
            return math.exp(_to_number(args[0]))
        if fname == "LN":
            v = _to_number(args[0])
            if v <= 0:
                raise FormulaError("#NUM!")
            return math.log(v)
        if fname in ("LOG", "LOG10"):
            v = _to_number(args[0])
            if v <= 0:
                raise FormulaError("#NUM!")
            base = _to_number(args[1]) if (fname == "LOG" and len(args) > 1) else 10.0
            return math.log(v, base)
        if fname == "PI":
            return math.pi
        if fname == "SIGN":
            v = _to_number(args[0])
            return 1.0 if v > 0 else (-1.0 if v < 0 else 0.0)
        if fname == "RAND":
            return random.random()
        if fname == "SIN":
            return math.sin(_to_number(args[0]))
        if fname == "COS":
            return math.cos(_to_number(args[0]))
        if fname == "TAN":
            return math.tan(_to_number(args[0]))

        # ---- Logical ----
        if fname == "AND":
            vals = _flatten_values(args)
            return all(_to_bool(v) for v in vals if v is not None)
        if fname == "OR":
            vals = _flatten_values(args)
            return any(_to_bool(v) for v in vals if v is not None)
        if fname == "NOT":
            return not _to_bool(args[0])
        if fname == "TRUE":
            return True
        if fname == "FALSE":
            return False
        if fname == "ISBLANK":
            return args[0] is None or args[0] == ""
        if fname == "ISNUMBER":
            return isinstance(args[0], (int, float)) and not isinstance(args[0], bool)
        if fname in ("ISTEXT", "ISSTRING"):
            return isinstance(args[0], str)

        # ---- Text / String ----
        if fname in ("LEN", "LENGTH"):
            return float(len(_to_str(args[0])))
        if fname == "LEFT":
            s = _to_str(args[0])
            n = int(_to_number(args[1])) if len(args) > 1 else 1
            return s[: max(0, n)]
        if fname == "RIGHT":
            s = _to_str(args[0])
            n = int(_to_number(args[1])) if len(args) > 1 else 1
            return s[-n:] if n > 0 else ""
        if fname == "MID":
            s = _to_str(args[0])
            start = max(1, int(_to_number(args[1]))) - 1
            length = max(0, int(_to_number(args[2])))
            return s[start : start + length]
        if fname == "UPPER":
            return _to_str(args[0]).upper()
        if fname == "LOWER":
            return _to_str(args[0]).lower()
        if fname == "PROPER":
            return _to_str(args[0]).title()
        if fname == "TRIM":
            return " ".join(_to_str(args[0]).split())
        if fname in ("CONCAT", "CONCATENATE"):
            return "".join(_to_str(v) for v in _flatten_values(args))
        if fname in ("REPT", "REPEAT"):
            return _to_str(args[0]) * max(0, int(_to_number(args[1])))
        if fname == "SUBSTITUTE":
            s, old, new = _to_str(args[0]), _to_str(args[1]), _to_str(args[2])
            return s.replace(old, new)
        if fname in ("FIND", "SEARCH"):
            needle, haystack = _to_str(args[0]), _to_str(args[1])
            start = max(1, int(_to_number(args[2]))) - 1 if len(args) > 2 else 0
            if fname == "SEARCH":
                idx = haystack.lower().find(needle.lower(), start)
            else:
                idx = haystack.find(needle, start)
            if idx < 0:
                raise FormulaError("#VALUE!")
            return float(idx + 1)
        if fname == "VALUE":
            return _to_number(args[0])
        if fname in ("TEXT", "STRING"):
            val = _to_number(args[0])
            fmt_arg = args[1] if len(args) > 1 else 0
            if isinstance(fmt_arg, (int, float)):
                return f"{val:.{max(0, int(fmt_arg))}f}"
            fstr = str(fmt_arg)
            if "%" in fstr:
                return f"{val * 100:.1f}%"
            if "$" in fstr:
                return f"${val:,.2f}"
            return f"{val:,.2f}"

        # ---- Lookup ----
        if fname == "VLOOKUP":
            if len(args) < 3 or not isinstance(args[1], list):
                raise FormulaError("#VALUE!")
            key = args[0]
            table: List[List[Any]] = args[1]
            col_idx = int(_to_number(args[2])) - 1
            approx = _to_bool(args[3]) if len(args) > 3 else True
            if col_idx < 0:
                raise FormulaError("#VALUE!")
            best_row = None
            for row in table:
                if not row or col_idx >= len(row):
                    continue
                first = row[0]
                if not approx:
                    if self._apply_compare("=", first, key):
                        return row[col_idx]
                else:
                    if self._apply_compare("<=", first, key):
                        best_row = row
            if approx and best_row is not None:
                return best_row[col_idx]
            raise FormulaError("#N/A")

        if fname == "HLOOKUP":
            if len(args) < 3 or not isinstance(args[1], list) or not args[1]:
                raise FormulaError("#VALUE!")
            key = args[0]
            table = args[1]
            row_idx = int(_to_number(args[2])) - 1
            approx = _to_bool(args[3]) if len(args) > 3 else True
            if row_idx < 0 or row_idx >= len(table):
                raise FormulaError("#REF!")
            header = table[0]
            best_c = None
            for c_i, val in enumerate(header):
                if not approx:
                    if self._apply_compare("=", val, key):
                        return table[row_idx][c_i]
                else:
                    if self._apply_compare("<=", val, key):
                        best_c = c_i
            if approx and best_c is not None:
                return table[row_idx][best_c]
            raise FormulaError("#N/A")

        if fname == "INDEX":
            if len(args) < 2 or not isinstance(args[0], list):
                raise FormulaError("#VALUE!")
            mat = args[0]
            r_i = int(_to_number(args[1])) - 1
            c_i = int(_to_number(args[2])) - 1 if len(args) > 2 else 0
            if 0 <= r_i < len(mat) and 0 <= c_i < len(mat[r_i]):
                return mat[r_i][c_i]
            raise FormulaError("#REF!")

        if fname == "MATCH":
            if len(args) < 2:
                raise FormulaError("#VALUE!")
            key = args[0]
            vals = _flatten_values([args[1]])
            mtype = int(_to_number(args[2])) if len(args) > 2 else 1
            for i, v in enumerate(vals):
                if mtype == 0 and self._apply_compare("=", v, key):
                    return float(i + 1)
            if mtype == 1:
                best = -1
                for i, v in enumerate(vals):
                    if self._apply_compare("<=", v, key):
                        best = i
                if best >= 0:
                    return float(best + 1)
            raise FormulaError("#N/A")

        if fname == "CHOOSE":
            idx = int(_to_number(args[0]))
            if 1 <= idx < len(args):
                return args[idx]
            raise FormulaError("#VALUE!")

        if fname == "ROWS":
            return float(len(args[0])) if isinstance(args[0], list) else 1.0
        if fname == "COLUMNS":
            if isinstance(args[0], list) and args[0] and isinstance(args[0][0], list):
                return float(len(args[0][0]))
            return 1.0

        # ---- Classic Lotus 1-2-3 Financial Functions ----
        if fname == "PMT":
            # Supports Google Sheets PMT(rate, nper, pv)
            rate, nper, pv = _to_number(args[0]), _to_number(args[1]), _to_number(args[2])
            if nper == 0:
                raise FormulaError("#DIV/0!")
            if rate == 0:
                return -pv / nper
            return -pv * (rate * (1 + rate) ** nper) / ((1 + rate) ** nper - 1)
        if fname == "FV":
            rate, nper, pmt = _to_number(args[0]), _to_number(args[1]), _to_number(args[2])
            pv = _to_number(args[3]) if len(args) > 3 else 0.0
            if rate == 0:
                return -(pv + pmt * nper)
            return -(pv * (1 + rate) ** nper + pmt * (((1 + rate) ** nper - 1) / rate))
        if fname == "PV":
            rate, nper, pmt = _to_number(args[0]), _to_number(args[1]), _to_number(args[2])
            fv = _to_number(args[3]) if len(args) > 3 else 0.0
            if rate == 0:
                return -(fv + pmt * nper)
            return -(fv + pmt * (((1 + rate) ** nper - 1) / rate)) / ((1 + rate) ** nper)
        if fname == "NPV":
            rate = _to_number(args[0])
            flows = _flatten_numbers(args[1:])
            return sum(cf / ((1.0 + rate) ** (i + 1)) for i, cf in enumerate(flows))
        if fname == "SLN":
            cost, salvage, life = _to_number(args[0]), _to_number(args[1]), _to_number(args[2])
            if life == 0:
                raise FormulaError("#DIV/0!")
            return (cost - salvage) / life

        # ---- Date & Time ----
        if fname == "TODAY":
            return datetime.date.today().strftime("%Y-%m-%d")
        if fname == "NOW":
            return datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
        if fname == "DATE":
            y, m, d = int(_to_number(args[0])), int(_to_number(args[1])), int(_to_number(args[2]))
            return datetime.date(y, m, d).strftime("%Y-%m-%d")
        if fname == "YEAR":
            dt = self._parse_date_arg(args[0])
            return float(dt.year)
        if fname == "MONTH":
            dt = self._parse_date_arg(args[0])
            return float(dt.month)
        if fname == "DAY":
            dt = self._parse_date_arg(args[0])
            return float(dt.day)

        raise FormulaError("#NAME?")

    def _parse_date_arg(self, val: Any) -> datetime.date:
        s = _to_str(val).strip()
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%d-%b-%y", "%d-%b-%Y"):
            try:
                return datetime.datetime.strptime(s[:10], fmt).date()
            except ValueError:
                continue
        raise FormulaError("#VALUE!")

    def _matches_criteria(self, val: Any, criteria: Any) -> bool:
        if isinstance(criteria, str):
            m = re.match(r"^(<=|>=|<>|!=|=|<|>)(.*)$", criteria.strip())
            if m:
                op, rhs_str = m.groups()
                rhs_num = try_parse_number(rhs_str)
                rhs: Any = rhs_num if rhs_num is not None else rhs_str
                return self._apply_compare(op, val, rhs)
        return self._apply_compare("=", val, criteria)


# =============================================================================
# Cell Display Formatting (Lotus 1-2-3 Style)
# =============================================================================

def format_number(val: float, fmt: str) -> str:
    """Formats a numeric value using Lotus 1-2-3 format codes."""
    code = (fmt or "G").upper()
    if abs(val) < 1e-12:
        val = 0.0

    if code == "G":
        if val == int(val) and abs(val) < 1e12:
            return str(int(val))
        return f"{val:.8g}"

    if code.startswith("F"):
        dec = int(code[1:]) if len(code) > 1 and code[1:].isdigit() else 2
        return f"{val:.{dec}f}"

    if code.startswith(","):
        dec = int(code[1:]) if len(code) > 1 and code[1:].isdigit() else 2
        return f"{val:,.{dec}f}"

    if code.startswith("C"):
        dec = int(code[1:]) if len(code) > 1 and code[1:].isdigit() else 2
        if val < 0:
            return f"(${abs(val):,.{dec}f})"
        return f"${val:,.{dec}f}"

    if code.startswith("P"):
        dec = int(code[1:]) if len(code) > 1 and code[1:].isdigit() else 1
        return f"{val * 100:.{dec}f}%"

    if code.startswith("S"):
        dec = int(code[1:]) if len(code) > 1 and code[1:].isdigit() else 2
        return f"{val:.{dec}E}"

    if code == "D":
        # Excel/Lotus serial date (days since 1899-12-30)
        try:
            dt = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(val))
            return dt.strftime("%d-%b-%y")
        except Exception:
            return str(val)

    return f"{val:.8g}"


def format_cell_display(cell: Optional[Cell], global_fmt: str, width: int) -> str:
    """
    Renders a cell's evaluated value into a string aligned for a column of `width`.
    Classic Lotus 1-2-3 behaviors:
      - Repeating label prefix `\\` fills entire column width (e.g. `\\-` -> `-----------`)
      - Label alignment prefixes `'` (left), `"` (right), `^` (center)
      - Numbers right-aligned; if formatted number exceeds width -> `*******`
    """
    if cell is None or (cell.raw == "" and cell.value is None):
        return " " * width

    if cell.error:
        return cell.error.center(width)[:width]

    # Lotus repeating fill label (e.g. \- or \=)
    if cell.raw.startswith("\\"):
        pattern = cell.raw[1:] or " "
        repeated = (pattern * (width // max(1, len(pattern)) + 2))[:width]
        return repeated

    val = cell.value
    if val is None:
        return " " * width

    if isinstance(val, bool):
        return ("TRUE" if val else "FALSE").center(width)[:width]

    if isinstance(val, (int, float)):
        fmt = cell.fmt or global_fmt or "G"
        s = format_number(float(val), fmt)
        if len(s) > width:
            # Try general fallback if General format was used and decimal precision caused overflow
            if fmt.upper() == "G" and len(f"{val:.4g}") <= width:
                s = f"{val:.4g}"
            else:
                return "*" * width
        return s.rjust(width)

    # Text / label
    text = str(val)
    align = "'"
    if cell.raw and cell.raw[0] in ("'", '"', "^"):
        align = cell.raw[0]
    if align == '"':
        return text[:width].rjust(width)
    if align == "^":
        return text[:width].center(width)
    return text[:width].ljust(width)


# =============================================================================
# CSV, Pure-Python Multi-Sheet XLSX, and ASCII Report Import / Export
# =============================================================================

def import_csv_to_worksheet(csv_text: str, name: str = "Sheet1") -> Worksheet:
    ws = Worksheet(name)
    reader = csv.reader(io.StringIO(csv_text))
    for r, row in enumerate(reader):
        if r >= MAX_ROWS:
            break
        for c, val in enumerate(row):
            if c >= MAX_COLS:
                break
            if val != "":
                ws.set_raw(r, c, val)
    return ws


def export_worksheet_to_csv(ws: Worksheet, formulas: bool = False) -> str:
    max_r, max_c = ws.get_bounds()
    if max_r < 0:
        return ""
    buf = io.StringIO()
    writer = csv.writer(buf)
    for r in range(max_r + 1):
        row_out = []
        for c in range(max_c + 1):
            cell = ws.get_cell(r, c)
            if not cell or cell.raw == "":
                row_out.append("")
            elif formulas:
                row_out.append(to_gsheets_formula(cell.raw))
            else:
                if cell.error:
                    row_out.append(cell.error)
                elif cell.value is None:
                    row_out.append("")
                elif isinstance(cell.value, float) and cell.value.is_integer():
                    row_out.append(str(int(cell.value)))
                else:
                    row_out.append(str(cell.value))
        writer.writerow(row_out)
    return buf.getvalue()


def export_worksheet_to_ascii(ws: Worksheet, title: str = "") -> str:
    """Renders a crisp monospace ASCII report suitable for printing or OV/400 Word Processing."""
    max_r, max_c = ws.get_bounds()
    lines: List[str] = []
    if title:
        lines.append(f"# {title} — {ws.name}")
        lines.append(f"Generated by Lotus 1-2-3 for OV/400 on {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}")
        lines.append("")
    if max_r < 0:
        lines.append("(Empty Worksheet)")
        return "\n".join(lines) + "\n"

    widths = [ws.get_col_width(c) for c in range(max_c + 1)]
    sep = "+" + "+".join("-" * (w + 2) for w in widths) + "+"
    hdr = "|" + "|".join(f" {idx_to_col(c).center(widths[c])} " for c in range(max_c + 1)) + "|"
    lines.append(sep)
    lines.append(hdr)
    lines.append(sep)
    for r in range(max_r + 1):
        row_cells = []
        for c in range(max_c + 1):
            cell = ws.get_cell(r, c)
            txt = format_cell_display(cell, ws.global_format, widths[c])
            row_cells.append(f" {txt} ")
        lines.append("|" + "|".join(row_cells) + "|")
    lines.append(sep)
    return "\n".join(lines) + "\n"


def _xml_escape(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def export_workbook_to_xlsx_bytes(wb: Workbook) -> bytes:
    """
    Generates a valid multi-sheet OpenXML .xlsx workbook in memory using only
    the Python standard library (zipfile), preserving live formulas (<f> tags),
    sheet tab names, numbers, and strings. Google Drive imports this cleanly
    into a native multi-tab Google Sheet!
    """
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        # [Content_Types].xml
        sheet_overrides = "\n".join(
            f'  <Override PartName="/xl/worksheets/sheet{i + 1}.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            for i in range(len(wb.sheets))
        )
        zf.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">\n'
            '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>\n'
            '  <Default Extension="xml" ContentType="application/xml"/>\n'
            '  <Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>\n'
            f"{sheet_overrides}\n"
            "</Types>",
        )

        # _rels/.rels
        zf.writestr(
            "_rels/.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            '  <Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/>\n'
            "</Relationships>",
        )

        # xl/workbook.xml
        sheet_tags = "\n".join(
            f'    <sheet name="{_xml_escape(ws.name[:31] or f"Sheet{i + 1}")}" '
            f'sheetId="{i + 1}" r:id="rId{i + 1}"/>'
            for i, ws in enumerate(wb.sheets)
        )
        zf.writestr(
            "xl/workbook.xml",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">\n'
            "  <sheets>\n"
            f"{sheet_tags}\n"
            "  </sheets>\n"
            "</workbook>",
        )

        # xl/_rels/workbook.xml.rels
        rel_tags = "\n".join(
            f'  <Relationship Id="rId{i + 1}" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{i + 1}.xml"/>'
            for i in range(len(wb.sheets))
        )
        zf.writestr(
            "xl/_rels/workbook.xml.rels",
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">\n'
            f"{rel_tags}\n"
            "</Relationships>",
        )

        # xl/worksheets/sheetN.xml
        for i, ws in enumerate(wb.sheets):
            rows_by_r: Dict[int, List[Tuple[int, Cell]]] = {}
            for (r, c), cell in ws.cells.items():
                if cell.raw != "":
                    rows_by_r.setdefault(r, []).append((c, cell))

            row_xmls: List[str] = []
            for r in sorted(rows_by_r.keys()):
                c_xmls: List[str] = []
                for c, cell in sorted(rows_by_r[r], key=lambda x: x[0]):
                    ref = format_cell_ref(r, c)
                    if is_formula_raw(cell.raw):
                        g_formula = to_gsheets_formula(cell.raw)
                        if g_formula.startswith("="):
                            g_formula = g_formula[1:]
                        v_tag = ""
                        if isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                            v_tag = f"<v>{cell.value}</v>"
                        c_xmls.append(f'<c r="{ref}"><f>{_xml_escape(g_formula)}</f>{v_tag}</c>')
                    else:
                        num = try_parse_number(cell.raw)
                        if num is not None:
                            c_xmls.append(f'<c r="{ref}"><v>{num}</v></c>')
                        else:
                            txt = to_gsheets_formula(cell.raw)
                            c_xmls.append(
                                f'<c r="{ref}" t="inlineStr"><is><t>{_xml_escape(txt)}</t></is></c>'
                            )
                row_xmls.append(f'    <row r="{r + 1}">{"".join(c_xmls)}</row>')

            sheet_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">\n'
                "  <sheetData>\n"
                + "\n".join(row_xmls)
                + "\n  </sheetData>\n"
                "</worksheet>"
            )
            zf.writestr(f"xl/worksheets/sheet{i + 1}.xml", sheet_xml)

    return buf.getvalue()


def import_xlsx_bytes_to_workbook(data: bytes, title: str = "IMPORTED") -> Workbook:
    """
    Parses a multi-sheet .xlsx binary (from Excel or Google Drive export) using
    only Python's zipfile + ElementTree, preserving sheet names, formulas (<f>),
    numbers, booleans, inline strings, and sharedStrings.
    """
    wb = Workbook(title=title)
    wb.sheets = []

    with zipfile.ZipFile(io.BytesIO(data), "r") as zf:
        names_in_zip = set(zf.namelist())

        # 1. Load sharedStrings.xml if present
        shared_strings: List[str] = []
        if "xl/sharedStrings.xml" in names_in_zip:
            ss_root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in ss_root.iter():
                if si.tag.endswith("}si") or si.tag == "si":
                    parts = []
                    for t in si.iter():
                        if t.tag.endswith("}t") or t.tag == "t":
                            if t.text:
                                parts.append(t.text)
                    shared_strings.append("".join(parts))

        # 2. Resolve sheet names and target XML paths from workbook.xml + rels
        rel_map: Dict[str, str] = {}
        if "xl/_rels/workbook.xml.rels" in names_in_zip:
            rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            for rel in rels_root:
                rid = rel.attrib.get("Id", "")
                target = rel.attrib.get("Target", "")
                if rid and target:
                    if not target.startswith("xl/"):
                        target = "xl/" + target.lstrip("/")
                    rel_map[rid] = target

        sheet_specs: List[Tuple[str, str]] = []
        if "xl/workbook.xml" in names_in_zip:
            wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
            for el in wb_root.iter():
                if el.tag.endswith("}sheet") or el.tag == "sheet":
                    sname = el.attrib.get("name", f"Sheet{len(sheet_specs) + 1}")
                    rid = ""
                    for ak, av in el.attrib.items():
                        if ak.endswith("}id") or ak == "r:id":
                            rid = av
                            break
                    target = rel_map.get(rid, f"xl/worksheets/sheet{len(sheet_specs) + 1}.xml")
                    sheet_specs.append((sname, target))

        if not sheet_specs:
            for nm in sorted(names_in_zip):
                if nm.startswith("xl/worksheets/sheet") and nm.endswith(".xml"):
                    sheet_specs.append((f"Sheet{len(sheet_specs) + 1}", nm))

        # 3. Parse each sheet
        for sname, spath in sheet_specs:
            if spath not in names_in_zip:
                continue
            ws = Worksheet(sname)
            s_root = ET.fromstring(zf.read(spath))
            for c_el in s_root.iter():
                if not (c_el.tag.endswith("}c") or c_el.tag == "c"):
                    continue
                ref = c_el.attrib.get("r", "")
                parsed = parse_cell_ref(ref)
                if not parsed:
                    continue
                row, col = parsed[0], parsed[1]
                ctype = c_el.attrib.get("t", "")

                f_text = None
                v_text = None
                inline_text = None
                for child in c_el:
                    tag = child.tag.split("}")[-1]
                    if tag == "f" and child.text:
                        f_text = child.text
                    elif tag == "v" and child.text is not None:
                        v_text = child.text
                    elif tag == "is":
                        parts = [
                            t.text for t in child.iter()
                            if (t.tag.endswith("}t") or t.tag == "t") and t.text
                        ]
                        inline_text = "".join(parts)

                if f_text:
                    ws.set_raw(row, col, f"={f_text}" if not f_text.startswith("=") else f_text)
                elif ctype == "s" and v_text is not None:
                    try:
                        sval = shared_strings[int(v_text)]
                    except Exception:
                        sval = ""
                    ws.set_raw(row, col, sval)
                elif ctype == "inlineStr" and inline_text is not None:
                    ws.set_raw(row, col, inline_text)
                elif ctype == "b" and v_text is not None:
                    ws.set_raw(row, col, "TRUE" if v_text == "1" else "FALSE")
                elif v_text is not None:
                    ws.set_raw(row, col, v_text)
            wb.sheets.append(ws)

    if not wb.sheets:
        wb.sheets = [Worksheet("Sheet1")]
    wb.recalculate()
    return wb


def build_sample_workbook() -> Workbook:
    """Creates a classic Lotus 1-2-3 sample budget workbook so first launch looks iconic."""
    wb = Workbook(title="Q4_BUDGET")
    ws = wb.active_sheet
    ws.name = "Budget"
    ws.set_col_width(0, 20)
    ws.set_col_width(1, 12)
    ws.set_col_width(2, 12)
    ws.set_col_width(3, 12)
    ws.set_col_width(4, 13)
    ws.set_col_width(5, 10)

    ws.set_raw(0, 0, "'LOTUS 1-2-3 FOR OV/400 - Q4 OPERATING BUDGET")
    ws.set_raw(1, 0, r"\=")
    for c in range(1, 6):
        ws.set_raw(1, c, r"\=")

    headers = ["'Category", '"October', '"November', '"December', '"Q4 Total', '"% Share']
    for c, h in enumerate(headers):
        ws.set_raw(3, c, h)
    for c in range(6):
        ws.set_raw(4, c, r"\-")

    rows = [
        ("Hardware & MDA", 1250, 980, 1420),
        ("Cloud & Hosting", 340, 340, 365),
        ("Software Licenses", 520, 480, 520),
        ("Communications", 185, 190, 195),
        ("Office Supplies", 210, 145, 290),
    ]
    for i, (cat, m1, m2, m3) in enumerate(rows):
        r = 5 + i
        ws.set_raw(r, 0, f"'{cat}")
        ws.set_raw(r, 1, str(m1), "C0")
        ws.set_raw(r, 2, str(m2), "C0")
        ws.set_raw(r, 3, str(m3), "C0")
        ws.set_raw(r, 4, f"@SUM(B{r + 1}..D{r + 1})", "C0")
        ws.set_raw(r, 5, f"+E{r + 1}/$E$12", "P1")

    for c in range(6):
        ws.set_raw(10, c, r"\-")

    ws.set_raw(11, 0, "'TOTAL EXPENSES")
    for c, col_l in [(1, "B"), (2, "C"), (3, "D"), (4, "E")]:
        ws.set_raw(11, c, f"@SUM({col_l}6..{col_l}10)", "C0")
    ws.set_raw(11, 5, "@SUM(F6..F10)", "P1")

    ws.set_raw(13, 0, "'Monthly Average")
    ws.set_raw(13, 1, "@AVG(B6..B10)", "C2")
    ws.set_raw(13, 2, "@AVG(C6..C10)", "C2")
    ws.set_raw(13, 3, "@AVG(D6..D10)", "C2")
    ws.set_raw(13, 4, "@AVG(E6..E10)", "C2")

    ws.set_raw(14, 0, "'Highest Line Item")
    ws.set_raw(14, 4, "@MAX(E6..E10)", "C0")

    ws.set_raw(16, 0, "'Tip: Press / for Lotus Menu, F10 for Google Sheets Sync, F1 for Help")

    # Second tab showcasing multi-sheet Release 3 + Google Sheets tabs
    ws2 = wb.add_sheet("Summary")
    ws2.set_col_width(0, 22)
    ws2.set_col_width(1, 15)
    ws2.set_raw(0, 0, "'EXECUTIVE SUMMARY")
    ws2.set_raw(1, 0, r"\=")
    ws2.set_raw(1, 1, r"\=")
    ws2.set_raw(3, 0, "'Q4 Total Spend")
    ws2.set_raw(3, 1, "=Budget!E12", "C2")
    ws2.set_raw(4, 0, "'Line Item Count")
    ws2.set_raw(4, 1, "=COUNT(Budget!E6:E10)", "F0")
    ws2.set_raw(5, 0, "'Avg Category Spend")
    ws2.set_raw(5, 1, "=AVERAGE(Budget!E6:E10)", "C2")

    wb.active_sheet_idx = 0
    wb.recalculate()
    wb.dirty = False
    return wb
