"""
Spreadsheet Search MCP Server

An MCP server for searching spreadsheets (CSV, CSV.GZ, Excel) without
loading the entire content into an AI agent's context.
"""

from __future__ import annotations

import os
from typing import Any

import pandas as pd
from fastmcp import FastMCP
from pydantic import BaseModel

mcp = FastMCP(
    name="Spreadsheet Search",
    instructions=(
        "Load a spreadsheet with 'load', then search with 'find_one' or 'find_all', "
        "and retrieve cell or row contents with 'get_single_cell', 'get_single_row', "
        "'get_all_cells', or 'get_all_rows'."
    ),
)

# ── Module-level state ────────────────────────────────────────────────────────
_sheets: dict[str, pd.DataFrame] = {}
_sheet_order: list[str] = []
# (sheet_idx, row_idx, col_idx) — all 0-based
_search_pos: tuple[int, int, int] = (0, 0, 0)


# ── Column-letter helpers ─────────────────────────────────────────────────────

def _col_idx_to_letter(idx: int) -> str:
    """Convert a 0-based column index to an Excel-style column letter (A, B, ..., AA, ...)."""
    letters = ""
    n = idx + 1
    while n > 0:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _col_letter_to_idx(letter: str) -> int:
    """Convert an Excel-style column letter to a 0-based column index."""
    result = 0
    for ch in letter.upper():
        result = result * 26 + (ord(ch) - 64)
    return result - 1


# ── Value helper ──────────────────────────────────────────────────────────────

def _safe_value(v: Any) -> Any:
    """Convert pandas NA / NaN to None; leave all other values unchanged."""
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v


# ── Core logic helpers (used by both tools and tests) ─────────────────────────

def _get_cell(sheet: str, row: int, col: str) -> dict[str, Any]:
    """Return {'value': …} for the given cell, or {'error': …} on invalid input."""
    if sheet not in _sheets:
        return {"error": f"Sheet '{sheet}' not found."}
    df = _sheets[sheet]
    row_idx = row - 1
    col_idx = _col_letter_to_idx(col)
    if not (0 <= row_idx < len(df)):
        return {"error": f"Row {row} is out of range (sheet has {len(df)} rows)."}
    if not (0 <= col_idx < len(df.columns)):
        return {"error": f"Column '{col}' is out of range (sheet has {len(df.columns)} columns)."}
    return {"value": _safe_value(df.iloc[row_idx, col_idx])}


def _get_row(sheet: str, row: int) -> dict[str, Any]:
    """Return {'values': […]} for the given row, or {'error': …} on invalid input."""
    if sheet not in _sheets:
        return {"error": f"Sheet '{sheet}' not found."}
    df = _sheets[sheet]
    row_idx = row - 1
    if not (0 <= row_idx < len(df)):
        return {"error": f"Row {row} is out of range (sheet has {len(df)} rows)."}
    return {"values": [_safe_value(v) for v in df.iloc[row_idx].tolist()]}


# ── Pydantic model for list-based tools ───────────────────────────────────────

class CellRef(BaseModel):
    """A reference to a single cell in the loaded spreadsheet."""

    sheet: str = ""
    row: int
    col: str


# ── MCP tools ─────────────────────────────────────────────────────────────────

@mcp.tool
def load(path: str) -> str:
    """
    Load a spreadsheet into memory.

    Supports CSV (.csv), gzip-compressed CSV (.csv.gz), and Excel files
    (.xlsx, .xlsm, .xls).  Must be called before any search or retrieval
    operation.  Returns a success or failure message but NEVER the
    spreadsheet contents.
    """
    global _sheets, _sheet_order, _search_pos

    if not os.path.exists(path):
        return f"Error: file not found: {path}"

    try:
        p = path.lower()
        if p.endswith(".xlsx") or p.endswith(".xls") or p.endswith(".xlsm"):
            all_sheets: dict[str, pd.DataFrame] = pd.read_excel(
                path, sheet_name=None, header=None
            )
            _sheet_order = list(all_sheets.keys())
            _sheets = dict(all_sheets)
        elif p.endswith(".csv") or p.endswith(".csv.gz"):
            df = pd.read_csv(path, header=None, dtype=str, keep_default_na=True)
            _sheet_order = [""]
            _sheets = {"": df}
        else:
            return (
                "Error: unsupported file format. "
                "Supported formats: .csv, .csv.gz, .xlsx, .xlsm, .xls"
            )

        _search_pos = (0, 0, 0)

        if _sheet_order == [""]:
            sheet_desc = "CSV"
        else:
            sheet_desc = (
                f"{len(_sheet_order)} sheet(s): "
                + ", ".join(repr(s) for s in _sheet_order)
            )
        total = sum(len(df) for df in _sheets.values())
        return f"Loaded {total} row(s), {sheet_desc}."
    except Exception as exc:
        return f"Error loading file: {exc}"


@mcp.tool
def find_one(search: str) -> dict[str, Any]:
    """
    Find the next cell whose string representation contains *search*
    (case-insensitive).

    Returns a dict with a ``result`` key containing a ``CellRef`` when a match
    is found, or ``None`` when no match exists.  Returns a dict with an
    ``error`` key when no spreadsheet is loaded.

    Successive calls advance the internal cursor so that repeated calls
    iterate through all matching cells.  When the end of the spreadsheet is
    reached without a match, ``result`` is ``None`` and the cursor resets to
    the beginning.
    """
    global _search_pos

    if not _sheets:
        return {"error": "No spreadsheet loaded. Call load() first."}

    needle = search.lower()
    si, ri, ci = _search_pos

    for sheet_idx in range(si, len(_sheet_order)):
        sheet = _sheet_order[sheet_idx]
        df = _sheets[sheet]
        row_start = ri if sheet_idx == si else 0

        for row_idx in range(row_start, len(df)):
            col_start = ci if (sheet_idx == si and row_idx == ri) else 0

            for col_idx in range(col_start, len(df.columns)):
                v = df.iloc[row_idx, col_idx]
                if pd.notna(v) and needle in str(v).lower():
                    # Advance cursor past this cell
                    next_ci = col_idx + 1
                    if next_ci >= len(df.columns):
                        next_ri = row_idx + 1
                        if next_ri >= len(df):
                            next_si = sheet_idx + 1
                            if next_si >= len(_sheet_order):
                                next_si = 0
                            _search_pos = (next_si, 0, 0)
                        else:
                            _search_pos = (sheet_idx, next_ri, 0)
                    else:
                        _search_pos = (sheet_idx, row_idx, next_ci)

                    return {
                        "result": CellRef(
                            sheet=sheet,
                            row=row_idx + 1,
                            col=_col_idx_to_letter(col_idx),
                        )
                    }

    _search_pos = (0, 0, 0)
    return {"result": None}


@mcp.tool
def find_all(search: str) -> dict[str, Any]:
    """
    Search the entire spreadsheet for cells whose string representation
    contains *search* (case-insensitive).

    Returns a dict with a ``result`` key containing a list of ``CellRef``
    objects.  All worksheets in an Excel file are searched.  Returns a dict
    with an ``error`` key when no spreadsheet is loaded.
    """
    if not _sheets:
        return {"error": "No spreadsheet loaded. Call load() first."}

    needle = search.lower()
    results: list[CellRef] = []

    for sheet in _sheet_order:
        df = _sheets[sheet]
        for row_idx in range(len(df)):
            for col_idx in range(len(df.columns)):
                v = df.iloc[row_idx, col_idx]
                if pd.notna(v) and needle in str(v).lower():
                    results.append(
                        CellRef(
                            sheet=sheet,
                            row=row_idx + 1,
                            col=_col_idx_to_letter(col_idx),
                        )
                    )

    return {"result": results}


@mcp.tool
def get_single_cell(cell: CellRef) -> dict[str, Any]:
    """
    Return the value of the cell identified by the given ``CellRef``.

    Result dict contains ``value`` on success or ``error`` on failure.
    For CSV files set ``sheet`` to an empty string in the ``CellRef``.
    """
    if not _sheets:
        return {"error": "No spreadsheet loaded. Call load() first."}
    return _get_cell(cell.sheet, cell.row, cell.col)


@mcp.tool
def get_single_row(cell: CellRef) -> dict[str, Any]:
    """
    Return all values in the row identified by the given ``CellRef``.

    The ``col`` field of the ``CellRef`` is ignored; the entire row is
    always returned.

    Result dict contains ``values`` (a list) on success or ``error`` on
    failure.  For CSV files set ``sheet`` to an empty string in the
    ``CellRef``.
    """
    if not _sheets:
        return {"error": "No spreadsheet loaded. Call load() first."}
    return _get_row(cell.sheet, cell.row)


@mcp.tool
def get_all_cells(cells: list[CellRef]) -> list[dict[str, Any]]:
    """
    Return the value of each cell in the provided list of cell references.

    Each item in the result corresponds to the cell reference at the same
    index.  Individual items contain ``value`` on success or ``error`` on
    failure.
    """
    if not _sheets:
        return [{"error": "No spreadsheet loaded. Call load() first."}]
    return [_get_cell(c.sheet, c.row, c.col) for c in cells]


@mcp.tool
def get_all_rows(cells: list[CellRef]) -> list[list[Any]]:
    """
    Return the full row for each cell reference in the provided list.

    The result is a list of lists — one list of values per input cell
    reference (the ``col`` field of each reference is ignored).
    """
    if not _sheets:
        return [[{"error": "No spreadsheet loaded. Call load() first."}]]
    rows: list[list[Any]] = []
    for c in cells:
        result = _get_row(c.sheet, c.row)
        if "error" in result:
            rows.append([result])
        else:
            rows.append(result["values"])
    return rows


# ── Entry point ───────────────────────────────────────────────────────────────

def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
