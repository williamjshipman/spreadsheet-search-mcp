"""Tests for the spreadsheet-search MCP server."""

from __future__ import annotations

import os
import sys

import pytest

# Ensure the src layout is importable when running tests directly
SRC_DIR = os.path.join(os.path.dirname(__file__), "..", "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR)

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")


def _reset() -> None:
    """Reset the module-level state between tests."""
    import spreadsheet_search_mcp.server as srv

    srv._sheets = {}
    srv._sheet_order = []
    srv._search_pos = (0, 0, 0)


# ── Column-letter helpers ─────────────────────────────────────────────────────


def test_col_idx_to_letter():
    from spreadsheet_search_mcp.server import _col_idx_to_letter

    assert _col_idx_to_letter(0) == "A"
    assert _col_idx_to_letter(1) == "B"
    assert _col_idx_to_letter(25) == "Z"
    assert _col_idx_to_letter(26) == "AA"
    assert _col_idx_to_letter(27) == "AB"
    assert _col_idx_to_letter(51) == "AZ"
    assert _col_idx_to_letter(52) == "BA"


def test_col_letter_to_idx():
    from spreadsheet_search_mcp.server import _col_letter_to_idx

    assert _col_letter_to_idx("A") == 0
    assert _col_letter_to_idx("B") == 1
    assert _col_letter_to_idx("Z") == 25
    assert _col_letter_to_idx("AA") == 26
    assert _col_letter_to_idx("AB") == 27
    assert _col_letter_to_idx("BA") == 52
    # Case-insensitive
    assert _col_letter_to_idx("a") == 0
    assert _col_letter_to_idx("aa") == 26


def test_roundtrip_col():
    from spreadsheet_search_mcp.server import _col_idx_to_letter, _col_letter_to_idx

    for i in range(100):
        assert _col_letter_to_idx(_col_idx_to_letter(i)) == i


# ── load ─────────────────────────────────────────────────────────────────────


def test_load_csv():
    _reset()
    from spreadsheet_search_mcp.server import load

    result = load(os.path.join(DATA_DIR, "test.csv"))
    assert "Error" not in result
    assert "5" in result  # 5 rows


def test_load_csv_gz():
    _reset()
    from spreadsheet_search_mcp.server import load

    result = load(os.path.join(DATA_DIR, "test.csv.gz"))
    assert "Error" not in result
    assert "5" in result


def test_load_xlsx():
    _reset()
    from spreadsheet_search_mcp.server import load

    result = load(os.path.join(DATA_DIR, "test.xlsx"))
    assert "Error" not in result
    assert "People" in result
    assert "Products" in result


def test_load_nonexistent():
    _reset()
    from spreadsheet_search_mcp.server import load

    result = load("/nonexistent/path/file.csv")
    assert "Error" in result


def test_load_unsupported_format():
    _reset()
    from spreadsheet_search_mcp.server import load

    result = load("/tmp/file.txt")
    assert "Error" in result


def test_load_resets_search_position():
    _reset()
    import spreadsheet_search_mcp.server as srv
    from spreadsheet_search_mcp.server import load

    load(os.path.join(DATA_DIR, "test.csv"))
    srv._search_pos = (0, 3, 2)
    load(os.path.join(DATA_DIR, "test.csv"))
    assert srv._search_pos == (0, 0, 0)


# ── find_one ─────────────────────────────────────────────────────────────────


def test_find_one_csv_first_match():
    _reset()
    from spreadsheet_search_mcp.server import find_one, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result = find_one("alice")
    assert result is not None
    assert result.sheet == ""
    assert result.row == 2  # row 1 is header "Name", row 2 is "Alice"
    assert result.col == "A"


def test_find_one_case_insensitive():
    _reset()
    from spreadsheet_search_mcp.server import find_one, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result = find_one("ALICE")
    assert result is not None
    assert result.row == 2


def test_find_one_advances_cursor():
    _reset()
    from spreadsheet_search_mcp.server import find_one, load

    load(os.path.join(DATA_DIR, "test.csv"))
    r1 = find_one("alice")
    r2 = find_one("alice")
    assert r1 is not None
    assert r2 is not None
    # Second result should be different from first (alice_duplicate row)
    assert r1.row != r2.row or r1.col != r2.col


def test_find_one_not_found_resets():
    _reset()
    import spreadsheet_search_mcp.server as srv
    from spreadsheet_search_mcp.server import find_one, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result = find_one("zzznomatchzzz")
    assert result is None
    assert srv._search_pos == (0, 0, 0)


def test_find_one_no_spreadsheet_loaded():
    _reset()
    from spreadsheet_search_mcp.server import find_one

    result = find_one("test")
    assert result is None


def test_find_one_excel_multiple_sheets():
    _reset()
    from spreadsheet_search_mcp.server import find_one, load

    load(os.path.join(DATA_DIR, "test.xlsx"))
    # "alice" appears in People sheet row 2 (Alice) and Products sheet (alice_excel)
    matches = []
    for _ in range(10):
        r = find_one("alice")
        if r is None:
            break
        matches.append(r)

    sheets_found = {m.sheet for m in matches}
    # Should have found alice in both sheets
    assert "People" in sheets_found
    assert "Products" in sheets_found


# ── find_all ─────────────────────────────────────────────────────────────────


def test_find_all_csv():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, find_all, load

    load(os.path.join(DATA_DIR, "test.csv"))
    results = find_all("alice")
    # "Alice" in row 2 col A AND "alice_duplicate" in row 5 col A
    assert len(results) == 2
    assert all(isinstance(r, CellRef) for r in results)
    for r in results:
        assert r.sheet == ""
        assert r.col == "A"


def test_find_all_returns_empty_list_when_no_match():
    _reset()
    from spreadsheet_search_mcp.server import find_all, load

    load(os.path.join(DATA_DIR, "test.csv"))
    results = find_all("zzznomatchzzz")
    assert results == []


def test_find_all_no_spreadsheet_loaded():
    _reset()
    from spreadsheet_search_mcp.server import find_all

    results = find_all("test")
    assert results == []


def test_find_all_excel_searches_all_sheets():
    _reset()
    from spreadsheet_search_mcp.server import find_all, load

    load(os.path.join(DATA_DIR, "test.xlsx"))
    results = find_all("alice")
    sheets_found = {r.sheet for r in results}
    assert "People" in sheets_found
    assert "Products" in sheets_found


# ── get_single_cell ───────────────────────────────────────────────────────────


def test_get_single_cell_csv():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_cell, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result = get_single_cell(CellRef(sheet="", row=1, col="A"))
    assert result["value"] == "Name"


def test_get_single_cell_csv_data_row():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_cell, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result = get_single_cell(CellRef(sheet="", row=2, col="B"))
    assert result["value"] == "New York"


def test_get_single_cell_invalid_row():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_cell, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result = get_single_cell(CellRef(sheet="", row=999, col="A"))
    assert "error" in result


def test_get_single_cell_invalid_col():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_cell, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result = get_single_cell(CellRef(sheet="", row=1, col="ZZ"))
    assert "error" in result


def test_get_single_cell_invalid_sheet():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_cell, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result = get_single_cell(CellRef(sheet="NonExistent", row=1, col="A"))
    assert "error" in result


def test_get_single_cell_no_spreadsheet():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_cell

    result = get_single_cell(CellRef(sheet="", row=1, col="A"))
    assert "error" in result


def test_get_single_cell_excel():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_cell, load

    load(os.path.join(DATA_DIR, "test.xlsx"))
    result = get_single_cell(CellRef(sheet="People", row=1, col="A"))
    assert result["value"] == "Name"


# ── get_single_row ────────────────────────────────────────────────────────────


def test_get_single_row_csv():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_row, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result = get_single_row(CellRef(sheet="", row=1, col="A"))
    assert result["values"] == ["Name", "City", "Score"]


def test_get_single_row_col_ignored():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_row, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result_a = get_single_row(CellRef(sheet="", row=2, col="A"))
    result_b = get_single_row(CellRef(sheet="", row=2, col="C"))
    assert result_a["values"] == result_b["values"]


def test_get_single_row_invalid():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_row, load

    load(os.path.join(DATA_DIR, "test.csv"))
    result = get_single_row(CellRef(sheet="", row=999, col="A"))
    assert "error" in result


def test_get_single_row_no_spreadsheet():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_single_row

    result = get_single_row(CellRef(sheet="", row=1, col="A"))
    assert "error" in result


# ── get_all_cells ─────────────────────────────────────────────────────────────


def test_get_all_cells_csv():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_all_cells, load

    load(os.path.join(DATA_DIR, "test.csv"))
    cells = [
        CellRef(sheet="", row=1, col="A"),
        CellRef(sheet="", row=2, col="B"),
        CellRef(sheet="", row=3, col="C"),
    ]
    results = get_all_cells(cells)
    assert len(results) == 3
    assert results[0]["value"] == "Name"
    assert results[1]["value"] == "New York"
    assert results[2]["value"] == "87"


def test_get_all_cells_no_spreadsheet():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_all_cells

    results = get_all_cells([CellRef(sheet="", row=1, col="A")])
    assert len(results) == 1
    assert "error" in results[0]


# ── get_all_rows ──────────────────────────────────────────────────────────────


def test_get_all_rows_csv():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_all_rows, load

    load(os.path.join(DATA_DIR, "test.csv"))
    cells = [
        CellRef(sheet="", row=1, col="A"),
        CellRef(sheet="", row=2, col="C"),
    ]
    results = get_all_rows(cells)
    assert len(results) == 2
    assert results[0] == ["Name", "City", "Score"]
    assert results[1] == ["Alice", "New York", "95"]


def test_get_all_rows_col_ignored():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_all_rows, load

    load(os.path.join(DATA_DIR, "test.csv"))
    cells = [CellRef(sheet="", row=2, col="Z")]
    results = get_all_rows(cells)
    assert results[0] == ["Alice", "New York", "95"]


def test_get_all_rows_no_spreadsheet():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_all_rows

    results = get_all_rows([CellRef(sheet="", row=1, col="A")])
    assert len(results) == 1
    assert "error" in results[0][0]


def test_get_all_rows_excel_multiple_sheets():
    _reset()
    from spreadsheet_search_mcp.server import CellRef, get_all_rows, load

    load(os.path.join(DATA_DIR, "test.xlsx"))
    cells = [
        CellRef(sheet="People", row=1, col="A"),
        CellRef(sheet="Products", row=1, col="A"),
    ]
    results = get_all_rows(cells)
    assert results[0] == ["Name", "City", "Score"]
    assert results[1] == ["Product", "Price", "Category"]
