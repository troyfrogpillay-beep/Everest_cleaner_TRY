import io
import re
import difflib
from copy import copy
from dataclasses import dataclass, field
from datetime import datetime
from typing import Tuple, List, Dict, Any, Optional, NamedTuple

import openpyxl
from openpyxl.utils import get_column_letter
import pandas as pd
from email_validator import validate_email, EmailNotValidError
import streamlit as st

st.set_page_config(page_title="Everest Data Cleaner", layout="wide")

# ==============================================================================
# PRE-CLEANING STEPS: CLIENT SECTION EXTRACTION (ADDITIVE - NEVER TOUCHES THE
# SOURCE WORKSHEET OR THE DATAFRAME USED BY THE CLEANER)
# ==============================================================================

SUMMARY_SHEET_NAME = "Client Premium Summary"
DETAIL_SHEET_NAME = "Client Detailed Schedule"

SUMMARY_HEADERS = ["Payment Date", "Month End", "Policy Name", "Client Name", "Risk Payment"]
DETAIL_HEADERS = ["Client Name", "Member ID", "Surname", "First Name", "Policy Description", "Risk Premium"]


class ExtractionError(Exception):
    """Raised with a user-readable message when an extraction cannot proceed."""


class CellSnapshot(NamedTuple):
    """Independent copy of one source cell (value + type + formatting)."""
    value: Any
    data_type: str
    number_format: str
    font: Any
    fill: Any
    border: Any
    alignment: Any
    protection: Any


@dataclass
class ExtractedSection:
    section_label: str
    default_sheet_name: str
    source_sheet: str
    header_row: int
    first_row: int
    last_row: int
    first_col: int
    last_col: int
    header: List[CellSnapshot] = field(default_factory=list)
    rows: List[List[CellSnapshot]] = field(default_factory=list)
    source_rows: List[int] = field(default_factory=list)
    col_widths: Dict[int, float] = field(default_factory=dict)

    @property
    def section_range(self) -> str:
        return (f"'{self.source_sheet}'!{get_column_letter(self.first_col)}{self.header_row}:"
                f"{get_column_letter(self.last_col)}{self.last_row}")


def _norm_header(val: Any) -> str:
    if not isinstance(val, str):
        return ""
    return re.sub(r"\s+", " ", val.strip()).casefold()


def _norm_client(val: Any) -> str:
    if val is None:
        return ""
    return str(val).strip().casefold()


def _is_blank(val: Any) -> bool:
    return val is None or (isinstance(val, str) and not val.strip())


def _sheet_grid(ws) -> List[tuple]:
    return list(ws.iter_rows(values_only=True))


def _header_positions(row: tuple) -> Dict[str, List[int]]:
    pos: Dict[str, List[int]] = {}
    for idx, v in enumerate(row, start=1):
        key = _norm_header(v)
        if key:
            pos.setdefault(key, []).append(idx)
    return pos


def _matching_header_rows(grid: List[tuple], required: List[str]) -> List[Tuple[int, Dict[str, List[int]]]]:
    req = [_norm_header(h) for h in required]
    hits = []
    for r_idx, row in enumerate(grid, start=1):
        pos = _header_positions(row)
        if all(h in pos for h in req):
            hits.append((r_idx, pos))
    return hits


def _best_partial(grid: List[tuple], required: List[str]) -> Tuple[int, int, List[str]]:
    req = [(h, _norm_header(h)) for h in required]
    best = (0, 0, [h for h, _ in req])
    for r_idx, row in enumerate(grid, start=1):
        pos = _header_positions(row)
        found = [h for h, n in req if n in pos]
        if len(found) > best[1]:
            best = (r_idx, len(found), [h for h, n in req if n not in pos])
    return best


def _extract_section(wb, client_name: str, required: List[str], other_required: List[str],
                     label: str, default_sheet_name: str, stop_at_first_blank_row: bool) -> ExtractedSection:
    if client_name is None or not str(client_name).strip():
        raise ExtractionError(f"Client name is blank. Enter a client name to extract the {label}.")
    target = _norm_client(client_name)

    found: List[Tuple[Any, List[tuple], int, Dict[str, List[int]]]] = []
    best_overall = (None, 0, 0, list(required))
    for ws in wb.worksheets:
        grid = _sheet_grid(ws)
        for r_idx, pos in _matching_header_rows(grid, required):
            found.append((ws, grid, r_idx, pos))
        r, cnt, missing = _best_partial(grid, required)
        if cnt > best_overall[2]:
            best_overall = (ws.title, r, cnt, missing)

    if not found:
        if best_overall[2] >= 2:
            raise ExtractionError(
                f"Could not find the {label} header. The closest row is '{best_overall[0]}' row "
                f"{best_overall[1]}, which is missing: {', '.join(best_overall[3])}. "
                f"Required headers: {', '.join(required)}.")
        raise ExtractionError(
            f"Could not find the {label} header. No row contains the required headers: {', '.join(required)}.")

    if len(found) > 1:
        where = ", ".join(f"'{ws.title}' row {r}" for ws, _, r, _ in found)
        raise ExtractionError(
            f"Ambiguous {label} header: the required headers appear in more than one place ({where}). "
            f"Cannot tell which one to use.")

    ws, grid, header_row, pos = found[0]
    dupes = [h for h in required if len(pos[_norm_header(h)]) > 1]
    if dupes:
        raise ExtractionError(
            f"Ambiguous {label} header in '{ws.title}' row {header_row}: "
            f"duplicate column(s) {', '.join(dupes)}.")

    header_row_vals = grid[header_row - 1]
    used_cols = [i for i, v in enumerate(header_row_vals, start=1) if not _is_blank(v)]
    first_col, last_col = min(used_cols), max(used_cols)
    client_col = pos[_norm_header("Client Name")][0]

    other_headers = [r for r, _ in _matching_header_rows(grid, other_required) if r > header_row]
    limit = (min(other_headers) - 1) if other_headers else len(grid)
    first_row = header_row + 1

    def row_blank(r: int) -> bool:
        return all(_is_blank(v) for v in grid[r - 1][first_col - 1:last_col])

    if stop_at_first_blank_row:
        end = first_row - 1
        for r in range(first_row, limit + 1):
            if row_blank(r):
                break
            end = r
    else:
        end = limit
        while end >= first_row and row_blank(end):
            end -= 1

    section = ExtractedSection(
        section_label=label, default_sheet_name=default_sheet_name, source_sheet=ws.title,
        header_row=header_row, first_row=first_row, last_row=max(end, header_row),
        first_col=first_col, last_col=last_col)

    def snap(r: int, c: int) -> CellSnapshot:
        cell = ws.cell(row=r, column=c)
        return CellSnapshot(cell.value, cell.data_type, cell.number_format, copy(cell.font),
                            copy(cell.fill), copy(cell.border), copy(cell.alignment), copy(cell.protection))

    section.header = [snap(header_row, c) for c in range(first_col, last_col + 1)]
    seen_clients: Dict[str, str] = {}
    for r in range(first_row, end + 1):
        row = grid[r - 1]
        raw = row[client_col - 1] if client_col - 1 < len(row) else None
        key = _norm_client(raw)
        if key:
            seen_clients.setdefault(key, str(raw).strip())
        if key == target:
            section.rows.append([snap(r, c) for c in range(first_col, last_col + 1)])
            section.source_rows.append(r)

    if not section.rows:
        close = difflib.get_close_matches(target, list(seen_clients.keys()), n=3, cutoff=0.6)
        hint = f" Did you mean: {'; '.join(seen_clients[k] for k in close)}?" if close else ""
        raise ExtractionError(
            f"Client '{str(client_name).strip()}' was not found in the {label} "
            f"({section.section_range}, {end - first_row + 1} rows scanned).{hint}")

    for c in range(first_col, last_col + 1):
        dim = ws.column_dimensions.get(get_column_letter(c))
        if dim is not None and dim.width:
            section.col_widths[c - first_col + 1] = dim.width
    return section


def _open_source(source_bytes: bytes):
    try:
        return openpyxl.load_workbook(io.BytesIO(source_bytes), data_only=True)
    except Exception as e:
        raise ExtractionError(f"Could not open the file for extraction (.xlsx required): {e}")


def extract_client_premium_summary(source_bytes: bytes, client_name: str, workbook=None) -> ExtractedSection:
    wb = workbook if workbook is not None else _open_source(source_bytes)
    return _extract_section(wb, client_name, SUMMARY_HEADERS, DETAIL_HEADERS,
                            "premium summary", SUMMARY_SHEET_NAME, stop_at_first_blank_row=True)


def extract_client_detailed_schedule(source_bytes: bytes, client_name: str, workbook=None) -> ExtractedSection:
    wb = workbook if workbook is not None else _open_source(source_bytes)
    return _extract_section(wb, client_name, DETAIL_HEADERS, SUMMARY_HEADERS,
                            "detailed schedule", DETAIL_SHEET_NAME, stop_at_first_blank_row=False)


def run_client_extractions(source_bytes: bytes, client_name: str) -> Tuple[List[ExtractedSection], List[str]]:
    errors: List[str] = []
    if client_name is None or not str(client_name).strip():
        return [], ["Client name is blank. Enter a client name before running."]
    try:
        wb = _open_source(source_bytes)
    except ExtractionError as e:
        return [], [str(e)]
    sections: List[ExtractedSection] = []
    for fn in (extract_client_premium_summary, extract_client_detailed_schedule):
        try:
            sections.append(fn(source_bytes, client_name, workbook=wb))
        except ExtractionError as e:
            errors.append(str(e))
    return sections, errors


def _unique_sheet_title(book, desired: str) -> str:
    existing = {t.casefold() for t in book.sheetnames}
    if desired.casefold() not in existing:
        return desired
    n = 2
    while True:
        suffix = f" ({n})"
        title = desired[:31 - len(suffix)] + suffix
        if title.casefold() not in existing:
            return title
        n += 1


def write_section_to_workbook(book, section: ExtractedSection) -> str:
    ws = book.create_sheet(_unique_sheet_title(book, section.default_sheet_name))
    for r_i, cells in enumerate([section.header] + section.rows, start=1):
        for c_i, s in enumerate(cells, start=1):
            cell = ws.cell(row=r_i, column=c_i)
            cell.value = s.value
            if s.data_type == "s" and isinstance(s.value, str):
                cell.data_type = "s"
            cell.number_format = s.number_format
            cell.font, cell.fill, cell.border = copy(s.font), copy(s.fill), copy(s.border)
            cell.alignment, cell.protection = copy(s.alignment), copy(s.protection)
    for c_i, width in section.col_widths.items():
        ws.column_dimensions[get_column_letter(c_i)].width = width
    return ws.title


# ==============================================================================
# NAME HANDLING (RULE 5): SA-AWARE NAME / TITLE / INITIALS / SURNAME SPLITTING
# ==============================================================================

SA_SURNAME_PARTICLES = {
    "van", "von", "der", "den", "de", "du", "di", "le", "la", "ten", "ter", "te",
    "op", "het", "da", "dos", "das", "do", "del", "della", "el", "al", "bin",
}

NAME_TITLES = {
    "mr": "Mr", "mrs": "Mrs", "ms": "Ms", "miss": "Miss", "mx": "Mx",
    "mnr": "Mnr", "mev": "Mev", "me": "Me", "dr": "Dr", "prof": "Prof",
    "professor": "Prof", "adv": "Adv", "advocate": "Adv", "rev": "Rev",
    "pastor": "Pastor", "hon": "Hon", "sir": "Sir", "ds": "Ds",
}

_NAME_CHARS = r"A-Za-z\u00C0-\u00FF"
_GLUED_PARTICLE = re.compile(r"^(van|von|du|de|le|la|di|ten|ter)([A-Z]["