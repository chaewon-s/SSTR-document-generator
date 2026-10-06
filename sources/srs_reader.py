from __future__ import annotations
from docx import Document


def _iter_table_rows(doc):
    for tbl in doc.tables:
        for row in tbl.rows:
            yield [c.text.strip() for c in row.cells]


def load_sw_versions(path: str) -> list:
    doc = Document(path)
    out, last = [], ""
    for cells in _iter_table_rows(doc):
        if len(cells) != 3:
            continue
        vehicle, phase, sw = cells
        if "HUD" not in sw and "CS." not in sw and "CW." not in sw:
            continue
        if vehicle:
            last = vehicle
        out.append({"vehicle": last, "phase": phase, "sw_version": sw})
    return out


def find_sw_version(path: str, vehicle: str, phase: str):
    n = lambda s: s.upper().replace(" ", "")
    for r in load_sw_versions(path):
        if n(r["vehicle"]) == n(vehicle) and n(r["phase"]) == n(phase):
            return r["sw_version"]
    return None


def load_latest_revision(path: str):
    doc = Document(path)
    latest = None
    for cells in _iter_table_rows(doc):
        if len(cells) < 4:
            continue
        ver = cells[0].strip()
        if ver and ver[0].isdigit() and "." in ver:
            latest = {"ver": ver, "contents": cells[1], "released": cells[3]}
    return latest
