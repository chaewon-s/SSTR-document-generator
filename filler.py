from __future__ import annotations
import datetime
from docx import Document

def _norm(s):
    return s.strip().lower().rstrip(":").strip()

def _fill_label_row(doc, label, value):
    target = _norm(label)
    for tbl in doc.tables:
        for row in tbl.rows:
            try:
                cells = row.cells
            except Exception:
                continue
            for i, c in enumerate(cells[:-1]):
                if _norm(c.text) == target:
                    filled = False
                    for j in range(i + 1, len(cells)):
                        nxt = cells[j]
                        if nxt.text.strip().endswith(":"):
                            break
                        if not nxt.text.strip():
                            nxt.text = value
                            filled = True
                    if filled:
                        return True
    return False

def _fill_paragraph(doc, label, value):
    lab = label.lower().rstrip(":").strip()
    for p in doc.paragraphs:
        txt = p.text.strip()
        if txt.lower().startswith(lab):
            after = txt.split(":", 1)[-1].strip() if ":" in txt else ""
            if not after:
                p.add_run(" " + value)
                return True
    return False

def _check_reason(doc, reason):
    for tbl in doc.tables:
        for row in tbl.rows:
            cells = row.cells
            for i, c in enumerate(cells):
                if reason.lower() in c.text.strip().lower() and i > 0:
                    if not cells[i - 1].text.strip():
                        cells[i - 1].text = "X"
                        return True
    return False

def _norm_td(tech_docs):
    """tech_docs를 (name, revision) 튜플 리스트로 정규화 (dict/튜플 모두 허용)."""
    out = []
    for d in tech_docs or []:
        if isinstance(d, dict):
            out.append((d.get("name", ""), d.get("revision", "")))
        elif isinstance(d, (list, tuple)) and len(d) >= 2:
            out.append((d[0], d[1]))
    return out

def _apply_tech_docs(doc, tech_docs):
    tds = _norm_td(tech_docs)
    if not tds:
        return
    for tbl in doc.tables:
        header = " ".join(c.text.strip().lower() for c in tbl.rows[0].cells)
        if "technical" in header or "revision" in header or "doors" in header:
            idx = 0
            for row in tbl.rows[1:]:
                if idx >= len(tds):
                    break
                cells = row.cells
                name, rev = tds[idx]
                if len(cells) >= 2 and not cells[0].text.strip():
                    cells[0].text = str(name)
                    cells[1].text = str(rev)
                    idx += 1
            return

def fill_sstr(template_path, out_path, data, paragraphs=None, reason=None,
              terra=None, tech_docs=None):
    doc = Document(template_path)
    missed = []
    for label, value in data.items():
        if value in (None, ""):
            continue
        if not _fill_label_row(doc, label, str(value)):
            missed.append(label)
    if paragraphs:
        for label, value in paragraphs.items():
            if value not in (None, ""):
                _fill_paragraph(doc, label, str(value))
    if reason:
        _check_reason(doc, reason)
    if tech_docs:
        _apply_tech_docs(doc, tech_docs)
    doc.save(out_path)
    if missed:
        print("[경고] 채우지 못한 라벨:", missed)
    return out_path
