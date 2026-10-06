from __future__ import annotations
import pandas as pd


def _detect_excel_engine(path: str) -> str:
    with open(path, "rb") as f:
        head = f.read(8)
    if head[:4] == b"\xd0\xcf\x11\xe0":
        return "xlrd"
    if head[:2] == b"PK":
        return "openpyxl"
    raise ValueError(f"지원하지 않는 엑셀 포맷 (head={head.hex()}).")


def load_variants(path: str, cfg: dict) -> dict:
    engine = _detect_excel_engine(path)
    ex = cfg["excel"]
    header_row0 = ex["header_row"] - 1
    part_row0 = ex["part_number_row"] - 1
    df = pd.read_excel(path, sheet_name=ex["variant_sheet"], header=None, engine=engine)
    cmap = cfg.get("customer_map", {})
    variants = {}
    hdr = df.iloc[header_row0]
    pns = df.iloc[part_row0]
    for col in range(df.shape[1]):
        code = hdr.get(col)
        if not isinstance(code, str) or "_" not in code:
            continue
        code = code.strip()
        prefix, _, model = code.partition("_")
        if prefix not in cmap:
            continue
        pn = pns.get(col)
        pn = str(pn).strip() if pd.notna(pn) else ""
        if not pn or pn.upper().endswith("XXXXX"):
            continue
        variants[code] = {
            "customer": cmap.get(prefix, prefix),
            "model": model,
            "part_number": pn,
            "column": col,
        }
    return variants
