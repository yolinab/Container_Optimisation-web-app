import pandas as pd
from typing import List, Tuple, Dict, Any, Optional
import re

# ── Column aliases ──────────────────────────────────────────────────────────
# Edit these lists when the source Excel renames a column.
# Each list is tried in order; the FIRST match found in your file is used.
# Only columns actually read by the optimizer are listed here.
# ────────────────────────────────────────────────────────────────────────────
COLUMN_ALIASES: Dict[str, List[str]] = {
    # Pallet / box type code: A1, A2, NP …
    "TYPE_CODE":    ["Pallet type", "Pallet size"],

    # Physical dimensions string, e.g. "1.15x1.15x1.20"
    "DIMENSIONS":   ["Pallet and packing size", "Pallet size", "size"],

    # Number of units ordered (pallets for A-rows, boxes for NP rows)
    "QUANTITY": [
        "Order External Packaging Quantity",
        "Ordered External Packaging Quantity",
        "External Packaging Quantity",
        "Total order full pallets",
        "Total number of pallets",
        "external packaging",
        "full pallets",
        "order full pallets",
        "number of pallets",
    ],

    # Human-readable product description
    "PRODUCT_NAME": ["Productname", "product name", "product"],

    # Optional supplementary columns
    "ITEM":         ["Item", "item"],
    "BARCODE":      ["Barcode", "bar code", "ean"],
    "CODE":         ["Code", "article", "sku"],
    "WEIGHT":       ["External Net weight", "Net weight", "weight"],
    "PRICE_FOB":    ["Item price FOB", "price fob", "fob price", "item price", "unit price"],
}
# ─────────────────────────────────────────────────────────────────────────────

def print_parsed_pallets(pallets_data):
    """
    Print pallet data in simple CSV-like rows:
    pallet_type,length,width,height,count
    And print total number of individual pallets.
    """
    total = 0

    print("pallet_type,length,width,height,count")  # header

    for p in pallets_data:
        print(f"{p['pallet_type']},{p['length']},{p['width']},{p['height']},{p['count']}")
        total += p["count"]

    print(f"\nTOTAL_PALLETS,{total}")


def _find_col(df, candidates):
    """
    Helper: find the first existing column whose normalized name
    starts with any of the candidate strings.
    """
    norm_cols = {c.lower().strip(): c for c in df.columns}

    for cand in candidates:
        cand = cand.lower().strip()
        for norm_name, orig_name in norm_cols.items():
            if norm_name.startswith(cand):
                return orig_name

    raise KeyError(f"None of the candidate columns {candidates} found in {list(df.columns)}")


def _parse_pallet_size_str(size_str: str) -> Tuple[int, int, int]:
    """
    Parse a pallet size string like '1.15x1.15x1.01', '1.15x1.15x1.01cm',
    or '1,15x1,15x1,01 ' into integer dimensions in centimetres.

    Assumes the numbers are in metres with decimal separators '.' or ','.
    """
    s = str(size_str).strip().lower()

    # Remove units and other trailing text
    s = s.replace("cm", "")

    # Normalise decimal comma to dot
    s = s.replace(",", ".")

    # Split on x / X / ×
    parts = re.split(r"[x×]", s)
    parts = [p.strip() for p in parts if p.strip()]

    if len(parts) != 3:
        raise ValueError(f"Cannot parse pallet size string: '{size_str}'")

    # Convert metres to centimetres and round
    def to_cm(p: str) -> int:
        val = float(p)
        return int(round(val * 100))

    L = to_cm(parts[0])
    W = to_cm(parts[1])
    H = to_cm(parts[2])
    return L, W, H


def parse_pallet_excel(
    excel_path: str,
    sheet_name=0
) -> Tuple[List[int], List[int], List[int], List[Dict]]:
    """
    Parse the pallet Excel file (Edelman order export) and return:

        lengths:  list[int]  (one entry per individual pallet)
        widths:   list[int]
        heights:  list[int]
        pallets_data: list[dict] with metadata per pallet *type* row

    Expected layout (based on current order export):

        Column (e.g. F): "Pallet size"            -> string like "1.15x1.15x1.01"
        Column (e.g. Q): "Total order full pallets" -> how many full pallets of this type are ordered
        Optional: "Productname"/"Item"/"Pallet type" used as human-readable type label.

    Parameters
    ----------
    excel_path : str
        Path to the Excel file.
    sheet_name : str | int, default 0
        Sheet name or index passed to pandas.read_excel.
    """
    # Read the sheet
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Identify columns robustly
    col_pallet_size = _find_col(df, ["Pallet size", "size"])
    col_count       = _find_col(df, ["Total order full pallets", "full pallets", "order full pallets"])
    col_pallet_type = _find_col(df, ["pallet type", "type", "productname", "product name", "item"])

    # Drop rows with empty/NaN pallet size
    df = df.dropna(subset=[col_pallet_size])

    # Drop rows where count is NaN or 0
    df = df.dropna(subset=[col_count])
    df = df[df[col_count] > 0]

    pallets_data: List[Dict] = []

    for _, row in df.iterrows():
        size_str = row[col_pallet_size]
        try:
            length, width, height = _parse_pallet_size_str(size_str)
        except Exception:
            # Skip rows with unparseable size strings
            continue

        count = int(row[col_count])

        pallets_data.append({
            "pallet_size": str(size_str).strip(),
            "length": length,
            "width": width,
            "height": height,
            "pallet_type": str(row[col_pallet_type]),
            "count": count,
        })

    # Expand into one entry per physical pallet
    lengths: List[int] = []
    widths:  List[int] = []
    heights: List[int] = []

    for p in pallets_data:
        n = p["count"]
        lengths.extend([p["length"]] * n)
        widths.extend([p["width"]] * n)
        heights.extend([p["height"]] * n)

    #################################
    print_parsed_pallets(pallets_data)
    #################################

    return lengths, widths, heights, pallets_data


def _find_col_optional(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    """
    Return first matching column name, or None if not found.
    Matching is prefix-based on normalised names (lower/strip/collapse-spaces).
    Multiple consecutive spaces in column headers (e.g. "Ordered External  Packaging Quantity")
    are collapsed to one before comparison.
    """
    import re
    def _norm(s: str) -> str:
        return re.sub(r'\s+', ' ', s.lower().strip())

    norm_cols = {_norm(c): c for c in df.columns}
    for cand in candidates:
        cand_norm = _norm(cand)
        for norm_name, orig_name in norm_cols.items():
            if norm_name.startswith(cand_norm):
                return orig_name
    return None


def _find_col_required(df: pd.DataFrame, candidates: List[str]) -> str:
    col = _find_col_optional(df, candidates)
    if col is None:
        raise KeyError(
            f"Required column not found. Tried: {candidates}.\n"
            f"Columns in your file: {list(df.columns)}"
        )
    return col


def _parse_pallet_size_str(size_str: str) -> Tuple[int, int, int]:
    """
    Parse strings like '1.15x1.15x1.01', '1,15x1,15x1,01', optionally with 'cm'.
    Assumes metres -> converts to cm.
    """
    s = str(size_str).strip().lower()
    s = s.replace("cm", "").replace(",", ".")
    parts = re.split(r"[x×]", s)
    parts = [p.strip() for p in parts if p.strip()]
    if len(parts) != 3:
        raise ValueError(f"Cannot parse pallet size string: '{size_str}'")

    def to_cm(p: str) -> int:
        return int(round(float(p) * 100))

    L = to_cm(parts[0])
    W = to_cm(parts[1])
    H = to_cm(parts[2])
    return L, W, H


def parse_pallet_excel_v2(
    excel_path: str,
    sheet_name: Any = 0,
    return_per_pallet_meta: bool = True,
) -> Tuple[List[int], List[int], List[int], List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    New format for multi-container/subset model:

    Returns:
      lengths, widths, heights: one entry per physical pallet (cm)
      pallets_data: aggregated per type row
      meta_per_pallet: one dict per physical pallet, aligned with lengths/widths/heights
                      so meta_per_pallet[i] describes pallet i.

    If return_per_pallet_meta=False, meta_per_pallet will be [].
    """
    df = pd.read_excel(excel_path, sheet_name=sheet_name)

    # Required columns
    col_pallet_size = _find_col_required(df, ["Pallet size", "size"])
    col_count = _find_col_required(df, ["Total order full pallets", "full pallets", "order full pallets"])

    # Optional columns (best-effort)
    col_productname = _find_col_optional(df, ["Productname", "product name", "product"])
    col_item = _find_col_optional(df, ["Item", "item"])
    col_barcode = _find_col_optional(df, ["Barcode", "bar code", "ean"])
    col_code = _find_col_optional(df, ["Code", "article", "sku"])
    col_pallet_type = _find_col_optional(df, ["pallet type", "type"])  # sometimes exists

    # Clean rows
    df = df.dropna(subset=[col_pallet_size])
    df = df.dropna(subset=[col_count])
    df = df[df[col_count] > 0]

    pallets_data: List[Dict[str, Any]] = []
    meta_per_pallet: List[Dict[str, Any]] = []

    lengths: List[int] = []
    widths: List[int] = []
    heights: List[int] = []

    pallet_global_id = 1  # stable running id across expanded pallets

    for _, row in df.iterrows():
        size_str = row[col_pallet_size]
        try:
            L_cm, W_cm, H_cm = _parse_pallet_size_str(size_str)
        except Exception:
            continue

        count = int(row[col_count])

        # Choose a human-readable label
        label_parts = []
        if col_productname and pd.notna(row[col_productname]):
            label_parts.append(str(row[col_productname]).strip())
        if col_item and pd.notna(row[col_item]):
            label_parts.append(str(row[col_item]).strip())
        if not label_parts and col_pallet_type and pd.notna(row[col_pallet_type]):
            label_parts.append(str(row[col_pallet_type]).strip())

        pallet_label = " | ".join(label_parts) if label_parts else "UNKNOWN"

        # Aggregated row info (type-level)
        type_row: Dict[str, Any] = {
            "pallet_size_raw": str(size_str).strip(),
            "length": L_cm,
            "width": W_cm,
            "height": H_cm,
            "count": count,
            "label": pallet_label,
        }
        # Keep any useful ids if present
        if col_barcode and pd.notna(row[col_barcode]):
            type_row["barcode"] = str(row[col_barcode]).strip()
        if col_code and pd.notna(row[col_code]):
            type_row["code"] = str(row[col_code]).strip()

        pallets_data.append(type_row)

        # Expand to per-physical-pallet entries
        for j in range(count):
            lengths.append(L_cm)
            widths.append(W_cm)
            heights.append(H_cm)

            if return_per_pallet_meta:
                meta: Dict[str, Any] = {
                    "pallet_id": pallet_global_id,     # 1..N expanded
                    "type_index": len(pallets_data)-1, # index into pallets_data
                    "within_type_index": j + 1,        # 1..count
                    "label": pallet_label,
                    "pallet_size_raw": str(size_str).strip(),
                    "length": L_cm,
                    "width": W_cm,
                    "height": H_cm,
                }
                if col_productname and pd.notna(row[col_productname]):
                    meta["productname"] = str(row[col_productname]).strip()
                if col_item and pd.notna(row[col_item]):
                    meta["item"] = str(row[col_item]).strip()
                if col_barcode and pd.notna(row[col_barcode]):
                    meta["barcode"] = str(row[col_barcode]).strip()
                if col_code and pd.notna(row[col_code]):
                    meta["code"] = str(row[col_code]).strip()

                meta_per_pallet.append(meta)

            pallet_global_id += 1

    if not return_per_pallet_meta:
        meta_per_pallet = []

    return lengths, widths, heights, pallets_data, meta_per_pallet


def parse_np_boxes_excel_v3(
    excel_path: str,
    sheet_name: Any = 0,
    count_col_override: Optional[str] = None,
) -> List[Dict[str, Any]]:
    """
    Parse NP (non-palletized / loose box) rows from the Excel.

    NP rows are identified by 'NP' in the 'Pallet size' type-code column.
    Dimensions come from 'Pallet and packing size'.
    Count is taken from the first available of:
      'External Packaging Quantity', 'Total pallets in row',
      'Total pallet in container', or the regular pallet count column.

    Returns list of dicts per box TYPE (not expanded per unit):
      {label, length_cm, width_cm, height_cm, quantity,
       weight_kg (per box or None), volume_cm3, total_volume_cm3, total_weight_kg}

    Returns [] if no NP rows found or required columns are missing.
    """
    header_row = _detect_header_row(excel_path, sheet_name)
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row)

    # Type-code column: contains A1, A2, NP, etc.
    col_type_code = _find_col_optional(df, COLUMN_ALIASES["TYPE_CODE"])
    # Physical dimension string e.g. "1.15x1.15x0.43"
    col_dimensions = _find_col_optional(df, COLUMN_ALIASES["DIMENSIONS"])

    if col_type_code is None:
        print("[NP boxes] No type-code column (e.g. 'Pallet type') found; skipping NP parsing.")
        return []
    if col_dimensions is None:
        print("[NP boxes] No dimensions column (e.g. 'Pallet and packing size') found; skipping NP parsing.")
        return []

    col_productname = _find_col_optional(df, COLUMN_ALIASES["PRODUCT_NAME"])
    col_item        = _find_col_optional(df, COLUMN_ALIASES["ITEM"])
    col_barcode     = _find_col_optional(df, COLUMN_ALIASES["BARCODE"])

    # Count column: use override if provided, otherwise fuzzy-match from COLUMN_ALIASES.
    if count_col_override is not None:
        if count_col_override not in df.columns:
            raise KeyError(
                f"count_col_override='{count_col_override}' not found in file.\n"
                f"Columns in your file: {list(df.columns)}"
            )
        count_cols: List[str] = [count_col_override]
    else:
        col_count_eq = _find_col_optional(df, COLUMN_ALIASES["QUANTITY"])
        count_cols = [c for c in [col_count_eq] if c]

    col_weight = _find_col_optional(df, COLUMN_ALIASES["WEIGHT"])

    # Filter to NP rows
    np_mask = df[col_type_code].astype(str).str.strip().str.upper() == "NP"
    df_np = df[np_mask].copy()

    if df_np.empty:
        print("[NP boxes] No NP rows found in Excel.")
        return []

    print(f"[NP boxes] Found {len(df_np)} NP row(s) in Excel.")

    np_box_types: List[Dict[str, Any]] = []

    for _, row in df_np.iterrows():
        dim_val = row.get(col_dimensions)
        if dim_val is None or pd.isna(dim_val):
            continue
        try:
            L_cm, W_cm, H_cm = _parse_pallet_size_str(str(dim_val))
        except Exception:
            continue

        # Quantity — try each candidate column in order, take first non-zero value
        qty = 0
        for col_count in count_cols:
            raw_qty = row.get(col_count)
            if raw_qty is not None and pd.notna(raw_qty):
                try:
                    qty = int(float(str(raw_qty).replace(",", ".")))
                except Exception:
                    qty = 0
            if qty > 0:
                break
        if qty <= 0:
            print(f"[NP boxes] Skipping NP row with zero/missing quantity (dims: {L_cm}x{W_cm}x{H_cm})")
            continue

        # Weight per box (optional)
        weight_kg: Optional[float] = None
        if col_weight:
            raw_w = row.get(col_weight)
            if raw_w is not None and pd.notna(raw_w):
                try:
                    weight_kg = float(str(raw_w).strip().replace(",", "."))
                except Exception:
                    pass

        # Human-readable label
        label_parts = []
        if col_productname and pd.notna(row.get(col_productname)):
            label_parts.append(str(row[col_productname]).strip())
        if col_item and pd.notna(row.get(col_item)):
            label_parts.append(str(row[col_item]).strip())
        if col_barcode and pd.notna(row.get(col_barcode)):
            label_parts.append(f"[{str(row[col_barcode]).strip()}]")
        label = " | ".join(label_parts) if label_parts else "NP box"

        vol_cm3 = L_cm * W_cm * H_cm
        np_box_types.append({
            "label": label,
            "length_cm": L_cm,
            "width_cm": W_cm,
            "height_cm": H_cm,
            "quantity": qty,
            "weight_kg": weight_kg,          # per box; None if unknown
            "volume_cm3": vol_cm3,           # per box
            "total_volume_cm3": vol_cm3 * qty,
            "total_weight_kg": (weight_kg or 0.0) * qty,
        })

    total_qty = sum(b["quantity"] for b in np_box_types)
    total_vol = sum(b["total_volume_cm3"] for b in np_box_types)
    total_wt  = sum(b["total_weight_kg"] for b in np_box_types)
    print(
        f"[NP boxes] Parsed {len(np_box_types)} NP box type(s): "
        f"{total_qty} boxes, {total_vol/1e6:.3f} m³"
        + (f", {total_wt:.0f} kg total" if total_wt else "")
    )
    return np_box_types


def _detect_header_row(excel_path: str, sheet_name: Any = 0) -> int:
    """
    Scan the raw sheet to find the row index of the actual column headers.
    Looks for a row containing at least 2 known header keywords.
    Returns 0 (first row) as a fallback.
    """
    markers = {
        "barcode", "pallet type", "pallet size", "pallet and packing size",
        "productname", "product name", "total order full pallets",
        "total number of pallets", "total pallet in container",
        "order external packaging quantity", "external packaging quantity",
    }
    df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
    for i, row in df_raw.iterrows():
        row_lower = {str(v).strip().lower() for v in row.values if pd.notna(v)}
        if len(row_lower & markers) >= 2:
            return int(i)
    return 0


def parse_pallet_excel_v3(
    excel_path: str,
    sheet_name: Any = 0,
    return_per_pallet_meta: bool = True,
    count_col_override: Optional[str] = None,
) -> Tuple[List[int], List[int], List[int], List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    New format for multi-container/subset model:

    Returns:
      lengths, widths, heights: one entry per physical pallet (cm)
      pallets_data: aggregated per type row
      meta_per_pallet: one dict per physical pallet, aligned with lengths/widths/heights
                      so meta_per_pallet[i] describes pallet i.

    If return_per_pallet_meta=False, meta_per_pallet will be [].

    count_col_override: if provided, use this exact column name for the order quantity
                        instead of fuzzy-matching from the candidate list.
    """
    header_row = _detect_header_row(excel_path, sheet_name)
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row)

    # Required columns
    col_pallet_size = _find_col_required(df, COLUMN_ALIASES["DIMENSIONS"])

    if count_col_override is not None:
        if count_col_override not in df.columns:
            raise KeyError(
                f"count_col_override='{count_col_override}' not found in file.\n"
                f"Columns in your file: {list(df.columns)}"
            )
        col_count = count_col_override
    else:
        col_count = _find_col_required(df, COLUMN_ALIASES["QUANTITY"])

    # Optional columns (best-effort)
    col_productname = _find_col_optional(df, COLUMN_ALIASES["PRODUCT_NAME"])
    col_item        = _find_col_optional(df, COLUMN_ALIASES["ITEM"])
    col_barcode     = _find_col_optional(df, COLUMN_ALIASES["BARCODE"])
    col_code        = _find_col_optional(df, COLUMN_ALIASES["CODE"])
    col_weight      = _find_col_optional(df, COLUMN_ALIASES["WEIGHT"])
    col_price_fob   = _find_col_optional(df, COLUMN_ALIASES["PRICE_FOB"])

    # Exclude NP (loose box) rows — handled separately by parse_np_boxes_excel_v3
    col_type_code = _find_col_optional(df, COLUMN_ALIASES["TYPE_CODE"])
    if col_type_code:
        np_mask = df[col_type_code].astype(str).str.strip().str.upper() == "NP"
        df = df[~np_mask]

    # Clean rows
    df = df.dropna(subset=[col_pallet_size])
    df = df.dropna(subset=[col_count])
    df = df[df[col_count] > 0]

    pallets_data: List[Dict[str, Any]] = []
    meta_per_pallet: List[Dict[str, Any]] = []

    lengths: List[int] = []
    widths: List[int] = []
    heights: List[int] = []

    pallet_global_id = 1  # stable running id across expanded pallets

    for _, row in df.iterrows():
        size_str = row[col_pallet_size]
        try:
            L_cm, W_cm, H_cm = _parse_pallet_size_str(size_str)
        except Exception:
            continue

        count = int(row[col_count])

        # NEW: parse weight (best-effort)
        weight_kg: Optional[float] = None
        if col_weight and pd.notna(row[col_weight]):
            try:
                # handle strings like "1.234,5" or "1234.5"
                raw = str(row[col_weight]).strip().replace(",", ".")
                weight_kg = float(raw)
            except Exception:
                weight_kg = None

        # NEW: parse FOB price (best-effort)
        price_fob: Optional[float] = None
        if col_price_fob and pd.notna(row.get(col_price_fob)):
            try:
                price_fob = float(str(row[col_price_fob]).strip().replace(",", "."))
            except Exception:
                price_fob = None

        # Choose a human-readable label
        label_parts = []
        if col_productname and pd.notna(row[col_productname]):
            label_parts.append(str(row[col_productname]).strip())
        if col_item and pd.notna(row[col_item]):
            label_parts.append(str(row[col_item]).strip())
        if not label_parts and col_type_code and pd.notna(row[col_type_code]):
            label_parts.append(str(row[col_type_code]).strip())

        pallet_label = " | ".join(label_parts) if label_parts else "UNKNOWN"

        # Aggregated row info (type-level)
        type_row: Dict[str, Any] = {
            "pallet_size_raw": str(size_str).strip(),
            "length": L_cm,
            "width": W_cm,
            "height": H_cm,
            "count": count,
            "label": pallet_label,
            # NEW
            "weight_kg": weight_kg,
            "price_fob": price_fob,
        }
        if col_barcode and pd.notna(row[col_barcode]):
            type_row["barcode"] = str(row[col_barcode]).strip()
        if col_code and pd.notna(row[col_code]):
            type_row["code"] = str(row[col_code]).strip()

        pallets_data.append(type_row)

        # Expand to per-physical-pallet entries
        for j in range(count):
            lengths.append(L_cm)
            widths.append(W_cm)
            heights.append(H_cm)

            if return_per_pallet_meta:
                meta: Dict[str, Any] = {
                    "pallet_id": pallet_global_id,     # 1..N expanded
                    "type_index": len(pallets_data)-1, # index into pallets_data
                    "within_type_index": j + 1,        # 1..count
                    "label": pallet_label,
                    "pallet_size_raw": str(size_str).strip(),
                    "length": L_cm,
                    "width": W_cm,
                    "height": H_cm,
                    # NEW
                    "weight_kg": weight_kg,
                    "price_fob": price_fob,
                }
                if col_productname and pd.notna(row[col_productname]):
                    meta["productname"] = str(row[col_productname]).strip()
                if col_item and pd.notna(row[col_item]):
                    meta["item"] = str(row[col_item]).strip()
                if col_barcode and pd.notna(row[col_barcode]):
                    meta["barcode"] = str(row[col_barcode]).strip()
                if col_code and pd.notna(row[col_code]):
                    meta["code"] = str(row[col_code]).strip()

                meta_per_pallet.append(meta)

            pallet_global_id += 1

    if not return_per_pallet_meta:
        meta_per_pallet = []

    return lengths, widths, heights, pallets_data, meta_per_pallet