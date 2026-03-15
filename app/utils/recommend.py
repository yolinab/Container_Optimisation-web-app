"""
Recommendation engine — 2-D greedy fill (Y × Z).

For each free zone in a solved container the engine recommends what items to
add to the order to fill the space.  Two zone types are analysed:

  - tail  : unused Y-length after the last pallet row  (z = 0, ceiling = Hdoor)
  - atop  : headroom above each placed pallet row-block (z = row_h,
             ceiling = Hdoor  i.e. items must still pass through the door)

Candidate item types
  - Pallet block types already in this container's order.
  - NP (loose) box types from the current order, treated as full-width rows:
    n_across boxes placed side-by-side along X, row length = box_L (or box_W).

Algorithm (per zone): iterative Z-layer greedy fill.
  Layer 0 fills the floor of the zone along Y with the tallest items that fit.
  Layer 1 fills the headroom left above layer 0 with smaller items.
  … continues until no items fit.
"""

from typing import List, Dict, Any, Optional, Tuple
from utils.oneDbuildblocks import build_block_type_table, BlockType


# ------------------------------------------------------------------ #
# Internal helpers                                                     #
# ------------------------------------------------------------------ #

def _compute_order_qty_by_key(containers: List[Dict[str, Any]]) -> Dict[str, int]:
    """Total pallets placed per block_type_key across ALL containers."""
    qty: Dict[str, int] = {}
    for c in containers:
        for row in c.get("rows", []):
            k = row.get("block_type", "")
            qty[k] = qty.get(k, 0) + int(row.get("pallet_count", 0))
    return qty


def _compute_type_price_fob(containers: List[Dict[str, Any]]) -> Dict[str, float]:
    """Average FOB price per pallet, keyed by block_type_key."""
    totals: Dict[str, float] = {}
    counts: Dict[str, int] = {}
    for c in containers:
        for row in c.get("rows", []):
            key = row.get("block_type", "")
            for pm in row.get("pallets", []):
                p = pm.get("price_fob")
                if p is not None:
                    totals[key] = totals.get(key, 0.0) + float(p)
                    counts[key] = counts.get(key, 0) + 1
    return {k: totals[k] / counts[k] for k in totals}


def _score_candidate(cand: Dict[str, Any], price_by_type: Dict[str, float],
                     objective: str) -> float:
    """Score a unified candidate dict (higher = better)."""
    length = cand["length_cm"]
    h      = cand["height_cm"]
    if objective == "min_leftover":
        return float(length)
    if objective == "min_pallets":
        units = cand.get("pallets_per_block") or cand.get("units_per_placement") or 1
        return length / units if units else 0.0
    if objective == "max_weight":
        return float(length * h)      # volume proxy (full width always)
    if objective == "max_value":
        avg_price = price_by_type.get(cand.get("key", ""), 0.0)
        ppb = cand.get("pallets_per_block", 0)
        return avg_price * ppb
    return float(length)


def _greedy_1d(
    avail_L: int,
    candidates: List[Dict[str, Any]],
    gap_cm: int,
    leading_gap_cm: int,
    objective: str,
    secondary: str,
    price_by_type: Dict[str, float],
    y_base: int = 0,
) -> Tuple[List[Dict[str, Any]], int]:
    """
    Greedy 1-D fill along Y for one Z-layer.

    Returns (placements, leftover_L).  Each placement dict has 'y_start_cm'
    set to its absolute position along the container length axis.
    """
    sorted_cands = sorted(
        candidates,
        key=lambda c: (
            -_score_candidate(c, price_by_type, objective),
            -_score_candidate(c, price_by_type, secondary),
        ),
    )

    placements: List[Dict[str, Any]] = []
    y_cursor  = y_base
    remaining = avail_L
    first     = True

    while remaining > 0:
        placed = False
        for cand in sorted_cands:
            gap_before = leading_gap_cm if first else gap_cm
            cost = cand["length_cm"] + gap_before
            if remaining >= cost:
                p = dict(cand)
                p["y_start_cm"] = y_cursor + gap_before
                placements.append(p)
                y_cursor  += cost
                remaining -= cost
                first      = False
                placed     = True
                break
        if not placed:
            break

    return placements, remaining


def _greedy_fill_2d(
    avail_L: int,
    H_avail: int,
    candidates: List[Dict[str, Any]],
    gap_cm: int,
    objective: str,
    secondary: str,
    price_by_type: Dict[str, float],
    leading_gap_cm: int = 0,
    y_base: int = 0,
    z_base: int = 0,
) -> Tuple[List[Dict[str, Any]], int]:
    """
    2-D greedy fill: iterate Z-layers, filling each layer along Y.

    Returns (all_placements, floor_leftover_L).

    leading_gap_cm — gap before the very first block of the zone (tail: gap_cm,
                     atop: 0 since the atop zone starts at the row's y_start).
    y_base / z_base — absolute origin of the zone inside the container.

    Each returned placement dict carries 'y_start_cm' and 'z_base_cm'.
    """
    z_cursor      = 0
    all_placements: List[Dict[str, Any]] = []
    floor_leftover = avail_L
    first_layer    = True

    while z_cursor < H_avail:
        remaining_h = H_avail - z_cursor
        layer_cands = [c for c in candidates if c["height_cm"] <= remaining_h]
        if not layer_cands:
            break

        leading = leading_gap_cm if first_layer else 0
        placements, leftover_L = _greedy_1d(
            avail_L, layer_cands, gap_cm, leading,
            objective, secondary, price_by_type, y_base=y_base,
        )
        if not placements:
            break

        layer_h = max(p["height_cm"] for p in placements)
        for p in placements:
            p["z_base_cm"] = z_base + z_cursor
        all_placements.extend(placements)

        if first_layer:
            floor_leftover = leftover_L
            first_layer    = False

        z_cursor += layer_h

    return all_placements, floor_leftover


def _build_pallet_candidates(
    used_keys: set,
    type_table: Dict[str, Any],
    actual_height_by_key: Dict[str, int],
) -> List[Dict[str, Any]]:
    """Build unified candidate dicts for pallet block types."""
    cands = []
    for key in used_keys:
        bt = type_table.get(key)
        if bt is None:
            continue
        actual_h = actual_height_by_key.get(key, bt.block_height_cm)
        for L_opt in bt.allowed_lengths:
            cands.append({
                "type":               "pallet",
                "key":                key,
                "length_cm":          int(L_opt),
                "height_cm":          actual_h,
                "pallets_per_block":  bt.pallets_per_block,
                "units_per_placement":bt.pallets_per_block,
                "footprint":          list(bt.footprint),
                "height_band":        key.split("|")[1] if "|" in key else "",
                "stack_count":        bt.stack_count,
            })
    return cands


def _build_np_box_candidates(
    np_boxes: List[Dict[str, Any]],
    W: int,
) -> List[Dict[str, Any]]:
    """
    Build unified candidate dicts for NP box types.

    Each candidate represents ONE ROW of boxes arranged across the container
    width (n_across boxes side-by-side along X).  Both orientations of each box
    are tried; zero-capacity or duplicate orientations are dropped.
    """
    cands = []
    seen: set = set()
    for box in (np_boxes or []):
        bL = int(box.get("length_cm", 0))
        bW = int(box.get("width_cm",  0))
        bH = int(box.get("height_cm", 0))
        label = box.get("label", "unknown")
        if bL <= 0 or bW <= 0 or bH <= 0:
            continue

        for row_L, across_W in [(bL, bW), (bW, bL)]:
            n_across = W // across_W if across_W > 0 else 0
            if n_across < 1:
                continue
            k = (row_L, across_W, bH)
            if k in seen:
                continue
            seen.add(k)
            dim_key = f"{bL}×{bW}×{bH}cm"
            cands.append({
                "type":               "np_box",
                "key":                dim_key,   # dimension string — used for grouping
                "label":              label,      # product name — used for display
                "length_cm":          row_L,
                "height_cm":          bH,
                "pallets_per_block":  0,
                "units_per_placement":n_across,
                "n_across":           n_across,
                "box_dims":           [bL, bW, bH],
            })
    return cands


def _aggregate_placements(
    placements: List[Dict[str, Any]],
    price_by_type: Dict[str, float],
) -> Dict[str, Any]:
    """Summarise a flat placement list into per-type totals."""
    pallet_agg: Dict[tuple, Dict] = {}
    np_agg:     Dict[tuple, Dict] = {}

    for p in placements:
        if p["type"] == "pallet":
            k = (p["key"], p["length_cm"], p["height_cm"])
            if k not in pallet_agg:
                pallet_agg[k] = {
                    "block_type_key":    p["key"],
                    "length_cm":         p["length_cm"],
                    "height_cm":         p["height_cm"],
                    "pallets_per_block": p.get("pallets_per_block", 0),
                    "count":             0,
                }
            pallet_agg[k]["count"] += 1
        elif p["type"] == "np_box":
            k = (p["key"], p["length_cm"], p["height_cm"])
            if k not in np_agg:
                np_agg[k] = {
                    "label":      p.get("label", p["key"]),  # product name
                    "dim":        p["key"],                   # "LxWxHcm" string
                    "length_cm":  p["length_cm"],
                    "height_cm":  p["height_cm"],
                    "n_across":   p.get("n_across", 1),
                    "count":      0,
                }
            np_agg[k]["count"] += 1

    pallet_list = []
    for entry in pallet_agg.values():
        total_pal = entry["count"] * entry["pallets_per_block"]
        avg_p = price_by_type.get(entry["block_type_key"])
        entry["total_pallets"] = total_pal
        entry["est_value_fob"] = round(avg_p * total_pal, 2) if avg_p else None
        pallet_list.append(entry)

    np_list = []
    for entry in np_agg.values():
        entry["total_boxes"] = entry["count"] * entry["n_across"]
        np_list.append(entry)

    return {
        "pallet_blocks":  pallet_list,
        "np_box_rows":    np_list,
        "total_pallets":  sum(e["total_pallets"] for e in pallet_list),
        "total_np_boxes": sum(e["total_boxes"]   for e in np_list),
    }


def _proportional_tail(
    avail_L: int,
    H_avail: int,
    pallet_cands: List[Dict[str, Any]],
    np_box_cands: List[Dict[str, Any]],
    gap_cm: int,
    leading_gap_cm: int,
    order_qty_by_key: Dict[str, int],
    y_base: int,
    objective: str,
    secondary: str,
    price_by_type: Dict[str, float],
) -> Tuple[List[Dict[str, Any]], int]:
    """
    Fill the tail zone with:
      1. Pallets at z=0, distributed proportionally to their global order quantities.
      2. NP boxes filling ALL remaining Y-space with full 2-D (Y × Z) stacking up
         to H_avail — so boxes use the full available height, not just one layer.

    Pallet type share = order_qty[type] / total_order_qty.
    Space per type   = floor(budget × share / slot_length) blocks.
    Surplus after rounding is redistributed to highest-share types.

    Returns (placements, leftover_L).
    """
    # One candidate per key — use the longest block for maximum pallet coverage.
    by_key: Dict[str, Dict] = {}
    for c in pallet_cands:
        k = c["key"]
        if k not in by_key or c["length_cm"] > by_key[k]["length_cm"]:
            by_key[k] = c

    # Restrict to keys that appear in the global order (qty > 0).
    active = {k: v for k, v in by_key.items() if order_qty_by_key.get(k, 0) > 0}
    if not active:
        active = by_key   # fallback: all candidates with equal shares

    budget = avail_L - leading_gap_cm
    if not active or budget <= 0:
        # No pallets — fill the entire tail with 2-D box stacking.
        box_pls, leftover = _greedy_fill_2d(
            avail_L=avail_L, H_avail=H_avail,
            candidates=np_box_cands, gap_cm=gap_cm,
            objective=objective, secondary=secondary,
            price_by_type=price_by_type,
            leading_gap_cm=leading_gap_cm, y_base=y_base, z_base=0,
        )
        return box_pls, leftover

    # Proportional shares.
    total_qty = sum(order_qty_by_key.get(k, 0) for k in active)
    if total_qty == 0:
        shares = {k: 1.0 / len(active) for k in active}
    else:
        shares = {k: order_qty_by_key[k] / total_qty for k in active}

    # Allocate block counts proportionally (slot = block length + inter-block gap).
    alloc: Dict[str, int] = {}
    for k, cand in active.items():
        slot = cand["length_cm"] + gap_cm
        alloc[k] = max(0, int(budget * shares[k] / slot))

    # Redistribute unallocated surplus to highest-share types.
    used_budget = sum(alloc[k] * (active[k]["length_cm"] + gap_cm) for k in active)
    surplus = budget - used_budget
    min_slot = min(active[k]["length_cm"] + gap_cm for k in active)
    for k in sorted(active, key=lambda k: -shares[k]):
        if surplus < min_slot:
            break
        slot = active[k]["length_cm"] + gap_cm
        extra = int(surplus // slot)
        if extra > 0:
            alloc[k] += extra
            surplus -= extra * slot

    # Build pallet placements at z=0, highest-share types first.
    placements: List[Dict[str, Any]] = []
    y_cursor = y_base + leading_gap_cm
    for k in sorted(active, key=lambda k: -shares[k]):
        n = alloc[k]
        if n == 0:
            continue
        cand = active[k]
        for _ in range(n):
            p = dict(cand)
            p["y_start_cm"] = y_cursor
            p["z_base_cm"]  = 0
            placements.append(p)
            y_cursor += cand["length_cm"] + gap_cm

    # Remaining Y-length after pallets → fill with 2-D box stacking (Y × Z).
    pallet_used = y_cursor - y_base - leading_gap_cm
    remaining   = budget - pallet_used

    if remaining > 0 and np_box_cands and H_avail > 0:
        box_pls, leftover = _greedy_fill_2d(
            avail_L=remaining, H_avail=H_avail,
            candidates=np_box_cands, gap_cm=gap_cm,
            objective=objective, secondary=secondary,
            price_by_type=price_by_type,
            leading_gap_cm=0, y_base=y_cursor, z_base=0,
        )
        placements.extend(box_pls)
    else:
        leftover = remaining

    return placements, leftover


# ------------------------------------------------------------------ #
# Public API                                                           #
# ------------------------------------------------------------------ #

def recommend_fill_containers(
    containers: List[Dict[str, Any]],
    Hdoor_cm: int,
    H_container_cm: int,
    W: int,
    gap_cm: int,
    objective: str = "min_leftover",
    secondary: str = "min_pallets",
    np_boxes: Optional[List[Dict[str, Any]]] = None,
) -> List[Dict[str, Any]]:
    """
    For each container recommend items to add to fill free space.

    Two zones per container:
      1. tail  — unused Y-length after the last row (2-D fill from z=0);
                 both pallet blocks and NP boxes are candidates here.
      2. atop  — headroom above each pallet row (2-D fill from z=row_h);
                 pallet blocks only — NP boxes cannot be placed on top of pallets.

    The door height is the binding ceiling constraint: items at height z must
    satisfy  z + item_height ≤ Hdoor_cm  (= items must pass through the door).
    This means the available height for atop zones is  Hdoor_cm − row_h,
    NOT  H_container_cm − row_h.

    np_boxes — list of NP box type dicts from the current order.  These are
               treated as full-width row placements alongside pallet blocks.
    """
    type_table          = build_block_type_table(Hdoor_cm)
    price_by_type       = _compute_type_price_fob(containers)
    np_box_cands_global = _build_np_box_candidates(np_boxes or [], W)
    order_qty_by_key    = _compute_order_qty_by_key(containers)

    results = []
    for container in containers:
        tail_L   = int(container.get("leftover_cm", 0))
        used_len = int(container.get("used_length_cm", 0))
        rows     = container.get("rows", [])

        # Actual max height per block type from placed rows (not type-table max)
        actual_height_by_key: Dict[str, int] = {}
        used_keys: set = set()
        for row in rows:
            k = row["block_type"]
            used_keys.add(k)
            h = int(row["height_cm"])
            if k not in actual_height_by_key or h > actual_height_by_key[k]:
                actual_height_by_key[k] = h

        pallet_cands = _build_pallet_candidates(used_keys, type_table, actual_height_by_key)

        # ── Account for NP boxes already placed in the tail ─────────────────
        # leftover_cm has already been reduced by box packing (see pipeline/main).
        # We still need np_tail_length to compute the correct y-offset so
        # recommendations are placed AFTER the box zone, not on top of it.
        box_zones = container.get("box_zones", [])
        np_tail_length = sum(
            z["length_cm"] for z in box_zones if z.get("zone_type") == "tail"
        )
        # tail_L already reflects space after boxes — no further subtraction needed.
        rec_tail_L  = tail_L
        rec_y_base  = used_len + np_tail_length

        # ---- Tail zone (z = 0, ceiling = Hdoor_cm) -------------------
        tail_placements: List[Dict[str, Any]] = []
        tail_leftover = rec_tail_L
        if rec_tail_L > gap_cm and (pallet_cands or np_box_cands_global):
            tail_placements, tail_leftover = _proportional_tail(
                avail_L=rec_tail_L,
                H_avail=Hdoor_cm,
                pallet_cands=pallet_cands,
                np_box_cands=np_box_cands_global,
                gap_cm=gap_cm,
                leading_gap_cm=gap_cm,
                order_qty_by_key=order_qty_by_key,
                y_base=rec_y_base,
                objective=objective,
                secondary=secondary,
                price_by_type=price_by_type,
            )
            for p in tail_placements:
                p["zone"] = "tail"

        tail_summary = _aggregate_placements(tail_placements, price_by_type)

        # ---- Atop zones (one per placed pallet row) ------------------
        atop_placements: List[Dict[str, Any]] = []
        for row in rows:
            row_h = int(row["height_cm"])
            row_L = int(row["length_cm"])
            row_y = int(row["y_start_cm"])

            # Door is the binding ceiling: combined height z + h ≤ Hdoor
            avail_h = Hdoor_cm - row_h
            if avail_h <= 0:
                continue

            # Only pallet blocks above pallets — NP boxes cannot go on top of pallets.
            atop_cands = [
                c for c in pallet_cands
                if c["height_cm"] <= avail_h and c["length_cm"] <= row_L
            ]
            if not atop_cands:
                continue

            raw, _ = _greedy_fill_2d(
                avail_L=row_L,
                H_avail=avail_h,
                candidates=atop_cands,
                gap_cm=0,                  # NP boxes pack tightly — no fork-lift gap
                objective=objective,
                secondary=secondary,
                price_by_type=price_by_type,
                leading_gap_cm=0,          # no leading gap in atop zone
                y_base=row_y,
                z_base=row_h,
            )
            for p in raw:
                p["zone"] = "atop"
                p["above_row_block_type"] = row["block_type"]
            atop_placements.extend(raw)

        atop_summary = _aggregate_placements(atop_placements, price_by_type)

        results.append({
            "container_index":       container["container_index"],
            "used_length_cm":        used_len,
            "np_tail_length_cm":     np_tail_length,
            "leftover_before_cm":    rec_tail_L,
            "leftover_after_cm":     tail_leftover,
            "fill_rate_pct": (
                round(100.0 * (rec_tail_L - tail_leftover) / rec_tail_L, 1)
                if rec_tail_L > 0 else 100.0
            ),
            "tail_placements":       tail_placements,
            "atop_placements":       atop_placements,
            "tail_summary":          tail_summary,
            "atop_summary":          atop_summary,
            "total_pallets_to_add":  (tail_summary["total_pallets"]
                                      + atop_summary["total_pallets"]),
            "total_np_boxes_to_add": (tail_summary["total_np_boxes"]
                                      + atop_summary["total_np_boxes"]),
        })

    return results


def print_recommendations(recs: List[Dict[str, Any]], objective: str) -> None:
    """Print a human-readable recommendation table to stdout."""
    sep = "=" * 68
    print(f"\n{sep}")
    print(f"  CONTAINER FILL RECOMMENDATIONS  (objective: {objective})")
    print(sep)

    for r in recs:
        idx    = r["container_index"]
        before = r["leftover_before_cm"]
        after  = r["leftover_after_cm"]
        rate   = r["fill_rate_pct"]
        n_pal  = r["total_pallets_to_add"]
        n_box  = r["total_np_boxes_to_add"]

        ts  = r.get("tail_summary",  {})
        as_ = r.get("atop_summary",  {})

        has_any = bool(r["tail_placements"] or r["atop_placements"])
        if not has_any:
            print(f"\n  Container {idx}: {before} cm tail — "
                  f"no items from current order fit any free zone")
            continue

        print(f"\n  Container {idx}:")
        if before > 0:
            print(f"    Tail: {before} cm → {after} cm after  ({rate}% tail filled)")
        else:
            print(f"    Tail: already full")

        def _print_pallet_rows(blocks):
            for b in blocks:
                val_s = (f"  ≈ FOB {b['est_value_fob']:,.0f}"
                         if b.get("est_value_fob") else "")
                print(f"      {b['count']}× {b['block_type_key']}"
                      f"  (L={b['length_cm']} cm, H={b['height_cm']} cm)"
                      f"  →  +{b['total_pallets']} pallets{val_s}")

        def _print_np_rows(rows):
            for b in rows:
                name = b["label"]
                if len(name) > 35:
                    name = name[:33] + "…"
                print(f"      {b['count']}× row of {b['n_across']}× {name}"
                      f"  [{b.get('dim', '')}]"
                      f"  →  +{b['total_boxes']} boxes")

        if ts.get("pallet_blocks") or ts.get("np_box_rows"):
            print("    Tail zone additions:")
            _print_pallet_rows(ts.get("pallet_blocks", []))
            _print_np_rows(ts.get("np_box_rows", []))

        if as_.get("pallet_blocks") or as_.get("np_box_rows"):
            print("    Atop-row additions (headroom above existing pallets):")
            _print_pallet_rows(as_.get("pallet_blocks", []))
            _print_np_rows(as_.get("np_box_rows", []))

        print(f"    Total: +{n_pal} pallets  +{n_box} NP boxes")

    print(f"\n{sep}\n")
