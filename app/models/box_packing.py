"""
box_packing.py — Column-based geometric box packer for NP (non-palletised) boxes.

Core improvement over the previous per-type sequential packer
=====================================================================
Old approach: each box type gets its own dedicated length section.
  → 10 pot types × 88 cm each = 880 cm of container tail consumed.

New approach: the tail zone is divided into *columns* along the length axis.
  Every column fills the full W×H cross-section with multiple box types stacked
  in horizontal height-strips.  All types share one column → 88 cm total.

Column algorithm (one iteration of the while-loop)
------------------------------------------------------
1. For each remaining box type find the orientation (bl, bw, bh) that maximises
   the cross-section density (n_across × n_high), with ties broken by smallest bl.
2. Column depth D = max(bl) across all remaining types — guarantees every type fits
   at least one layer deep.  D is capped at the available tail length.
3. Within the W×H face, assign horizontal height-strips to each type, tallest first
   (heavy/tall items on the floor).  Each strip uses h_used = ceil(fit / (n_across×n_deep)) × bh.
4. Advance the length cursor by D.

Data stored per zone
---------------------
  placed   — per-type aggregate list  (backward-compatible with Excel report & recommender)
  columns  — per-column detail list   (consumed by the visualiser to draw box outlines)
"""

import math
from typing import List, Dict, Any, Optional, Tuple


# All 6 axis permutations: (depth_axis, width_axis, height_axis)
_ORIENT_PERMS: List[Tuple[int, int, int]] = [
    (0, 1, 2), (0, 2, 1),
    (1, 0, 2), (1, 2, 0),
    (2, 0, 1), (2, 1, 0),
]


def _best_orientation(
    box: Dict[str, Any],
    zone_W: int,
    zone_H: int,
) -> Optional[Tuple[int, int, int, int, int]]:
    """
    Choose the orientation (bl, bw, bh) that maximises cross-section density
    n_across × n_high subject to bw ≤ zone_W and bh ≤ zone_H.
    On equal density prefer the smallest bl (less column depth consumed).

    Returns (bl, bw, bh, n_across, n_high) or None if no orientation fits.
    """
    dims = [
        int(box.get("length_cm", 0)),
        int(box.get("width_cm",  0)),
        int(box.get("height_cm", 0)),
    ]
    best_density = 0
    best_bl      = 10 ** 9
    best: Optional[Tuple[int, int, int, int, int]] = None

    for d0, d1, d2 in _ORIENT_PERMS:
        bl, bw, bh = dims[d0], dims[d1], dims[d2]
        if bl <= 0 or bw <= 0 or bh <= 0:
            continue
        if bw > zone_W or bh > zone_H:
            continue
        n_across = zone_W // bw
        n_high   = zone_H // bh
        if n_across == 0 or n_high == 0:
            continue
        density = n_across * n_high
        if density > best_density or (density == best_density and bl < best_bl):
            best_density = density
            best_bl      = bl
            best         = (bl, bw, bh, n_across, n_high)

    return best


class BoxPacker:
    """
    Column-based geometric packer.

    Usage::

        pool   = [[dict(box), int(box["quantity"])] for box in np_boxes]
        packer = BoxPacker()

        placed, columns, vol_cm3, wt_kg, length_used = packer.pack(
            zone_L, zone_W, zone_H, pool, weight_budget
        )

    ``pool`` is mutated in-place (quantities decremented as boxes are placed),
    so the same list can be passed sequentially for multiple containers and
    remaining stock is carried over correctly.
    """

    def pack(
        self,
        zone_L:        int,
        zone_W:        int,
        zone_H:        int,
        pool:          List[list],     # [[box_dict, qty], …] — mutated in-place
        weight_budget: float,
    ) -> Tuple[List[Dict], List[Dict], float, float, int]:
        """
        Pack boxes into a single zone (tail only — no atop logic).

        Returns
        -------
        placed       list[dict]  per-type aggregate (backward-compat with report/recommender)
        columns      list[dict]  per-column placement detail (for visualiser)
        vol_cm3      float       total volume placed (cm³)
        wt_kg        float       total weight placed (kg)
        length_used  int         actual length consumed by the cursor (cm)
        """
        placed_totals: Dict[str, Dict] = {}   # label → aggregate
        columns:       List[Dict]      = []
        length_cursor  = 0
        w_budget       = float(weight_budget)

        while length_cursor < zone_L:
            remaining_L = zone_L - length_cursor

            # ── Orient every remaining box type for this cross-section ────────
            # (pool_idx, bl, bw, bh, n_across, n_high)
            orients: List[Tuple[int, int, int, int, int, int]] = []
            for idx, (box, qty) in enumerate(pool):
                if qty <= 0:
                    continue
                result = _best_orientation(box, zone_W, zone_H)
                if result is None:
                    continue            # no valid orientation for this zone
                bl, bw, bh, n_across, n_high = result
                if bl > remaining_L:
                    continue            # even 1 layer won't fit in remaining space
                orients.append((idx, bl, bw, bh, n_across, n_high))

            if not orients:
                break

            # ── Column depth = largest bl so every type fits ≥ 1 layer deep ──
            D = min(max(o[1] for o in orients), remaining_L)

            # ── Fill W×H face: assign horizontal strips, tallest type first ──
            # Taller items on the floor (z = 0) for physical stability.
            orients.sort(key=lambda o: -o[3])   # sort by bh descending

            z_cursor    = 0
            col_strips: List[Dict] = []
            any_placed  = False

            for idx, bl, bw, bh, n_across, _ in orients:
                box, qty = pool[idx]
                remaining_H = zone_H - z_cursor
                if bh > remaining_H or qty <= 0:
                    continue

                n_high = remaining_H // bh
                n_deep = D // bl        # layers along column depth
                if n_deep == 0 or n_high == 0:
                    continue

                per_col  = n_across * n_high * n_deep
                fit      = min(qty, per_col)

                # Weight budget constraint
                wt_per = float(box.get("weight_kg") or 0.0)
                if wt_per > 0:
                    fit = min(fit, int(w_budget // wt_per))
                if fit <= 0:
                    continue

                # How many z-layers does this strip actually consume?
                boxes_per_z_layer = n_across * n_deep
                h_layers  = math.ceil(fit / boxes_per_z_layer)
                h_used    = h_layers * bh
                wt_used   = wt_per  * fit
                vol_used  = bl * bw * bh * fit

                pool[idx][1] -= fit
                w_budget     -= wt_used
                any_placed    = True

                z_start   = z_cursor
                z_cursor += h_used

                # ── Accumulate per-type totals (for Excel / recommender) ──────
                label = box["label"]
                if label not in placed_totals:
                    placed_totals[label] = {
                        "label":            label,
                        "length_cm":        bl,
                        "width_cm":         bw,
                        "height_cm":        bh,
                        "quantity":         0,
                        "weight_kg_total":  0.0,
                        "volume_cm3_total": 0.0,
                    }
                placed_totals[label]["quantity"]         += fit
                placed_totals[label]["weight_kg_total"]  += wt_used
                placed_totals[label]["volume_cm3_total"] += vol_used

                col_strips.append({
                    "label":      label,
                    "bl":         bl,
                    "bw":         bw,
                    "bh":         bh,
                    "n_across":   n_across,
                    "n_deep":     n_deep,
                    "n_high":     h_layers,
                    "z_start_cm": z_start,
                    "h_used_cm":  h_used,
                    "quantity":   fit,
                    "volume_cm3": vol_used,
                    "weight_kg":  wt_used,
                })

            if not any_placed:
                break

            columns.append({
                "y_start_cm": length_cursor,   # relative to zone y_start
                "depth_cm":   D,
                "width_cm":   zone_W,
                "strips":     col_strips,
            })
            length_cursor += D

        placed    = list(placed_totals.values())
        total_vol = sum(p["volume_cm3_total"] for p in placed)
        total_wt  = sum(p["weight_kg_total"]  for p in placed)
        return placed, columns, total_vol, total_wt, length_cursor
