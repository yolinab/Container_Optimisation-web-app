"""pipeline.py — Web-facing pipeline for the Container Packing Optimizer.

Called by api.py.  No matplotlib dependency.
"""

from typing import List, Dict, Any
from pathlib import Path

from utils.parse_xlsx import parse_pallet_excel_v3, parse_np_boxes_excel_v3
from utils.oneDbuildblocks import build_row_blocks_from_pallets
from models.A_1D_multi_container_placement import RowBlock1DOrderModel
from utils.recommend import recommend_fill_containers
from utils.export_excel import export_excel_report
from config import (
    CONTAINER_LENGTH_CM, CONTAINER_WIDTH_CM, CONTAINER_HEIGHT_CM,
    CONTAINER_DOOR_HEIGHT_CM, CONTAINER_MAX_WEIGHT_KG, ROW_GAP_CM,
    SOLVER_TIME_LIMIT_SEC,
    RECOMMEND_OBJECTIVE, RECOMMEND_SECONDARY_OBJECTIVE,
)


MAX_CONTAINERS = 30


def _humanize_block_key(key: str) -> str:
    try:
        foot, band = key.split("|")
        L, W = foot.split("x")
        return f"{L}×{W} cm footprint, height {band} cm"
    except Exception:
        return key


def select_one_variant_per_block(blocks):
    """Keep exactly one variant per physical block_id (shortest length)."""
    best = {}
    for b in blocks:
        bid = b.block_id
        if bid not in best or b.length_cm < best[bid].length_cm:
            best[bid] = b
    return [best[k] for k in sorted(best.keys())]


def assign_boxes_to_containers(
    containers: List[Dict[str, Any]],
    np_boxes: List[Dict[str, Any]],
    W: int,
    Hdoor: int,
    L: int,
    Wmax_kg: int,
) -> List[Dict[str, Any]]:
    pool: List[list] = [[dict(b), b["quantity"]] for b in np_boxes]

    def _fill_zone(zone_L, zone_W, zone_H, weight_budget):
        vol_remaining = float(zone_L * zone_W * zone_H)
        w_remaining   = float(weight_budget)
        placed = []
        for entry in pool:
            box, qty_left = entry
            if qty_left <= 0 or vol_remaining <= 0:
                continue
            bL, bW, bH = box["length_cm"], box["width_cm"], box["height_cm"]
            best_geo = 0
            best_chosen = None
            for fh, fl, fw in [
                (bH, bL, bW), (bH, bW, bL),
                (bL, bH, bW), (bL, bW, bH),
                (bW, bH, bL), (bW, bL, bH),
            ]:
                if fh > zone_H or fl <= 0 or fw <= 0:
                    continue
                geo = (zone_L // fl) * (zone_W // fw) * (zone_H // fh)
                if geo > best_geo:
                    best_geo = geo
                    best_chosen = (fl, fw, fh)
            if best_chosen is None or best_geo == 0:
                continue
            fl, fw, fh = best_chosen
            box_vol = fl * fw * fh
            vol_fit = int(vol_remaining // box_vol)
            fit = min(qty_left, best_geo, vol_fit)
            if box.get("weight_kg"):
                fit = min(fit, int(w_remaining // box["weight_kg"]))
            if fit <= 0:
                continue
            wt_used  = (box["weight_kg"] or 0.0) * fit
            vol_used = box_vol * fit
            entry[1] -= fit
            vol_remaining -= vol_used
            w_remaining   -= wt_used
            placed.append({
                "label":            box["label"],
                "length_cm":        fl,
                "width_cm":         fw,
                "height_cm":        fh,
                "quantity":         fit,
                "weight_kg_total":  wt_used,
                "volume_cm3_total": vol_used,
            })
        vol_used_total = sum(p["volume_cm3_total"] for p in placed)
        wt_used_total  = sum(p["weight_kg_total"]  for p in placed)
        return placed, vol_used_total, wt_used_total

    for container in containers:
        container["box_zones"] = []
        weight_budget = float(Wmax_kg) - float(container.get("loaded_weight", 0))
        tail_L = L - int(container.get("used_length_cm", 0))
        if tail_L > 0 and any(entry[1] > 0 for entry in pool):
            placed, vol_used, wt_used = _fill_zone(tail_L, W, Hdoor, weight_budget)
            weight_budget -= wt_used
            container["loaded_weight"] = container.get("loaded_weight", 0) + wt_used
            if placed:
                container["box_zones"].append({
                    "zone_type":       "tail",
                    "y_start_cm":      int(container.get("used_length_cm", 0)),
                    "z_base_cm":       0,
                    "length_cm":       tail_L,
                    "width_cm":        W,
                    "height_cm":       Hdoor,
                    "volume_used_cm3": vol_used,
                    "placed":          placed,
                    "total_weight_kg": wt_used,
                })

    return [{"box": entry[0], "remaining_qty": entry[1]} for entry in pool if entry[1] > 0]


def run_pipeline(
    excel_path,
    out_dir=None,
    count_col_override=None,
    L_cm: int = CONTAINER_LENGTH_CM,
    gap_cm: int = ROW_GAP_CM,
    Wmax_kg: int = CONTAINER_MAX_WEIGHT_KG,
    Hdoor_cm: int = CONTAINER_DOOR_HEIGHT_CM,
    solver: str = "ortools",
    time_limit: int = SOLVER_TIME_LIMIT_SEC,
):
    excel_path = Path(excel_path)
    if out_dir is None:
        out_dir = excel_path.parent / "outputs"
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # ── 1) Parse Excel ────────────────────────────────────────────────────────
    print("Parsing Excel...")
    lengths, widths, heights, pallets_data, meta_per_pallet = parse_pallet_excel_v3(
        str(excel_path),
        sheet_name=0,
        return_per_pallet_meta=True,
        count_col_override=count_col_override,
    )

    if not meta_per_pallet:
        raise RuntimeError(
            "No pallets were parsed from the Excel file.\n"
            "Check that the file contains rows with a valid pallet size string "
            "(e.g. '1,15x1,15x1,27') and a non-zero order quantity."
        )

    np_boxes = parse_np_boxes_excel_v3(
        str(excel_path), sheet_name=0, count_col_override=count_col_override
    )

    # ── 2) Build row-block instances ──────────────────────────────────────────
    print("Building row-blocks...")
    blocks, recommendations, warnings = build_row_blocks_from_pallets(
        meta_per_pallet,
        Hdoor_cm=Hdoor_cm,
        require_multiples=True,
    )

    if recommendations:
        lines = []
        for k, v in recommendations.items():
            human = _humanize_block_key(k)
            lines.append(f"  {human}: add {v} pallet{'s' if v != 1 else ''}")
        detail = "\n".join(lines)
        raise RuntimeError(
            "Pallet counts are not exact multiples — cannot build complete blocks.\n"
            "Add the following pallets to your order:\n" + detail
        )

    blocks = select_one_variant_per_block(blocks)

    if not blocks:
        raise RuntimeError(
            "No valid pallet blocks could be built from the input.\n"
            "Check pallet size strings and footprint dimensions."
        )

    door_ok_check = [b for b in blocks if b.height_cm <= Hdoor_cm]
    if not door_ok_check:
        heights_str = ", ".join(str(h) for h in sorted({b.height_cm for b in blocks}))
        raise RuntimeError(
            f"No pallet blocks fit through the container door ({Hdoor_cm} cm).\n"
            f"Stacked block heights in your order: {heights_str} cm."
        )

    # ── 3) Multi-container greedy loop ────────────────────────────────────────
    print("Solving containers...")
    remaining_blocks = blocks[:]
    containers: List[Dict[str, Any]] = []
    container_idx = 1

    while remaining_blocks:
        if container_idx > MAX_CONTAINERS:
            raise RuntimeError(
                f"Stopped after {MAX_CONTAINERS} containers — something may be wrong with the input."
            )
        print(f"  Solving container {container_idx} ({len(remaining_blocks)} blocks remaining)...")

        # Reserve all but one door-compatible block for future containers' door rows.
        door_ok   = [b for b in remaining_blocks if b.height_cm <= Hdoor_cm]
        door_over = [b for b in remaining_blocks if b.height_cm > Hdoor_cm]
        blocks_for_solver = (door_over + door_ok[:1]) if door_over else remaining_blocks

        lens = [b.length_cm for b in blocks_for_solver]
        hs   = [b.height_cm for b in blocks_for_solver]
        ws   = [b.weight_kg for b in blocks_for_solver]
        vals = [b.value     for b in blocks_for_solver]

        model = RowBlock1DOrderModel(
            lengths_cm=lens,
            heights_cm=hs,
            weights_kg=ws,
            values=vals,
            L_cm=L_cm,
            gap_cm=gap_cm,
            Wmax_kg=Wmax_kg,
            Hdoor_cm=Hdoor_cm,
        )

        solved = model.solve(solver=solver, time_limit=time_limit)
        if not solved:
            raise RuntimeError(f"No feasible solution for container {container_idx}")

        chosen_variant_indices = model.loaded_indices_in_order()
        chosen_blocks = [blocks_for_solver[i - 1] for i in chosen_variant_indices]
        used_block_ids = {b.block_id for b in chosen_blocks}

        if not chosen_blocks:
            raise RuntimeError(
                "Solver returned empty selection. No feasible non-empty packing exists "
                "under current constraints (often because no remaining door-compliant blocks)."
            )

        y_cursor = 0
        rows = []
        for b in chosen_blocks:
            rows.append({
                "block_id":    b.block_id,
                "block_type":  b.block_type_key,
                "length_cm":   b.length_cm,
                "height_cm":   b.height_cm,
                "weight_kg":   b.weight_kg,
                "pallet_count": b.value,
                "y_start_cm":  y_cursor,
                "pallets":     b.pallets,
            })
            y_cursor += b.length_cm + gap_cm

        used_len = model.usedLen.value()
        containers.append({
            "container_index": container_idx,
            "rows":            rows,
            "used_length_cm":  used_len,
            "leftover_cm":     L_cm - used_len,
            "loaded_value":    model.loadedValue.value(),
            "loaded_weight":   model.loadedWeight.value(),
        })

        remaining_blocks = [b for b in remaining_blocks if b.block_id not in used_block_ids]
        container_idx += 1

    # ── 4) Assign NP boxes ────────────────────────────────────────────────────
    unplaced = []
    if np_boxes:
        print("Assigning NP boxes...")
        unplaced = assign_boxes_to_containers(
            containers, np_boxes,
            W=CONTAINER_WIDTH_CM, Hdoor=Hdoor_cm, L=L_cm, Wmax_kg=Wmax_kg,
        )

    # ── 5) Fill recommendations ───────────────────────────────────────────────
    print("Computing recommendations...")
    recs = recommend_fill_containers(
        containers,
        Hdoor_cm=Hdoor_cm,
        H_container_cm=CONTAINER_HEIGHT_CM,
        W=CONTAINER_WIDTH_CM,
        gap_cm=ROW_GAP_CM,
        objective=RECOMMEND_OBJECTIVE,
        secondary=RECOMMEND_SECONDARY_OBJECTIVE,
        np_boxes=np_boxes if np_boxes else None,
    )

    # ── 6) Excel report ───────────────────────────────────────────────────────
    print("Generating Excel report...")
    _config = {
        "CONTAINER_LENGTH_CM":      CONTAINER_LENGTH_CM,
        "CONTAINER_WIDTH_CM":       CONTAINER_WIDTH_CM,
        "CONTAINER_HEIGHT_CM":      CONTAINER_HEIGHT_CM,
        "CONTAINER_DOOR_HEIGHT_CM": CONTAINER_DOOR_HEIGHT_CM,
        "CONTAINER_MAX_WEIGHT_KG":  CONTAINER_MAX_WEIGHT_KG,
        "ROW_GAP_CM":               ROW_GAP_CM,
        "RECOMMEND_OBJECTIVE":      RECOMMEND_OBJECTIVE,
    }
    report_path = export_excel_report(
        containers=containers,
        recs=recs,
        np_boxes=np_boxes if np_boxes else None,
        unplaced=unplaced if unplaced else None,
        out_dir=out_dir,
        config=_config,
    )

    print(f"Done. {len(containers)} container(s) packed. Report: {report_path}")

    return {
        "containers":      containers,
        "recommendations": recs,
        "report_path":     Path(report_path),
    }
