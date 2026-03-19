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
from utils.validate import validate_packing_result, report_validation_issues
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


def _explain_container(
    chosen_blocks,
    blocks_for_solver,
    door_over,
    used_len_cm: int,
    loaded_weight_kg: float,
    L_cm: int,
    Wmax_kg: int,
    Hdoor_cm: int,
    gap_cm: int,
    blocks_remaining_after: int,
) -> Dict[str, Any]:
    """
    Build a human-readable decisions dict for one container.
    Returned as container["decisions"] and surfaced in the UI.
    """
    n_rows    = len(chosen_blocks)
    n_pallets = sum(b.value for b in chosen_blocks)
    len_pct   = round(100 * used_len_cm / L_cm)
    wt_pct    = round(100 * loaded_weight_kg / Wmax_kg)

    # Which constraint actually limited how much we packed?
    leftover_cm = L_cm - used_len_cm
    # Estimate whether one more block would have fit
    min_next_len = min((b.length_cm for b in blocks_for_solver
                        if b.block_id not in {x.block_id for x in chosen_blocks}),
                       default=None)
    length_limited = (min_next_len is not None) and (leftover_cm < min_next_len + gap_cm)
    weight_limited = loaded_weight_kg > 0.95 * Wmax_kg

    reasons = []
    if length_limited:
        reasons.append(f"container length ({L_cm} cm) reached — "
                       f"only {leftover_cm} cm left, next block needs ≥{(min_next_len or 0)+gap_cm} cm")
    if weight_limited:
        reasons.append(f"weight limit ({Wmax_kg} kg) nearly reached — "
                       f"{int(loaded_weight_kg)} kg loaded ({wt_pct}%)")
    if door_over and not length_limited and not weight_limited:
        reasons.append(
            f"door-height rule: {len(door_over)} tall-pallet row(s) loaded from the rear, "
            f"1 shorter row placed at the door ({Hdoor_cm} cm opening)"
        )
    if not reasons:
        reasons.append("all available pallet blocks for this pass were packed")

    # Heights in this container — explain ordering rule if mixed
    heights = sorted({b.height_cm for b in chosen_blocks})
    if len(heights) > 1:
        max_h, min_h = max(heights), min(heights)
        height_note = (
            f"Mixed stack heights ({'/'.join(str(h)+' cm' for h in heights)}). "
            f"Taller rows ({max_h} cm) are loaded first (rear of container); "
            f"shorter rows ({min_h} cm) are nearest the door — required by the "
            f"non-increasing height rule so everything fits through the "
            f"{Hdoor_cm} cm door opening."
        )
    else:
        height_note = (
            f"All rows are {heights[0]} cm tall. "
            f"Fits through the {Hdoor_cm} cm door opening."
        )

    return {
        "rows_packed":          n_rows,
        "pallets_packed":       n_pallets,
        "length_used_cm":       used_len_cm,
        "length_capacity_cm":   L_cm,
        "length_pct":           len_pct,
        "weight_kg":            int(loaded_weight_kg),
        "weight_capacity_kg":   Wmax_kg,
        "weight_pct":           wt_pct,
        "tall_rows_at_rear":    len(door_over) > 0,
        "blocks_deferred":      blocks_remaining_after,
        "why_not_all_fit":      "; ".join(reasons),
        "height_note":          height_note,
    }


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
    """
    Column-based geometric box packing into the tail zone only.

    Boxes may NOT be placed on top of pallets.  Delegates to BoxPacker which
    fills the W×H cross-section with multiple box types per column before
    advancing the length cursor — eliminating the one-section-per-type waste.

    Modifies containers in-place.  Returns unplaced boxes.
    """
    from models.box_packing import BoxPacker

    pool: List[list] = sorted(
        [[dict(b), int(b["quantity"])] for b in np_boxes],
        key=lambda e: e[0]["length_cm"] * e[0]["width_cm"] * e[0]["height_cm"],
        reverse=True,
    )
    packer = BoxPacker()

    for container in containers:
        container["box_zones"] = []
        weight_budget = float(Wmax_kg) - float(container.get("loaded_weight", 0))

        tail_L = L - int(container.get("used_length_cm", 0))
        if tail_L > 0 and any(e[1] > 0 for e in pool):
            placed, columns, vol_cm3, wt_kg, length_used = packer.pack(
                tail_L, W, Hdoor, pool, weight_budget
            )
            container["loaded_weight"] = container.get("loaded_weight", 0) + wt_kg
            if placed:
                container["box_zones"].append({
                    "zone_type":       "tail",
                    "y_start_cm":      int(container.get("used_length_cm", 0)),
                    "z_base_cm":       0,
                    "length_cm":       length_used,
                    "width_cm":        W,
                    "height_cm":       Hdoor,
                    "volume_used_cm3": vol_cm3,
                    "placed":          placed,
                    "columns":         columns,
                    "total_weight_kg": wt_kg,
                })

    return [{"box": e[0], "remaining_qty": e[1]} for e in pool if e[1] > 0]


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

        door_ok   = [b for b in remaining_blocks if b.height_cm <= Hdoor_cm]
        door_over = [b for b in remaining_blocks if b.height_cm > Hdoor_cm]

        if door_over and not door_ok:
            heights_str = ", ".join(str(h) for h in sorted({b.height_cm for b in door_over}))
            raise RuntimeError(
                f"Cannot pack remaining blocks into container {container_idx}: "
                f"all remaining rows ({heights_str} cm) are taller than the door "
                f"opening ({Hdoor_cm} cm) and no shorter rows are left to place at "
                f"the door. Reduce pallet stack heights or split the order."
            )

        def _solve_candidate(candidate_blocks):
            m = RowBlock1DOrderModel(
                lengths_cm=[b.length_cm for b in candidate_blocks],
                heights_cm=[b.height_cm for b in candidate_blocks],
                weights_kg=[b.weight_kg for b in candidate_blocks],
                values    =[b.value     for b in candidate_blocks],
                L_cm=L_cm, gap_cm=gap_cm, Wmax_kg=Wmax_kg, Hdoor_cm=Hdoor_cm,
            )
            ok = m.solve(solver=solver, time_limit=time_limit)
            return m, ok

        if door_over:
            # Determine whether all remaining door_over blocks fit length-wise
            # in this single container.  If they do, this is the last door_over
            # container — offer all door_ok freely so none are wasted.
            # If not, at least one future container will still need a door_ok row,
            # so reserve one per future container (conservative: 1 per door_over
            # block remaining beyond what fits here).
            door_over_total_len = (
                sum(b.length_cm for b in door_over)
                + max(0, len(door_over) - 1) * gap_cm
            )
            if door_over_total_len <= L_cm:
                # All door_over fit here — no future door_over containers needed.
                blocks_for_solver = door_over + door_ok
            else:
                # More containers will follow; reserve 1 door_ok per extra container.
                # Estimate extra containers = ceil(excess length / container length).
                excess = door_over_total_len - L_cm
                extra_containers = -(-excess // L_cm)   # ceiling division
                n_reserve = int(extra_containers)
                door_ok_to_offer = door_ok[:max(1, len(door_ok) - n_reserve)]
                blocks_for_solver = door_over + door_ok_to_offer
        else:
            blocks_for_solver = remaining_blocks

        model, solved = _solve_candidate(blocks_for_solver)
        if not solved:
            raise RuntimeError(
                f"No feasible solution for container {container_idx}. "
                f"Block heights: {sorted({b.height_cm for b in blocks_for_solver})} cm, "
                f"container L={L_cm} cm, Hdoor={Hdoor_cm} cm, Wmax={Wmax_kg} kg."
            )

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

        used_len    = model.usedLen.value()
        loaded_wt   = model.loadedWeight.value()
        remaining_blocks = [b for b in remaining_blocks if b.block_id not in used_block_ids]

        decisions = _explain_container(
            chosen_blocks=chosen_blocks,
            blocks_for_solver=blocks_for_solver,
            door_over=door_over,
            used_len_cm=used_len,
            loaded_weight_kg=loaded_wt,
            L_cm=L_cm,
            Wmax_kg=Wmax_kg,
            Hdoor_cm=Hdoor_cm,
            gap_cm=gap_cm,
            blocks_remaining_after=len(remaining_blocks),
        )

        containers.append({
            "container_index": container_idx,
            "rows":            rows,
            "used_length_cm":  used_len,
            "leftover_cm":     L_cm - used_len,
            "loaded_value":    model.loadedValue.value(),
            "loaded_weight":   loaded_wt,
            "decisions":       decisions,
        })

        container_idx += 1

    # ── 4) Assign NP boxes ────────────────────────────────────────────────────
    unplaced = []
    if np_boxes:
        print("Assigning NP boxes...")
        unplaced = assign_boxes_to_containers(
            containers, np_boxes,
            W=CONTAINER_WIDTH_CM, Hdoor=Hdoor_cm, L=L_cm, Wmax_kg=Wmax_kg,
        )
        # Update leftover_cm and decisions length stats to include NP box zones
        for container in containers:
            box_len = sum(z["length_cm"] for z in container.get("box_zones", []))
            container["leftover_cm"] = max(0, container["leftover_cm"] - box_len)
            if box_len > 0:
                total_used = container["used_length_cm"] + box_len
                container["decisions"]["length_used_cm"] = total_used
                container["decisions"]["length_pct"] = round(100 * total_used / L_cm)

    # ── 4b) Validation ────────────────────────────────────────────────────────
    print("Validating packing result...")
    validation_issues = validate_packing_result(
        containers=containers,
        original_blocks=blocks,
        np_boxes=np_boxes,
        L_cm=L_cm,
        Hdoor_cm=Hdoor_cm,
        Wmax_kg=Wmax_kg,
        gap_cm=gap_cm,
    )
    has_errors = report_validation_issues(validation_issues)
    if has_errors:
        error_messages = [
            f"[{i['code']}] {i['message']}"
            for i in validation_issues
            if i["level"] == "ERROR"
        ]
        raise RuntimeError(
            "Packing validation failed — the computed solution violates physical "
            "constraints and cannot be used:\n" + "\n".join(error_messages)
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

    total_pallets = sum(c["loaded_value"] for c in containers)
    avg_len_pct   = round(sum(c["decisions"]["length_pct"] for c in containers) / len(containers))
    avg_wt_pct    = round(sum(c["decisions"]["weight_pct"]  for c in containers) / len(containers))

    # Build a plain-language reason why multiple containers were needed
    has_tall       = any(c["decisions"]["tall_rows_at_rear"]  for c in containers)
    len_bottleneck = any("length" in c["decisions"]["why_not_all_fit"] for c in containers[:-1])
    wt_bottleneck  = any("weight" in c["decisions"]["why_not_all_fit"] for c in containers[:-1])
    nc             = len(containers)
    if nc == 1:
        overall_reason = "All pallets fit in a single container."
    else:
        parts = [f"{total_pallets} pallets required {nc} containers."]
        if len_bottleneck:
            parts.append(
                f"Container length ({L_cm} cm) was the main limiting factor "
                f"(average {avg_len_pct}% used)."
            )
        if wt_bottleneck:
            parts.append(
                f"Weight limit ({Wmax_kg} kg) was reached in some containers "
                f"(average {avg_wt_pct}% used)."
            )
        if has_tall:
            parts.append(
                "Tall pallet rows (height > door) were loaded from the rear of each "
                "container; a shorter door row was placed last to allow loading/unloading."
            )
        overall_reason = " ".join(parts)

    overall_decisions = {
        "containers_needed":  nc,
        "total_pallets":      total_pallets,
        "avg_length_pct":     avg_len_pct,
        "avg_weight_pct":     avg_wt_pct,
        "has_tall_blocks":    has_tall,
        "overall_reason":     overall_reason,
        "constraints": {
            "container_length_cm":   L_cm,
            "container_width_cm":    CONTAINER_WIDTH_CM,
            "container_height_cm":   CONTAINER_HEIGHT_CM,
            "door_height_cm":        Hdoor_cm,
            "max_weight_kg":         Wmax_kg,
            "row_gap_cm":            gap_cm,
        },
    }

    return {
        "containers":         containers,
        "recommendations":    recs,
        "report_path":        Path(report_path),
        "validation_issues":  validation_issues,
        "overall_decisions":  overall_decisions,
    }
