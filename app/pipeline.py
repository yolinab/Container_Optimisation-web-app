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
        # Update leftover_cm to reflect space consumed by NP boxes
        for container in containers:
            box_len = sum(z["length_cm"] for z in container.get("box_zones", []))
            container["leftover_cm"] = max(0, container["leftover_cm"] - box_len)

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
