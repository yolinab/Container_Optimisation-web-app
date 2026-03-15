"""MAIN PIPELINE — 1D ROW-BLOCK CONTAINER PACKING

Demo-friendly runner:
- When frozen (PyInstaller), uses the executable folder as base directory.
- When not frozen, uses this file's folder as base directory.
- Default input: `input.xlsx` next to the executable/script.
- Outputs written to `outputs/` next to the executable/script:
    - summary.txt
    - containers.json
    - run.log

IMPORTANT:
CPMpy tries to import optional solvers (e.g., Gurobi) if installed.
For a stable executable, build in an environment WITHOUT `gurobipy`
(or build with PyInstaller excluding `gurobipy`).
"""

from typing import List, Dict, Any
from pathlib import Path
import argparse
import datetime
import json
import sys

from utils.parse_xlsx import parse_pallet_excel_v3, parse_np_boxes_excel_v3
from utils.oneDbuildblocks import build_row_blocks_from_pallets
from models.A_1D_multi_container_placement import RowBlock1DOrderModel
from utils.visualize_row_blocks import plot_all_row_block_containers_pallets
from utils.recommend import recommend_fill_containers, print_recommendations
from utils.export_excel import export_excel_report
from config import (
    CONTAINER_LENGTH_CM, CONTAINER_WIDTH_CM, CONTAINER_HEIGHT_CM,
    CONTAINER_DOOR_HEIGHT_CM, CONTAINER_MAX_WEIGHT_KG, ROW_GAP_CM,
    SOLVER_TIME_LIMIT_SEC,
    RECOMMEND_OBJECTIVE, RECOMMEND_SECONDARY_OBJECTIVE,
    _CONFIG_SOURCE, _USING_DEFAULTS,
)


def _base_dir(user_base_dir: str | None = None) -> Path:
    """Resolve the base directory for inputs/outputs."""
    if user_base_dir:
        return Path(user_base_dir).expanduser().resolve()
    # PyInstaller frozen app/exe
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # Normal python execution
    return Path(__file__).resolve().parent


def _setup_outputs(base: Path) -> Path:
    out = base / "outputs"
    out.mkdir(parents=True, exist_ok=True)
    return out


def _log(out_dir: Path, msg: str) -> None:
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    with (out_dir / "run.log").open("a", encoding="utf-8") as f:
        f.write(line + "\n")


def _log_assumptions(out_dir: Path) -> None:
    """Print and log every assumption the optimizer makes so nothing is silent."""
    sep = "=" * 68
    lines = [
        sep,
        "  OPTIMIZER ASSUMPTIONS & CONFIGURATION",
        sep,
        "",
        "  Config source:",
        f"    {'Built-in defaults (optimizer_config.json not found — see warning above)' if _USING_DEFAULTS else _CONFIG_SOURCE}",
        "",
        "  Container dimensions (internal usable space):",
        f"    Length : {CONTAINER_LENGTH_CM} cm",
        f"    Width  : {CONTAINER_WIDTH_CM} cm",
        f"    Height : {CONTAINER_HEIGHT_CM} cm",
        f"    Door H : {CONTAINER_DOOR_HEIGHT_CM} cm  ← ceiling constraint for loading",
        f"    Max wt : {CONTAINER_MAX_WEIGHT_KG:,} kg",
        "",
        "  Packing parameters:",
        f"    Row gap         : {ROW_GAP_CM} cm  (fork-lift clearance between pallet rows)",
        f"    Solver time cap : {SOLVER_TIME_LIMIT_SEC} s per container",
        "",
        "  Recommendation objective:",
        f"    Primary   : {RECOMMEND_OBJECTIVE}",
        f"    Secondary : {RECOMMEND_SECONDARY_OBJECTIVE}",
        "",
        "  HARDCODED PALLET STANDARDS (not user-configurable — requires code change):",
        "    Recognised footprints : 115×115 cm, 115×108 cm, 115×77 cm, 77×77 cm",
        "    Footprint tolerance   : ±2 cm  (raw dims snapped to nearest standard)",
        "    Height bands          : <66 cm, 66–89 cm, 89–130 cm, >130 cm, 230 cm",
        "    Stacking rules per band are fixed (see utils/oneDbuildblocks.py)",
        "",
        "  HARDCODED COLUMN NAME ALIASES (Excel parsing):",
        "    Pallet size  : 'Pallet and packing size', 'Pallet size', 'size'",
        "    Qty          : 'External Packaging Quantity', 'Total order full pallets', ...",
        "    Weight       : 'External Net weight', 'Net weight', 'Weight', ...",
        "    NP type flag : 'NP' anywhere in pallet-type column triggers loose-box logic",
        "",
        sep,
    ]
    for line in lines:
        _log(out_dir, line)


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

    # Pool sorted largest-volume first so big items get first pick of space.
    # Shared across containers — quantities decrease as boxes are placed.
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


def select_one_variant_per_block(blocks):
    """Keep exactly one variant per physical block_id (currently the shortest length)."""
    best = {}
    for b in blocks:
        bid = b.block_id
        if bid not in best:
            best[bid] = b
        else:
            # choose shorter length variant to increase chance of fitting
            if b.length_cm < best[bid].length_cm:
                best[bid] = b
    # preserve stable ordering by block_id
    return [best[k] for k in sorted(best.keys())]


MAX_CONTAINERS = 30  # safety cap to prevent infinite loops on bad input


def _humanize_block_key(key: str) -> str:
    """Convert '115x77|>130' → '115×77 cm footprint, height >130 cm'."""
    try:
        foot, band = key.split("|")
        L, W = foot.split("x")
        return f"{L}×{W} cm footprint, height {band} cm"
    except Exception:
        return key


def main(
    excel_path: str = "input_final.xlsx",
    sheet_name=0,
    L_cm: int = CONTAINER_LENGTH_CM,
    gap_cm: int = ROW_GAP_CM,
    Wmax_kg: int = CONTAINER_MAX_WEIGHT_KG,
    Hdoor_cm: int = CONTAINER_DOOR_HEIGHT_CM,
    solver: str = "ortools",
    time_limit: int = SOLVER_TIME_LIMIT_SEC,
    base_dir: str | None = None,
    no_plot: bool = False,
    count_col_override: str | None = None,
):
    base = _base_dir(base_dir)
    out_dir = _setup_outputs(base)
    _log_assumptions(out_dir)

    excel_p = Path(excel_path)
    if not excel_p.is_absolute():
        excel_p = (base / excel_p).resolve()

    _log(out_dir, f"Base directory: {base}")
    _log(out_dir, f"Excel input: {excel_p}")

    # ------------------------------------------------------------
    # 1) Parse Excel
    # ------------------------------------------------------------
    _log(out_dir, "=== STEP 1: Parsing Excel ===")

    lengths, widths, heights, pallets_data, meta_per_pallet = parse_pallet_excel_v3(
        str(excel_p),
        sheet_name=sheet_name,
        return_per_pallet_meta=True,
        count_col_override=count_col_override,
    )

    if not meta_per_pallet:
        raise RuntimeError(
            "No pallets were parsed from the Excel file.\n"
            "Check that the file contains rows with a valid pallet size string "
            "(e.g. '1,15x1,15x1,27') and a non-zero order quantity."
        )

    print(f"Parsed {len(meta_per_pallet)} physical pallets")
    print(f"Distinct pallet rows: {len(pallets_data)}")

    # Parse NP (loose box) rows from the same file
    np_boxes = parse_np_boxes_excel_v3(
        str(excel_p), sheet_name=sheet_name, count_col_override=count_col_override
    )

    # ------------------------------------------------------------
    # 2) Build row-block instances (and validate multiples)
    # ------------------------------------------------------------
    _log(out_dir, "=== STEP 2: Building Row-Blocks ===")

    blocks, recommendations, warnings = build_row_blocks_from_pallets(
        meta_per_pallet,
        Hdoor_cm=Hdoor_cm,
        require_multiples=True,   # HARD requirement
    )





    if warnings:
        print("\nWARNINGS during block construction:")
        for w in warnings:
            print(" -", w)

    if recommendations:
        lines = []
        for k, v in recommendations.items():
            human = _humanize_block_key(k)
            lines.append(f"  {human}: add {v} pallet{'s' if v != 1 else ''}")
        detail = "\n".join(lines)
        print("\nORDER NOT VALID — pallet counts are not exact multiples:")
        print(detail)
        summary_path = out_dir / "summary.txt"
        with summary_path.open("w", encoding="utf-8") as f:
            f.write("ORDER NOT VALID — pallet counts are not exact multiples.\n")
            f.write("Add the following pallets to reach valid block sizes:\n\n")
            for line in lines:
                f.write(line.strip() + "\n")
        _log(out_dir, f"Wrote summary: {summary_path}")
        raise RuntimeError(
            "Pallet counts are not exact multiples — cannot build complete blocks.\n"
            "Add the following pallets to your order:\n" + detail
        )

    print(f"Constructed {len(blocks)} row-block VARIANTS")
    physical_blocks = len(set(b.block_id for b in blocks))
    print(f"Corresponding to {physical_blocks} physical row-blocks")

    # IMPORTANT: current model cannot enforce mutual exclusion across rotation variants.
    # So we keep only one variant per physical block_id.
    blocks = select_one_variant_per_block(blocks)
    print(f"After choosing ONE variant per block_id: {len(blocks)} blocks")

    if not blocks:
        raise RuntimeError(
            "No valid pallet blocks could be built from the input.\n"
            "All pallets had unknown footprints or unrecognised dimensions.\n"
            "Check pallet size strings (expected format: '1,15x0,77x1,27') "
            "and footprint dimensions (recognised: 115×115, 115×108, 115×77, 77×77 cm)."
        )

    # Pre-solver: verify at least one block fits through the door
    door_ok = [b for b in blocks if b.height_cm <= Hdoor_cm]
    if not door_ok:
        heights_str = ", ".join(str(h) for h in sorted({b.height_cm for b in blocks}))
        raise RuntimeError(
            f"No pallet blocks fit through the container door ({Hdoor_cm} cm).\n"
            f"Stacked block heights in your order: {heights_str} cm.\n"
            f"Fix: increase CONTAINER_DOOR_HEIGHT_CM in optimizer_config.json "
            f"(standard 40ft High-Cube door = 259 cm), or check pallet heights in the Excel."
        )

    # ------------------------------------------------------------
    # 3) Multi-container loop
    # ------------------------------------------------------------
    _log(out_dir, "=== STEP 3: Solving Containers ===")

    remaining_blocks = blocks[:]  # copy
    containers: List[Dict[str, Any]] = []
    container_idx = 1

    while remaining_blocks:
        if container_idx > MAX_CONTAINERS:
            raise RuntimeError(
                f"Stopped after {MAX_CONTAINERS} containers — something may be wrong with the input. "
                f"Check for blocks that are too heavy or too long to ever be packed."
            )
        print(f"\n--- Solving container {container_idx} ---")

        # Reserve all but one door-compatible block for future containers' door rows.
        door_ok   = [b for b in remaining_blocks if b.height_cm <= Hdoor_cm]
        door_over = [b for b in remaining_blocks if b.height_cm > Hdoor_cm]
        if door_over:
            blocks_for_solver = door_over + door_ok[:1]
        else:
            blocks_for_solver = remaining_blocks

        print(f"  Blocks offered to solver: {len(blocks_for_solver)} "
              f"({len(door_over)} tall, {min(len(door_ok), 1)} door-row)")

        lens = [b.length_cm for b in blocks_for_solver]
        hs   = [b.height_cm for b in blocks_for_solver]
        ws   = [b.weight_kg for b in blocks_for_solver]
        vals = [b.value for b in blocks_for_solver]

        # ---- Build model ----
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

        solved = model.solve(
            solver=solver,
            time_limit=time_limit,
        )

        if not solved:
            raise RuntimeError(f"No feasible solution for container {container_idx}")

        # ----------------------------------------------------
        # 4) Extract solution
        # ----------------------------------------------------
        chosen_variant_indices = model.loaded_indices_in_order()
        chosen_blocks = [blocks_for_solver[i - 1] for i in chosen_variant_indices]

        # Physical block IDs used
        used_block_ids = {b.block_id for b in chosen_blocks}

        if len(chosen_blocks) == 0:
            print("\n!! EMPTY CONTAINER SOLUTION RETURNED !!")
            print(f"Remaining blocks: {len(remaining_blocks)}")

            door_ok = [b for b in remaining_blocks if b.height_cm <= Hdoor_cm]
            print(f"Door-OK blocks (height <= {Hdoor_cm}): {len(door_ok)}")

            heights_unique = sorted({b.height_cm for b in remaining_blocks})
            print(f"Remaining heights (unique): {heights_unique[:30]}{'...' if len(heights_unique) > 30 else ''}")

            raise RuntimeError(
                "Solver returned empty selection. Likely no feasible non-empty packing exists "
                "under current constraints (often because no remaining door-allowed blocks)."
            )

        # Reconstruct y-coordinates (back -> door)
        y_cursor = 0
        rows = []
        for b in chosen_blocks:
            rows.append({
                "block_id": b.block_id,
                "block_type": b.block_type_key,
                "length_cm": b.length_cm,
                "height_cm": b.height_cm,
                "weight_kg": b.weight_kg,
                "pallet_count": b.value,
                "y_start_cm": y_cursor,
                "pallets": b.pallets,
            })
            y_cursor += b.length_cm + gap_cm

        used_len = model.usedLen.value()
        leftover = L_cm - used_len

        container_info = {
            "container_index": container_idx,
            "rows": rows,
            "used_length_cm": used_len,
            "leftover_cm": leftover,
            "loaded_value": model.loadedValue.value(),
            "loaded_weight": model.loadedWeight.value(),
        }
        containers.append(container_info)

        # ----------------------------------------------------
        # 5) Print container summary
        # ----------------------------------------------------
        print(f"Loaded blocks: {len(rows)}")
        print(f"Used length: {used_len} / {L_cm} cm")
        print(f"Leftover length: {leftover} cm")
        print(f"Loaded pallets: {model.loadedValue.value()}")
        print(f"Loaded weight: {model.loadedWeight.value()} kg")

        print("\nRow layout (back → door):")
        for r in rows:
            print(
                f"  y={r['y_start_cm']:>4} cm | "
                f"{r['block_type']:>12} | "
                f"L={r['length_cm']:>3} | "
                f"H={r['height_cm']:>3} | "
                f"pallets={r['pallet_count']}"
            )

        # ----------------------------------------------------
        # 6) Remove used physical blocks
        # ----------------------------------------------------
        remaining_blocks = [b for b in remaining_blocks if b.block_id not in used_block_ids]
        container_idx += 1

    # ------------------------------------------------------------
    # 7) Assign NP boxes into leftover container space
    # ------------------------------------------------------------
    unplaced = []
    if np_boxes:
        _log(out_dir, "=== STEP 7: Assigning NP Boxes ===")
        unplaced = assign_boxes_to_containers(
            containers,
            np_boxes,
            W=CONTAINER_WIDTH_CM,
            Hdoor=Hdoor_cm,
            L=L_cm,
            Wmax_kg=Wmax_kg,
        )

        print("\n--- NP Box Assignment Summary ---")
        for c in containers:
            zones = c.get("box_zones", [])
            n_boxes = sum(p["quantity"] for z in zones for p in z["placed"])
            vol_m3  = sum(z["volume_used_cm3"] for z in zones) / 1e6
            wt_kg   = sum(z["total_weight_kg"] for z in zones)
            print(
                f"  Container {c['container_index']}: "
                f"{n_boxes} boxes in {len(zones)} zone(s), "
                f"{vol_m3:.3f} m³, {wt_kg:.0f} kg"
            )
        if unplaced:
            total_unplaced = sum(e["remaining_qty"] for e in unplaced)
            print(f"  Unplaced boxes: {total_unplaced} (need additional container or space)")
        else:
            print("  All NP boxes assigned.")

    # ------------------------------------------------------------
    # 8) Fill recommendations
    # ------------------------------------------------------------
    _log(out_dir, "=== STEP 8: Fill Recommendations ===")
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
    print_recommendations(recs, RECOMMEND_OBJECTIVE)
    recs_path = out_dir / "recommendations.json"
    with recs_path.open("w", encoding="utf-8") as f:
        json.dump(recs, f, ensure_ascii=False, indent=2)
    _log(out_dir, f"Wrote recommendations: {recs_path}")

    # ------------------------------------------------------------
    # 9) Final output
    # ------------------------------------------------------------
    _log(out_dir, "=== ALL CONTAINERS SOLVED ===")
    print(f"Total containers used: {len(containers)}")

    containers_path = out_dir / "containers.json"
    with containers_path.open("w", encoding="utf-8") as f:
        json.dump(containers, f, ensure_ascii=False, indent=2)
    _log(out_dir, f"Wrote containers: {containers_path}")

    summary_path = out_dir / "summary.txt"
    total_pallets = sum(c["loaded_value"] for c in containers)
    total_weight  = sum(c["loaded_weight"] for c in containers)
    with summary_path.open("w", encoding="utf-8") as f:
        f.write("CONTAINER PACKING SUMMARY\n")
        f.write("========================\n\n")
        f.write(f"Containers used: {len(containers)}\n")
        f.write(f"Total pallets loaded: {total_pallets}\n")
        f.write(f"Total weight loaded (kg): {total_weight}\n\n")
        for c in containers:
            zones    = c.get("box_zones", [])
            n_boxes  = sum(p["quantity"] for z in zones for p in z["placed"])
            box_vol  = sum(z["volume_used_cm3"] for z in zones) / 1e6
            box_wt   = sum(z["total_weight_kg"] for z in zones)
            f.write(
                f"Container {c['container_index']}: "
                f"pallets={c['loaded_value']}, "
                f"weight={c['loaded_weight']:.0f} kg, "
                f"used_length={c['used_length_cm']} cm, "
                f"leftover={c['leftover_cm']} cm"
            )
            if n_boxes:
                f.write(f", NP_boxes={n_boxes} ({box_vol:.3f} m³, {box_wt:.0f} kg)")
            f.write("\n")
        if unplaced:
            total_unplaced = sum(e["remaining_qty"] for e in unplaced)
            f.write(f"\nUnplaced NP boxes: {total_unplaced}\n")
    _log(out_dir, f"Wrote summary: {summary_path}")

    # ------------------------------------------------------------
    # 10) Excel report
    # ------------------------------------------------------------
    _log(out_dir, "=== STEP 10: Excel Report ===")
    _config = {
        "CONTAINER_LENGTH_CM":    CONTAINER_LENGTH_CM,
        "CONTAINER_WIDTH_CM":     CONTAINER_WIDTH_CM,
        "CONTAINER_HEIGHT_CM":    CONTAINER_HEIGHT_CM,
        "CONTAINER_DOOR_HEIGHT_CM": CONTAINER_DOOR_HEIGHT_CM,
        "CONTAINER_MAX_WEIGHT_KG":  CONTAINER_MAX_WEIGHT_KG,
        "ROW_GAP_CM":             ROW_GAP_CM,
        "RECOMMEND_OBJECTIVE":    RECOMMEND_OBJECTIVE,
    }
    report_path = export_excel_report(
        containers=containers,
        recs=recs,
        np_boxes=np_boxes if np_boxes else None,
        unplaced=unplaced if unplaced else None,
        out_dir=out_dir,
        config=_config,
    )
    if report_path:
        _log(out_dir, f"Wrote Excel report: {report_path}")

    # ------------------------------------------------------------
    # 8) Visualization of all containers
    # ------------------------------------------------------------
    # containers = main("sample_instances/input_large.xlsx")
    if not no_plot:
        plot_all_row_block_containers_pallets(containers, W=CONTAINER_WIDTH_CM, L=CONTAINER_LENGTH_CM, H=CONTAINER_HEIGHT_CM, recs=recs)
        # Keep plot windows open when running as a script
        try:
            import matplotlib.pyplot as plt
            plt.show()
        except Exception:
            pass

    return containers


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Container optimizer (double-click friendly)")
    # Default: in frozen executable expect input.xlsx next to the exe; in dev use sample_instances
    default_excel = "input_final.xlsx" if getattr(sys, "frozen", False) else "sample_instances/input_final.xlsx"
    # In dev, show plots by default; in frozen app, disable plots by default
    default_no_plot = True if getattr(sys, "frozen", False) else False
    parser.add_argument(
        "--excel",
        default=default_excel,
        help="Excel file (frozen: input.xlsx next to exe; dev: sample_instances/input.xlsx)",
    )
    parser.add_argument("--sheet", default=0, help="Sheet index or name")
    parser.add_argument("--no_plot", action="store_true", help="Disable plotting")
    parser.add_argument("--base_dir", default=None, help="Base directory for input/output")

    args = parser.parse_args()

    sheet_val = args.sheet
    try:
        sheet_val = int(sheet_val)
    except Exception:
        pass

    main(
        excel_path=args.excel,
        sheet_name=sheet_val,
        base_dir=args.base_dir,
        no_plot=(args.no_plot or default_no_plot),
    )
