"""
validate.py — Packing result sanity checks.

Runs after the solver and NP-box assignment steps.  Returns a list of issues;
each issue is a dict:  {"level": "ERROR"|"WARNING", "code": str, "message": str}

Call  validate_packing_result()  then  report_validation_issues()  in both
main.py and pipeline.py.

Severity guide
--------------
ERROR   — a physical or mathematical impossibility.  The result is wrong and
          must not be trusted.  (Container over-length, pallet dropped, etc.)
WARNING — suspicious but not necessarily fatal.  Should be reviewed.
          (Zero weights, low fill on non-last container, unplaced boxes, etc.)
"""

from typing import List, Dict, Any


# ───────────────────────────────────────────────────────────────────────────
# Internal thresholds
# ───────────────────────────────────────────────────────────────────────────

_LOW_FILL_THRESHOLD   = 0.50   # warn if non-last container < 50 % length-filled
_LENGTH_TOL_CM        = 1      # 1 cm rounding tolerance for geometry checks


def validate_packing_result(
    containers:       List[Dict[str, Any]],
    original_blocks:  list,              # BlockInstance objects post-selection
    np_boxes:         List[Dict[str, Any]],
    L_cm:             int,
    Hdoor_cm:         int,
    Wmax_kg:          int,
    gap_cm:           int,
) -> List[Dict[str, Any]]:
    """
    Run all sanity checks and return a list of issue dicts.
    Does NOT raise — callers decide how to surface issues.
    """
    issues: List[Dict[str, Any]] = []

    def error(code: str, msg: str) -> None:
        issues.append({"level": "ERROR", "code": code, "message": msg})

    def warn(code: str, msg: str) -> None:
        issues.append({"level": "WARNING", "code": code, "message": msg})

    # ── 1. Block coverage ─────────────────────────────────────────────────
    # Every block that entered the solver must appear in exactly one container.
    original_ids = {b.block_id for b in original_blocks}

    packed_id_to_containers: Dict[int, List[int]] = {}
    for c in containers:
        for row in c.get("rows", []):
            bid = row.get("block_id")
            if bid is not None:
                packed_id_to_containers.setdefault(bid, []).append(c["container_index"])

    missing = original_ids - set(packed_id_to_containers.keys())
    if missing:
        sample = sorted(missing)[:10]
        suffix = f" … ({len(missing) - 10} more)" if len(missing) > 10 else ""
        error(
            "BLOCK_DROPPED",
            f"{len(missing)} block(s) from the input were never packed: "
            f"{sample}{suffix}",
        )

    duplicates = {bid: cs for bid, cs in packed_id_to_containers.items() if len(cs) > 1}
    if duplicates:
        sample = [(bid, cs) for bid, cs in list(duplicates.items())[:5]]
        error(
            "BLOCK_DUPLICATED",
            f"{len(duplicates)} block(s) appear in multiple containers: "
            + "; ".join(f"block {bid} → containers {cs}" for bid, cs in sample),
        )

    # ── 2. Pallet count conservation ─────────────────────────────────────
    input_pallets  = sum(b.value for b in original_blocks)
    packed_pallets = sum(
        row.get("pallet_count", 0)
        for c in containers
        for row in c.get("rows", [])
    )
    if packed_pallets != input_pallets:
        error(
            "PALLET_COUNT_MISMATCH",
            f"Input had {input_pallets} pallets but containers hold {packed_pallets}. "
            f"Difference: {packed_pallets - input_pallets:+d}",
        )

    # ── 3. Container geometry — length ────────────────────────────────────
    for c in containers:
        idx  = c["container_index"]
        used = c.get("used_length_cm", 0)

        if used > L_cm + _LENGTH_TOL_CM:
            error(
                "GEOMETRY_OVERLENGTH",
                f"Container {idx}: used_length={used} cm exceeds container length "
                f"{L_cm} cm by {used - L_cm} cm",
            )

        leftover = c.get("leftover_cm", 0)
        if leftover < -_LENGTH_TOL_CM:
            error(
                "GEOMETRY_NEGATIVE_LEFTOVER",
                f"Container {idx}: leftover_cm={leftover} cm (should be ≥ 0)",
            )

        # Accounting cross-check: used + box_zones + leftover should equal L_cm
        box_len = sum(z.get("length_cm", 0) for z in c.get("box_zones", []))
        accounted = used + box_len + leftover
        if abs(accounted - L_cm) > _LENGTH_TOL_CM:
            warn(
                "LENGTH_ACCOUNTING_MISMATCH",
                f"Container {idx}: used({used}) + boxes({box_len}) + leftover({leftover}) "
                f"= {accounted} cm ≠ L_cm ({L_cm} cm)",
            )

    # ── 4. Door-height constraint ─────────────────────────────────────────
    for c in containers:
        idx  = c["container_index"]
        rows = c.get("rows", [])
        if not rows:
            continue

        # Last row must fit through the door
        last_h = rows[-1].get("height_cm", 0)
        if last_h > Hdoor_cm:
            error(
                "DOOR_HEIGHT_VIOLATION",
                f"Container {idx}: last (door) row height {last_h} cm "
                f"exceeds door height {Hdoor_cm} cm",
            )

        # Non-last rows may legitimately exceed the door height (loaded from back).
        # Emit one summary warning per container rather than one per row.
        tall_non_last = [
            (i + 1, row.get("height_cm", 0))
            for i, row in enumerate(rows[:-1])
            if row.get("height_cm", 0) > Hdoor_cm
        ]
        if tall_non_last:
            max_h = max(h for _, h in tall_non_last)
            warn(
                "TALL_ROWS_BACK_LOADED",
                f"Container {idx}: {len(tall_non_last)} back row(s) exceed door height "
                f"(max {max_h} cm > {Hdoor_cm} cm door) — these are loaded from the rear "
                "before the door row and require tilt-loading or specialist equipment.",
            )

    # ── 5. Weight constraint ──────────────────────────────────────────────
    all_zero_weight = all(c.get("loaded_weight", 0) == 0 for c in containers)
    if all_zero_weight and containers:
        warn(
            "ZERO_WEIGHT",
            "All containers report loaded_weight=0 kg. "
            "Pallet weights may be missing from the Excel file — "
            "the weight limit is currently inactive.",
        )
    else:
        for c in containers:
            idx = c["container_index"]
            wt  = c.get("loaded_weight", 0)
            if wt > Wmax_kg + 0.5:
                error(
                    "WEIGHT_EXCEEDED",
                    f"Container {idx}: loaded_weight={wt:.1f} kg exceeds limit "
                    f"{Wmax_kg} kg by {wt - Wmax_kg:.1f} kg",
                )

    # ── 6. Row Y-coordinate consistency ───────────────────────────────────
    for c in containers:
        idx  = c["container_index"]
        rows = c.get("rows", [])
        y_expected = 0
        for i, row in enumerate(rows):
            y_actual = row.get("y_start_cm", 0)
            if abs(y_actual - y_expected) > _LENGTH_TOL_CM:
                warn(
                    "ROW_Y_INCONSISTENT",
                    f"Container {idx}: row {i+1} (block_id={row.get('block_id')}) "
                    f"y_start={y_actual} cm, expected {y_expected} cm "
                    f"(delta={y_actual - y_expected:+d} cm)",
                )
            y_expected = y_actual + row.get("length_cm", 0) + gap_cm

    # ── 7. NP box quantity integrity ──────────────────────────────────────
    if np_boxes:
        ordered_qty: Dict[str, int] = {}
        for box in np_boxes:
            label = box.get("label", "?")
            ordered_qty[label] = ordered_qty.get(label, 0) + int(box.get("quantity", 0))

        placed_qty: Dict[str, int] = {}
        for c in containers:
            for zone in c.get("box_zones", []):
                for p in zone.get("placed", []):
                    label = p.get("label", "?")
                    placed_qty[label] = placed_qty.get(label, 0) + int(p.get("quantity", 0))

        # Over-count: placed more than ordered (bug in pool mutation logic)
        for label, placed in placed_qty.items():
            ordered = ordered_qty.get(label, 0)
            if placed > ordered + 0:
                error(
                    "NP_BOX_OVERCOUNT",
                    f"NP box '{label}': placed {placed} but only {ordered} ordered",
                )

        # Under-count: some boxes never placed
        total_ordered = sum(ordered_qty.values())
        total_placed  = sum(placed_qty.values())
        if total_placed < total_ordered:
            warn(
                "NP_BOX_UNPLACED",
                f"{total_ordered - total_placed} of {total_ordered} ordered NP boxes "
                f"could not be placed across {len(containers)} container(s). "
                "Consider a dedicated NP overflow container.",
            )

    # ── 8. Low fill on non-last containers ────────────────────────────────
    for c in containers[:-1]:
        idx  = c["container_index"]
        used = c.get("used_length_cm", 0)
        if L_cm > 0:
            fill = used / L_cm
            if fill < _LOW_FILL_THRESHOLD:
                warn(
                    "LOW_FILL_NON_LAST",
                    f"Container {idx}: only {used}/{L_cm} cm used "
                    f"({fill:.0%}) — unusually low for a non-final container",
                )

    return issues


def report_validation_issues(
    issues:   List[Dict[str, Any]],
    log_func=print,
) -> bool:
    """
    Print all issues and return True if any ERRORs were found.

    log_func — callable(str), defaults to print.  Pass _log from main.py for
               file-backed logging.
    """
    if not issues:
        log_func("[VALIDATION] All checks passed.")
        return False

    errors   = [i for i in issues if i["level"] == "ERROR"]
    warnings = [i for i in issues if i["level"] == "WARNING"]

    sep = "=" * 68
    log_func(sep)
    log_func(f"  VALIDATION RESULTS  ({len(errors)} error(s), {len(warnings)} warning(s))")
    log_func(sep)

    for issue in issues:
        tag = "!! ERROR  " if issue["level"] == "ERROR" else "   WARNING"
        log_func(f"{tag} [{issue['code']}]  {issue['message']}")

    log_func(sep)
    return len(errors) > 0
