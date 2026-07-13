"""
Tests for app/utils/oneDbuildblocks.py
"""
import pytest
import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "app"))

from utils.oneDbuildblocks import (
    build_row_blocks_from_pallets,
    classify_height_band,
    canonical_footprint,
    build_block_type_table,
)

HDOOR = 259  # standard HC door used in tests


# ---------------------------------------------------------------------------
# Helpers to build meta_per_pallet lists
# ---------------------------------------------------------------------------

def _make_pallets(l_cm, w_cm, h_cm, count):
    return [
        {"pallet_id": i + 1, "length": l_cm, "width": w_cm, "height": h_cm, "weight_kg": 50.0}
        for i in range(count)
    ]


# ---------------------------------------------------------------------------
# classify_height_band
# ---------------------------------------------------------------------------

class TestClassifyHeightBand:
    def test_below_66(self):
        assert classify_height_band(65) == "<66"
        assert classify_height_band(1) == "<66"

    def test_66_to_89(self):
        assert classify_height_band(66) == "66-89"
        assert classify_height_band(89) == "66-89"

    def test_89_to_130(self):
        assert classify_height_band(90) == "89-130"
        assert classify_height_band(130) == "89-130"

    def test_above_130(self):
        assert classify_height_band(131) == ">130"
        assert classify_height_band(254) == ">130"

    def test_exactly_230(self):
        assert classify_height_band(230) == "230"


# ---------------------------------------------------------------------------
# canonical_footprint
# ---------------------------------------------------------------------------

class TestCanonicalFootprint:
    def test_115x115(self):
        assert canonical_footprint(115, 115) == (115, 115)

    def test_115x108(self):
        assert canonical_footprint(115, 108) == (115, 108)

    def test_115x77(self):
        assert canonical_footprint(115, 77) == (115, 77)

    def test_77x77(self):
        assert canonical_footprint(77, 77) == (77, 77)

    def test_within_tolerance(self):
        # 113 is within 2cm of 115
        assert canonical_footprint(113, 115) == (115, 115)

    def test_unknown_footprint(self):
        assert canonical_footprint(200, 200) is None

    def test_ordering_canonical(self):
        # Larger dimension always first
        fp = canonical_footprint(77, 115)
        assert fp == (115, 77)


# ---------------------------------------------------------------------------
# build_row_blocks_from_pallets
# ---------------------------------------------------------------------------

class TestBuildRowBlocks:

    def test_8_a2_pallets_make_1_block(self):
        """8 pallets of 115×115×120 → 1 block instance (89-130 band, stack=2)."""
        pallets = _make_pallets(115, 115, 120, 8)
        blocks, recs, warnings = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        assert len(blocks) >= 1
        assert recs == {}
        assert warnings == []

    def test_4_a2_pallets_make_1_block(self):
        """4 pallets of 115×115×120 → 1 block (4 per block in 89-130 band)."""
        pallets = _make_pallets(115, 115, 120, 4)
        blocks, recs, warnings = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        assert len(blocks) >= 1
        assert recs == {}

    def test_7_pallets_trigger_multiples(self):
        """7 A2 pallets (115×115×120) — 7 mod 8 ≠ 0 → recommendations, no blocks."""
        pallets = _make_pallets(115, 115, 120, 7)
        blocks, recs, warnings = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        assert blocks == []
        assert len(recs) > 0
        # Must say to add 1 more pallet
        assert any(v == 1 for v in recs.values())

    def test_multiples_not_required_keeps_partial(self):
        """require_multiples=False → partial chunks are dropped but no early exit."""
        pallets = _make_pallets(115, 115, 120, 7)
        blocks, recs, warnings = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=False)
        # recs still populated but blocks are built from complete chunks only
        # 7 pallets → 1 complete block of 4, 3 left over (ignored)
        # Depends on block_type: 89-130 has pallets_per_block=4 → 1 block from 4, 3 leftover
        assert recs  # still has recommendations about the partial 3

    def test_unknown_footprint_produces_warning(self):
        """200×200 pallets are not a recognised footprint → warning, no block."""
        pallets = _make_pallets(200, 200, 80, 8)
        blocks, recs, warnings = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        assert len(warnings) > 0
        assert any("footprint" in w.lower() or "100" in w for w in warnings)

    def test_c2_pallets_produce_blocks(self):
        """4 pallets of 115×77×120 → band 89-130, 4 per block → 1 block."""
        pallets = _make_pallets(115, 77, 120, 4)
        blocks, recs, warnings = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        assert len(blocks) >= 1
        assert recs == {}

    def test_77x77_pallets_produce_blocks(self):
        """6 pallets of 77×77×100cm → band '89-130', pallets_per_block=6 → 1 block."""
        # NOTE: The block type table uses key '77x77|<89' for short pallets, but
        # classify_height_band() never returns '<89' (returns '66-89' or '<66' instead).
        # 77×77 pallets below 89cm are therefore silently skipped — a known limitation.
        # This test uses height=100cm (band '89-130') which matches correctly.
        pallets = _make_pallets(77, 77, 100, 6)
        blocks, recs, warnings = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        assert len(blocks) >= 1
        assert recs == {}

    def test_77x77_short_pallets_now_pack_correctly(self):
        """77×77 pallets below 89cm now work — exact-height logic removed the old band-table limitation."""
        # 259 // 80 = 3 stacks, 235 // 77 = 3 across → k = 9
        pallets = _make_pallets(77, 77, 80, 9)
        blocks, recs, warnings = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        assert recs == {}
        assert warnings == []
        assert len(blocks) == 1
        assert blocks[0].height_cm == 3 * 80   # 3 stacks × 80 cm
        assert blocks[0].pallets_across == 3

    def test_block_height_uses_actual_pallet_height(self):
        """Block height = stack_count × max_pallet_h (not the table conservative value)."""
        # 115×115×120 → band 89-130 → stack_count=2 → height=240
        pallets = _make_pallets(115, 115, 120, 8)
        blocks, _, _ = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        heights = {b.height_cm for b in blocks}
        assert 240 in heights

    def test_single_block_result_n1(self):
        """Producing exactly 1 block should not crash (regression for _BoolVarImpl bug)."""
        pallets = _make_pallets(115, 115, 120, 4)
        blocks, _, _ = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        assert len(blocks) == 1

    def test_no_door_valid_blocks_flagged(self):
        """
        115×77×135cm pallets → band '>130', stack=1, height=135 → fits through 259cm door.
        Verify that these ARE produced (no false rejection).
        """
        pallets = _make_pallets(115, 77, 135, 2)
        blocks, recs, warnings = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        assert recs == {}
        assert len(blocks) >= 1
        # All blocks should be ≤ door height
        assert all(b.height_cm <= HDOOR for b in blocks)

    def test_block_type_keys_valid(self):
        """All returned block type keys should follow the 'LxW|band' format."""
        pallets = _make_pallets(115, 115, 120, 8) + _make_pallets(115, 77, 120, 4)
        blocks, _, _ = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        for b in blocks:
            assert "|" in b.block_type_key
            parts = b.block_type_key.split("|")
            assert "x" in parts[0]

    def test_weight_summed_correctly(self):
        """Block weight = sum of pallet weights in the chunk."""
        pallets = _make_pallets(115, 115, 120, 8)  # 8 pallets × 50 kg = 400 kg
        blocks, _, _ = build_row_blocks_from_pallets(pallets, W_cm=235, H_cm=269, Hdoor_cm=HDOOR, require_multiples=True)
        total_weight = sum(b.weight_kg for b in blocks)
        assert abs(total_weight - 400.0) < 0.01


# ---------------------------------------------------------------------------
# 120×100 footprint — comprehensive tests added June 2026
#
# Geometry for 40HC (W=235, H=269, Hdoor=259, usable_H=260), 103cm pallets:
#   pa_A = 235 // 100 = 2  (row depth = 120, across dim = 100)
#   pa_B = 235 // 120 = 1  (row depth = 100, across dim = 120)
#   s    = min(260//103, 259//103) = min(2, 2) = 2   → block height = 206 cm
#   k_A  = 2 × 2 = 4,  k_B = 1 × 2 = 2
#   GCD(k_A, k_B) = 2  → valid n must be even; add 1 for any odd n
# ---------------------------------------------------------------------------

class TestFootprint120x100:

    # ── Footprint snapping ──────────────────────────────────────────────────

    def test_snap_exact(self):
        assert canonical_footprint(120, 100) == (120, 100)

    def test_snap_reversed_input(self):
        """Canonical form always has L >= W."""
        assert canonical_footprint(100, 120) == (120, 100)

    def test_snap_within_tolerance_low(self):
        assert canonical_footprint(119, 101) == (120, 100)
        assert canonical_footprint(118, 100) == (120, 100)   # 118 → 120 (dist 2)
        assert canonical_footprint(120, 98)  == (120, 100)   # 98  → 100 (dist 2)

    def test_snap_within_tolerance_high(self):
        assert canonical_footprint(121, 99)  == (120, 100)
        assert canonical_footprint(122, 100) == (120, 100)   # 122 → 120 (dist 2)
        assert canonical_footprint(120, 102) == (120, 100)   # 102 → 100 (dist 2)

    def test_snap_outside_tolerance_rejected(self):
        """123 cm is 3 away from 120 — exceeds ±2 tolerance → None."""
        assert canonical_footprint(123, 100) is None

    def test_snap_does_not_collide_with_115(self):
        """Dimensions that should snap to 115, not 120."""
        fp = canonical_footprint(115, 115)
        assert fp == (115, 115)   # 115 is closer to 115 than to 120

    def test_snap_does_not_collide_with_108(self):
        """107 cm → 108 (dist 1), not 100 (dist 7)."""
        fp = canonical_footprint(107, 120)
        assert fp is not None and fp[0] in (108, 120) and fp[1] in (108, 120)

    # ── Block creation (40HC) ───────────────────────────────────────────────

    def test_blocks_created_8_pallets(self):
        """8 pallets → 2 blocks of k_A=4 each, no recommendations."""
        pallets = _make_pallets(120, 100, 103, 8)
        blocks, recs, warnings = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert recs == {}
        assert warnings == []
        assert len(blocks) == 2
        assert all("120x100" in b.block_type_key for b in blocks)

    def test_block_height_double_stack(self):
        """103cm pallets in 40HC: s=2 → block height = 2 × 103 = 206 cm."""
        pallets = _make_pallets(120, 100, 103, 4)
        blocks, _, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert all(b.height_cm == 206 for b in blocks)

    def test_block_height_single_stack_tall_pallet(self):
        """200cm pallet: usable_H=260 → s=1 → block height = 200 cm."""
        pallets = _make_pallets(120, 100, 200, 2)
        blocks, recs, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert recs == {}
        assert all(b.height_cm == 200 for b in blocks)

    def test_key_format(self):
        """Block type key follows 'LxW|Hcm' format."""
        pallets = _make_pallets(120, 100, 103, 4)
        blocks, _, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        for b in blocks:
            assert b.block_type_key == "120x100|103cm"

    def test_weight_summed(self):
        """Block weight = sum of pallet weights in the chunk (4 × 50 kg = 200 kg)."""
        pallets = _make_pallets(120, 100, 103, 4)
        blocks, _, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert abs(sum(b.weight_kg for b in blocks) - 200.0) < 0.01

    # ── Rotation ────────────────────────────────────────────────────────────

    def test_orientation_A_row_depth_120(self):
        """4 pallets → k_A=4 fits perfectly → one block, row depth = 120 cm."""
        pallets = _make_pallets(120, 100, 103, 4)
        blocks, _, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert any(b.length_cm == 120 for b in blocks)

    def test_orientation_B_row_depth_100(self):
        """2 pallets → k_B=2 fits → one block, row depth = 100 cm."""
        pallets = _make_pallets(120, 100, 103, 2)
        blocks, _, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert any(b.length_cm == 100 for b in blocks)

    def test_both_orientations_appear_with_6_pallets(self):
        """6 pallets → greedy: 1 A-block (4) + 1 B-block (2) → both depths present."""
        pallets = _make_pallets(120, 100, 103, 6)
        blocks, recs, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert recs == {}
        depths = {b.length_cm for b in blocks}
        assert 120 in depths
        assert 100 in depths

    def test_pallets_across_orientation_A(self):
        """Orientation A (row=120, across=100): pa = 235 // 100 = 2."""
        pallets = _make_pallets(120, 100, 103, 4)
        blocks, _, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        a_blocks = [b for b in blocks if b.length_cm == 120]
        assert a_blocks
        assert all(b.pallets_across == 2 for b in a_blocks)

    def test_pallets_across_orientation_B(self):
        """Orientation B (row=100, across=120): pa = 235 // 120 = 1."""
        pallets = _make_pallets(120, 100, 103, 6)
        blocks, _, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        b_blocks = [b for b in blocks if b.length_cm == 100]
        assert b_blocks
        assert all(b.pallets_across == 1 for b in b_blocks)

    # ── Multiples ───────────────────────────────────────────────────────────

    def test_multiples_2_ok(self):
        """2 pallets → k_B=2 → valid, no recommendation."""
        _, recs, _ = build_row_blocks_from_pallets(
            _make_pallets(120, 100, 103, 2), 235, 269, 259
        )
        assert recs == {}

    def test_multiples_4_ok(self):
        """4 pallets → k_A=4 → valid."""
        _, recs, _ = build_row_blocks_from_pallets(
            _make_pallets(120, 100, 103, 4), 235, 269, 259
        )
        assert recs == {}

    def test_multiples_6_ok(self):
        """6 = k_A + k_B → valid."""
        _, recs, _ = build_row_blocks_from_pallets(
            _make_pallets(120, 100, 103, 6), 235, 269, 259
        )
        assert recs == {}

    def test_multiples_8_ok(self):
        """8 = 2 × k_A → valid."""
        _, recs, _ = build_row_blocks_from_pallets(
            _make_pallets(120, 100, 103, 8), 235, 269, 259
        )
        assert recs == {}

    def test_multiples_3_fails_add_1(self):
        """3 pallets: GCD(4,2)=2, 3 is odd → need +1 pallet."""
        blocks, recs, _ = build_row_blocks_from_pallets(
            _make_pallets(120, 100, 103, 3), 235, 269, 259
        )
        assert blocks == []
        assert recs.get("120x100|103cm") == 1

    def test_multiples_5_fails_add_1(self):
        """5 pallets: odd → need +1."""
        blocks, recs, _ = build_row_blocks_from_pallets(
            _make_pallets(120, 100, 103, 5), 235, 269, 259
        )
        assert blocks == []
        assert recs.get("120x100|103cm") == 1

    # ── Container types ─────────────────────────────────────────────────────

    def test_stacking_40ft(self):
        """103cm pallets in 40FT (Hdoor=230, usable_H=230): s=2 → H=206."""
        pallets = _make_pallets(120, 100, 103, 4)
        blocks, recs, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=239, Hdoor_cm=230, require_multiples=True
        )
        assert recs == {}
        assert all(b.height_cm == 206 for b in blocks)

    def test_stacking_20ft(self):
        """103cm pallets in 20FT (Hdoor=230, usable_H=230): same as 40FT → H=206."""
        pallets = _make_pallets(120, 100, 103, 4)
        blocks, recs, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=239, Hdoor_cm=230, require_multiples=True
        )
        assert recs == {}
        assert all(b.height_cm == 206 for b in blocks)

    def test_stacking_40hc(self):
        """103cm pallets in 40HC (Hdoor=259, usable_H=260): s=2 → H=206."""
        pallets = _make_pallets(120, 100, 103, 4)
        blocks, recs, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert recs == {}
        assert all(b.height_cm == 206 for b in blocks)

    # ── Mixed footprints ────────────────────────────────────────────────────

    def test_mixed_with_115x115_no_cross_contamination(self):
        """120×100 and 115×115 pallets coexist; each gets its own blocks."""
        pallets = _make_pallets(120, 100, 103, 4) + _make_pallets(115, 115, 103, 4)
        blocks, recs, warnings = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert recs == {}
        assert warnings == []
        keys = {b.block_type_key for b in blocks}
        assert any("120x100" in k for k in keys)
        assert any("115x115" in k for k in keys)
        # No block should contain pallets of mixed footprints
        for b in blocks:
            pallet_lengths = {int(p["length"]) for p in b.pallets}
            assert len(pallet_lengths) == 1, "block mixed different footprint lengths"

    def test_mixed_with_115x77(self):
        """120×100 alongside 115×77 — both recognised, no mutual interference."""
        pallets = _make_pallets(120, 100, 103, 4) + _make_pallets(115, 77, 103, 4)
        blocks, recs, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert recs == {}
        keys = {b.block_type_key for b in blocks}
        assert any("120x100" in k for k in keys)
        assert any("115x77" in k for k in keys)

    # ── Existing footprints unaffected ──────────────────────────────────────

    def test_existing_115x115_unchanged(self):
        """Adding 120×100 must not change 115×115 block count or height."""
        pallets = _make_pallets(115, 115, 103, 4)
        blocks, recs, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert recs == {}
        assert all("115x115" in b.block_type_key for b in blocks)
        assert all(b.height_cm == 206 for b in blocks)

    def test_existing_77x77_unchanged(self):
        """77×77 pallets still produce 3-across × 2-high = k=6 blocks."""
        pallets = _make_pallets(77, 77, 103, 6)
        blocks, recs, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert recs == {}
        assert all(b.pallets_across == 3 for b in blocks)
        assert all(b.height_cm == 206 for b in blocks)


# ---------------------------------------------------------------------------
# Cross-height leftover reconciliation (Niels van Lingen's request, July 2026)
#
# Same footprint, different exact heights, each too few to complete a clean
# block on its own — should combine into mixed-height rows instead of each
# independently asking to "add N pallets".
# ---------------------------------------------------------------------------

def _pallet(l_cm, w_cm, h_cm, pallet_id):
    return {"pallet_id": pallet_id, "length": l_cm, "width": w_cm, "height": h_cm, "weight_kg": 50.0}


class TestMixedHeightReconciliation:

    def test_niels_example_combines_into_one_block(self):
        """
        115x115 leftovers: 1@122cm, 2@93cm, 1@104cm — none complete a clean
        k=4 block alone (pa=2, s=2 => k=4 for each height in 40HC), but
        pooled together (4 pallets, pa=2 stacks-per-row) they form exactly
        one full mixed-height row instead of 3 separate "add N" rejections.
        """
        pallets = (
            [_pallet(115, 115, 122, 1)]
            + [_pallet(115, 115, 93, i) for i in (2, 3)]
            + [_pallet(115, 115, 104, 4)]
        )
        blocks, recs, warnings = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert warnings == []
        assert recs == {}, f"expected no recommendations, got {recs}"
        assert len(blocks) == 1
        b = blocks[0]
        assert b.block_type_key == "115x115|mixedcm"
        assert b.value == 4
        assert b.pallets_across == 2
        # tallest stack (122+104=226) sets the row height, capped well under Hdoor
        assert b.height_cm == 226
        assert b.height_cm <= 259
        assert abs(b.weight_kg - 200.0) < 0.01  # 4 pallets x 50kg

    def test_single_height_leftover_still_asks_to_add(self):
        """
        No cross-height partner exists (only one height present) — behaviour
        must be identical to before this feature: reject with "add N pallets".
        """
        pallets = _make_pallets(115, 115, 122, 1)
        blocks, recs, _ = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=True
        )
        assert blocks == []
        assert recs.get("115x115|122cm") == 3   # 1 pallet -> needs 3 more to reach k=4

    def test_reconciled_block_never_exceeds_door_height(self):
        """Even with many wildly different heights pooled, no stack may exceed Hdoor_cm."""
        heights = [250, 240, 230, 60, 55, 50, 45, 199, 12, 8]
        pallets = [_pallet(115, 115, h, i) for i, h in enumerate(heights, start=1)]
        blocks, recs, warnings = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=False
        )
        assert warnings == []
        assert all(b.height_cm <= 259 for b in blocks)
        # every pallet must be accounted for across the mixed blocks
        assert sum(b.value for b in blocks) == len(heights)

    def test_different_footprints_do_not_cross_contaminate(self):
        """Leftovers from 115x115 must never combine with leftovers from 115x77."""
        pallets = (
            [_pallet(115, 115, 122, 1), _pallet(115, 115, 93, 2)]
            + [_pallet(115, 77, 121, 3), _pallet(115, 77, 128, 4)]
        )
        blocks, recs, warnings = build_row_blocks_from_pallets(
            pallets, W_cm=235, H_cm=269, Hdoor_cm=259, require_multiples=False
        )
        assert warnings == []
        for b in blocks:
            footprints = {(int(p["length"]), int(p["width"])) for p in b.pallets}
            assert len(footprints) == 1, "block mixed pallets from different footprints"
