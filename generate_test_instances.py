"""
generate_test_instances.py
Run once:  python generate_test_instances.py
Outputs test Excel files in  test_instances/

pallets_per_block rules:
  115x115 | 89-130 → 4 per block  (H 89-130 cm stacked 2 high)
  115x115 | >130   → 2 per block  (H >130 cm, single stack)
  115x108 | 89-130 → 4 per block
  115x77  | 89-130 → 4 per block
  77x77   | 89-130 → 6 per block

Container: 1203 cm long.
  115-cm block + 5-cm gap = 120 cm → 10 blocks per container (length-limited)
  Weight limit: 18 000 kg per container

Error-case files intentionally violate the multiple rule to confirm the
"add X pallets" error message works in the webapp.
"""

from pathlib import Path
import pandas as pd

OUT = Path("test_instances")
OUT.mkdir(exist_ok=True)


# ── helpers ──────────────────────────────────────────────────────────────────

def _row(product, pallet_size_code, dims_m, qty, weight=None):
    """One spreadsheet row."""
    return {
        "Productname":                       product,
        "Pallet type":                       pallet_size_code,
        "Pallet and packing size":           dims_m,
        "Order External Packaging Quantity": qty,
        "External Net weight":               weight or "",
    }


def _np(product, dims_m, qty, weight=None):
    return _row(product, "NP", dims_m, qty, weight)


def save(name, rows, note=""):
    df = pd.DataFrame(rows)
    path = OUT / name
    df.to_excel(path, index=False)
    total = sum(r["Order External Packaging Quantity"] for r in rows)
    tag = f"  [{note}]" if note else ""
    print(f"  {path.name:52s}  {len(rows):3d} row(s)  {total:5d} units{tag}")


# ─────────────────────────────────────────────────────────────────────────────
# T1 — tiny, 1 container
# 10 blocks × 4 = 40 pallets → 10×115 + 9×5 = 1195 cm  (fits perfectly)
# ─────────────────────────────────────────────────────────────────────────────
save("T1_tiny_1container.xlsx", [
    _row("Bravo Vase Tall White",       "A2", "1.15x1.15x1.20", 40),
])

# ─────────────────────────────────────────────────────────────────────────────
# T2 — small mix, 2 containers + NP boxes
# ─────────────────────────────────────────────────────────────────────────────
save("T2_small_2containers.xlsx", [
    _row("Bravo Vase Round Large",      "A2", "1.15x1.15x1.20", 40),
    _row("Bravo Vase Round Medium",     "A2", "1.15x1.15x1.20", 40),
    _np( "Bravo Bowl Small",                  "0.36x0.36x0.43", 60),
])

# ─────────────────────────────────────────────────────────────────────────────
# T3 — mixed footprints, ~3 containers
# 115x115:10 + 115x77:8 + 77x77:4 = 22 blocks → ~2.2 containers → 3
# ─────────────────────────────────────────────────────────────────────────────
save("T3_mixed_footprints.xlsx", [
    _row("Bravo Pot Round Large Grey",  "A2", "1.15x1.15x1.20", 40),
    _row("Bravo Pot Round Medium Grey", "A2", "1.15x0.77x1.20", 32),
    _row("Bravo Pot Square Small",      "A2", "0.77x0.77x1.20", 24),
    _np( "Bravo Bowl Round Large",            "0.46x0.46x0.87", 40),
    _np( "Bravo Lantern Small",               "0.20x0.20x0.35", 80),
])

# ─────────────────────────────────────────────────────────────────────────────
# T4 — tall + short mix (height-ordering logic)
# 115x115|>130 → 2/block: 20 pallets = 10 blocks  H=150 cm (single-stack)
# 115x115|89-130→4/block: 60 pallets = 15 blocks  H=240 cm (double-stack)
# Result: 240-cm rows go to rear, 150-cm rows toward door (non-increasing rule)
# ─────────────────────────────────────────────────────────────────────────────
save("T4_tall_short_mix.xlsx", [
    _row("Bravo Shelf Unit Tall",       "A2", "1.15x1.15x1.50", 20, weight=85),
    _row("Bravo Vase Round Large",      "A2", "1.15x1.15x1.20", 60, weight=32),
    _np( "Bravo Bowl Small",                  "0.36x0.36x0.43", 40),
])

# ─────────────────────────────────────────────────────────────────────────────
# T5 — large order, ~5 containers
# ─────────────────────────────────────────────────────────────────────────────
save("T5_large_5containers.xlsx", [
    _row("Bravo Vase Tall White",       "A2", "1.15x1.15x1.20", 80),
    _row("Bravo Vase Tall Black",       "A2", "1.15x1.15x1.20", 80),
    _row("Bravo Vase Tall Bronze",      "A2", "1.15x1.15x1.20", 40),
    _np( "Bravo Pot Round L Grey Zinc",       "0.46x0.46x0.87", 60),
    _np( "Bravo Bowl Round L Grey Zinc",      "0.36x0.36x0.43", 60),
    _np( "Bravo Lantern Outdoor S",           "0.20x0.20x0.35", 120),
])

# ─────────────────────────────────────────────────────────────────────────────
# T6 — full stress, all types, ~6 containers
# ─────────────────────────────────────────────────────────────────────────────
save("T6_full_stress.xlsx", [
    _row("Bravo Vase Round Xl White",   "A2", "1.15x1.15x1.20", 80, weight=28),
    _row("Bravo Vase Round Xl Black",   "A2", "1.15x1.15x1.20", 80, weight=28),
    _row("Bravo Pot Oval Medium",       "A2", "1.15x0.77x1.20", 40, weight=22),
    _row("Bravo Pot Square Small",      "A2", "0.77x0.77x1.20", 24, weight=18),
    _row("Bravo Shelf Unit Tall White", "A2", "1.15x1.15x1.50", 20, weight=90),
    _np( "Bravo Pot Round L Grey Zinc",       "0.46x0.46x0.87", 60),
    _np( "Bravo Bowl Round L Grey Zinc",      "0.36x0.36x0.43", 60),
    _np( "Bravo Bowl Round L Blue",           "0.36x0.36x0.43", 40),
    _np( "Bravo Pot Round L Brown",           "0.46x0.46x0.87", 30),
    _np( "Bravo Lantern Outdoor S",           "0.20x0.20x0.35", 60),
    _np( "Bravo Lantern Outdoor M",           "0.25x0.25x0.45", 40),
])

# ─────────────────────────────────────────────────────────────────────────────
# T7 — weight-constrained order
# Heavy pallets: 480 kg each → 4/block = 1920 kg/block.
# 18000 / 1920 = 9.375 → weight limits each container to 9 blocks
# (length allows 10 blocks, so WEIGHT is the bottleneck here)
# 3 containers worth: 9+9+9 = 27 blocks × 4 = 108 pallets
# ─────────────────────────────────────────────────────────────────────────────
save("T7_weight_constrained.xlsx", [
    _row("Bravo Stone Planter XL Grey",  "A2", "1.15x1.15x1.20", 36, weight=480),
    _row("Bravo Stone Planter XL White", "A2", "1.15x1.15x1.20", 36, weight=480),
    _row("Bravo Stone Planter XL Black", "A2", "1.15x1.15x1.20", 36, weight=480),
    _np( "Bravo Lantern Outdoor S",            "0.20x0.20x0.35", 40, weight=2),
    _np( "Bravo Bowl Round S",                 "0.25x0.25x0.18", 60, weight=1),
], note="weight bottleneck (~3 containers)")

# ─────────────────────────────────────────────────────────────────────────────
# T8 — chaotic real-world mix
# Simulates a messy actual order: 6 pallet types, 8 NP box types,
# different footprints, tall + standard, heavy + light, small NP fill
# ~5-6 containers expected
# ─────────────────────────────────────────────────────────────────────────────
save("T8_chaotic_realworld.xlsx", [
    # 115x115 standard (4/block)
    _row("Bravo Planter Cylinder L White",  "A2", "1.15x1.15x1.20", 20, weight=35),
    _row("Bravo Planter Cylinder L Grey",   "A2", "1.15x1.15x1.20", 12, weight=35),
    _row("Bravo Planter Cylinder L Black",  "A2", "1.15x1.15x1.20",  8, weight=35),
    # 115x115 tall single-stack (2/block)
    _row("Bravo Shelf Tower Natural",       "A2", "1.15x1.15x1.60", 10, weight=72),
    _row("Bravo Shelf Tower Dark",          "A2", "1.15x1.15x1.60",  4, weight=72),
    # 115x77 (4/block)
    _row("Bravo Trough Planter L",          "A2", "1.15x0.77x1.20", 16, weight=29),
    _row("Bravo Trough Planter M",          "A2", "1.15x0.77x1.20",  8, weight=22),
    # 77x77 (6/block)
    _row("Bravo Cube Planter S",            "A2", "0.77x0.77x1.10", 24, weight=18),
    # NP boxes — many small types
    _np( "Bravo Tealight S",                     "0.15x0.15x0.12", 200, weight=0.3),
    _np( "Bravo Tealight M",                     "0.18x0.18x0.15", 150, weight=0.5),
    _np( "Bravo Candle Holder S",                "0.12x0.12x0.20",  80, weight=0.4),
    _np( "Bravo Candle Holder Tall",             "0.10x0.10x0.35",  60, weight=0.3),
    _np( "Bravo Vase S White",                   "0.22x0.22x0.32",  40, weight=0.8),
    _np( "Bravo Vase S Black",                   "0.22x0.22x0.32",  40, weight=0.8),
    _np( "Bravo Bowl Deco S",                    "0.28x0.28x0.15",  30, weight=0.6),
    _np( "Bravo Frame A4",                       "0.32x0.05x0.42",  20, weight=0.5),
], note="chaotic real-world, ~5-6 containers")

# ─────────────────────────────────────────────────────────────────────────────
# T9 — mostly NP boxes, very few pallets
# Just 1 pallet block (4 pallets), the rest is NP overflow.
# Tests: NP box zone fills almost the entire container.
# ─────────────────────────────────────────────────────────────────────────────
save("T9_mostly_np_boxes.xlsx", [
    _row("Bravo Vase Tall White",   "A2", "1.15x1.15x1.20",  4),
    _np( "Bravo Tealight S",              "0.15x0.15x0.12", 500, weight=0.3),
    _np( "Bravo Candle Holder S",         "0.12x0.12x0.20", 300, weight=0.4),
    _np( "Bravo Vase S White",            "0.22x0.22x0.32", 200, weight=0.8),
    _np( "Bravo Bowl Deco S",             "0.28x0.28x0.15", 150, weight=0.6),
    _np( "Bravo Frame A4",                "0.32x0.05x0.42", 100, weight=0.5),
], note="mostly NP boxes — 1 pallet block + huge NP pool")

# ─────────────────────────────────────────────────────────────────────────────
# T10 — very large order, 10+ containers
# 400 standard pallets + heavy NP pool → stress test for solver + render
# ─────────────────────────────────────────────────────────────────────────────
save("T10_large_10containers.xlsx", [
    _row("Bravo Vase XL White",          "A2", "1.15x1.15x1.20", 80, weight=30),
    _row("Bravo Vase XL Black",          "A2", "1.15x1.15x1.20", 80, weight=30),
    _row("Bravo Vase XL Grey",           "A2", "1.15x1.15x1.20", 80, weight=30),
    _row("Bravo Vase XL Beige",          "A2", "1.15x1.15x1.20", 80, weight=30),
    _row("Bravo Planter Round XL Grey",  "A2", "1.15x0.77x1.20", 40, weight=24),
    _row("Bravo Planter Square M Grey",  "A2", "0.77x0.77x1.20", 48, weight=16),
    _np( "Bravo Bowl S",                       "0.28x0.28x0.15", 200, weight=0.6),
    _np( "Bravo Lantern S",                    "0.20x0.20x0.35",  80, weight=0.8),
], note="large order ~10 containers")

# ─────────────────────────────────────────────────────────────────────────────
# ERR1 — pallet count NOT a multiple (expected error: "add X pallets")
# 115x115|89-130 needs multiples of 4.  3 pallets → needs 1 more.
# 77x77|89-130 needs multiples of 6.  10 pallets → needs 2 more.
# Upload this in the webapp to verify the friendly error message appears.
# ─────────────────────────────────────────────────────────────────────────────
save("ERR1_non_multiple_counts.xlsx", [
    _row("Bravo Vase Test",    "A2", "1.15x1.15x1.20",  3),   # need 1 more  (4-1=3 → add 1)
    _row("Bravo Pot Test",     "A2", "0.77x0.77x1.20", 10),   # need 2 more  (12-10=2)
    _np( "Bravo Bowl Test",          "0.36x0.36x0.43", 20),
], note="ERROR CASE — non-multiples, expect friendly error in webapp")

# ─────────────────────────────────────────────────────────────────────────────
# ERR2 — duplicate pallet types spread across multiple rows (common user mistake)
# Same footprint/height across 4 rows that individually aren't multiples,
# but the total IS a multiple.  Parser should aggregate per footprint+height.
# ─────────────────────────────────────────────────────────────────────────────
save("ERR2_split_rows_same_type.xlsx", [
    _row("Bravo Vase S SKU-001",  "A2", "1.15x1.15x1.20",  8),
    _row("Bravo Vase S SKU-002",  "A2", "1.15x1.15x1.20", 12),
    _row("Bravo Vase S SKU-003",  "A2", "1.15x1.15x1.20",  8),
    _row("Bravo Vase S SKU-004",  "A2", "1.15x1.15x1.20", 12),
    _np( "Bravo Bowl Test",             "0.36x0.36x0.43", 30),
], note="split rows same type — total=40 pallets, should work fine")


print("\nDone. Files written to  test_instances/")
print("  T1-T6 : core regression cases")
print("  T7    : weight-bottleneck")
print("  T8    : chaotic real-world order")
print("  T9    : mostly NP boxes")
print("  T10   : large 10+ container order")
print("  ERR1  : non-multiple counts  → expect 'add X pallets' error in webapp")
print("  ERR2  : split rows same type → should pack correctly")
