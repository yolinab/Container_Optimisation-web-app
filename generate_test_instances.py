"""
generate_test_instances.py
Run once:  python generate_test_instances.py
Outputs 6 test Excel files in  test_instances/

Rules enforced:
  - Pallet counts must be exact multiples of pallets_per_block for each type
  - pallets_per_block table (subset used here):
      115x115 | 89-130 → 4   (H 89-130 cm)
      115x115 | >130   → 2   (H 131-229 cm)
      115x77  | 89-130 → 4
      77x77   | 89-130 → 6
  - Container 1203 cm long; one 115-cm block + 5-cm gap = 120 cm → 10 blocks/container
  - NP rows: Pallet size = "NP", dims in Pallet and packing size, qty in External Packaging Quantity
"""

from pathlib import Path
import pandas as pd

OUT = Path("test_instances")
OUT.mkdir(exist_ok=True)


# ── helpers ──────────────────────────────────────────────────────────────────

def _row(product, pallet_size_code, dims_m, qty, weight=None):
    """One spreadsheet row."""
    return {
        "Productname":              product,
        "Pallet size":              pallet_size_code,   # "A2", "A1", "NP" …
        "Pallet and packing size":  dims_m,             # "1.15x1.15x1.20"
        "External Packaging Quantity": qty,
        "External Net weight":      weight or "",
    }


def _np(product, dims_m, qty):
    return _row(product, "NP", dims_m, qty)


def save(name, rows):
    df = pd.DataFrame(rows)
    path = OUT / name
    df.to_excel(path, index=False)
    total = sum(r["External Packaging Quantity"] for r in rows)
    print(f"  {path.name:45s}  {len(rows):3d} row(s)  {total:4d} units")


# ── T1 — tiny, 1 container ───────────────────────────────────────────────────
# 10 blocks × 4 pallets = 40 pallets  →  10×115 + 9×5 = 1195 cm  (fits)
save("T1_tiny_1container.xlsx", [
    _row("Bravo Vase Tall White",       "A2", "1.15x1.15x1.20", 40),
])

# ── T2 — small mix, ~2 containers ────────────────────────────────────────────
# 20 blocks 115x115 × 4 = 80 pallets  →  2 containers
# 10 NP boxes to fill tail
save("T2_small_2containers.xlsx", [
    _row("Bravo Vase Round Large",      "A2", "1.15x1.15x1.20", 40),
    _row("Bravo Vase Round Medium",     "A2", "1.15x1.15x1.20", 40),
    _np("Bravo Bowl Small",                   "0.36x0.36x0.43", 60),
])

# ── T3 — mixed footprints, ~3 containers ─────────────────────────────────────
# 115x115|89-130 → 4/block:  40 pallets = 10 blocks
# 115x77 |89-130 → 4/block:  32 pallets =  8 blocks
# 77x77  |89-130 → 6/block:  24 pallets =  4 blocks
# blocks: 10+8+4 = 22 → ~2.2 containers  (expect 3 with NP fill)
save("T3_mixed_footprints.xlsx", [
    _row("Bravo Pot Round Large Grey",  "A2", "1.15x1.15x1.20", 40),
    _row("Bravo Pot Round Medium Grey", "A2", "1.15x0.77x1.20", 32),
    _row("Bravo Pot Square Small",      "A2", "0.77x0.77x1.20", 24),
    _np("Bravo Bowl Round Large",             "0.46x0.46x0.87", 40),
    _np("Bravo Lantern Small",                "0.20x0.20x0.35", 80),
])

# ── T4 — tall + short mix (tests back-loading logic) ─────────────────────────
# 115x115|>130  → 2/block:  20 pallets = 10 tall blocks  H=230cm (door_over)
# 115x115|89-130→ 4/block:  60 pallets = 15 door_ok blocks H=240cm
# Expect: tall blocks loaded from rear, 1 door_ok per container until last tall container
save("T4_tall_short_mix.xlsx", [
    _row("Bravo Shelf Unit Tall",       "A2", "1.15x1.15x1.50", 20, weight=85),
    _row("Bravo Vase Round Large",      "A2", "1.15x1.15x1.20", 60, weight=32),
    _np("Bravo Bowl Small",                   "0.36x0.36x0.43", 40),
])

# ── T5 — large order, 5 containers ───────────────────────────────────────────
# 200 pallets 115x115|89-130 → 50 blocks → 5 containers of 10 blocks each
# plenty of NP boxes
save("T5_large_5containers.xlsx", [
    _row("Bravo Vase Tall White",       "A2", "1.15x1.15x1.20", 80),
    _row("Bravo Vase Tall Black",       "A2", "1.15x1.15x1.20", 80),
    _row("Bravo Vase Tall Bronze",      "A2", "1.15x1.15x1.20", 40),
    _np("Bravo Pot Round L Grey Zinc",        "0.46x0.46x0.87", 60),
    _np("Bravo Bowl Round L Grey Zinc",       "0.36x0.36x0.43", 60),
    _np("Bravo Lantern Outdoor S",            "0.20x0.20x0.35", 120),
])

# ── T6 — full stress, all types, ~7 containers ───────────────────────────────
# 115x115|89-130 → 4/block: 160 pallets = 40 blocks
# 115x77 |89-130 → 4/block:  40 pallets = 10 blocks
# 77x77  |89-130 → 6/block:  24 pallets =  4 blocks
# 115x115|>130   → 2/block:  20 pallets = 10 tall blocks
# total blocks: 40+10+4+10 = 64 → ~6.4 containers
# large NP pool
save("T6_full_stress.xlsx", [
    _row("Bravo Vase Round Xl White",   "A2", "1.15x1.15x1.20", 80, weight=28),
    _row("Bravo Vase Round Xl Black",   "A2", "1.15x1.15x1.20", 80, weight=28),
    _row("Bravo Pot Oval Medium",       "A2", "1.15x0.77x1.20", 40, weight=22),
    _row("Bravo Pot Square Small",      "A2", "0.77x0.77x1.20", 24, weight=18),
    _row("Bravo Shelf Unit Tall White", "A2", "1.15x1.15x1.50", 20, weight=90),
    _np("Bravo Pot Round L Grey Zinc",        "0.46x0.46x0.87", 60),
    _np("Bravo Bowl Round L Grey Zinc",       "0.36x0.36x0.43", 60),
    _np("Bravo Bowl Round L Blue",            "0.36x0.36x0.43", 40),
    _np("Bravo Pot Round L Brown",            "0.46x0.46x0.87", 30),
    _np("Bravo Lantern Outdoor S",            "0.20x0.20x0.35", 60),
    _np("Bravo Lantern Outdoor M",            "0.25x0.25x0.45", 40),
])


print("\nDone. Upload any file from  test_instances/  to the web app.")
