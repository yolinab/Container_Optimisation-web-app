from dataclasses import dataclass
from typing import Any, Dict, List, Tuple, Optional
import math

@dataclass
class BlockType:
    key: str                 # e.g. "115x115|<66"
    footprint: Tuple[int,int]# canonical (L,W) in cm, unordered ok
    pallets_per_block: int   # e.g. 8, 6, 4, 2, 9
    block_height_cm: int     # conservative height for constraints
    stack_count: int        # how many pallets stacked in this block type
    allowed_lengths: Tuple[int, ...]  # possible row lengths (rotation variants), e.g. (115,) or (115,77)

@dataclass
class BlockInstance:
    block_id: int
    block_type_key: str
    length_cm: int       # row depth in container (along Y axis)
    height_cm: int
    weight_kg: float
    value: int           # total pallets in this block
    pallets: List[Dict[str,Any]]
    pallets_across: int = 0   # pallets side-by-side along container width (X axis)


# ---------- 1) Canonical mapping ----------
def normalize_dim_to_set(x: int, allowed: List[int], tol: int = 2) -> Optional[int]:
    """Map x to nearest allowed value within tolerance; else None."""
    best = None
    best_dist = 10**9
    for a in allowed:
        d = abs(x - a)
        if d < best_dist:
            best_dist = d
            best = a
    return best if best is not None and best_dist <= tol else None


def canonical_footprint(L: int, W: int, tol: int = 2) -> Optional[Tuple[int,int]]:
    """
    Map raw (L,W) to canonical footprint.
    Returns (Lcan, Wcan) with Lcan>=Wcan convention for keys.
    """
    # Allowed sides we expect
    allowed = [115, 108, 77]
    a = normalize_dim_to_set(L, allowed, tol=tol)
    b = normalize_dim_to_set(W, allowed, tol=tol)
    if a is None or b is None:
        return None
    Lcan, Wcan = max(a,b), min(a,b)
    return (Lcan, Wcan)


# ---------- 2) Height band classification ----------
def classify_height_band(h_cm: int) -> str:
    """Return band label used in the block type table."""
    if h_cm == 230:
        return "230"
    if h_cm < 66:
        return "<66"
    if h_cm <= 89:
        return "66-89"
    if h_cm <= 130:
        return "89-130"
    return ">130"


# ---------- 3) Build block type table ----------

# Standard EU/ISO pallet footprints: (L, W, allow_rotate).
# L >= W by convention.  allow_rotate=True means both orientations
# produce valid row lengths (the pallet can face either way along the
# container depth axis).
_PALLET_FOOTPRINTS: List[Tuple[int, int, bool]] = [
    (115, 115, False),
    (115, 108, True),
    (115,  77, True),
    ( 77,  77, False),
]

# Height bands: (name, band_max_h_cm).
# band_max_h is used to compute how many layers fit in the container
# and as a conservative block-height estimate.
_HEIGHT_BANDS: List[Tuple[str, int]] = [
    ("<66",     65),
    ("66-89",   89),
    ("89-130", 130),
    (">130",   230),  # single-stack tall pallets; 230 cm is conservative
    ("230",    230),  # special band for near-container-height pallets
]


def build_block_type_table(W_cm: int, H_cm: int, Hdoor_cm: int) -> Dict[str, BlockType]:
    """
    Build the block-type lookup table from container dimensions.

    For each pallet footprint (L, W) and height band:
      pallets_across    = W_cm // L          (conservative: large footprint dim as depth)
      stack_count       = max(1, H_cm // band_max_h)
      pallets_per_block = pallets_across * stack_count
      block_height_cm   = stack_count * band_max_h   (conservative upper estimate)

    The "230" special band is only added when H_cm >= 230.
    """
    table: Dict[str, BlockType] = {}

    def add(foot: Tuple[int, int], band: str, pallets_per_block: int,
            stack_count: int, block_height_cm: int, allow_rotate: bool) -> None:
        L, W = foot
        lengths = (L, W) if allow_rotate and L != W else (L,)
        key = f"{L}x{W}|{band}"
        table[key] = BlockType(
            key=key,
            footprint=foot,
            pallets_per_block=pallets_per_block,
            block_height_cm=block_height_cm,
            stack_count=stack_count,
            allowed_lengths=tuple(sorted(set(lengths), reverse=True)),
        )

    usable_h = H_cm - 10  # 10 cm clearance buffer between top of stack and ceiling

    for L, W, allow_rotate in _PALLET_FOOTPRINTS:
        # Conservative: use the larger footprint dimension as row depth,
        # so pallets_across uses the smaller clearance direction.
        pallets_across = W_cm // L
        if pallets_across < 1:
            continue  # pallet too wide for this container

        for band_name, band_max_h in _HEIGHT_BANDS:
            if band_name == "230" and H_cm < 230:
                continue  # container too short for near-container-height pallets
            stack_count       = max(1, usable_h // band_max_h)
            pallets_per_block = pallets_across * stack_count
            block_height_cm   = stack_count * band_max_h  # actual pallet height, no buffer added
            add((L, W), band_name, pallets_per_block, stack_count, block_height_cm, allow_rotate)

    return table


# ---------- 4) Main function: pallets -> blocks ----------

_CEILING_BUFFER_CM = 9   # clearance between top of stack and container ceiling


def build_row_blocks_from_pallets(
    meta_per_pallet: List[Dict[str,Any]],
    W_cm: int,
    H_cm: int,
    Hdoor_cm: int,
    tol_cm: int = 2,
    require_multiples: bool = True
) -> Tuple[List[BlockInstance], Dict[str,int], List[str]]:
    """
    Returns:
      blocks: list of BlockInstance
      recommendations: dict display_key -> pallets to add
      warnings: list of issue strings

    Key design decisions
    --------------------
    * Exact pallet heights — each (footprint, exact_height) is its own bucket so
      a 130 cm pallet never contaminates 103 cm stacking decisions.
    * Rotation — for rectangular footprints BOTH orientations are created with the
      CORRECT pallets_across for each.  A 115×77 pallet oriented with 77 cm facing
      across the 235 cm container width fits 3 across (not 2).
    * Stack count uses Hdoor_cm (per-pallet, not per-bucket max) so the stacked
      block always fits through the door.
    * Multiples: with two orientations, a valid partition a*k_A + b*k_B = n is
      found; if none exists the minimum shortfall is recommended.
    """
    usable_H = H_cm - _CEILING_BUFFER_CM

    def _stacks(Hraw: int) -> int:
        """How many pallets of height Hraw stack, limited by door AND ceiling."""
        return max(1, min(Hdoor_cm // Hraw, usable_H // Hraw))

    def _pa(across_dim: int) -> int:
        """Pallets that fit across the container width when each is across_dim cm wide."""
        return max(1, W_cm // across_dim)

    def _display_key(L: int, W: int, Hraw: int) -> str:
        return f"{L}x{W}|{Hraw}cm"

    def _find_split(n: int, k_A: int, k_B: int) -> Optional[Tuple[int, int]]:
        """
        Find a>=0, b>=0 maximising a such that a*k_A + b*k_B = n.
        Returns None if no solution exists.
        """
        for a in range(n // k_A, -1, -1):
            rem = n - a * k_A
            if rem >= 0 and rem % k_B == 0:
                return a, rem // k_B
        return None

    def _min_to_add(n: int, k_A: int, k_B: int) -> int:
        """Minimum r >= 0 such that _find_split(n+r, k_A, k_B) is not None."""
        g = math.gcd(k_A, k_B)
        r = (g - n % g) % g
        while n + r < min(k_A, k_B):
            r += g
        return r

    # ── Bucket by (canonical_L, canonical_W, exact_height) ────────────────
    buckets: Dict[Tuple[int,int,int], List[Dict[str,Any]]] = {}
    warnings: List[str] = []

    for pm in meta_per_pallet:
        Lraw, Wraw, Hraw = int(pm["length"]), int(pm["width"]), int(pm["height"])
        fp = canonical_footprint(Lraw, Wraw, tol=tol_cm)
        if fp is None:
            warnings.append(
                f"Unknown footprint for pallet_id={pm.get('pallet_id')} "
                f"dims=({Lraw},{Wraw}). Skipping."
            )
            continue
        L, W = fp
        # Accept pallet if at least one orientation fits in container width
        if _pa(L) < 1 and _pa(W) < 1:
            warnings.append(
                f"Pallet {L}×{W} cm too wide for container width {W_cm} cm. Skipping."
            )
            continue
        buckets.setdefault((L, W, Hraw), []).append(pm)

    # ── Multiples check ────────────────────────────────────────────────────
    recommendations: Dict[str, int] = {}
    for (L, W, Hraw), plist in buckets.items():
        s  = _stacks(Hraw)
        dk = _display_key(L, W, Hraw)
        n  = len(plist)

        if L == W:
            # Square: single orientation
            k = _pa(W) * s
            if n % k != 0:
                recommendations[dk] = recommendations.get(dk, 0) + (k - n % k)
        else:
            # Rectangular: A = row L-deep (W across), B = row W-deep (L across)
            k_A = _pa(W) * s
            k_B = _pa(L) * s
            if k_A == k_B:
                k = k_A
                if n % k != 0:
                    recommendations[dk] = recommendations.get(dk, 0) + (k - n % k)
            else:
                if _find_split(n, k_A, k_B) is None:
                    recommendations[dk] = (
                        recommendations.get(dk, 0) + _min_to_add(n, k_A, k_B)
                    )

    if require_multiples and any(recommendations.values()):
        return [], recommendations, warnings

    # ── Build block instances ──────────────────────────────────────────────
    blocks: List[BlockInstance] = []
    block_id = 1

    for (L, W, Hraw), plist in buckets.items():
        s       = _stacks(Hraw)
        block_h = s * Hraw
        dk      = _display_key(L, W, Hraw)

        def _make_block(chunk, row_depth, pa):
            nonlocal block_id
            w_sum = sum(float(pm.get("weight_kg") or 0.0) for pm in chunk)
            blocks.append(BlockInstance(
                block_id=block_id,
                block_type_key=dk,
                length_cm=int(row_depth),
                height_cm=int(block_h),
                weight_kg=w_sum,
                value=int(len(chunk)),
                pallets=chunk,
                pallets_across=int(pa),
            ))
            block_id += 1

        if L == W:
            # Square footprint: one orientation
            pa = _pa(W)
            k  = pa * s
            for start in range(0, len(plist), k):
                chunk = plist[start:start+k]
                if len(chunk) < k:
                    continue
                _make_block(chunk, L, pa)
        else:
            k_A = _pa(W) * s   # orientation A: row_depth=L, across_dim=W
            k_B = _pa(L) * s   # orientation B: row_depth=W, across_dim=L

            if k_A == k_B:
                # Same block size — alternate orientations to give solver both depths
                k   = k_A
                pa_A = _pa(W)
                pa_B = _pa(L)
                for i, start in enumerate(range(0, len(plist) // k * k, k)):
                    chunk = plist[start:start+k]
                    if i % 2 == 0:
                        _make_block(chunk, L, pa_A)
                    else:
                        _make_block(chunk, W, pa_B)
            else:
                # Different block sizes — partition via _find_split
                pa_A = _pa(W)
                pa_B = _pa(L)
                a, b = _find_split(len(plist), k_A, k_B)
                idx = 0
                for _ in range(a):
                    chunk = plist[idx:idx+k_A];  idx += k_A
                    _make_block(chunk, L, pa_A)
                for _ in range(b):
                    chunk = plist[idx:idx+k_B];  idx += k_B
                    _make_block(chunk, W, pa_B)

    if blocks:
        print("[oneDbuildblocks] unique block heights:",
              sorted({b.height_cm for b in blocks})[:20])

    return blocks, recommendations, warnings