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
    * Exact pallet heights — each (footprint, exact_height) is its own bucket.
    * Rotation — rectangular footprints generate blocks for BOTH orientations with
      the CORRECT pallets_across per orientation.
    * Dual stack heights — when usable container height allows more stacks than the
      door opening, BOTH a full-ceiling block (for back rows) and a door-limited
      block (for the last/door row) are created from the same pallet pool.
      The solver's C9 constraint places full-stack blocks at the back automatically.
    * Greedy allocation — pallets are assigned to the largest available block
      configuration first, guaranteeing full coverage with no leftover.
    """
    usable_H = H_cm - _CEILING_BUFFER_CM

    def _pa(across_dim: int) -> int:
        return max(1, W_cm // across_dim)

    def _display_key(L: int, W: int, Hraw: int) -> str:
        return f"{L}x{W}|{Hraw}cm"

    # ── Build all block configurations for a (footprint, height) bucket ───
    # A configuration is (row_depth, pallets_across, stack_count).
    # We generate every combination of orientation × stack-height variant,
    # deduplicating by k = pa * stacks (keep the one with more stacks for same k).
    def _all_configs(L: int, W: int, Hraw: int) -> List[Tuple[int,int,int]]:
        s_back = max(1, usable_H  // Hraw)   # stacks fitting within ceiling
        s_door = max(1, Hdoor_cm  // Hraw)   # stacks fitting through door
        stacks = [s_back] if s_back == s_door else [s_back, s_door]

        orientations = [(L, _pa(W))] if L == W else [(L, _pa(W)), (W, _pa(L))]

        seen: Dict[int, Tuple[int,int,int]] = {}
        for row_depth, pa in orientations:
            for s in stacks:
                k = pa * s
                if k not in seen or s > seen[k][2]:
                    seen[k] = (row_depth, pa, s)

        # Sort descending by k so greedy allocation prefers larger blocks
        return sorted(seen.values(), key=lambda c: -(c[1] * c[2]))

    # ── Greedy allocation: fill n pallets with blocks from configs ─────────
    def _allocate(n: int, configs: List[Tuple[int,int,int]]) -> Optional[List[int]]:
        """Returns per-config block counts (descending k order), or None if impossible."""
        counts = []
        remaining = n
        for row_depth, pa, s in configs:
            k = pa * s
            c = remaining // k
            counts.append(c)
            remaining -= c * k
        return counts if remaining == 0 else None

    def _min_to_add_multi(n: int, configs: List[Tuple[int,int,int]]) -> int:
        """Minimum pallets to add so _allocate returns non-None."""
        k_vals = [c[1] * c[2] for c in configs]
        g = k_vals[0]
        for k in k_vals[1:]:
            g = math.gcd(g, k)
        r = (g - n % g) % g
        min_k = min(k_vals)
        while n + r < min_k:
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
        if _pa(L) < 1 and _pa(W) < 1:
            warnings.append(
                f"Pallet {L}×{W} cm too wide for container width {W_cm} cm. Skipping."
            )
            continue
        buckets.setdefault((L, W, Hraw), []).append(pm)

    # ── Multiples check ────────────────────────────────────────────────────
    recommendations: Dict[str, int] = {}
    for (L, W, Hraw), plist in buckets.items():
        configs = _all_configs(L, W, Hraw)
        dk = _display_key(L, W, Hraw)
        n  = len(plist)
        if _allocate(n, configs) is None:
            recommendations[dk] = (
                recommendations.get(dk, 0) + _min_to_add_multi(n, configs)
            )

    if require_multiples and any(recommendations.values()):
        return [], recommendations, warnings

    # ── Build block instances ──────────────────────────────────────────────
    blocks: List[BlockInstance] = []
    block_id = 1

    for (L, W, Hraw), plist in buckets.items():
        configs = _all_configs(L, W, Hraw)
        dk      = _display_key(L, W, Hraw)
        counts = _allocate(len(plist), configs)
        if counts is None:
            continue   # require_multiples=False: partial chunk — drop silently

        idx = 0
        for (row_depth, pa, s), n_blocks in zip(configs, counts):
            k       = pa * s
            block_h = s * Hraw
            for _ in range(n_blocks):
                chunk = plist[idx:idx+k];  idx += k
                w_sum = sum(float(pm.get("weight_kg") or 0.0) for pm in chunk)
                blocks.append(BlockInstance(
                    block_id=block_id,
                    block_type_key=dk,
                    length_cm=int(row_depth),
                    height_cm=int(block_h),
                    weight_kg=w_sum,
                    value=int(k),
                    pallets=chunk,
                    pallets_across=int(pa),
                ))
                block_id += 1

    if blocks:
        print("[oneDbuildblocks] unique block heights:",
              sorted({b.height_cm for b in blocks})[:20])

    return blocks, recommendations, warnings