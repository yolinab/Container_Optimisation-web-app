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
    length_cm: int
    height_cm: int
    weight_kg: float
    value: int
    pallets: List[Dict[str,Any]]  # meta dicts of pallets included


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
      blocks: list of BlockInstance (each is a full row-block instance)
      recommendations: dict type_key -> how many pallets to add to reach next multiple
      warnings: list of strings for any rejected pallets/types
    """
    type_table = build_block_type_table(W_cm, H_cm, Hdoor_cm)
    # bucket_key = "<footprint>|<band>|s<stack_count>" so pallets with different
    # effective stack counts never share a bucket (fixes inconsistent stacking).
    buckets:     Dict[str, List[Dict[str,Any]]] = {}
    bucket_info: Dict[str, Dict] = {}   # bucket_key -> {base_key, stack_count, pallets_across, allowed_lengths}
    warnings: List[str] = []

    # 1) assign each pallet to a bucket keyed by (footprint, band, per-pallet stack count)
    for pm in meta_per_pallet:
        Lraw, Wraw, Hraw = int(pm["length"]), int(pm["width"]), int(pm["height"])
        fp = canonical_footprint(Lraw, Wraw, tol=tol_cm)
        if fp is None:
            warnings.append(f"Unknown footprint for pallet_id={pm.get('pallet_id')} dims=({Lraw},{Wraw}).")
            continue

        band = classify_height_band(Hraw)
        base_key = f"{fp[0]}x{fp[1]}|{band}"
        if base_key not in type_table:
            warnings.append(f"No block type rule for pallet_id={pm.get('pallet_id')} key={base_key}.")
            continue

        # Each pallet's own stack count — how many of THIS pallet fit under the door.
        # Using per-pallet height means 103 cm and 130 cm pallets in the same band
        # go into separate buckets and always get the correct stack count.
        per_pallet_stack = max(1, Hdoor_cm // Hraw)
        bucket_key = f"{base_key}|s{per_pallet_stack}"

        buckets.setdefault(bucket_key, []).append(pm)
        if bucket_key not in bucket_info:
            bt = type_table[base_key]
            pallets_across = max(1, bt.pallets_per_block // bt.stack_count)
            bucket_info[bucket_key] = {
                "base_key":        base_key,
                "stack_count":     per_pallet_stack,
                "pallets_across":  pallets_across,
                "allowed_lengths": bt.allowed_lengths,
            }

    # 2) compute recommendations for multiples (keyed by base_key for display)
    recommendations: Dict[str,int] = {}
    for bucket_key, plist in buckets.items():
        info = bucket_info[bucket_key]
        k    = info["pallets_across"] * info["stack_count"]
        rem  = len(plist) % k
        if rem != 0:
            base_key = info["base_key"]
            # accumulate in case multiple sub-buckets map to the same base_key
            recommendations[base_key] = recommendations.get(base_key, 0) + (k - rem)

    if require_multiples and any(recommendations.values()):
        # return no blocks; caller should show recommendations and stop
        return [], recommendations, warnings

    # 3) build block instances by chunking
    blocks: List[BlockInstance] = []
    block_id = 1
    for bucket_key, plist in buckets.items():
        info = bucket_info[bucket_key]
        base_key        = info["base_key"]
        es              = info["stack_count"]
        pallets_across  = info["pallets_across"]
        k               = pallets_across * es
        allowed_lengths = info["allowed_lengths"]

        for start in range(0, len(plist), k):
            chunk = plist[start:start+k]
            if len(chunk) < k:
                continue

            w_sum = sum(float(pm.get("weight_kg") or 0.0) for pm in chunk)
            max_pallet_h    = max(int(pm["height"]) for pm in chunk)
            block_height_cm = es * max_pallet_h

            for Lopt in allowed_lengths:
                blocks.append(BlockInstance(
                    block_id=block_id,
                    block_type_key=base_key,   # display key — no stack suffix
                    length_cm=int(Lopt),
                    height_cm=int(block_height_cm),
                    weight_kg=w_sum,
                    value=int(k),
                    pallets=chunk
                ))
            block_id += 1


    if blocks:
        print("[oneDbuildblocks] unique block heights:", sorted({b.height_cm for b in blocks})[:20])


    return blocks, recommendations, warnings