import math
import matplotlib.pyplot as plt
from mpl_toolkits.mplot3d import Axes3D
from collections import Counter
from typing import List, Dict, Any, Optional


# ------------------------------------------------------------
# Helpers: summarize pallet composition inside a row-block
# ------------------------------------------------------------

def summarize_pallets(pallets, max_items=3):
    """
    Create a compact summary string of pallet composition.
    Groups by (length x width x height).
    Robust to key naming style: 'length' vs 'length_cm', etc.
    """
    if not pallets:
        return "empty"

    def _dim(pm, k1, k2):
        return pm.get(k1, pm.get(k2, "?"))

    counter = Counter(
        f"{_dim(p,'length','length_cm')}x{_dim(p,'width','width_cm')}x{_dim(p,'height','height_cm')}"
        for p in pallets
    )

    parts = []
    for k, v in counter.most_common(max_items):
        parts.append(f"{k}×{v}")

    if len(counter) > max_items:
        parts.append("...")

    return ", ".join(parts)


# ------------------------------------------------------------
# Build 3D boxes from container rows
# ------------------------------------------------------------

def build_boxes_from_row_blocks(container_rows, container_width_cm):
    """
    Convert row-blocks into 3D boxes compatible with plot_boxes_3d().
    """
    boxes = []

    for i, r in enumerate(container_rows):
        boxes.append({
            "id": i + 1,
            "x": 0,  # full width blocks start at x=0
            "y": r["y_start_cm"],
            "z": 0,
            "w": container_width_cm,
            "l": r["length_cm"],
            "h": r["height_cm"],
            "block_type": r["block_type"],
            "components": summarize_pallets(r["pallets"]),
        })

    return boxes


# ------------------------------------------------------------
# NP box zone helpers
# ------------------------------------------------------------

# Distinct warm colours for box zones (different palette from pallets)
_BOX_ZONE_FILL_COLORS = ["khaki", "lightcyan", "palegreen", "thistle", "peachpuff"]


def _draw_box_wireframe(
    ax,
    x: float, y: float, z: float,
    w: float, l: float, h: float,
    color: str = "darkorange",
    linestyle: str = "--",
    linewidth: float = 1.1,
    alpha: float = 0.85,
) -> None:
    """Draw all 12 edges of a 3-D box as styled lines (dashed by default)."""
    x1, y1, z1 = x + w, y + l, z + h
    edges = [
        # bottom face
        ([x, x1], [y,  y ], [z,  z ]),
        ([x1, x1], [y, y1], [z,  z ]),
        ([x1, x ], [y1,y1], [z,  z ]),
        ([x,  x ], [y1, y], [z,  z ]),
        # top face
        ([x, x1], [y,  y ], [z1, z1]),
        ([x1, x1], [y, y1], [z1, z1]),
        ([x1, x ], [y1,y1], [z1, z1]),
        ([x,  x ], [y1, y], [z1, z1]),
        # vertical edges
        ([x,  x ], [y,  y ], [z, z1]),
        ([x1, x1], [y,  y ], [z, z1]),
        ([x1, x1], [y1, y1], [z, z1]),
        ([x,  x ], [y1, y1], [z, z1]),
    ]
    for xs, ys, zs in edges:
        ax.plot3D(xs, ys, zs,
                  color=color, linestyle=linestyle,
                  linewidth=linewidth, alpha=alpha)


def build_box_zone_visuals(
    container_box_zones: Optional[List[Dict[str, Any]]],
) -> List[Dict[str, Any]]:
    """
    Convert container['box_zones'] into renderable dicts for plot_boxes_3d().

    Each returned dict has:
      x, y, z, w, l, h_fill   — filled-volume cuboid (volume-equivalent height)
      h_zone                   — full available zone height (for the wireframe outline)
      color, zone_type, n_boxes, label, legend_line
    """
    if not container_box_zones:
        return []

    visuals = []
    for zi, zone in enumerate(container_box_zones):
        zone_L = zone["length_cm"]
        zone_W = zone["width_cm"]
        vol_used = zone.get("volume_used_cm3", 0.0)
        zone_H_max = zone["height_cm"]

        # zone_L is now the actual length used by placed boxes (from the length cursor),
        # so render the zone at that size. Fill height = vol / (length × width).
        dense_length = max(1, zone_L)
        footprint = dense_length * zone_W
        h_fill = (vol_used / footprint) if footprint > 0 else 0.0
        h_fill = min(h_fill, zone_H_max)

        n_boxes = sum(p["quantity"] for p in zone.get("placed", []))
        wt_kg   = zone.get("total_weight_kg", 0.0)
        vol_m3  = vol_used / 1e6

        # Build "product name (LxWxHcm)" summary — group by product name
        name_info: dict = {}  # name -> {"dim": str, "count": int}
        for p in zone.get("placed", []):
            nm  = p["label"]
            dim = f"{p['length_cm']}×{p['width_cm']}×{p['height_cm']}cm"
            if nm not in name_info:
                name_info[nm] = {"dim": dim, "count": 0}
            name_info[nm]["count"] += p["quantity"]

        def _short(s, n=20):
            return s if len(s) <= n else s[:n - 1] + "…"

        legend_parts = [
            f"{v['count']}× {_short(nm)} ({v['dim']})"
            for nm, v in list(name_info.items())[:2]
        ]
        if len(name_info) > 2:
            legend_parts.append(f"+{len(name_info) - 2} more")
        dim_summary  = ", ".join(legend_parts)
        label_short  = dim_summary or f"{n_boxes} boxes"

        visuals.append({
            "x":        0,
            "y":        zone["y_start_cm"],
            "z":        zone["z_base_cm"],
            "w":        zone_W,
            "l":        dense_length,
            "h_fill":   h_fill,
            "h_zone":   zone_H_max,
            "color":    _BOX_ZONE_FILL_COLORS[zi % len(_BOX_ZONE_FILL_COLORS)],
            "zone_type": zone["zone_type"],
            "n_boxes":  n_boxes,
            "label":    label_short,
            "legend_line": (
                f"  NP ({zone['zone_type']}): {n_boxes} boxes | "
                f"{vol_m3:.3f} m³"
                + (f" | {wt_kg:.0f} kg" if wt_kg else "")
                + f" | {dim_summary}"
            ),
        })
    return visuals


# ------------------------------------------------------------
# Recommended-block visualization helpers
# ------------------------------------------------------------

_REC_PALLET_FILL_COLOR = "palegreen"    # recommended pallet blocks
_REC_PALLET_WIRE_COLOR = "forestgreen"
_REC_NP_BOX_FILL_COLOR = "lightcyan"   # recommended NP box rows
_REC_NP_BOX_WIRE_COLOR = "steelblue"

# Legacy aliases kept so nothing else breaks
_REC_FILL_COLOR = _REC_PALLET_FILL_COLOR
_REC_WIRE_COLOR = _REC_PALLET_WIRE_COLOR


def build_rec_block_visuals(
    rec: Optional[Dict[str, Any]],
    container: Dict[str, Any],
    W: int,
    gap_cm: int,
) -> List[Dict[str, Any]]:
    """
    Convert a recommendation dict into individual renderable box dicts.

    The new recommendation format stores flat lists 'tail_placements' and
    'atop_placements' with pre-computed y_start_cm and z_base_cm, so no
    cursor reconstruction is needed here.
    """
    if rec is None:
        return []

    visuals: List[Dict[str, Any]] = []

    all_placements = (
        rec.get("tail_placements", []) + rec.get("atop_placements", [])
    )

    for p in all_placements:
        is_np = p.get("type") == "np_box"
        fill_color = _REC_NP_BOX_FILL_COLOR if is_np else _REC_PALLET_FILL_COLOR
        wire_color = _REC_NP_BOX_WIRE_COLOR if is_np else _REC_PALLET_WIRE_COLOR

        if is_np:
            n_units = p.get("units_per_placement", p.get("n_across", 1))
            label   = f"+{n_units}b"
        else:
            ppb   = p.get("pallets_per_block", 0)
            label = f"+{ppb}p"

        # For NP boxes: carry product name for legend; fall back to dim key
        label_display = p.get("label", p["key"]) if is_np else p["key"]

        visuals.append({
            "x":               0,
            "y":               p["y_start_cm"],
            "z":               p["z_base_cm"],
            "w":               W,
            "l":               p["length_cm"],
            "h":               p["height_cm"],
            "color":           fill_color,
            "wire_color":      wire_color,
            "zone":            p.get("zone", "tail"),
            "type":            p.get("type", "pallet"),
            "block_type_key":  p["key"],
            "label_display":   label_display,
            "pallets_per_block": p.get("pallets_per_block", 0),
            "n_boxes":         n_units if is_np else 0,
            "label":           label,
        })

    return visuals


# ------------------------------------------------------------
# Main 3D plotting function (refactored from your original)
# ------------------------------------------------------------

def plot_boxes_3d(W, L, H, boxes, box_zone_visuals=None, rec_box_visuals=None, title=None):
    fig = plt.figure(figsize=(12, 7))
    ax = fig.add_subplot(111, projection="3d")

    ax.set_xlim(0, W)
    ax.set_ylim(0, L)
    ax.set_zlim(0, H)
    ax.set_box_aspect((W, L, H))

    colors = [
        "tab:blue", "tab:orange", "tab:green",
        "tab:red", "tab:purple", "tab:brown",
        "tab:pink", "tab:gray", "tab:olive", "tab:cyan"
    ]

    # --- Pallets ---
    for i, b in enumerate(boxes):
        ax.bar3d(
            b["x"], b["y"], b["z"],
            b["w"], b["l"], b["h"],
            alpha=0.55,
            color=b.get("color", colors[i % len(colors)]),
            edgecolor="k",
            linewidth=0.6,
            shade=True,
        )

        cx = b["x"] + b["w"] / 2
        cy = b["y"] + b["l"] / 2
        cz = b["z"] + b["h"] / 2

        ax.text(cx, cy, cz, str(b["id"]), color="k", fontsize=9, ha="center")

    # --- NP box zones ---
    if box_zone_visuals:
        for bz in box_zone_visuals:
            x, y, z = bz["x"], bz["y"], bz["z"]
            w, l    = bz["w"], bz["l"]
            h_fill  = bz["h_fill"]

            # Draw actual box volume only (volume-equivalent height bar)
            if h_fill > 0.5:
                ax.bar3d(x, y, z, w, l, h_fill,
                         alpha=0.55, color=bz["color"],
                         edgecolor="darkorange", linewidth=0.8, shade=True)
            else:
                # Nearly nothing placed — just draw a thin marker line
                _draw_box_wireframe(ax, x, y, z, w, l, 2,
                                    color="darkorange", linestyle="--", linewidth=0.8)

            # Label just above the bar
            label_z = z + max(h_fill, 3) + 2
            ax.text(x + w / 2, y + l / 2, label_z,
                    f"NP ×{bz['n_boxes']}\n({bz['zone_type']})",
                    color="darkorange", fontsize=7,
                    ha="center", va="bottom", fontweight="bold")

    # --- Recommended additions ---
    if rec_box_visuals:
        for rb in rec_box_visuals:
            x, y, z = rb["x"], rb["y"], rb["z"]
            w, l, h = rb["w"], rb["l"], rb["h"]
            wc = rb.get("wire_color", _REC_PALLET_WIRE_COLOR)

            # Dotted wireframe — green for pallets, blue for NP box rows
            _draw_box_wireframe(ax, x, y, z, w, l, h,
                                color=wc, linestyle=":", linewidth=1.5)

            # Semi-transparent fill
            ax.bar3d(x, y, z, w, l, h,
                     alpha=0.28, color=rb["color"],
                     edgecolor="none", shade=False)

            # Label at block centre
            ax.text(x + w / 2, y + l / 2, z + h / 2,
                    rb["label"],
                    color=wc, fontsize=8,
                    ha="center", va="center", fontweight="bold")

    ax.set_xlabel("X — container width (cm)")
    ax.set_ylabel("Y — container length (cm)")
    ax.set_zlabel("Z — height (cm)")

    if title:
        ax.set_title(title, fontsize=14, pad=12)

    # ---------------- Legend text ----------------
    legend_lines = []
    seen = set()

    for b in boxes:
        legend_id   = b.get("legend_id")
        legend_line = b.get("legend_line")
        if legend_id is not None and legend_line is not None:
            if legend_id not in seen:
                legend_lines.append(legend_line)
                seen.add(legend_id)
            continue
        # Fallback: original per-box legend
        line = (
            f"{b.get('id','?'):>2}: {b.get('block_type','')} | "
            f"L={b.get('l','?')} H={b.get('h','?')} | "
            f"{b.get('components','')}"
        )
        legend_lines.append(line)

    # Append NP box zone legend entries
    if box_zone_visuals:
        legend_lines.append("")  # blank separator
        for bz in box_zone_visuals:
            legend_lines.append(bz["legend_line"])

    # Append recommended block legend entries
    if rec_box_visuals:
        legend_lines.append("")
        legend_lines.append("  ++ RECOMMENDED ADDITIONS:")

        # Tally counts and units; also capture display name per (key, zone)
        pal_counts:    dict = {}
        pal_units:     dict = {}
        np_counts:     dict = {}
        np_units:      dict = {}
        np_name:       dict = {}   # (key, zone) -> product name
        for rb in rec_box_visuals:
            k = (rb["block_type_key"], rb.get("zone", ""))
            if rb.get("type") == "np_box":
                np_counts[k] = np_counts.get(k, 0) + 1
                np_units[k]  = np_units.get(k, 0)  + rb.get("n_boxes", 0)
                np_name[k]   = rb.get("label_display", k[0])
            else:
                pal_counts[k] = pal_counts.get(k, 0) + 1
                pal_units[k]  = pal_units.get(k, 0)  + rb.get("pallets_per_block", 0)

        def _trunc(s, n=28):
            return s if len(s) <= n else s[:n - 1] + "…"

        seen_rec: set = set()
        for rb in rec_box_visuals:
            k = (rb["block_type_key"], rb.get("zone", ""))
            if k in seen_rec:
                continue
            seen_rec.add(k)
            if rb.get("type") == "np_box":
                name = _trunc(np_name.get(k, k[0]))
                legend_lines.append(
                    f"     (boxes) +{np_units[k]}× {name}"
                    f"  [{k[0]}]  [{k[1]}]"
                )
            else:
                legend_lines.append(
                    f"     (pallet) +{pal_units[k]}× {k[0]}"
                    f"  [{k[1]}]"
                )

    fig.text(
        0.02,
        0.02,
        "\n".join(legend_lines),
        fontsize=9,
        family="monospace",
        va="bottom",
        ha="left",
    )
    # ------------------------------------------------

    plt.tight_layout()
    plt.show()


# ------------------------------------------------------------
# Public API — what you actually call
# ------------------------------------------------------------

def plot_row_block_container(container_info, W, L, H):
    """
    Plot a single container solution (one figure).
    """
    boxes = build_boxes_from_row_blocks(container_info["rows"], W)
    title = f"Container {container_info['container_index']} — Row-Block Layout"
    plot_boxes_3d(W, L, H, boxes, title=title)


def plot_all_row_block_containers(containers, W, L, H):
    """
    Plot all containers, one figure per container.
    """
    for c in containers:
        plot_row_block_container(c, W, L, H)


# ------------------------------------------------------------
# NEW: Pallet-level plotting (pallets colored by row-block)
# ------------------------------------------------------------

def build_pallet_boxes_from_row_blocks(container_rows, container_width_cm, gap_cm=5):
    """
    Expand each row-block into individual pallet cuboids.

    - All pallets within the same row-block share the same color.
    - Layout is deterministic and simple:
        * across = 3 if footprint is 77x77 else 2
        * pallets are placed left-to-right across X
        * pallets are stacked in layers along Z
        * all pallets in a row-block share the same Y start (the row start)

    This is a visualisation convenience: it is not a physics re-check.
    """

    palette = [
        "tab:blue", "tab:orange", "tab:green",
        "tab:red", "tab:purple", "tab:brown",
        "tab:pink", "tab:gray", "tab:olive", "tab:cyan"
    ]

    pallet_boxes = []

    for block_idx, r in enumerate(container_rows, start=1):
        pallets = r.get("pallets", [])
        if not pallets:
            continue

        # Row start and row dims
        y0 = int(r.get("y_start_cm", 0))
        row_len = int(r.get("length_cm", 0))
        row_h = int(r.get("height_cm", 0))
        block_type = str(r.get("block_type", ""))

        # Determine pallet footprint from first pallet meta
        p0 = pallets[0]
        Lp = int(p0.get("length", p0.get("length_cm", 0)))
        Wp = int(p0.get("width",  p0.get("width_cm",  0)))

        # Choose orientation so pallet length along Y matches row_len when possible
        if row_len == Lp:
            pallet_len_y, pallet_wid_x = Lp, Wp
        elif row_len == Wp:
            pallet_len_y, pallet_wid_x = Wp, Lp
        else:
            pallet_len_y, pallet_wid_x = Lp, Wp

        # Across heuristic (matches your business assumption)
        across = 3 if (Lp == 77 and Wp == 77) else 2

        # Z step: use tallest pallet height in the block to avoid overlaps
        heights = [int(pm.get("height", pm.get("height_cm", 0))) for pm in pallets]
        z_step = max(heights) if heights else 0

        color = palette[(block_idx - 1) % len(palette)]

        # Block-level legend line (one per block)
        comps = summarize_pallets(pallets)
        legend_line = (
            f"{block_idx:>2}: {block_type} | row L={row_len} H={row_h} | {comps}"
        )

        for i, pm in enumerate(pallets):
            layer = i // across
            pos = i % across

            x = pos * pallet_wid_x
            y = y0
            z = layer * z_step

            h_cm = int(pm.get("height", pm.get("height_cm", 0)))

            pallet_boxes.append({
                "id": i + 1,
                "x": x,
                "y": y,
                "z": z,
                "w": pallet_wid_x,
                "l": pallet_len_y,
                "h": h_cm,
                "color": color,
                # provide grouped legend info
                "legend_id": block_idx,
                "legend_line": legend_line,
            })

    return pallet_boxes


def plot_row_block_container_pallets(container_info, W, L, H, gap_cm=5, rec=None):
    """Plot a single container showing individual pallets + NP box zones + recommended additions."""
    rows = container_info.get("rows", [])
    boxes = build_pallet_boxes_from_row_blocks(rows, W, gap_cm=gap_cm)
    box_zone_visuals = build_box_zone_visuals(container_info.get("box_zones"))
    rec_box_visuals  = build_rec_block_visuals(rec, container_info, W, gap_cm)
    title = f"Container {container_info.get('container_index','')} — Pallet + NP Box View"
    plot_boxes_3d(W, L, H, boxes,
                  box_zone_visuals=box_zone_visuals,
                  rec_box_visuals=rec_box_visuals,
                  title=title)
    return boxes


def plot_all_row_block_containers_pallets(containers, W, L, H, gap_cm=5, recs=None):
    """Plot each container showing individual pallets (colored by row-block)."""
    recs_by_idx = {r["container_index"]: r for r in (recs or [])}
    all_boxes = []
    for c in containers:
        rec = recs_by_idx.get(c.get("container_index"))
        all_boxes.append(plot_row_block_container_pallets(c, W, L, H, gap_cm=gap_cm, rec=rec))
    return all_boxes