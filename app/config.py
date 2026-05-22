"""config.py — Configuration sourced from environment variables.

Container type is selected via CONTAINER_TYPE (default: "40HC").
All container dimensions default to the selected preset but can be
individually overridden with env vars if needed.

  CONTAINER_TYPE            20FT | 40FT | 40HC  (default: 40HC)
  CONTAINER_LENGTH_CM       override preset length
  CONTAINER_WIDTH_CM        override preset width
  CONTAINER_HEIGHT_CM       override preset internal height
  CONTAINER_DOOR_HEIGHT_CM  override preset door opening height
  CONTAINER_MAX_WEIGHT_KG   override preset max payload weight
  ROW_GAP_CM                gap between pallet row-blocks (default: 5)
  SOLVER_TIME_LIMIT_SEC     CP-SAT time limit per container (default: 3)

On Render (or any cloud host): set these in the Environment tab.
Locally: export them in your shell before starting.
"""

import os

# ── Container presets ──────────────────────────────────────────────────────
# Internal dimensions (cm) and standard door opening heights.
# Door heights are the ISO standard opening; use CONTAINER_DOOR_HEIGHT_CM
# env var to set a more conservative value if needed.
CONTAINER_PRESETS = {
    "20FT": dict(length_cm=590,  width_cm=235, height_cm=239, door_height_cm=229, max_weight_kg=18000),
    "40FT": dict(length_cm=1203, width_cm=235, height_cm=239, door_height_cm=229, max_weight_kg=18000),
    "40HC": dict(length_cm=1203, width_cm=235, height_cm=269, door_height_cm=259, max_weight_kg=18000),
}

# Active container type — change with CONTAINER_TYPE env var
_raw_type = os.environ.get("CONTAINER_TYPE", "40HC").strip().upper()
if _raw_type not in CONTAINER_PRESETS:
    print(f"[config] WARNING: unknown CONTAINER_TYPE={_raw_type!r} — falling back to '40HC'")
    _raw_type = "40HC"
ACTIVE_CONTAINER_TYPE: str = _raw_type
_preset = CONTAINER_PRESETS[ACTIVE_CONTAINER_TYPE]

# ── Defaults (sourced from preset) ────────────────────────────────────────
_DEFAULTS: dict = {
    "CONTAINER_LENGTH_CM":           _preset["length_cm"],
    "CONTAINER_WIDTH_CM":            _preset["width_cm"],
    "CONTAINER_HEIGHT_CM":           _preset["height_cm"],
    "CONTAINER_DOOR_HEIGHT_CM":      _preset["door_height_cm"],
    "CONTAINER_MAX_WEIGHT_KG":       _preset["max_weight_kg"],
    "ROW_GAP_CM":                    5,
    "SOLVER_TIME_LIMIT_SEC":         3,
    "RECOMMEND_OBJECTIVE":           "min_leftover",
    "RECOMMEND_SECONDARY_OBJECTIVE": "min_pallets",
}


def _get_int(key: str) -> int:
    val = os.environ.get(key)
    if val is not None:
        try:
            return int(val)
        except ValueError:
            print(f"[config] WARNING: env var {key}={val!r} is not a valid integer — using default {_DEFAULTS[key]}")
    return int(_DEFAULTS[key])


def _get_str(key: str) -> str:
    return os.environ.get(key, str(_DEFAULTS[key]))


CONTAINER_LENGTH_CM           = _get_int("CONTAINER_LENGTH_CM")
CONTAINER_WIDTH_CM            = _get_int("CONTAINER_WIDTH_CM")
CONTAINER_HEIGHT_CM           = _get_int("CONTAINER_HEIGHT_CM")
CONTAINER_DOOR_HEIGHT_CM      = _get_int("CONTAINER_DOOR_HEIGHT_CM")
CONTAINER_MAX_WEIGHT_KG       = _get_int("CONTAINER_MAX_WEIGHT_KG")
ROW_GAP_CM                    = _get_int("ROW_GAP_CM")
SOLVER_TIME_LIMIT_SEC         = _get_int("SOLVER_TIME_LIMIT_SEC")
RECOMMEND_OBJECTIVE           = _get_str("RECOMMEND_OBJECTIVE")
RECOMMEND_SECONDARY_OBJECTIVE = _get_str("RECOMMEND_SECONDARY_OBJECTIVE")

# Compat exports used by main.py logging
_CONFIG_SOURCE  = "environment variables"
_USING_DEFAULTS = False
