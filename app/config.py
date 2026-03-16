"""config.py — Configuration sourced from environment variables.

Set any of these environment variables to override the built-in defaults:
  CONTAINER_LENGTH_CM, CONTAINER_WIDTH_CM, CONTAINER_HEIGHT_CM,
  CONTAINER_DOOR_HEIGHT_CM, CONTAINER_MAX_WEIGHT_KG, ROW_GAP_CM,
  SOLVER_TIME_LIMIT_SEC, RECOMMEND_OBJECTIVE, RECOMMEND_SECONDARY_OBJECTIVE

On Render (or any cloud host): set these in the Environment tab.
Locally: export them in your shell, or create a .env file and load it before starting.
"""

import os

_DEFAULTS: dict = {
    # Container internal dimensions (cm)
    "CONTAINER_LENGTH_CM":           1203,
    "CONTAINER_WIDTH_CM":            235,
    "CONTAINER_HEIGHT_CM":           270,
    "CONTAINER_DOOR_HEIGHT_CM":      250,
    # Weight limit (kg)
    "CONTAINER_MAX_WEIGHT_KG":       18000,
    # Gap between consecutive pallet row-blocks along the container length (cm)
    "ROW_GAP_CM":                    5,
    # Solver wall-clock time limit per container (seconds)
    # 3 s is enough for typical orders; raise via env var for very large/complex ones
    "SOLVER_TIME_LIMIT_SEC":         3,
    # Recommendation objective (see recommend.py for valid values)
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
