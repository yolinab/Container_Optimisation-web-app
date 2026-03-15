# Container Packing Optimizer

A web application that takes an Excel order file and produces an optimised container loading plan — pallet row layout, NP (loose) box placement, fill recommendations, and an Excel report.

Deployed on **Render** at the URL configured in your Render dashboard.

---

## Architecture

```
/
├── api.py                   FastAPI app (web entry point)
├── static/index.html        Single-page browser UI
├── Procfile                 Render/Heroku start command
├── requirements.txt
├── runtime.txt              Python 3.10.14
└── app/
    ├── pipeline.py          Web-facing pipeline (no matplotlib)
    ├── main.py              CLI runner (adds plotting + file logging)
    ├── config.py            Config loader (env vars or optimizer_config.json)
    ├── models/
    │   ├── A_1D_multi_container_placement.py   OR-Tools CP-SAT pallet solver
    │   └── box_packing.py                      Column-based NP box packer
    └── utils/
        ├── parse_xlsx.py          Excel ingestion (pallets + NP boxes)
        ├── oneDbuildblocks.py     Pallet → row-block construction + type table
        ├── recommend.py           Fill-recommendation engine
        ├── export_excel.py        Excel report writer
        ├── validate.py            Post-solve sanity checks
        └── visualize_row_blocks.py  3-D matplotlib visualisation (CLI only)
```

---

## How it works

1. **Parse** — Excel file is read; pallets and NP (loose) boxes are extracted separately.
2. **Build row-blocks** — pallets are grouped by footprint × height band into `BlockInstance` objects representing one physical row in a container.
3. **Solve** — an OR-Tools CP-SAT model packs row-blocks into containers one by one (greedy loop). Tall blocks go at the back; the last (door) row must fit within the door height.
4. **Pack NP boxes** — a column-based geometric packer (`BoxPacker`) fills the tail zone of each container with loose boxes. All box types share columns, stacked vertically, eliminating per-type section waste.
5. **Validate** — `validate_packing_result()` checks block coverage, pallet count conservation, geometry bounds, weight limits, and NP box quantities. Issues are returned as structured dicts and shown in the UI.
6. **Recommend** — for each container's free tail the engine recommends additional pallets (distributed proportionally to order quantities) and NP boxes (2-D stacked to door height).
7. **Report** — an Excel workbook is generated with per-container layout sheets, a recommendations sheet, and a config sheet.

---

## Pallet block type rules

Recognised footprints: **115×115 cm**, **115×108 cm**, **115×77 cm**, **77×77 cm** (±2 cm tolerance).

Height bands: `<66 cm`, `66–89 cm`, `89–130 cm`, `>130 cm`, `230 cm`.

Each footprint × band combination has fixed stacking rules (pallets per block, stack count) defined in `utils/oneDbuildblocks.py`. Pallets that don't match a known footprint are silently skipped with a warning.

---

## Configuration

Container dimensions and solver settings are read from environment variables or `optimizer_config.json` (placed next to `main.py`). Environment variables take precedence.

| Variable | Default | Description |
|---|---|---|
| `CONTAINER_LENGTH_CM` | 1203 | Internal usable length |
| `CONTAINER_WIDTH_CM`  | 235  | Internal width |
| `CONTAINER_HEIGHT_CM` | 270  | Internal height |
| `CONTAINER_DOOR_HEIGHT_CM` | 250 | Door opening height (loading constraint) |
| `CONTAINER_MAX_WEIGHT_KG`  | 18000 | Max payload weight |
| `ROW_GAP_CM` | 5 | Fork-lift clearance between pallet rows |
| `SOLVER_TIME_LIMIT_SEC` | 5 | CP-SAT time limit per container |
| `RECOMMEND_OBJECTIVE` | min_leftover | Primary fill objective |
| `RECOMMEND_SECONDARY_OBJECTIVE` | min_pallets | Tiebreaker |

---

## API endpoints

| Method | Path | Description |
|---|---|---|
| `GET`  | `/`           | Serves the browser UI |
| `GET`  | `/health`     | Health check → `{"status": "ok"}` |
| `POST` | `/optimize`   | Upload Excel → JSON with base64 report + validation issues |
| `POST` | `/report-bug` | Send a bug report email |

### POST /optimize

**Form fields:** `file` (Excel), `count_col` (optional column name override).

**Response JSON:**
```json
{
  "report_b64":      "<base64 Excel>",
  "filename":        "packing_report.xlsx",
  "container_count": 7,
  "total_pallets":   256,
  "warnings":        ["All containers report loaded_weight=0 kg …"],
  "errors":          []
}
```

### POST /report-bug

**JSON body:**
```json
{
  "message":         "Describe the problem…",
  "container_count": 7,
  "total_pallets":   256,
  "warnings":        […],
  "errors":          […]
}
```
Sends an email via Office 365 SMTP. Requires env vars (see below).

---

## Bug report email setup

Uses [Resend](https://resend.com) (free, 3,000 emails/month) — no corporate SMTP or app passwords needed.

1. Sign up at **resend.com** (free)
2. Dashboard → API Keys → Create key → copy it
3. Add to **Render → Environment**:

```
RESEND_API_KEY  =  re_xxxxxxxxxxxx
BUG_EMAIL       =  yolina.yordanova@edelman.nl   (optional, this is the default)
```

If `RESEND_API_KEY` is not set, reports are logged to Render's stdout instead of emailed.

**Outlook folder rule** — to auto-sort bug reports:
Settings → View all Outlook settings → Mail → Rules → New rule:
- Condition: *Subject contains* `[Bug Report] Container Optimizer`
- Action: *Move to* → create folder `Bugs`

---

## Running locally (CLI)

```bash
# Create and activate environment
conda create -n cpmpy-env python=3.10
conda activate cpmpy-env
pip install -r requirements.txt

# Run with default sample input
cd app
python main.py

# Run with a specific file
python main.py --excel path/to/order.xlsx

# Suppress plots (useful in scripts)
python main.py --no_plot
```

Output is written to `app/outputs/`: `report.xlsx`, `containers.json`, `recommendations.json`, `summary.txt`, `run.log`.

## Running the web server locally

```bash
uvicorn api:app --reload --port 8000
# Open http://localhost:8000
```

---

## Validation checks

`utils/validate.py` runs after every solve and returns structured issues:

| Code | Level | Description |
|---|---|---|
| `BLOCK_DROPPED` | ERROR | A row-block from the input never appeared in any container |
| `BLOCK_DUPLICATED` | ERROR | Same block appears in multiple containers |
| `PALLET_COUNT_MISMATCH` | ERROR | Total packed pallets ≠ input pallet count |
| `GEOMETRY_OVERLENGTH` | ERROR | Container used-length exceeds physical limit |
| `GEOMETRY_NEGATIVE_LEFTOVER` | ERROR | Leftover space is negative |
| `DOOR_HEIGHT_VIOLATION` | ERROR | Last (door) row taller than door opening |
| `WEIGHT_EXCEEDED` | ERROR | Container weight exceeds limit |
| `NP_BOX_OVERCOUNT` | ERROR | More boxes placed than ordered |
| `LENGTH_ACCOUNTING_MISMATCH` | WARNING | used + boxes + leftover ≠ container length |
| `TALL_ROWS_BACK_LOADED` | WARNING | Non-last rows exceed door height (expected — loaded from rear) |
| `ZERO_WEIGHT` | WARNING | All weights are 0 (weights missing from Excel) |
| `ROW_Y_INCONSISTENT` | WARNING | Row y-positions don't match expected gaps |
| `NP_BOX_UNPLACED` | WARNING | Some ordered NP boxes could not be placed |
| `LOW_FILL_NON_LAST` | WARNING | Non-final container less than 50% filled |

Errors and warnings are surfaced in the browser UI and included in bug reports.

---

## Key design decisions

- **No CP-SAT for box packing** — NP boxes use a fast column-based geometric algorithm (`BoxPacker`). All box types share one column of depth `D = max(bl)`, stacked in horizontal strips, which eliminates the per-type section waste of the previous approach.
- **Tail-only box placement** — NP boxes are never placed on top of pallets. This simplifies the loading sequence and avoids stability concerns.
- **Proportional fill recommendations** — tail space is allocated to pallet types in proportion to their order quantities, not winner-takes-all. Remaining space after pallets is filled with NP boxes using full 2-D (Y × Z) stacking.
- **Door-row constraint** — only the last row in a container needs to fit within the door height. All other rows are loaded from the rear before the door row and may be taller.
