# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this project does

Automates copying data from **Tekla Structures** `.xls` export files into the **Allied official macro** (also `.xls`). The output must remain `.xls` — never `.xlsx` — because the Allied macro and its structure depend on the original format.

## Key commands

```bash
# Start Docker services (first time needs --build)
docker compose up -d --build
docker compose up -d           # subsequent runs

# Start the host XLS server (must run on WSL host, not inside Docker)
python3 scripts/xls_host_server.py   # listens on port 5055

# Trigger a full run manually (no N8N needed)
curl -X POST http://localhost:5055/run
curl http://localhost:5055/health

# Run the full .xls flow directly (bypasses the HTTP server)
scripts/run_xls_host.sh data/tekla data/macro/Allied_Macro_original.xls data/output

# Run only the Python/payload step inside Docker
docker compose exec python-runner python3 /scripts/export_tekla_payload.py

# View logs
docker compose logs -f n8n
```

## Architecture

The pipeline has two distinct stages, intentionally separated because Excel COM (Windows-only) cannot run inside Docker:

**Stage 1 — Python (runs in Docker container `python-runner`):**
- `scripts/export_tekla_payload.py` reads all Tekla `.xls` files from `data/tekla/`, applies all business rules, and writes `data/output/tekla_payload.json`.
- It imports `scripts/tekla_to_allied.py` as a library (file loading, header parsing, piece extraction, text transforms). Do not delete `tekla_to_allied.py` — it is not an entrypoint but a dependency.

**Stage 2 — PowerShell + Excel COM (runs on the Windows/WSL host):**
- `scripts/write_allied_xls.ps1` receives the payload JSON, copies `data/macro/Allied_Macro_original.xls` to `data/output/`, and uses Excel COM to write each tab with the correct data and formatting.

**Orchestration:**
- `scripts/run_xls_host.sh` chains both stages: runs Stage 1 via `docker compose exec`, extracts the `job_number`, then calls Stage 2 via `powershell.exe`.
- `scripts/xls_host_server.py` wraps `run_xls_host.sh` in an HTTP server (port 5055) so N8N can trigger it.
- `scripts/server.py` keeps the `python-runner` container alive and serves a healthcheck on port 5000. Do not delete it — Docker depends on it.

## Data flow

```
data/tekla/*.xls  →  export_tekla_payload.py  →  data/output/tekla_payload.json
                                                              ↓
data/macro/Allied_Macro_original.xls  →  write_allied_xls.ps1  →  data/output/[JOB]_Secondary_Shipper.xls
```

After processing, Tekla source files are moved to `data/tekla/procesados/<timestamp>/` so they are not reprocessed.

## Tab mapping (Tekla → Allied)

| Tekla file | Allied tab |
|------------|-----------|
| `SBS_Eave_Struts_Shipper` | `Eave Struts` |
| `SBS_CEE_Secondary_Shipper` | `Cold Form Members (CEE)` + pieces to `Misc. Cold Form` |
| `SBS_ZEE_Secondary_Shipper` | `Cold Form Members (ZEE)` / `(ZEE) (2)` / `(ZEE) (3)` |
| `SBS_Miscellaneous_Shipper` | `Misc. Cold Form` |
| `SBS_Clips_Shipper` | `Clips` |
| `SBS_Pre_Galv_Clips_Shipper` | `Pre-Galv Clips` |
| `Standing_Seam_Hardware_Shipper` | `Standing Seam Hardware` |

`Screws` tab is always excluded (`ALWAYS_EXCLUDED_TABS`). `Pre-Galv Clips` is excluded for residential projects (`RESIDENTIAL_EXCLUDED_TABS`).

## Functional conditions of the shipper

These rules define the expected output for every generated shipper. They apply in `export_tekla_payload.py` and must be preserved whenever the transformation logic is modified.

### SBS standard (PROJECT_STANDARD = SBS)

Descriptions must match the Allied manual standard even when Tekla exports different text:

- `Z Girt` → `Wall Girt` (implemented in `normalize_sbs_piece`)
- `CCF BRC-7/9/11` / `CCF BRC 7/9/11` → `DESCRIPTION = Clip`, `DWG # = BRC-11`, `COLOR = Pre-Galvanized`

### Residential projects (PROJECT_TYPE = Residential)

The following must be **excluded entirely** from the shipper:

- `Pre-Galv Clips` tab (the whole tab)
- Any piece whose mark matches `\d+BRZ_EXT` or `28SA_EXT1`
- Any piece whose description contains: `Extra Material`, `Extra Clip`, `Extra Clips`, `Clip Extra`, `Clips Extra`

### Eave Struts

- Tekla appends `(?)` to every MARK in the Eave Struts export — strip it (`transform_eave_pieces`).
- Tekla encodes the roof pitch inside the PART column as a `_(N)` suffix (e.g. `825E14_(3)`). Strip the suffix from PART and write the pitch as `N:12` in the PITCH column (e.g. `825E14_(3)` → PART = `8.25E14`, PITCH = `3:12`).
- Tekla omits the decimal point in the eave strut size: `825` → `8.25` (insert decimal 2 positions from the right of the leading digit group). Pattern: `(\d{3,})(E\d+)` → `digits[:-2] + "." + digits[-2:] + rest`.
- COLOR must be set to `Pre-Galvanized` for all Eave Struts pieces (Tekla exports it blank).

### Cold Form Members (CEE)

- If COLOR is empty for any piece, it must be set to `Pre-Galvanized`.
- `140BC*` marks are **not** written to CEE — they are routed to `Misc. Cold Form` as Base Angle.
- `140DJ*` marks get `DESCRIPTION = Framed Opening Jamb / Sub Jamb`.

### Cold Form Members (ZEE)

- If COLOR is empty for any piece, it must be set to `Pre-Galvanized`.
- Residential extra material must not be included (same rule as above).
- Bridging pieces must be moved to `Standing Seam Hardware` (see below).

### Bridging reclassification

Pieces with "BRIDGING" in their description must **not** remain in Cold Form Members or Misc. Cold Form. They must be moved to `Standing Seam Hardware` with:

- `DESCRIPTION = Bridging Zee`
- `DWG # = BRZ-1`
- `COLOR = Pre-Galvanized`

Exception: marks matching `\d+BRZ_EXT` are extra material and must be **excluded entirely** (not moved, not included anywhere).

### Standing Seam Hardware

- `Rolls of Strapping` description is preserved. `COLOR` must be `White` (not Pre-Galvanized).
- Strapping qty and length are calculated from the total linear footage Tekla exports: `rolls = ceil(total_feet / 175)`, `wt = rolls × 22`, `length = 175'- 0"`. Tekla sometimes omits the QTY column for STRAP — the piece is still extracted because `leer_piezas` triggers on a non-empty MARK even when QTY is blank.
- Normal Bridging pieces arrive here via reclassification (see above).

### Misc. Cold Form

- Only valid descriptions in this tab: `Sheeting Angle`, `Base Angle`, `8" Girt Header (8 1/4CX6X4)`, `Eave Strut Spacer  ( 3 : 12)`.
- `Sa Ext Sec` or any Sheeting Angle variant must be normalized to `Sheeting Angle`.
- Bridging pieces must be moved to `Standing Seam Hardware`, not written here.
- Sort order in sheet: `140SA1`, `140SA2`, `T1` (SA_EXT), `140GH1`, `140SSL1`.

### Row limit per tab

Each tab can have a **maximum of 38 total rows** (the Allied template size). This translates to `MAX_SHEET_DATA_ROWS = 28` data rows in `export_tekla_payload.py` (template has 7-row header section + 4-row raw_rows overhead = 11 rows of fixed structure; 38 − 11 = 27 available, but the weight row accounts for a gap, giving capacity = 28 before PowerShell inserts extra rows). If a tab exceeds the limit, `enforce_sheet_row_limit()` splits it into sequentially numbered tabs (e.g., `Cold Form Members (ZEE)`, `(ZEE) (2)`, `(ZEE) (3)`).

## Critical business rules

**CEE → Misc split:** `split_cee_for_macro()` routes `140BC*` marks from CEE into `Misc. Cold Form` as Base Angle. All other CEE pieces stay in CEE. `140DJ*` marks get description `Framed Opening Jamb / Sub Jamb`.

**Misc. Cold Form transforms (`transform_misc_piece`):**
- `140BC*`: stock qty = `ceil(actual_length_inches / 240)`, `length = 20'- 0"`, `wt = qty × 55.9`. EXT variant uses 10-foot stock.
- `140SA*` (non-EXT): stock qty = `ceil(tekla_qty × actual_length_inches / 240)`, `length = 20'- 0"`, `wt = qty × 24.19`. This recalculates from the raw Tekla count × real length into 20-foot stock units.
- `140SA_EXT*`: mark becomes `T1`, qty and wt kept from Tekla.
- `140GH*`: desc `8" Girt Header (8 1/4CX6X4)`, part `17 7/8X14Ga`.
- `140SSL*`: length forced to `6"`.

**Tab ordering:** When a tab exceeds 28 rows and is split (e.g. `Cold Form Members (CEE) (2)`), the cloned tab is placed at the end of the workbook by Excel COM's `Copy` method. `Get-WorksheetOrClone` receives the previously-processed worksheet (`$prevWorksheet`) and moves the clone to come right after it (`$created.Move(After:=$AfterWorksheet)`), preserving payload order.

**IFlg / Web / Flg detail rows (Eave Struts and ZEE):** Tekla exports punch/dimension data as detail rows that span all columns immediately after a piece row. Two formats exist depending on the Tekla version/export setting:
- `IFlg=3 326.5` / `Web=22.1875 57.75` — ZEE and some Eave Struts exports
- `Web : 1.75 16.125` / `Flg : 71.9375 291.8125` — alternate Eave Struts export format

`_DETALLE_PREFIXES` in `tekla_to_allied.py` lists all recognized prefixes. `primer_detalle()` and `es_detalle()` use this tuple. `agregar_pulgadas()` formats numeric values by appending `"` (values that already have `"` are preserved). Rows are stored in `piece["detalles"]` and written by `Write-Pieces` in PowerShell. If a Tekla file does not export those rows for a given job, no detail rows will appear in the output — this is correct behavior, not a bug.

**ZEE overrides:** `ZEE_DETAIL_OVERRIDES` replaces Web detail lines for specific marks. `ZEE_FINAL_DETAIL_OVERRIDES` applies after the main override. `140P_EXT[1-3]` marks become `T1/T2/T3`. The PART column for all ZEE pieces must come directly from Tekla — never hardcode a numeric value for it.

**_EXT exclusion (global):** Any piece whose MARK contains `_EXT` (case-insensitive) must be excluded from all generated tabs. The filter `filter_ext_material()` is applied in `construir_payload` immediately after `leer_piezas`, before any tab-specific transform runs. This order is critical: transforms like `transform_zee_pieces` rename `_EXT` marks to `T1/T2/T3`, so filtering after would miss them.

**Bridging reclassification:** Any piece with "BRIDGING" in its description is moved from its source tab into `Standing Seam Hardware` as `Bridging Zee`. Marks matching `\d+BRZ_EXT` are excluded entirely, not moved.

## Tekla file format

Tekla exports `.xls` files as HTML tables disguised with `.xls` extension. `cargar_archivo_excel()` in `tekla_to_allied.py` detects the real format by reading the first bytes: HTML (Tekla output), ZIP/PK (real `.xlsx`), or OLE `\xd0\xcf` (real `.xls` binary). The function always returns an `openpyxl` workbook regardless of input format.

## Environment variables

| Variable | Default (container path) | Used by |
|----------|--------------------------|---------|
| `TEKLA_FOLDER` | `/data/tekla` | `export_tekla_payload.py` |
| `OUTPUT_FOLDER` | `/data/output` | `export_tekla_payload.py` |
| `PROJECT_STANDARD` | `SBS` | `export_tekla_payload.py` |
| `PROJECT_TYPE` | `Residential` | `export_tekla_payload.py` (affects Pre-Galv Clips exclusion) |
| `WORKBOOK_STRUCTURE_PASSWORD` | — | `write_allied_xls.ps1` (needed if workbook structure is protected) |
| `XLS_HOST_PORT` | `5055` | `xls_host_server.py` |

## Payload JSON structure

`data/output/tekla_payload.json` is the contract between Stage 1 and Stage 2. Each entry in `sheets[]` contains: `tab_macro`, `template_tab_macro` (for cloned tabs), `shipper_number`, `total_shippers`, `encabezado` (job/building header fields), `piezas` (list of piece objects with qty/mark/desc/part/punch/dwg/color/length/wt/detalles), `peso` (page weight), and optionally `raw_rows` (pre-built row array used by some tabs to preserve source cell spans).
