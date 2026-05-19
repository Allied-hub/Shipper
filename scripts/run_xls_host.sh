#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$ROOT_DIR"

SOURCE_FOLDER="${1:-data/tekla}"
TEMPLATE_XLS="${2:-data/macro/Allied_Macro_original.xls}"
OUTPUT_DIR="${3:-/mnt/c/Users/EGuerrero/Downloads}"

if [[ ! -s "$TEMPLATE_XLS" ]]; then
  echo "La macro plantilla no existe o esta vacia: $TEMPLATE_XLS" >&2
  echo "Restaura data/macro/Allied_Macro_original.xls antes de generar el shipper." >&2
  exit 1
fi

TEKLA_ROOT="$(realpath data/tekla)"
SOURCE_ABS="$(realpath "$SOURCE_FOLDER")"

case "$SOURCE_ABS" in
  "$TEKLA_ROOT"*)
    CONTAINER_SOURCE="/data/tekla${SOURCE_ABS#$TEKLA_ROOT}"
    ;;
  *)
    echo "La carpeta fuente debe estar dentro de data/tekla para que Docker la vea: $SOURCE_FOLDER" >&2
    exit 1
    ;;
esac

PAYLOAD_WSL="$ROOT_DIR/data/output/tekla_payload.json"
PAYLOAD_CONTAINER="/data/output/tekla_payload.json"

docker compose exec -T \
  -e TEKLA_FOLDER="$CONTAINER_SOURCE" \
  -e PROJECT_TYPE="${PROJECT_TYPE:-Residential}" \
  -e PROJECT_STANDARD="${PROJECT_STANDARD:-SBS}" \
  python-runner \
  python3 /scripts/export_tekla_payload.py \
    --tekla-folder "$CONTAINER_SOURCE" \
    --output "$PAYLOAD_CONTAINER"

JOB_NUMBER="$(python3 - "$PAYLOAD_WSL" <<'PY'
import json
import sys
with open(sys.argv[1], encoding="utf-8") as f:
    payload = json.load(f)
print(payload.get("job_number") or "SIN_NUMERO")
PY
)"

mkdir -p "$OUTPUT_DIR"
OUTPUT_XLS="$OUTPUT_DIR/${JOB_NUMBER}_Secondary_Shipper.xls"

POWERSHELL_ARGS=(
  -NoProfile
  -ExecutionPolicy Bypass
  -File "$(wslpath -w "$ROOT_DIR/scripts/write_allied_xls.ps1")"
  -PayloadPath "$(wslpath -w "$PAYLOAD_WSL")"
  -TemplatePath "$(wslpath -w "$TEMPLATE_XLS")"
  -OutputPath "$(wslpath -w "$OUTPUT_XLS")"
)

if [[ -n "${WORKBOOK_STRUCTURE_PASSWORD:-}" ]]; then
  POWERSHELL_ARGS+=(-WorkbookStructurePassword "$WORKBOOK_STRUCTURE_PASSWORD")
fi

powershell.exe "${POWERSHELL_ARGS[@]}"

echo "Archivo .xls generado:"
echo "$OUTPUT_XLS"
