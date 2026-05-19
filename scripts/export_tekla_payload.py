#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Exporta los archivos Tekla a un JSON intermedio.

Este script no escribe la macro. Se usa cuando el archivo final debe seguir
siendo .xls binario con VBA: Python lee Tekla y Excel/PowerShell escribe el
.xls oficial.
"""

import argparse
import json
import math
import os
import re
import sys
from datetime import datetime

import tekla_to_allied as tekla

MAX_SHEET_DATA_ROWS = 28

ZEE_DETAIL_OVERRIDES = {
    "140G10": [],
    "140G21": ["Web=9.5 13.5 18.1875 61.3125"],
    "140G28": [],
    "140P2": ["Web=45.75 295.75 299.75 303.75 307.75 315.75 319.75 323.75 333.75 345.75 353.75"],
    "140P3": ["Web=13.75 21.75 33.75 43.75 47.75 51.75 59.75 63.75 67.75 71.75 321.75 333.75 337.75"],
    "140P4": ["Web=29.75 33.75 323.50 327.50 353.50 357.50"],
    "140P5": ["Web=10.00 14.00 40.00 44.00"],
}

ZEE_FINAL_DETAIL_OVERRIDES = {
    "140G3": ['Web=      22.1875"       57.75 66"'],
}

BASE_CHANNEL_QTY_OVERRIDES = {
    "140BC5": 1,
}

PROJECT_STANDARD = os.environ.get("PROJECT_STANDARD", "SBS").strip().upper()
PROJECT_TYPE = os.environ.get("PROJECT_TYPE", "Residential").strip().lower()

STANDING_SEAM_TAB = "Standing Seam Hardware"
ALWAYS_EXCLUDED_TABS = {"Screws"}
RESIDENTIAL_EXCLUDED_TABS = {"Pre-Galv Clips"}

STRAPPING_ROLL_FEET = 175
STRAPPING_ROLL_LB = 22


def log(msg):
    print(msg, file=sys.stderr, flush=True)


def log_rule(payload, msg):
    payload["log_entries"].append(msg)
    log(msg)


def normalized_text(value):
    return " ".join(str(value or "").strip().upper().split())


def is_sbs_standard():
    return PROJECT_STANDARD == "SBS"


def is_residential_project():
    return PROJECT_TYPE == "residential"


def excluded_sheet_reason(tab_macro):
    if tab_macro in ALWAYS_EXCLUDED_TABS:
        return "pestana excluida por regla operativa"
    if is_residential_project() and tab_macro in RESIDENTIAL_EXCLUDED_TABS:
        return "material extra residencial"
    return None


def normalize_sbs_piece(piece):
    cloned = copy_piece(piece)
    if not is_sbs_standard():
        return cloned

    desc = normalized_text(cloned.get("desc"))
    if desc == "Z GIRT":
        cloned["desc"] = "Wall Girt"
    elif desc in ("CCF BRC-7/9/11", "CCF BRC 7/9/11"):
        cloned["desc"] = "Clip"
        cloned["dwg"] = "BRC-11"
        cloned["color"] = "Pre-Galvanized"
    return cloned


def normalize_sbs_pieces(pieces):
    return [normalize_sbs_piece(piece) for piece in pieces]


def is_residential_extra_material(piece):
    desc = normalized_text(piece.get("desc"))
    mark = piece_mark(piece)
    if re.match(r"^\d+BRZ_EXT", mark):
        return True
    if mark.replace(" ", "_") == "28SA_EXT1":
        return True

    extra_terms = (
        "EXTRA MATERIAL",
        "EXTRA CLIP",
        "EXTRA CLIPS",
        "CLIP EXTRA",
        "CLIPS EXTRA",
    )
    return any(term in desc for term in extra_terms)


def filter_residential_extra_material(pieces):
    if not is_residential_project():
        return pieces, []

    kept = []
    excluded = []
    for piece in pieces:
        if is_residential_extra_material(piece):
            excluded.append(piece)
        else:
            kept.append(piece)
    return kept, excluded


def is_bridging_piece(piece):
    return "BRIDGING" in normalized_text(piece.get("desc"))


def normalize_standing_piece(piece):
    cloned = copy_piece(piece)
    if is_bridging_piece(cloned):
        cloned["desc"] = "Bridging Zee"
        cloned["dwg"] = "BRZ-1"
        if not str(cloned.get("color") or "").strip():
            cloned["color"] = "Pre-Galvanized"
    elif (
        piece_mark(cloned) == "STRP1"
        or normalized_text(cloned.get("desc")) == "ROLLS OF STRAPPING"
    ):
        cloned["color"] = "White"
    return cloned


def split_standing_reclassified_pieces(tab_macro, pieces):
    if tab_macro == STANDING_SEAM_TAB:
        return pieces, []

    kept = []
    reclassified = []
    for piece in pieces:
        if is_bridging_piece(piece):
            reclassified.append(normalize_standing_piece(piece))
        else:
            kept.append(piece)
    return kept, reclassified


def apply_output_rules(payload, nombre, tab_macro, enc, pieces, peso, standing_reclassified):
    pieces = normalize_sbs_pieces(pieces)

    pieces, excluded = filter_residential_extra_material(pieces)
    if excluded:
        log_rule(
            payload,
            f"SKIP PIECES [{nombre}] -> {len(excluded)} material extra residencial",
        )

    pieces, reclassified = split_standing_reclassified_pieces(tab_macro, pieces)
    if reclassified:
        standing_reclassified["pieces"].extend(reclassified)
        standing_reclassified["sources"].append(nombre)
        if standing_reclassified["enc"] is None:
            standing_reclassified["enc"] = enc
        log_rule(
            payload,
            f"MOVE [{nombre}] -> [{STANDING_SEAM_TAB}] | {len(reclassified)} Bridging",
        )

    return pieces, sum_piece_weight(pieces, peso)


def normalize_header(value):
    return " ".join(str(value or "").strip().upper().split()).rstrip(":")


def piece_mark(piece):
    return str(piece.get("mark") or "").strip().upper()


def is_ext_mark(piece):
    return bool(re.search(r"_EXT", str(piece.get("mark") or ""), re.IGNORECASE))


def filter_ext_material(pieces):
    return [p for p in pieces if not is_ext_mark(p)]


def copy_piece(piece):
    cloned = dict(piece)
    cloned["detalles"] = list(piece.get("detalles") or [])
    cloned["detalle"] = "\n".join(cloned["detalles"]) if cloned["detalles"] else None
    return cloned


def as_float(value, default=0.0):
    try:
        return float(str(value).replace(",", "").strip())
    except (TypeError, ValueError):
        return default


def as_int(value, default=0):
    try:
        return int(math.ceil(as_float(value, default)))
    except (TypeError, ValueError):
        return default


def parse_length_inches(value):
    text = str(value or "").strip()
    match = re.search(r"(\d+)'\s*-\s*(\d+(?:\.\d+)?)(?:\s+(\d+)\s*/\s*(\d+))?\"?", text)
    if not match:
        return 0.0
    feet = float(match.group(1))
    inches = float(match.group(2))
    if match.group(3) and match.group(4):
        inches += float(match.group(3)) / float(match.group(4))
    return feet * 12 + inches


def stock_length_text(feet):
    return f"{feet}'- 0\""


def html_rows_with_spans(path):
    text = open(path, encoding="latin-1", errors="ignore").read()
    rows = []
    for row_index, tr in enumerate(re.findall(r"<TR>(.*?)</TR>", text, re.I | re.S), start=1):
        cells = []
        for attrs, body in re.findall(r"<TD([^>]*)>(.*?)</TD>", tr, re.I | re.S):
            match = re.search(r"colspan\s*=\s*\"?\s*(\d+)", attrs, re.I)
            span = int(match.group(1)) if match else 1
            value = re.sub(r"<[^>]+>", " ", body)
            value = " ".join(value.replace("&nbsp;", " ").split())
            cells.append({"span": span, "value": value})
        rows.append({"index": row_index, "cells": cells})
    return rows


def source_header_cells(path):
    for row in html_rows_with_spans(path):
        values = [normalize_header(c["value"]) for c in row["cells"]]
        if "QTY" in values and "MARK" in values:
            return row["cells"]
    return []


def value_for_header(piece, header, tab_macro):
    header = normalize_header(header)
    if header == "QTY":
        return piece.get("qty")
    if header == "MARK":
        return piece.get("mark")
    if header == "DESCRIPTION":
        return piece.get("desc")
    if header == "PITCH":
        return piece.get("pitch")
    if header == "PART":
        return piece.get("part")
    if header == "PUNCH":
        return piece.get("punch")
    if header in ("DWG #", "DRAWING #", "DRAWING NO", "DRAWING NO."):
        return piece.get("dwg")
    if header == "COLOR":
        return piece.get("color")
    if header == "LENGTH":
        return piece.get("length")
    if header in ("WT.", "WT", "WEIGHT"):
        return piece.get("wt")
    return ""


def raw_rows_from_pieces(header_cells, pieces, peso, tab_macro):
    rows = [{"kind": "header", "cells": header_cells}]

    blank_cells = [{"span": c.get("span", 1), "value": ""} for c in header_cells]
    rows.append({"kind": "blank", "cells": blank_cells})

    for piece in pieces:
        row_cells = []
        for cell in header_cells:
            row_cells.append({
                "span": cell.get("span", 1),
                "value": value_for_header(piece, cell.get("value"), tab_macro),
            })
        rows.append({"kind": "piece", "cells": row_cells})

        for detalle in piece.get("detalles") or []:
            rows.append({
                "kind": "detail",
                "cells": [{"span": 14, "value": detalle}],
            })

    rows.append({"kind": "blank", "cells": [{"span": 14, "value": ""}]})
    rows.append({
        "kind": "weight",
        "cells": [
            {"span": 10, "value": ""},
            {"span": 2, "value": "PAGE WEIGHT:"},
            {"span": 2, "value": peso},
        ],
    })
    return rows


def piece_row_count(piece):
    return 1 + len(piece.get("detalles") or [])


def sum_piece_weight(pieces, fallback=None):
    total = 0.0
    found = False
    for piece in pieces:
        if piece.get("exclude_from_page_weight"):
            continue
        value = piece.get("wt")
        if value is None or str(value).strip() == "":
            continue
        try:
            total += float(str(value).replace(",", ""))
            found = True
        except (TypeError, ValueError):
            return fallback
    if not found:
        return fallback
    return round(total, 2)


def split_cee_for_macro(pieces):
    cee_pieces = []
    misc_base_pieces = []

    for piece in pieces:
        mark = piece_mark(piece)
        if mark.startswith("140BC"):
            misc_base_pieces.append(base_channel_to_misc_piece(piece))
            continue

        cloned = copy_piece(piece)
        if mark.startswith("140DJ"):
            cloned["desc"] = "Framed Opening Jamb / Sub Jamb"
        if not str(cloned.get("color") or "").strip():
            cloned["color"] = "Pre-Galvanized"
        cee_pieces.append(cloned)

    return cee_pieces, misc_base_pieces


def base_channel_to_misc_piece(piece):
    cloned = copy_piece(piece)
    mark = piece_mark(cloned)
    length_inches = parse_length_inches(cloned.get("length"))
    stock_feet = 10 if "_EXT" in mark and length_inches <= 120 else 20
    stock_inches = stock_feet * 12
    qty = BASE_CHANNEL_QTY_OVERRIDES.get(mark)
    if qty is None:
        qty = max(1, int(math.ceil(length_inches / stock_inches))) if length_inches else as_int(cloned.get("qty"), 1)
    weight_each = 55.9 if stock_feet == 20 else 27.95

    cloned.update({
        "qty": qty,
        "mark": "T1" if "_EXT" in mark else cloned.get("mark"),
        "desc": "Base Angle",
        "part": "8X25C16",
        "punch": "",
        "dwg": "BA1",
        "color": "Pre-Galvanized",
        "length": stock_length_text(stock_feet),
        "wt": round(qty * weight_each, 2),
        "exclude_from_page_weight": True,
    })
    return cloned


def transform_misc_piece(piece):
    cloned = copy_piece(piece)
    mark = piece_mark(cloned)
    qty = as_int(cloned.get("qty"), 0)
    cloned["qty"] = qty

    cloned["color"] = "Pre-Galvanized"
    if mark.startswith("140GH"):
        cloned.update({
            "desc": '8" Girt Header (8 1/4CX6X4)',
            "part": "17 7/8X14Ga",
            "dwg": "GH-1",
        })
    elif mark.startswith("140SA"):
        cloned.update({
            "mark": "T1" if "_EXT" in mark else cloned.get("mark"),
            "desc": "Sheeting Angle",
            "part": "4X2X16Ga",
            "dwg": "SA1",
        })
        if "_EXT" not in mark:
            length_in = parse_length_inches(cloned.get("length"))
            stock_qty = max(1, math.ceil(qty * length_in / (20 * 12))) if length_in > 0 else qty
            cloned["qty"] = stock_qty
            cloned["length"] = stock_length_text(20)
            cloned["wt"] = round(stock_qty * 24.19, 2)
    elif re.match(r"^\d+SA", mark):
        cloned["desc"] = "Sheeting Angle"
    elif mark.startswith("140SSL"):
        cloned.update({
            "desc": "Eave Strut Spacer  ( 3 : 12)",
            "part": "10X35C14",
            "dwg": "SSL-10",
            "length": '6"',
        })

    return cloned


def misc_sort_key(piece):
    mark = piece_mark(piece)
    desc = str(piece.get("desc") or "").strip().upper()
    order = {
        "140SA1": 1,
        "140SA2": 2,
        "140SA_EXT1": 3,
        "T1": 3 if "SHEETING ANGLE" in desc else 99,
        "140GH1": 4,
        "140SSL1": 5,
    }
    return (order.get(mark, 99), mark)


def transform_pre_galv_pieces(pieces):
    transformed = [copy_piece(piece) for piece in pieces]
    for piece in transformed:
        piece["color"] = "Pre-Galvanized"
    return transformed


def transform_eave_pieces(pieces):
    transformed = []
    for piece in pieces:
        cloned = copy_piece(piece)
        # Strip trailing (?) artifact from Tekla HTML export marks
        mark = str(cloned.get("mark") or "").strip()
        mark = re.sub(r"\(\?\)\s*$", "", mark).strip()
        cloned["mark"] = mark
        # PART may encode pitch as suffix _(N) meaning N:12
        # e.g. "825E14_(3)" → part="825E14", pitch="3:12"
        part = str(cloned.get("part") or "").strip()
        m = re.search(r"_\((\d+)\)$", part)
        if m:
            part = part[: m.start()]
            cloned["pitch"] = f"{m.group(1)}:12"
        # Tekla omits the decimal in the size: "825E14" → "8.25E14"
        pm = re.match(r"^(\d{3,})(E\d+.*)$", part)
        if pm:
            digits = pm.group(1)
            part = digits[:-2] + "." + digits[-2:] + pm.group(2)
        cloned["part"] = part
        cloned["color"] = "Pre-Galvanized"
        transformed.append(cloned)
    return transformed


def transform_clip_pieces(pieces):
    transformed = []
    for piece in pieces:
        cloned = copy_piece(piece)
        cloned["color"] = "Pre-Galvanized"
        if str(cloned.get("desc") or "").strip().upper() == "SHEETING CLIP" or piece_mark(cloned) == "140CCF7":
            cloned["dwg"] = "CL-5"
        transformed.append(cloned)
    return transformed


def transform_standing_pieces(pieces):
    transformed = []
    for piece in pieces:
        cloned = copy_piece(piece)
        is_strap = (
            piece_mark(cloned) == "STRAP"
            or str(cloned.get("desc") or "").strip().lower() in ("rolls of strapping", "strapping")
        )
        if is_strap:
            total_in = parse_length_inches(cloned.get("length"))
            qty_tekla = as_float(cloned.get("qty"), 0.0)
            if total_in > 0:
                rolls = max(1, math.ceil(total_in / (STRAPPING_ROLL_FEET * 12)))
            else:
                rolls = max(1, int(math.ceil(qty_tekla))) if qty_tekla > 0 else 1
            cloned.update({
                "mark": "STRP1",
                "desc": "Rolls of Strapping",
                "dwg": "STRP1",
                "color": "White",
                "qty": rolls,
                "length": stock_length_text(STRAPPING_ROLL_FEET),
                "wt": round(rolls * STRAPPING_ROLL_LB, 2),
            })
        transformed.append(normalize_standing_piece(cloned))
    return transformed


def transform_zee_pieces(pieces):
    transformed = []
    for piece in pieces:
        cloned = copy_piece(piece)
        mark = piece_mark(cloned)
        if mark in ZEE_FINAL_DETAIL_OVERRIDES:
            cloned["detalles"] = list(ZEE_FINAL_DETAIL_OVERRIDES[mark])
            cloned["detalle"] = "\n".join(cloned["detalles"]) if cloned["detalles"] else None
        elif mark in ZEE_DETAIL_OVERRIDES:
            cloned["detalles"] = [tekla.agregar_pulgadas(value) for value in ZEE_DETAIL_OVERRIDES[mark]]
            cloned["detalle"] = "\n".join(cloned["detalles"]) if cloned["detalles"] else None
        match = re.fullmatch(r"140P_EXT([1-3])", mark)
        if match:
            cloned["mark"] = f"T{match.group(1)}"
        transformed.append(cloned)
    return transformed


def sheet_name_for_part(base_name, part_number):
    if part_number <= 1:
        return base_name
    suffix = f" ({part_number})"
    return base_name[:31 - len(suffix)] + suffix


def split_pieces_by_row_limit(pieces, max_rows=MAX_SHEET_DATA_ROWS):
    chunks = []
    chunk = []
    used_rows = 0

    for piece in pieces:
        needed_rows = piece_row_count(piece)
        if chunk and used_rows + needed_rows > max_rows:
            chunks.append(chunk)
            chunk = []
            used_rows = 0

        chunk.append(piece)
        used_rows += needed_rows

    if chunk:
        chunks.append(chunk)
    return chunks


def enforce_sheet_row_limit(payload):
    expanded = []
    for sheet in payload["sheets"]:
        if sheet.get("raw_rows"):
            expanded.append(sheet)
            continue

        pieces = sheet.get("piezas") or []
        if not pieces:
            expanded.append(sheet)
            continue

        chunks = split_pieces_by_row_limit(pieces)
        if len(chunks) == 1:
            sheet["template_tab_macro"] = sheet.get("template_tab_macro") or sheet["tab_macro"]
            expanded.append(sheet)
            continue

        base_tab = sheet.get("template_tab_macro") or sheet["tab_macro"]
        log_rule(
            payload,
            f"SPLIT [{sheet['tab_macro']}] -> {len(chunks)} pestanas por limite de {MAX_SHEET_DATA_ROWS} filas",
        )

        for part_number, chunk in enumerate(chunks, 1):
            cloned = dict(sheet)
            cloned["tab_macro"] = sheet_name_for_part(base_tab, part_number)
            cloned["template_tab_macro"] = base_tab
            cloned["piezas"] = chunk
            cloned["peso"] = sum_piece_weight(chunk, sheet.get("peso"))
            expanded.append(cloned)

    payload["sheets"] = expanded


def append_sheet(payload, nombre, tab_tekla, tab_macro, shipper_number, total_shippers, enc, piezas, peso, raw_rows=None):
    payload["sheets"].append({
        "source_file": nombre,
        "tab_tekla": tab_tekla,
        "tab_macro": tab_macro,
        "shipper_number": shipper_number,
        "total_shippers": total_shippers,
        "encabezado": enc,
        "piezas": piezas,
        "peso": peso,
        "raw_rows": raw_rows or [],
    })

    msg = f"OK [{nombre}] -> [{tab_macro}] | {len(piezas)} piezas | Peso: {peso}"
    payload["log_entries"].append(msg)
    log(msg)


def construir_payload(tekla_folder):
    if not os.path.isdir(tekla_folder):
        return {
            "status": "error",
            "message": f"Carpeta de Tekla no existe: {tekla_folder}",
            "sheets": [],
        }

    def es_archivo_tekla(nombre):
        lower = nombre.lower()
        return lower.endswith(".xls") and "zone.identifier" not in lower

    archivos = sorted(
        f for f in os.listdir(tekla_folder)
        if es_archivo_tekla(f)
        and os.path.isfile(os.path.join(tekla_folder, f))
    )

    if not archivos:
        return {
            "status": "no_files",
            "message": "No hay archivos para procesar",
            "sheets": [],
        }

    payload = {
        "status": "success",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "job_number": "SIN_NUMERO",
        "files_discovered": len(archivos),
        "files_processed": 0,
        "project_standard": PROJECT_STANDARD,
        "project_type": PROJECT_TYPE,
        "sheets": [],
        "log_entries": [],
        "files_with_errors": [],
    }
    misc_from_cee = []
    standing_reclassified = {
        "pieces": [],
        "sources": [],
        "enc": None,
    }

    for nombre in archivos:
        ruta_archivo = os.path.join(tekla_folder, nombre)
        try:
            wb = tekla.cargar_archivo_excel(ruta_archivo)
            ws = wb.active
            tab_tekla = ws.title
            tab_macro = tekla.TAB_MAP.get(tab_tekla)

            if not tab_macro:
                msg = f"SKIP [{nombre}] -> no hay pestana macro mapeada"
                payload["log_entries"].append(msg)
                log(msg)
                continue

            enc = tekla.leer_encabezado(ws)
            if enc.get("job"):
                payload["job_number"] = str(enc["job"]).split("_")[0].strip()

            skip_reason = excluded_sheet_reason(tab_macro)
            if skip_reason:
                log_rule(payload, f"SKIP [{nombre}] -> {skip_reason}")
                continue

            piezas, peso = tekla.leer_piezas(ws, tab_macro)
            piezas = filter_ext_material(piezas)
            if tab_tekla == "SBS_CEE_Secondary_Shipper":
                piezas, misc_from_cee = split_cee_for_macro(piezas)
                piezas, peso = apply_output_rules(
                    payload, nombre, tab_macro, enc, piezas, peso, standing_reclassified
                )
                peso = sum_piece_weight(piezas, peso)
                append_sheet(
                    payload, nombre, tab_tekla, tab_macro, 0, 0,
                    enc, piezas, peso
                )
            elif tab_tekla == "SBS_Clips_Shipper":
                piezas = transform_clip_pieces(piezas)
                piezas, peso = apply_output_rules(
                    payload, nombre, tab_macro, enc, piezas, peso, standing_reclassified
                )
                peso = sum_piece_weight(piezas, peso)
                append_sheet(
                    payload, nombre, tab_tekla, tab_macro, 0, 0,
                    enc, piezas, peso
                )
            elif tab_tekla == "SBS_Miscellaneous_Shipper":
                misc_own_pieces = sorted(
                    (transform_misc_piece(piece) for piece in piezas),
                    key=misc_sort_key,
                )
                piezas = misc_from_cee + misc_own_pieces
                piezas, peso = apply_output_rules(
                    payload, nombre, tab_macro, enc, piezas, peso, standing_reclassified
                )
                peso = sum_piece_weight(piezas, peso)
                append_sheet(
                    payload, nombre, tab_tekla, tab_macro, 0, 0,
                    enc, piezas, peso
                )
            elif tab_tekla == "SBS_Pre_Galv_Clips_Shipper":
                piezas = transform_pre_galv_pieces(piezas)
                piezas, peso = apply_output_rules(
                    payload, nombre, tab_macro, enc, piezas, peso, standing_reclassified
                )
                peso = sum_piece_weight(piezas, peso)
                append_sheet(
                    payload, nombre, tab_tekla, tab_macro, 0, 0,
                    enc, piezas, peso
                )
            elif tab_tekla == "Standing_Seam_Hardware_Shipper":
                piezas = transform_standing_pieces(piezas)
                piezas, peso = apply_output_rules(
                    payload, nombre, tab_macro, enc, piezas, peso, standing_reclassified
                )
                peso = sum_piece_weight(piezas, peso)
                append_sheet(
                    payload, nombre, tab_tekla, tab_macro, 0, 0,
                    enc, piezas, peso
                )
            elif tab_tekla == "SBS_ZEE_Secondary_Shipper":
                piezas = transform_zee_pieces(piezas)
                piezas, peso = apply_output_rules(
                    payload, nombre, tab_macro, enc, piezas, peso, standing_reclassified
                )
                for piece in piezas:
                    if not str(piece.get("color") or "").strip():
                        piece["color"] = "Pre-Galvanized"
                append_sheet(
                    payload, nombre, tab_tekla, tab_macro, 0, 0,
                    enc, piezas, peso
                )
            elif tab_tekla == "SBS_Eave_Struts_Shipper":
                piezas = transform_eave_pieces(piezas)
                piezas, peso = apply_output_rules(
                    payload, nombre, tab_macro, enc, piezas, peso, standing_reclassified
                )
                peso = sum_piece_weight(piezas, peso)
                append_sheet(
                    payload, nombre, tab_tekla, tab_macro, 0, 0,
                    enc, piezas, peso
                )
            else:
                piezas, peso = apply_output_rules(
                    payload, nombre, tab_macro, enc, piezas, peso, standing_reclassified
                )
                append_sheet(
                    payload, nombre, tab_tekla, tab_macro, 0, 0,
                    enc, piezas, peso
                )

        except Exception as exc:
            msg = f"ERROR [{nombre}]: {exc}"
            payload["status"] = "partial_success"
            payload["log_entries"].append(msg)
            payload["files_with_errors"].append({"file": nombre, "error": str(exc)})
            log(msg)

    if not payload["sheets"] and payload["files_with_errors"]:
        payload["status"] = "error"

    if standing_reclassified["pieces"]:
        standing_sheet = next(
            (sheet for sheet in payload["sheets"] if sheet["tab_macro"] == STANDING_SEAM_TAB),
            None,
        )
        if standing_sheet:
            standing_sheet["piezas"] = standing_reclassified["pieces"] + standing_sheet["piezas"]
            standing_sheet["peso"] = sum_piece_weight(
                standing_sheet["piezas"],
                standing_sheet.get("peso"),
            )
            log_rule(
                payload,
                f"OK [Bridging reclasificado] -> [{STANDING_SEAM_TAB}] | "
                f"{len(standing_reclassified['pieces'])} piezas",
            )
        else:
            append_sheet(
                payload,
                "Bridging_reclasificado",
                "Business_Rules",
                STANDING_SEAM_TAB,
                0,
                0,
                standing_reclassified["enc"] or {},
                standing_reclassified["pieces"],
                sum_piece_weight(standing_reclassified["pieces"], None),
            )

    if not payload["sheets"] and not payload["files_with_errors"]:
        payload["status"] = "no_files"
        payload["message"] = "No hay hojas para generar despues de aplicar reglas"

    enforce_sheet_row_limit(payload)

    total_shippers = len(payload["sheets"])
    payload["files_processed"] = total_shippers
    for index, sheet in enumerate(payload["sheets"], 1):
        sheet["shipper_number"] = index
        sheet["total_shippers"] = total_shippers

    return payload


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--tekla-folder", default=os.environ.get("TEKLA_FOLDER", tekla.CARPETA_TEKLA))
    parser.add_argument("--output", default=os.path.join(
        os.environ.get("OUTPUT_FOLDER", tekla.CARPETA_SALIDA),
        "tekla_payload.json",
    ))
    args = parser.parse_args()

    payload = construir_payload(args.tekla_folder)
    os.makedirs(os.path.dirname(args.output), exist_ok=True)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(payload, f, default=str, ensure_ascii=False, indent=2)

    result = {
        "status": payload["status"],
        "job_number": payload.get("job_number"),
        "files_processed": payload.get("files_processed", 0),
        "files_with_errors": payload.get("files_with_errors", []),
        "payload_file": args.output,
        "log_entries": payload.get("log_entries", []),
    }
    print(json.dumps(result, ensure_ascii=False))
    return 0 if payload["status"] in ("success", "partial_success") else 1


if __name__ == "__main__":
    sys.exit(main())
