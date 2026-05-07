#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tekla -> Macro Allied
=====================
Procesa archivos Excel exportados desde Tekla Structures
y los integra en la macro oficial de Allied.

Pensado para correr desde el nodo "Execute Command" de N8N:
    python tekla_to_allied.py

- Logs informativos -> stderr (visibles en N8N como log de ejecucion)
- Resultado final (JSON) -> stdout (lo parsea el siguiente nodo)
- Codigo de salida 0 si todo bien, 1 si hubo error fatal

Configuracion: editar las constantes al inicio o usar variables de entorno
    TEKLA_FOLDER, MACRO_PATH, OUTPUT_FOLDER
"""

import openpyxl
import os
import sys
import json
import shutil
import traceback
from datetime import datetime

# ── CONFIGURACION ─────────────────────────────────────────────
# Las rutas se toman de variables de entorno definidas en docker-compose.yml.
# Los valores por defecto son las rutas DENTRO del contenedor.
CARPETA_TEKLA      = os.environ.get("TEKLA_FOLDER",  "/data/tekla")
RUTA_MACRO         = os.environ.get("MACRO_PATH",    "/data/macro/Allied_Macro.xlsx")
CARPETA_SALIDA     = os.environ.get("OUTPUT_FOLDER", "/data/output")
CARPETA_PROCESADOS = os.path.join(CARPETA_TEKLA, "procesados")

# ── MAPEOS ────────────────────────────────────────────────────
TAB_MAP = {
    "SBS_Eave_Struts_Shipper":        "Eave Struts",
    "SBS_CEE_Secondary_Shipper":      "Cold Form Members (CEE)",
    "SBS_ZEE_Secondary_Shipper":      "Cold Form Members (ZEE)",
    "SBS_Miscellaneous_Shipper":      "Misc. Cold Form",
    "SBS_Clips_Shipper":              "Clips",
    "SBS_Pre_Galv_Clips_Shipper":     "Pre-Galv Clips",
    "Standing_Seam_Hardware_Shipper": "Standing Seam Hardware",
    "SBS_Screws_Shipper":             "Screws",
}
TAB_ZEE_EXTRA = ["Cold Form Members (ZEE) (2)", "Cold Form Members (ZEE) (3)"]

DESC_MAP = {
    "EAVE_STRUT":                      "Eave Strut (LSSS)",
    "SHEETING_BASE_CHANNEL":           "Base Angle",
    "SHEETING_ANGLE":                  "Sheeting Angle",
    "SHEETING_BASE_ANGLE":             "Base Angle",
    "C_WRAP_CHANNEL":                  "Girt Header (8 1/4CX6X4)",
    "C_STRUT_SPACER_LOW":              "Eave Strut Spacer  ( 3 : 12)",
    "STRAPPING":                       "Rolls of Strapping",
    "CCF_CLIP":                        "Clip",
    "CCF_CL5":                         "Sheeting Clip",
    "CCF_CL103":                       "Girt / Jamb Clip",
    "CCF_CL104":                       "Jamb Base Clip",
    "CCF_CL100":                       "Header to Jamb Clip",
    "FRAMED OPENING JAMB / SUB JAMB":  "Jamb",
    "FRAME OPENING JAMB / SUB JAMB":   "Jamb",
    "WALL GIRT":                       "Wall Girt",
    "ROOF PURLIN":                     "Roof Purlin",
    "FRAME OPENING HEADER":            "Frame Opening Header",
}

PART_MAP = {
    "L1225E14":     "12.25E14",
    "L4X2X16GA":    "4X2X16Ga",
    "SSL10_(3)":    "10X35C14",
    "WRAP8X4X14GA": "8X25C16",
}

# ── UTILIDADES DE LOG ─────────────────────────────────────────
def log(msg):
    """Logs informativos -> stderr (no contaminan el JSON de stdout)."""
    print(msg, file=sys.stderr, flush=True)

def output_json(data):
    """Resultado final -> stdout para que N8N lo parsee."""
    print(json.dumps(data, default=str, ensure_ascii=False))

# ── TRANSFORMACIONES ──────────────────────────────────────────
def t_desc(v):
    if not v: return ""
    k = str(v).strip().upper()
    if k in DESC_MAP: return DESC_MAP[k]
    return str(v).strip().replace("_", " ").title()

def t_part(v):
    if not v: return ""
    k = str(v).strip().upper()
    if k in PART_MAP: return PART_MAP[k]
    if k.startswith("L") and len(k) > 1 and k[1].isdigit():
        return str(v).strip()[1:]
    return str(v).strip()

def t_color(v):
    if v is None or str(v).strip() in ("0", ""): return ""
    return str(v).strip()

def t_dwg(v):
    if v is None or str(v).strip() in ("0", ""): return ""
    return str(v).strip()

def to_float(v):
    """Convierte a float si es posible (para columna WT.)."""
    if v is None or v == "":
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return v

def agregar_pulgadas(texto):
    """Convierte 'IFlg=3 10.75 13.25' -> 'IFlg=  3\"   10.75\"   13.25\"'"""
    if not texto: return texto
    texto = str(texto).strip()
    prefijo, resto = "", texto
    for p in ["IFlg=", "Web=", "IFlg =", "Web ="]:
        if texto.upper().startswith(p.upper()):
            prefijo = texto[:len(p)]
            resto = texto[len(p):].strip()
            break
    partes = resto.split()
    resultado = []
    for parte in partes:
        try:
            float(parte)
            resultado.append(f'{parte}"')
        except ValueError:
            resultado.append(parte)
    return prefijo + "  " + "   ".join(resultado)

def es_detalle(fila):
    for c in fila:
        if c and (str(c).upper().startswith("IFLG=") or str(c).upper().startswith("WEB=")):
            return True
    return False

def es_peso(fila):
    return any(c and "PAGE WEIGHT" in str(c).upper() for c in fila)

def extraer_peso(fila):
    encontrado = False
    for c in fila:
        if c and "PAGE WEIGHT" in str(c).upper():
            encontrado = True
            continue
        if encontrado and c is not None:
            try: return float(c)
            except: return c
    return None

def concat_detalle(fila):
    return " ".join(str(c).strip() for c in fila if c and str(c).strip())

# ── LECTURA DE ARCHIVO EXCEL (.xls o .xlsx) ───────────────────
def cargar_archivo_excel(ruta):
    """
    Carga un archivo Excel y devuelve un workbook de openpyxl.
    Si es .xls (Excel 97-2003), convierte en memoria usando xlrd.
    Si es .xlsx, lo abre directamente.
    """
    if ruta.lower().endswith('.xls'):
        import xlrd
        xls_book = xlrd.open_workbook(ruta)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)  # quitar la hoja por defecto

        for sheet_name in xls_book.sheet_names():
            xls_sheet = xls_book.sheet_by_name(sheet_name)
            ws = wb.create_sheet(sheet_name)
            for row_idx in range(xls_sheet.nrows):
                for col_idx in range(xls_sheet.ncols):
                    cell_type = xls_sheet.cell_type(row_idx, col_idx)
                    value = xls_sheet.cell_value(row_idx, col_idx)

                    if cell_type in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
                        continue
                    if cell_type == xlrd.XL_CELL_DATE:
                        value = xlrd.xldate.xldate_as_datetime(value, xls_book.datemode)
                    elif cell_type == xlrd.XL_CELL_BOOLEAN:
                        value = bool(value)
                    elif cell_type == xlrd.XL_CELL_ERROR:
                        value = None

                    ws.cell(row=row_idx + 1, column=col_idx + 1, value=value)
        return wb
    else:
        return openpyxl.load_workbook(ruta, data_only=True)


# ── LECTURA ENCABEZADO ────────────────────────────────────────
def leer_encabezado(ws):
    enc = {}
    for fila in ws.iter_rows(min_row=1, max_row=7):
        for c in fila:
            if not c.value: continue
            v = str(c.value).upper()
            if "JOB NUMBER"      in v: enc["job"]      = ws.cell(c.row, c.column+1).value
            if "ISSUE DATE"      in v: enc["fecha"]    = ws.cell(c.row, c.column+1).value
            if "BUILDING NUMBER" in v: enc["edificio"] = ws.cell(c.row, c.column+1).value
            if "BLDG DESCRIP"    in v: enc["descrip"]  = ws.cell(c.row, c.column+1).value
            if "CUSTOMER"        in v: enc["cliente"]  = ws.cell(c.row, c.column+2).value
    return enc

# ── LECTURA DE PIEZAS ─────────────────────────────────────────
def leer_piezas(ws, tipo):
    piezas, pieza, peso = [], None, None
    for fila in ws.iter_rows(min_row=11, values_only=True):
        if es_peso(fila):
            peso = extraer_peso(fila)
            break
        if es_detalle(fila):
            if pieza:
                pieza["detalle"] = agregar_pulgadas(concat_detalle(fila))
            continue
        cols = list(fila) + [None]*15
        qty = cols[0]
        if qty is not None and str(qty).strip() not in ("", "QTY"):
            if pieza: piezas.append(pieza)
            pieza = {"detalle": None}

            if tipo == "Eave Struts":
                pieza.update({"qty":cols[0],"mark":cols[1],"desc":t_desc(cols[2]),
                    "pitch":cols[3],"part":t_part(cols[4]),"punch":cols[5],
                    "color":t_color(cols[6]),"dwg":t_dwg(cols[7]),
                    "length":cols[8],"wt":to_float(cols[9])})

            elif tipo in ("Cold Form Members (CEE)","Cold Form Members (ZEE)",
                          "Cold Form Members (ZEE) (2)","Cold Form Members (ZEE) (3)"):
                pieza.update({"qty":cols[0],"mark":cols[1],"desc":t_desc(cols[2]),
                    "part":t_part(cols[3]),"punch":cols[4],
                    "dwg":t_dwg(cols[5]),"color":t_color(cols[6]),
                    "length":cols[7],"wt":to_float(cols[8])})

            elif tipo == "Misc. Cold Form":
                pieza.update({"qty":cols[0],"mark":cols[1],"desc":t_desc(cols[2]),
                    "part":t_part(cols[3]),"dwg":t_dwg(cols[4]),
                    "color":t_color(cols[5]),"length":cols[6],"wt":to_float(cols[7])})

            elif tipo in ("Clips","Pre-Galv Clips"):
                pieza.update({"qty":cols[0],"mark":cols[1],"desc":t_desc(cols[2]),
                    "color":t_color(cols[3]),"dwg":t_dwg(cols[4]),"wt":to_float(cols[5])})

            elif tipo == "Standing Seam Hardware":
                pieza.update({"desc":t_desc(cols[0]),"dwg":t_dwg(cols[1]),
                    "color":t_color(cols[2]),"length":cols[3],"wt":to_float(cols[4])})

            elif tipo == "Screws":
                pieza.update({"qty":cols[0],"mark":cols[1],"desc":t_desc(cols[2]),
                    "color":t_color(cols[3]),"length":cols[4],"wt":to_float(cols[5])})

    if pieza: piezas.append(pieza)
    return piezas, peso

# ── ESCRITURA ─────────────────────────────────────────────────
def escribir_enc(mws, enc, num, total):
    for fila in mws.iter_rows(min_row=1, max_row=7):
        for c in fila:
            if not c.value: continue
            v = str(c.value).upper()
            if "SHIPPER NUMBER" in v: mws.cell(c.row, c.column+2).value = num
            if v.strip() == "OF":     mws.cell(c.row, c.column+1).value = total
            if "JOB NUMBER"      in v: mws.cell(c.row, c.column+1).value = enc.get("job")
            if "ISSUE DATE"      in v: mws.cell(c.row, c.column+1).value = enc.get("fecha")
            if "BUILDING NUMBER" in v: mws.cell(c.row, c.column+1).value = enc.get("edificio")
            if "BLDG DESCRIP"    in v: mws.cell(c.row, c.column+1).value = enc.get("descrip")
            if "CUSTOMER"        in v: mws.cell(c.row, c.column+2).value = enc.get("cliente")

def escribir_piezas(mws, piezas, peso, tipo):
    r = 11
    for p in piezas:
        if tipo == "Eave Struts":
            mws.cell(r,1).value=p.get("qty");   mws.cell(r,2).value=p.get("mark")
            mws.cell(r,3).value=p.get("desc");  mws.cell(r,4).value=p.get("pitch")
            mws.cell(r,5).value=p.get("part");  mws.cell(r,6).value=p.get("punch")
            mws.cell(r,7).value=p.get("color"); mws.cell(r,8).value=p.get("dwg")
            mws.cell(r,9).value=p.get("length");mws.cell(r,10).value=p.get("wt")

        elif tipo in ("Cold Form Members (CEE)","Cold Form Members (ZEE)",
                      "Cold Form Members (ZEE) (2)","Cold Form Members (ZEE) (3)"):
            mws.cell(r,1).value=p.get("qty");   mws.cell(r,2).value=p.get("mark")
            mws.cell(r,3).value=p.get("desc");  mws.cell(r,4).value=p.get("part")
            mws.cell(r,5).value=p.get("punch"); mws.cell(r,6).value=p.get("dwg")
            mws.cell(r,7).value=p.get("color"); mws.cell(r,8).value=p.get("length")
            mws.cell(r,9).value=p.get("wt")

        elif tipo == "Misc. Cold Form":
            mws.cell(r,1).value=p.get("qty");   mws.cell(r,2).value=p.get("mark")
            mws.cell(r,3).value=p.get("desc");  mws.cell(r,4).value=p.get("part")
            mws.cell(r,5).value=p.get("dwg");   mws.cell(r,6).value=p.get("color")
            mws.cell(r,7).value=p.get("length");mws.cell(r,8).value=p.get("wt")

        elif tipo in ("Clips","Pre-Galv Clips"):
            mws.cell(r,1).value=p.get("qty");   mws.cell(r,2).value=p.get("mark")
            mws.cell(r,3).value=p.get("desc");  mws.cell(r,7).value=p.get("color")
            mws.cell(r,8).value=p.get("dwg");   mws.cell(r,10).value=p.get("wt")

        elif tipo == "Standing Seam Hardware":
            mws.cell(r,3).value=p.get("desc");  mws.cell(r,6).value=p.get("dwg")
            mws.cell(r,7).value=p.get("color"); mws.cell(r,9).value=p.get("length")
            mws.cell(r,10).value=p.get("wt")

        elif tipo == "Screws":
            mws.cell(r,1).value=p.get("qty");   mws.cell(r,2).value=p.get("mark")
            mws.cell(r,3).value=p.get("desc");  mws.cell(r,7).value=p.get("color")
            mws.cell(r,9).value=p.get("length");mws.cell(r,10).value=p.get("wt")

        r += 1
        if p.get("detalle"):
            mws.cell(r, 6).value = p["detalle"]
            r += 1

    if peso is not None:
        r += 1
        for fila in mws.iter_rows(min_row=max(1, r-5), max_row=r+3):
            for c in fila:
                if c.value and "PAGE WEIGHT" in str(c.value).upper():
                    mws.cell(c.row, c.column+1).value = peso
                    return
        mws.cell(r, 9).value  = "PAGE WEIGHT:"
        mws.cell(r, 10).value = peso

# ── PROCESAMIENTO PRINCIPAL ───────────────────────────────────
def procesar():
    inicio = datetime.now()
    log_entries, errores, files_with_errors = [], [], []

    # 1) Validar carpeta y archivos
    if not os.path.isdir(CARPETA_TEKLA):
        return {
            "status": "error",
            "message": f"Carpeta de Tekla no existe: {CARPETA_TEKLA}",
            "duration_seconds": 0
        }

    archivos = sorted([f for f in os.listdir(CARPETA_TEKLA)
                       if f.lower().endswith((".xlsx", ".xls"))
                       and os.path.isfile(os.path.join(CARPETA_TEKLA, f))])

    if not archivos:
        return {
            "status": "no_files",
            "message": "No hay archivos para procesar",
            "duration_seconds": 0
        }

    log(f"Encontrados {len(archivos)} archivos para procesar")

    # 2) Validar macro
    if not os.path.isfile(RUTA_MACRO):
        return {
            "status": "error",
            "message": f"Macro Allied no existe: {RUTA_MACRO}",
            "duration_seconds": 0
        }

    # 3) Cargar macro (preservando VBA)
    try:
        macro_wb = openpyxl.load_workbook(RUTA_MACRO)
    except Exception as e:
        return {
            "status": "error",
            "message": f"No se pudo abrir la macro: {e}",
            "duration_seconds": (datetime.now()-inicio).seconds
        }

    zee_count = 0
    job_number = "SIN_NUMERO"
    total = len(archivos)
    procesados_ok = []  # archivos a mover a /procesados al final

    # 4) Procesar cada archivo
    for i, nombre in enumerate(archivos, 1):
        ruta_archivo = os.path.join(CARPETA_TEKLA, nombre)
        try:
            wb = cargar_archivo_excel(ruta_archivo)
            ws = wb.active
            tab_tekla = ws.title
            tab_macro = TAB_MAP.get(tab_tekla)

            # Manejo especial de multiples ZEE
            if tab_tekla == "SBS_ZEE_Secondary_Shipper":
                if zee_count > 0 and zee_count-1 < len(TAB_ZEE_EXTRA):
                    tab_macro = TAB_ZEE_EXTRA[zee_count-1]
                zee_count += 1

            if not tab_macro or tab_macro not in macro_wb.sheetnames:
                msg = f"SKIP [{nombre}] -> pestana '{tab_macro}' no encontrada"
                log_entries.append(msg); log(msg)
                continue

            enc = leer_encabezado(ws)
            if enc.get("job"):
                job_number = str(enc["job"]).split("_")[0].strip()

            piezas, peso = leer_piezas(ws, tab_macro)
            mws = macro_wb[tab_macro]
            escribir_enc(mws, enc, i, total)
            escribir_piezas(mws, piezas, peso, tab_macro)

            msg = f"OK [{nombre}] -> [{tab_macro}] | {len(piezas)} piezas | Peso: {peso}"
            log_entries.append(msg); log(msg)
            procesados_ok.append(nombre)

        except Exception as e:
            msg = f"ERROR [{nombre}]: {e}"
            log_entries.append(msg); errores.append(msg); log(msg)
            files_with_errors.append({"file": nombre, "error": str(e)})
            log(traceback.format_exc())

    # 5) Pestana COVER
    if "Cover" in macro_wb.sheetnames:
        cov = macro_wb["Cover"]
        for fila in cov.iter_rows():
            for c in fila:
                if not c.value: continue
                if "JOB NUMBER" in str(c.value).upper():
                    cov.cell(c.row, c.column+1).value = job_number
                if "TOTAL NUMBER" in str(c.value).upper():
                    cov.cell(c.row, c.column+1).value = total

    # 6) Pestana LOG
    if "LOG" in macro_wb.sheetnames:
        del macro_wb["LOG"]
    lws = macro_wb.create_sheet("LOG")
    lws["A1"] = "REPORTE DE PROCESAMIENTO - TEKLA -> MACRO ALLIED"
    lws["A2"] = f"Fecha: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    lws["A3"] = f"Archivos procesados: {total}"
    lws["A4"] = f"Errores: {len(errores)}"
    lws["A5"] = ""
    for idx, linea in enumerate(log_entries, 6):
        lws.cell(idx, 1).value = linea

    # 7) Guardar
    os.makedirs(CARPETA_SALIDA, exist_ok=True)
    salida = os.path.join(CARPETA_SALIDA, f"{job_number}_Secondary_Shipper.xlsx")
    macro_wb.save(salida)
    log(f"Guardado en: {salida}")

    # 8) Archivar archivos procesados (para no reprocesarlos en la siguiente corrida)
    os.makedirs(CARPETA_PROCESADOS, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    archivo_destino_ts = os.path.join(CARPETA_PROCESADOS, timestamp)
    os.makedirs(archivo_destino_ts, exist_ok=True)
    for nombre in procesados_ok:
        try:
            shutil.move(
                os.path.join(CARPETA_TEKLA, nombre),
                os.path.join(archivo_destino_ts, nombre)
            )
        except Exception as e:
            log(f"No se pudo mover {nombre}: {e}")

    duracion = (datetime.now() - inicio).seconds

    return {
        "status": "success" if not errores else "partial_success",
        "job_number": job_number,
        "files_processed": total,
        "files_with_errors": files_with_errors,
        "output_file": salida,
        "archive_folder": archivo_destino_ts,
        "duration_seconds": duracion,
        "log_entries": log_entries
    }


# ── ENTRY POINT ───────────────────────────────────────────────
if __name__ == "__main__":
    try:
        resultado = procesar()
        output_json(resultado)
        # Exit 0 siempre que el script haya terminado de forma controlada.
        # Que el IF de N8N decida si seguir (success/partial_success) o detener (error/no_files).
        sys.exit(0)
    except Exception as e:
        output_json({
            "status": "fatal_error",
            "message": str(e),
            "traceback": traceback.format_exc()
        })
        sys.exit(1)
