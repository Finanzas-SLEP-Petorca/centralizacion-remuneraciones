from __future__ import annotations

import argparse
import csv
import glob
import io
import re
import unicodedata
from collections import defaultdict
from contextlib import redirect_stderr, redirect_stdout
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, List, Tuple

import xlrd
from openpyxl import Workbook
from openpyxl.styles import Font
from pypdf import PdfReader


INFO_CODES = {"998", "1001", "1002", "30001", "30002"}
LIQUIDO_CODE = "30003"
EMPLOYER_PREFIX = "320"
TOTAL_PREFIX = "300"

FUENTE_COLS = ["GENERAL", "SEP", "PIE", "APORTE FISCAL", "FAEP"]
FUENTE_TO_PROYECTO = {
    "GENERAL": "SUBVENCION NORMAL",
    "SEP": "PROYECTO SEP",
    "PIE": "PROYECTO PIE",
    "APORTE FISCAL": "APORTE FISCAL",
    "FAEP": "FAEP",
}
_CAS_PROYECTOS = frozenset({"SUBVENCION NORMAL", "PROYECTO SEP", "PROYECTO PIE"})
# normalized_key("ASISTENTE EDUCACION") = "asistenteeducacion"
_TIPO_TO_ESCALAFON = {
    "asistenteeducacion": "ASISTENTE DE LA EDUCACION",
    "docente": "DOCENTE",
}


@dataclass
class Paths:
    process_files: List[str]
    gasto_files: List[str]
    central_files: List[str]
    cuentas_pdf: str = ""


def normalize_text(value: object) -> str:
    text = "" if value is None else str(value)
    return re.sub(r"\s+", " ", text).strip()


def normalized_key(value: object) -> str:
    text = normalize_text(value).lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^a-z0-9]+", "", text)


def parse_code(header: str) -> str | None:
    match = re.search(r"\((\d+)\)\s*$", header)
    return match.group(1) if match else None


def process_name_from_file(path: str) -> str:
    full_name = str(path)
    name = Path(path).stem
    process_match = re.search(r"PROCESO(?:\s|[-_])*(NORMAL|\d+)", full_name, re.IGNORECASE)
    if process_match:
        value = process_match.group(1).upper()
        if value == "NORMAL":
            return "Proceso Normal"
        return f"Proceso {int(value)}"
    upper_name = name.upper()
    if upper_name.startswith("MR2026_") or upper_name.startswith("ASIENTO REMUNERACIONES_") or upper_name.startswith("INFORMEGASTOFINANCIAMIENTO_") or upper_name.startswith("INFORMECENTRALIZACION"):
        return "Proceso Normal"
    if name.lower().startswith("informegastofinanciamiento_"):
        name = name.replace("InformeGastoFinanciamiento_", "", 1)
    if name.lower().startswith("informecentralizacionespecial_2_3_2026_11_20_03 "):
        name = name.replace("InformeCentralizacionEspecial_2_3_2026_11_20_03 ", "", 1)
    name = re.sub(r"\s+202602$", "", name)
    return normalize_text(name)


def safe_open_workbook(path: str):
    buffer = io.StringIO()
    with redirect_stdout(buffer), redirect_stderr(buffer):
        return xlrd.open_workbook(path)


def discover_paths() -> Paths:
    process_files = sorted(
        glob.glob(
            r"C:\Users\wilson.rojas\OneDrive*PETORCA\Contraparte Subtitulo 21\202602\Proceso * Educacion 202602.xls"
        )
    )
    gasto_files = sorted(
        glob.glob(r"C:\Users\wilson.rojas\Downloads\InformeGastoFinanciamiento_Proceso * Educacion.xls")
    )
    central_files = sorted(
        glob.glob(r"C:\Users\wilson.rojas\Downloads\InformeCentralizacionEspecial*Proceso * Educacion.xls")
    )
    cuentas = glob.glob(r"C:\Users\wilson.rojas\Downloads\Cuentas contables.pdf")
    cuentas_pdf = cuentas[0] if cuentas else ""
    return Paths(process_files, gasto_files, central_files, cuentas_pdf)


def open_sheet(path: str, sheet_name: str | None = None, sheet_index: int | None = None):
    book = safe_open_workbook(path)
    if sheet_name is not None:
        return book.sheet_by_name(sheet_name)
    return book.sheet_by_index(sheet_index or 0)


def rows_from_sheet(sheet) -> Tuple[List[str], Iterable[List[object]]]:
    headers = [normalize_text(sheet.cell_value(0, c)) for c in range(sheet.ncols)]
    rows = ([sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(1, sheet.nrows))
    return headers, rows


def find_header_index(headers: List[str], *candidates: str) -> int:
    normalized = {normalized_key(header): idx for idx, header in enumerate(headers)}
    for candidate in candidates:
        key = normalized_key(candidate)
        if key in normalized:
            return normalized[key]
    raise KeyError(f"No se encontró encabezado para {candidates!r}")


def aggregate_process(path: str, process_name_override: str = ""):
    process = process_name_override or process_name_from_file(path)
    hab_sheet = open_sheet(path, sheet_name="Haberes")
    des_sheet = open_sheet(path, sheet_name="Descuentos")
    hab_headers, hab_rows = rows_from_sheet(hab_sheet)
    des_headers, des_rows = rows_from_sheet(des_sheet)

    debit_rows = []
    employer_credit_rows = []
    discount_credit_rows = []
    liquid_rows = []

    debit_total = 0.0
    employer_total = 0.0
    credit_total = 0.0
    liquid_total = 0.0

    hab_code_cols = [(idx, parse_code(header), header) for idx, header in enumerate(hab_headers) if parse_code(header)]
    des_code_cols = [(idx, parse_code(header), header) for idx, header in enumerate(des_headers) if parse_code(header)]

    for row in hab_rows:
        for idx, code, header in hab_code_cols:
            amount = row[idx] or 0
            if not isinstance(amount, (int, float)) or amount == 0:
                continue
            if code.startswith(TOTAL_PREFIX):
                continue
            description = normalize_text(re.sub(r"\s*\(\d+\)\s*$", "", header))
            if code.startswith(EMPLOYER_PREFIX):
                employer_credit_rows.append((process, code, description, float(amount)))
                employer_total += float(amount)
            debit_rows.append((process, code, description, float(amount)))
            debit_total += float(amount)

    for row in des_rows:
        for idx, code, header in des_code_cols:
            amount = row[idx] or 0
            if not isinstance(amount, (int, float)) or amount == 0:
                continue
            description = normalize_text(re.sub(r"\s*\(\d+\)\s*$", "", header))
            if code == LIQUIDO_CODE:
                liquid_rows.append((process, code, "LIQUIDO A PAGO", float(amount)))
                liquid_total += float(amount)
                continue
            if code in INFO_CODES:
                continue
            if code == "30000":
                discount_credit_rows.append((process, code, description, float(amount)))
                credit_total += float(amount)
                continue
            if code.startswith(TOTAL_PREFIX):
                continue
            discount_credit_rows.append((process, code, description, float(amount)))
            credit_total += float(amount)

    return {
        "process": process,
        "debit_rows": debit_rows,
        "employer_credit_rows": employer_credit_rows,
        "discount_credit_rows": discount_credit_rows,
        "liquid_rows": liquid_rows,
        "totals": {
            "debe_proceso": round(debit_total, 2),
            "haber_descuentos": round(credit_total, 2),
            "haber_liquido": round(liquid_total, 2),
            "haber_aportes_patronales": round(employer_total, 2),
            "diferencia": round(debit_total - credit_total - liquid_total - employer_total, 2),
        },
    }


def aggregate_gasto(path: str, process_name_override: str = ""):
    process = process_name_override or process_name_from_file(path)
    sheet = open_sheet(path, sheet_index=0)
    headers, rows = rows_from_sheet(sheet)
    code_idx = find_header_index(headers, "Código", "Codigo")
    description_idx = find_header_index(headers, "Descripción", "Descripcion")
    project_idx = find_header_index(headers, "Proyecto")
    centro_costo_idx = find_header_index(headers, "Centro Costo")
    tipo_cargo_idx = find_header_index(headers, "Tipo de Cargo")
    monto_idx = find_header_index(headers, "Monto Debe")

    grouped: Dict[Tuple[str, str, str, str], float] = defaultdict(float)
    grouped_por_centro: Dict[Tuple[str, str, str, str, str], float] = defaultdict(float)
    raw_total = 0.0
    for row in rows:
        raw_code = row[code_idx]
        code = str(int(raw_code)) if isinstance(raw_code, float) else normalize_text(raw_code)
        description = normalize_text(row[description_idx])
        project = normalize_text(row[project_idx])
        centro_costo = normalize_text(row[centro_costo_idx])
        tipo_cargo = normalize_text(row[tipo_cargo_idx]).upper()
        if not code or not description or description.upper() == "TOTALES:":
            continue
        amount = float(row[monto_idx] or 0)
        raw_total += amount
        grouped[(project, tipo_cargo, code, description)] += amount
        grouped_por_centro[(centro_costo, project, tipo_cargo, code, description)] += amount

    detail = [
        (process, project, tipo_cargo, code, description, round(amount, 2))
        for (project, tipo_cargo, code, description), amount in sorted(grouped.items())
    ]
    detail_por_centro = [
        (process, cc, project, tipo_cargo, code, description, round(amount, 2))
        for (cc, project, tipo_cargo, code, description), amount in sorted(grouped_por_centro.items())
    ]
    return {
        "process": process,
        "detail": detail,
        "detail_por_centro": detail_por_centro,
        "total_gasto_original": round(raw_total, 2),
        "total_gasto_real": round(raw_total, 2),
    }


def aggregate_asiento_reference(path: str, process_name_override: str = ""):
    process = process_name_override or process_name_from_file(path)
    book = safe_open_workbook(path)
    totals = {"Haberes": 0.0, "Descuentos": 0.0}
    for sheet_name in ("Haberes", "Descuentos"):
        if sheet_name not in book.sheet_names():
            continue
        sh = book.sheet_by_name(sheet_name)
        if sh.nrows > 0 and sh.ncols > 1 and isinstance(sh.cell_value(0, 1), (int, float)):
            totals[sheet_name] = float(sh.cell_value(0, 1))
    return {"process": process, "asiento_haberes": totals["Haberes"], "asiento_descuentos": totals["Descuentos"]}


def aggregate_central(path: str, process_name_override: str = ""):
    process = process_name_override or process_name_from_file(path)
    sheet = open_sheet(path, sheet_index=0)
    debe = 0.0
    haber = 0.0
    for r in range(sheet.nrows):
        concept = normalize_text(sheet.cell_value(r, 2)) if sheet.ncols > 2 else ""
        if concept.upper() == "TOTALES:":
            continue
        if sheet.ncols > 3 and isinstance(sheet.cell_value(r, 3), (int, float)):
            debe += float(sheet.cell_value(r, 3))
        if sheet.ncols > 4 and isinstance(sheet.cell_value(r, 4), (int, float)):
            haber += float(sheet.cell_value(r, 4))
    return {
        "process": process,
        "debe_caschile": round(debe, 2),
        "haber_caschile": round(haber, 2),
        "diferencia_caschile": round(debe - haber, 2),
    }


def infer_section(tipo_cargo: str) -> str:
    if tipo_cargo == "PLANTA":
        return "PLANTA"
    if tipo_cargo == "CONTRATA":
        return "CONTRATA"
    if tipo_cargo == "SUPLENCIA":
        return "SUPLENCIA"
    if tipo_cargo in {"ASISTENTE EDUCACION", "CODIGO DEL TRABAJO", "CODIGO DEL TRABAJO LEY 18.620", "CODIGO DEL TRABAJO LEY 18.620 "}:
        return "CODIGOTRABAJO"
    return tipo_cargo


def section_label(section: str) -> str:
    labels = {
        "PLANTA": "PLANTA",
        "CONTRATA": "CONTRATA",
        "SUPLENCIA": "SUPLENCIA",
        "CODIGOTRABAJO": "CODIGO DEL TRABAJO",
    }
    return labels.get(section, section)


def parse_existing_budget_map(central_paths: List[str]) -> Dict[Tuple[str, str], str]:
    mapping: Dict[Tuple[str, str], str] = {}
    for path in central_paths:
        sheet = open_sheet(path, sheet_index=0)
        section = ""
        for r in range(sheet.nrows):
            col0 = normalize_text(sheet.cell_value(r, 0)) if sheet.ncols > 0 else ""
            col1 = normalize_text(sheet.cell_value(r, 1)) if sheet.ncols > 1 else ""
            col2 = normalize_text(sheet.cell_value(r, 2)) if sheet.ncols > 2 else ""
            if col0 in {"PLANTA", "CONTRATA", "SUPLENCIA", "CODIGOTRABAJO"}:
                section = col0
            if not section or not col1 or not col2 or "TOTALES" in col2 or col1 in {"-", "0-"}:
                continue
            mapping.setdefault((section, col2.upper()), col1)
    return mapping


def read_cuentas_pdf(cuentas_pdf: str) -> str:
    reader = PdfReader(cuentas_pdf)
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def extract_pdf_budget_map(cuentas_pdf: str) -> Dict[str, Dict[str, str]]:
    if not cuentas_pdf or not Path(cuentas_pdf).exists():
        return {}
    text = " ".join((page.extract_text() or "") for page in PdfReader(cuentas_pdf).pages)
    account_pattern = r"\d{3}-\d{2}-\d{2,3}-\d{3}-\d{3}-\d{3}"
    results: Dict[str, Dict[str, str]] = {}
    for code in sorted(set(re.findall(r"(?<!\d)(\d{3,5})(?=[A-Z( ])", text))):
        match = re.search(rf"{code}\s*.*?((?:{account_pattern}){{1,4}})", text)
        if not match:
            continue
        accounts = re.findall(account_pattern, match.group(1))
        if len(accounts) < 4:
            continue
        results[code] = {
            "CONTRATA": accounts[0],
            "SUPLENCIA": accounts[1],
            "ASISTENTE EDUCACION": accounts[2],
            "PLANTA": accounts[3],
        }
    return results


def load_mapping_csv(path: Path) -> Dict[Tuple[str, str, str, str], Dict[str, str]]:
    if not path.exists():
        return {}
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        return {
            (
                row["lado"],
                row["origen"],
                row["tipo_cargo"],
                row["codigo"],
            ): row
            for row in reader
        }


def load_mapping_workbook(path: Path) -> Dict[Tuple[str, str, str, str], Dict[str, str]]:
    if not path.exists():
        return {}
    wb = Workbook()
    try:
        from openpyxl import load_workbook

        wb = load_workbook(path, data_only=False)
    except Exception:
        return {}
    if "CompletarAqui" not in wb.sheetnames:
        return {}
    ws = wb["CompletarAqui"]
    headers = [normalize_text(cell.value) for cell in ws[1]]
    mapping = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        data = {headers[i]: ("" if row[i] is None else str(row[i])) for i in range(len(headers))}
        key = (
            data.get("lado", ""),
            data.get("origen", ""),
            data.get("tipo_cargo", ""),
            data.get("codigo", ""),
        )
        if not any(key):
            continue
        mapping[key] = {
            "lado": data.get("lado", ""),
            "origen": data.get("origen", ""),
            "tipo_cargo": data.get("tipo_cargo", ""),
            "codigo": data.get("codigo", ""),
            "descripcion": data.get("descripcion", ""),
            "cuenta_contable": data.get("cuenta_contable", ""),
            "cuenta_presupuestaria_sugerida": data.get("cuenta_presupuestaria_sugerida", ""),
            "glosa_sugerida": data.get("glosa_sugerida", ""),
            "observacion": data.get("observacion", ""),
        }
    return mapping


def load_diccionario(
    path: str,
) -> Tuple[Dict[Tuple[str, str], Dict[str, int]], Dict[str, Dict[str, int]]]:
    """
    Carga el diccionario de haberes por fuente de financiamiento.
    Retorna (primary, fallback):
      primary  → {(codigo, normalized_descripcion): {GENERAL, SEP, PIE, APORTE FISCAL, FAEP}}
      fallback → {codigo: OR de todas las filas con ese código}
    """
    if not path or not Path(path).exists():
        return {}, {}
    from openpyxl import load_workbook as _lw

    wb = _lw(path, data_only=True)
    ws = wb.active
    raw_headers = [normalize_text(c.value) for c in ws[1]]
    col_idx = {normalized_key(h): i for i, h in enumerate(raw_headers)}

    codigo_col = col_idx.get("codigo")
    desc_col = col_idx.get("descripcion")
    if codigo_col is None:
        return {}, {}

    primary: Dict[Tuple[str, str], Dict[str, int]] = {}
    fallback: Dict[str, Dict[str, int]] = {}

    for row in ws.iter_rows(min_row=2, values_only=True):
        raw_code = row[codigo_col]
        if raw_code is None:
            continue
        codigo = str(int(raw_code)) if isinstance(raw_code, float) else normalize_text(str(raw_code))
        if not codigo:
            continue
        desc = normalize_text(row[desc_col]) if desc_col is not None else ""
        flags: Dict[str, int] = {}
        for fuente in FUENTE_COLS:
            cidx = col_idx.get(normalized_key(fuente))
            val = (row[cidx] if cidx is not None else None) or 0
            flags[fuente] = int(val)
        primary[(codigo, normalized_key(desc))] = flags
        existing = fallback.get(codigo, {f: 0 for f in FUENTE_COLS})
        fallback[codigo] = {f: max(existing.get(f, 0), flags[f]) for f in FUENTE_COLS}

    return primary, fallback


def update_diccionario(path: str, new_entries: List[Dict[str, str]]) -> int:
    """
    Agrega haberes nuevos al diccionario con todos los flags en 0.
    Nunca modifica filas existentes.
    Las filas nuevas se resaltan en amarillo para que el equipo las identifique.
    Retorna la cantidad de filas agregadas.
    """
    if not path or not Path(path).exists() or not new_entries:
        return 0
    from openpyxl import load_workbook as _lw
    from openpyxl.styles import PatternFill

    wb = _lw(path)
    ws = wb.active
    raw_headers = [normalize_text(c.value) for c in ws[1]]
    col_idx = {normalized_key(h): i for i, h in enumerate(raw_headers)}
    codigo_col = col_idx.get("codigo")
    if codigo_col is None:
        return 0

    # Detectar columnas de flags (contienen solo 0, 1 o None en las primeras 20 filas)
    flag_col_indices: set[int] = set()
    sample_end = min(ws.max_row + 1, 22)
    for row in ws.iter_rows(min_row=2, max_row=sample_end, values_only=True):
        for i, val in enumerate(row):
            if val in (0, 1, 0.0, 1.0):
                flag_col_indices.add(i)

    # Recopilar códigos existentes
    existing: set[str] = set()
    for row in ws.iter_rows(min_row=2, values_only=True):
        rc = row[codigo_col]
        if rc is not None:
            existing.add(str(int(rc)) if isinstance(rc, float) else normalize_text(str(rc)))

    amarillo = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")
    ncols = ws.max_column
    added = 0

    for entry in new_entries:
        code = entry["codigo"]
        if code in existing:
            continue

        new_row: List[object] = [None] * ncols
        # Código
        new_row[col_idx.get("codigo", 0)] = int(code) if code.isdigit() else code
        # Descripción
        if "descripcion" in col_idx:
            new_row[col_idx["descripcion"]] = entry.get("descripcion", "")
        # Todos los flags detectados → 0
        for cidx in flag_col_indices:
            if cidx < ncols:
                new_row[cidx] = 0

        ws.append(new_row)
        # Resaltar la fila en amarillo
        new_row_num = ws.max_row
        for col in range(1, ncols + 1):
            ws.cell(row=new_row_num, column=col).fill = amarillo

        existing.add(code)
        added += 1

    if added > 0:
        wb.save(path)
    return added


def build_jornada_map(maestro_paths: List[str]) -> Dict[Tuple[str, str], Dict[str, float]]:
    """
    Lee la JORNADA del Encabezado de cada Maestro.
    Retorna {(centro_costo, escalafon): {proyecto: total_jornada}}
    """
    result: Dict[Tuple[str, str], Dict[str, float]] = defaultdict(lambda: defaultdict(float))
    for path in maestro_paths:
        book = safe_open_workbook(path)
        encabezado = None
        for name in book.sheet_names():
            if "encabezado" in normalized_key(name):
                encabezado = book.sheet_by_name(name)
                break
        if encabezado is None:
            continue
        headers = [normalize_text(encabezado.cell_value(0, c)) for c in range(encabezado.ncols)]
        try:
            cc_idx = find_header_index(headers, "CENTRO DE COSTOS")
            esc_idx = find_header_index(headers, "ESCALAFON")
            proj_idx = find_header_index(headers, "PROYECTO")
            jorn_idx = find_header_index(headers, "JORNADA")
        except KeyError:
            continue
        for r in range(1, encabezado.nrows):
            cc = normalize_text(encabezado.cell_value(r, cc_idx))
            esc = normalize_text(encabezado.cell_value(r, esc_idx))
            proj = normalize_text(encabezado.cell_value(r, proj_idx))
            jorn = encabezado.cell_value(r, jorn_idx)
            if not cc or not esc or not proj or not isinstance(jorn, (int, float)):
                continue
            result[(cc, esc)][proj] += float(jorn)
    return {k: dict(v) for k, v in result.items()}


def recalculate_fuentes(
    gasto_data: Dict[str, dict],
    jornada_map: Dict[Tuple[str, str], Dict[str, float]],
    dic_primary: Dict[Tuple[str, str], Dict[str, int]],
    dic_fallback: Dict[str, Dict[str, int]],
) -> List[List[object]]:
    """
    Recalcula la distribución por fuente de financiamiento usando el diccionario.

    Principio: CAS Chile ya calcula correctamente la proporción de cada haber por
    las horas reales de cada RUT en cada proyecto (InformeGastoFinanciamiento).
    El problema es que asigna el haber a la fuente equivocada.

    Algoritmo para fuentes CAS (GENERAL/SEP/PIE):
      - Tomamos los montos que CAS Chile puso en los proyectos ELEGIBLES para este
        haber y los usamos como pesos para redistribuir TODO el monto (incluido el
        que CAS puso en proyectos NO elegibles).
      - Ejemplo: haber elegible solo en [GENERAL, SEP]
        InformeGasto: NORMAL=1000, SEP=500, PIE=300 → total=1800
        peso_NORMAL = 1000 / (1000+500) = 0.667 → corregido_NORMAL = 1800*0.667 = 1200
        peso_SEP    = 500  / (1000+500) = 0.333 → corregido_SEP    = 1800*0.333 =  600
        Si ningún elegible tiene monto en InformeGasto → reparte equitativamente.

    Para fuentes no-CAS (APORTE FISCAL / FAEP):
      - Si el haber es exclusivo de fuentes no-CAS, el total completo va allí.
      - Usa jornada_map como fallback de ponderación si hay más de una fuente no-CAS.

    Retorna filas para la hoja FuenteCorregida:
    [Proceso, Centro_Costo, Tipo_Cargo, Codigo, Descripcion,
     Fuente, Monto_CAS, Monto_Corregido, Diferencia]
    """
    rows = []
    _default_flags = {f: 0 for f in FUENTE_COLS}

    def _distribute_with_weights(
        total: float,
        proyectos: List[str],
        weights: Dict[str, float],
    ) -> Dict[str, float]:
        """
        Distribuye `total` entre `proyectos` proporcional a `weights`.
        Absorbe el residuo de redondeo en el proyecto con mayor peso.
        """
        w_sum = sum(weights.get(p, 0.0) for p in proyectos)
        result: Dict[str, float] = {}
        if w_sum <= 0:
            # Sin pesos → reparto equitativo
            share = round(total / len(proyectos), 2) if proyectos else 0.0
            for p in proyectos:
                result[p] = share
            # Absorber redondeo en el primero
            if proyectos:
                result[proyectos[0]] = round(total - sum(result[p] for p in proyectos[1:]), 2)
            return result
        # Orden descendente por peso para que el mayor absorba el residuo de redondeo
        ordered = sorted(proyectos, key=lambda p: -weights.get(p, 0.0))
        allocated = 0.0
        for i, p in enumerate(ordered):
            if i == len(ordered) - 1:
                result[p] = round(total - allocated, 2)
            else:
                share = round(total * weights.get(p, 0.0) / w_sum, 2)
                result[p] = share
                allocated += share
        return result

    for process, gdata in gasto_data.items():
        cas_amounts: Dict[Tuple[str, str, str, str, str], float] = defaultdict(float)
        total_amounts: Dict[Tuple[str, str, str, str], float] = defaultdict(float)

        for _, cc, project, tipo_cargo, codigo, descripcion, monto in gdata.get("detail_por_centro", []):
            key4 = (cc, tipo_cargo, codigo, descripcion)
            cas_amounts[(cc, tipo_cargo, codigo, descripcion, project)] += monto
            total_amounts[key4] += monto

        for cc, tipo_cargo, codigo, descripcion in sorted(total_amounts):
            total = total_amounts[(cc, tipo_cargo, codigo, descripcion)]
            entry = (
                dic_primary.get((codigo, normalized_key(descripcion)))
                or dic_fallback.get(codigo)
                or _default_flags
            )
            eligible = [f for f in FUENTE_COLS if entry.get(f, 0) == 1]
            if not eligible:
                eligible = ["GENERAL"]

            eligible_cas = [f for f in eligible if FUENTE_TO_PROYECTO[f] in _CAS_PROYECTOS]
            eligible_new = [f for f in eligible if FUENTE_TO_PROYECTO[f] not in _CAS_PROYECTOS]

            corrected: Dict[str, float] = {}

            if eligible_cas:
                # ── Fuentes CAS: usar proporciones del InformeGasto (= horas RUT) ──
                eligible_proyectos = [FUENTE_TO_PROYECTO[f] for f in eligible_cas]
                # Pesos = montos que CAS ya puso en los proyectos ELEGIBLES
                weights = {
                    p: cas_amounts.get((cc, tipo_cargo, codigo, descripcion, p), 0.0)
                    for p in eligible_proyectos
                }
                corrected.update(_distribute_with_weights(total, eligible_proyectos, weights))

            if eligible_new and not eligible_cas:
                # ── Solo fuentes no-CAS (APORTE FISCAL / FAEP) ──────────────────
                non_cas_proyectos = [FUENTE_TO_PROYECTO[f] for f in eligible_new]
                if len(non_cas_proyectos) == 1:
                    corrected[non_cas_proyectos[0]] = round(total, 2)
                else:
                    # Usar jornada como ponderador entre fuentes no-CAS
                    escalafon = _TIPO_TO_ESCALAFON.get(normalized_key(tipo_cargo), tipo_cargo)
                    jornadas = jornada_map.get((cc, escalafon), {})
                    weights_nc = {p: jornadas.get(p, 0.0) for p in non_cas_proyectos}
                    corrected.update(_distribute_with_weights(total, non_cas_proyectos, weights_nc))

            # Si eligible_new Y eligible_cas: el total ya está cubierto en CAS;
            # las fuentes no-CAS quedan en 0 (CAS Chile no separa esos fondos).

            all_fuentes: set[str] = set()
            for k5 in cas_amounts:
                if k5[:4] == (cc, tipo_cargo, codigo, descripcion):
                    all_fuentes.add(k5[4])
            all_fuentes.update(corrected.keys())

            for fuente in sorted(all_fuentes):
                monto_cas = round(cas_amounts.get((cc, tipo_cargo, codigo, descripcion, fuente), 0.0), 2)
                monto_corr = round(corrected.get(fuente, 0.0), 2)
                rows.append([
                    process, cc, tipo_cargo, codigo, descripcion,
                    fuente, monto_cas, monto_corr, round(monto_corr - monto_cas, 2),
                ])

    return rows


def collect_mapping_candidates(
    debit_detail: List[Tuple[str, str, str, str, str, float]],
    discount_rows: List[Tuple[str, str, str, float]],
    employer_rows: List[Tuple[str, str, str, float]],
    existing_budget_map: Dict[Tuple[str, str], str],
    pdf_budget_map: Dict[str, Dict[str, str]],
) -> Dict[Tuple[str, str, str, str], Dict[str, str]]:
    """Recolecta los códigos presentes en el mes actual sin tocar el maestro."""
    candidates: Dict[Tuple[str, str, str, str], Dict[str, str]] = {}

    for _, _project, tipo_cargo, codigo, descripcion, _ in debit_detail:
        key = ("DEBE", "GASTO_FINANCIAMIENTO", tipo_cargo, codigo)
        if key in candidates:
            continue
        suggested = pdf_budget_map.get(codigo, {}).get(tipo_cargo, "")
        if not suggested:
            suggested = existing_budget_map.get((infer_section(tipo_cargo), descripcion.upper()), "")
        candidates[key] = {
            "lado": "DEBE",
            "origen": "GASTO_FINANCIAMIENTO",
            "tipo_cargo": tipo_cargo,
            "codigo": codigo,
            "descripcion": descripcion,
            "cuenta_contable": "",
            "cuenta_presupuestaria_sugerida": suggested,
            "glosa_sugerida": descripcion,
            "observacion": "",
        }

    for _, codigo, descripcion, _ in discount_rows:
        key = ("HABER", "DESCUENTO", "", codigo)
        if key not in candidates:
            candidates[key] = {
                "lado": "HABER",
                "origen": "DESCUENTO",
                "tipo_cargo": "",
                "codigo": codigo,
                "descripcion": descripcion,
                "cuenta_contable": "",
                "cuenta_presupuestaria_sugerida": "",
                "glosa_sugerida": descripcion,
                "observacion": "",
            }

    for _, codigo, descripcion, _ in employer_rows:
        key = ("HABER", "APORTE_PATRONAL", "", codigo)
        if key not in candidates:
            candidates[key] = {
                "lado": "HABER",
                "origen": "APORTE_PATRONAL",
                "tipo_cargo": "",
                "codigo": codigo,
                "descripcion": descripcion,
                "cuenta_contable": "",
                "cuenta_presupuestaria_sugerida": "",
                "glosa_sugerida": descripcion,
                "observacion": "",
            }

    key = ("HABER", "LIQUIDO", "", LIQUIDO_CODE)
    candidates.setdefault(
        key,
        {
            "lado": "HABER",
            "origen": "LIQUIDO",
            "tipo_cargo": "",
            "codigo": LIQUIDO_CODE,
            "descripcion": "LIQUIDO A PAGO",
            "cuenta_contable": "",
            "cuenta_presupuestaria_sugerida": "",
            "glosa_sugerida": "LIQUIDO A PAGO",
            "observacion": "",
        },
    )
    return candidates


def merge_into_master(
    master: Dict[Tuple[str, str, str, str], Dict[str, str]],
    candidates: Dict[Tuple[str, str, str, str], Dict[str, str]],
) -> Tuple[List[Dict[str, str]], int]:
    """
    Fusiona candidatos en el maestro. Solo agrega claves que no existen.
    Nunca modifica entradas ya presentes. Retorna (filas ordenadas, cantidad nuevas).
    """
    merged = dict(master)
    new_count = 0
    for key, entry in candidates.items():
        if key not in merged:
            merged[key] = entry
            new_count += 1
    return [merged[k] for k in sorted(merged)], new_count


def write_mapping_csv(path: Path, rows: List[Dict[str, str]]):
    path.parent.mkdir(parents=True, exist_ok=True)
    fieldnames = [
        "lado",
        "origen",
        "tipo_cargo",
        "codigo",
        "descripcion",
        "cuenta_contable",
        "cuenta_presupuestaria_sugerida",
        "glosa_sugerida",
        "observacion",
    ]
    with path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def write_mapping_workbook(path: Path, rows: List[Dict[str, str]]):
    wb = Workbook()
    ws = wb.active
    ws.title = "CompletarAqui"
    headers = [
        "estado",
        "lado",
        "origen",
        "tipo_cargo",
        "codigo",
        "descripcion",
        "cuenta_contable",
        "cuenta_presupuestaria_sugerida",
        "glosa_sugerida",
        "observacion",
    ]
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in rows:
        ws.append(
            [
                "OK" if row["cuenta_contable"] else "COMPLETAR",
                row["lado"],
                row["origen"],
                row["tipo_cargo"],
                row["codigo"],
                row["descripcion"],
                row["cuenta_contable"],
                row["cuenta_presupuestaria_sugerida"],
                row["glosa_sugerida"],
                row["observacion"],
            ]
        )
    autosize(ws)
    ws.freeze_panes = "A2"

    guia = wb.create_sheet("ComoUsarlo")
    guia["A1"] = "Este archivo es el MAPEO MAESTRO de cuentas contables. Sus entradas nunca se borran ni sobreescriben."
    guia["A2"] = "1. Completa la columna cuenta_contable en filas con estado COMPLETAR y guarda este archivo."
    guia["A3"] = "2. Al procesar un nuevo mes, solo se agregan filas nuevas (códigos no vistos antes). Las existentes no cambian."
    guia["A4"] = "3. Si un código cambia de cuenta contable, edita directamente la fila correspondiente en este archivo."
    guia["A5"] = "4. La cuenta_presupuestaria_sugerida se precarga automáticamente cuando se detecta en archivos de referencia."
    autosize(guia)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


def autosize(sheet):
    for column_cells in sheet.columns:
        width = max(len(str(cell.value or "")) for cell in column_cells) + 2
        sheet.column_dimensions[column_cells[0].column_letter].width = min(width, 40)


def append_sheet(workbook: Workbook, title: str, headers: List[str], data: Iterable[Iterable[object]]):
    ws = workbook.create_sheet(title=title)
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
    for row in data:
        ws.append(list(row))
    autosize(ws)
    ws.freeze_panes = "A2"


def build_workbook(
    output_path: Path,
    process_data: Dict[str, dict],
    gasto_data: Dict[str, dict],
    central_data: Dict[str, dict],
    mapping_rows: List[Dict[str, str]],
    asiento_data: Dict[str, dict] | None = None,
    title_text: str = "CENTRALIZACION REMUNERACIONES",
    fuente_corregida_rows: List[List[object]] | None = None,
):
    wb = Workbook()
    wb.remove(wb.active)

    resumen = []
    debit_detail = []
    discount_detail = []
    employer_detail = []
    liquid_detail = []
    asiento = []

    mapping_lookup = {
        (row["lado"], row["origen"], row["tipo_cargo"], row["codigo"]): row for row in mapping_rows
    }

    asiento_data = asiento_data or {}

    for process in sorted(process_data):
        p = process_data[process]
        g = gasto_data[process]
        c = central_data.get(
            process,
            {"debe_caschile": 0.0, "haber_caschile": 0.0, "diferencia_caschile": 0.0},
        )
        a = asiento_data.get(process, {"asiento_haberes": 0.0, "asiento_descuentos": 0.0})
        resumen.append(
            [
                process,
                p["totals"]["debe_proceso"],
                g["total_gasto_original"],
                g["total_gasto_real"],
                c["debe_caschile"],
                c["haber_caschile"],
                c["diferencia_caschile"],
                p["totals"]["haber_descuentos"],
                p["totals"]["haber_liquido"],
                p["totals"]["haber_aportes_patronales"],
                a["asiento_haberes"],
                a["asiento_descuentos"],
                p["totals"]["diferencia"],
            ]
        )

        for row in g["detail"]:
            process_name, project, tipo_cargo, codigo, descripcion, monto = row
            debit_detail.append(row)
            mapping = mapping_lookup.get(("DEBE", "GASTO_FINANCIAMIENTO", tipo_cargo, codigo), {})
            asiento.append(
                [
                    process_name,
                    "DEBE",
                    "GASTO_FINANCIAMIENTO",
                    project,
                    tipo_cargo,
                    codigo,
                    descripcion,
                    monto,
                    mapping.get("cuenta_contable", ""),
                    mapping.get("cuenta_presupuestaria_sugerida", ""),
                    mapping.get("glosa_sugerida", descripcion),
                ]
            )

        discount_grouped: Dict[Tuple[str, str], float] = defaultdict(float)
        for process_name, codigo, descripcion, monto in p["discount_credit_rows"]:
            discount_grouped[(codigo, descripcion)] += monto
        for (codigo, descripcion), monto in sorted(discount_grouped.items()):
            discount_detail.append([process_name, codigo, descripcion, round(monto, 2)])
            mapping = mapping_lookup.get(("HABER", "DESCUENTO", "", codigo), {})
            asiento.append(
                [
                    process_name,
                    "HABER",
                    "DESCUENTO",
                    "",
                    "",
                    codigo,
                    descripcion,
                    round(monto, 2),
                    mapping.get("cuenta_contable", ""),
                    "",
                    mapping.get("glosa_sugerida", descripcion),
                ]
            )

        # ── Aportes patronales HABER ──────────────────────────────────────────
        # Usamos InformeGasto (misma fuente que el DEBE) para los códigos 320xx.
        # Esto evita diferencias cuando CAS Chile proratea de forma ligeramente
        # distinta entre el Maestro y el InformeGasto (ej. COTIZ. EXPECTATIVAS).
        employer_grouped: Dict[Tuple[str, str], float] = defaultdict(float)
        for _, _proj, _tc, codigo, descripcion, monto in g["detail"]:
            if codigo.startswith(EMPLOYER_PREFIX):
                employer_grouped[(codigo, descripcion)] += monto
        # Fallback: si InformeGasto no tiene patronales, usar Maestro
        if not employer_grouped:
            for _pn, codigo, descripcion, monto in p["employer_credit_rows"]:
                employer_grouped[(codigo, descripcion)] += monto
        for (codigo, descripcion), monto in sorted(employer_grouped.items()):
            employer_detail.append([process_name, codigo, descripcion, round(monto, 2)])
            mapping = mapping_lookup.get(("HABER", "APORTE_PATRONAL", "", codigo), {})
            asiento.append(
                [
                    process_name,
                    "HABER",
                    "APORTE_PATRONAL",
                    "",
                    "",
                    codigo,
                    descripcion,
                    round(monto, 2),
                    mapping.get("cuenta_contable", ""),
                    "",
                    mapping.get("glosa_sugerida", descripcion),
                ]
            )

        # ── Líquido a pago ────────────────────────────────────────────────────
        # Calculamos el líquido como residuo: DEBE_total – descuentos – patronal,
        # lo que absorbe diferencias de redondeo internas de CAS Chile (≤ ~10 $).
        debe_process = round(g["total_gasto_real"], 2)
        desc_process = round(sum(m for _, c, d, m in p["discount_credit_rows"]), 2)
        patron_process = round(sum(employer_grouped.values()), 2)
        liquid_amount_maestro = round(sum(row[3] for row in p["liquid_rows"]), 2)
        residuo = round(debe_process - desc_process - patron_process - liquid_amount_maestro, 2)
        # Solo absorber diferencias pequeñas de redondeo (≤ 10 pesos).
        # Si la diferencia es mayor, dejar el monto original (algo más está mal).
        liquid_amount = round(liquid_amount_maestro + residuo, 2) if abs(residuo) <= 10 else liquid_amount_maestro

        liquid_detail.append([process, LIQUIDO_CODE, "LIQUIDO A PAGO", liquid_amount])
        mapping = mapping_lookup.get(("HABER", "LIQUIDO", "", LIQUIDO_CODE), {})
        asiento.append(
            [
                process,
                "HABER",
                "LIQUIDO",
                "",
                "",
                LIQUIDO_CODE,
                "LIQUIDO A PAGO",
                liquid_amount,
                mapping.get("cuenta_contable", ""),
                "",
                mapping.get("glosa_sugerida", "LIQUIDO A PAGO"),
            ]
        )

    pendientes = [
        [
            row["lado"],
            row["origen"],
            row["tipo_cargo"],
            row["codigo"],
            row["descripcion"],
            row["cuenta_contable"],
            row["cuenta_presupuestaria_sugerida"],
            row["observacion"],
        ]
        for row in mapping_rows
        if not row["cuenta_contable"]
    ]

    final_format_rows = []
    funding_rows = []
    debit_grouped: Dict[Tuple[str, str, str], float] = defaultdict(float)
    credit_grouped: Dict[Tuple[str, str], float] = defaultdict(float)
    funding_grouped: Dict[Tuple[str, str, str, str, str], float] = defaultdict(float)

    preferred_process = "Proceso Normal Educacion" if any(row[0] == "Proceso Normal Educacion" for row in asiento) else None
    for row in asiento:
        process_name, lado, origen, _, tipo_cargo, _, descripcion, monto, cuenta_contable, cuenta_presupuestaria, _ = row
        if preferred_process and process_name != preferred_process:
            continue
        if lado == "DEBE":
            section = infer_section(tipo_cargo)
            cuenta = cuenta_presupuestaria or cuenta_contable
            concept = descripcion
            if section == "PLANTA":
                concept = f"{descripcion} PLANTA"
            elif section == "CONTRATA":
                concept = f"{descripcion} CONTRATA"
            elif section == "SUPLENCIA":
                concept = f"{descripcion} SUPLENCIA"
            elif section == "CODIGOTRABAJO":
                concept = f"{descripcion} CODIGO DEL TRABAJO"
            debit_grouped[(section, cuenta, concept)] += float(monto)
            funding_grouped[(process_name, row[3] or "SIN FUENTE", section, cuenta, concept)] += float(monto)
        else:
            cuenta = cuenta_contable
            concept = descripcion
            credit_grouped[(cuenta, concept)] += float(monto)

    ordered_sections = ["PLANTA", "CONTRATA", "SUPLENCIA", "CODIGOTRABAJO"]
    for section in ordered_sections:
        section_rows = [
            (cuenta, concept, amount)
            for (sec, cuenta, concept), amount in debit_grouped.items()
            if sec == section
        ]
        if not section_rows:
            continue
        first = True
        for cuenta, concept, amount in sorted(section_rows):
            final_format_rows.append(
                [section_label(section) if first else "", cuenta, concept, round(amount, 2), ""]
            )
            first = False

    for cuenta, concept in sorted(credit_grouped):
        final_format_rows.append(["", cuenta, concept, "", round(credit_grouped[(cuenta, concept)], 2)])

    total_debe = round(sum(row[3] for row in final_format_rows if isinstance(row[3], (int, float))), 2)
    total_haber = round(sum(row[4] for row in final_format_rows if isinstance(row[4], (int, float))), 2)
    final_format_rows.append(["", "", "TOTALES:", total_debe, total_haber])
    for (process_name, fuente, section, cuenta, concept), amount in sorted(funding_grouped.items()):
        funding_rows.append([process_name, fuente, section_label(section), cuenta, concept, round(amount, 2)])

    append_sheet(
        wb,
        "Resumen",
        [
            "Proceso",
            "Debe_proceso_real",
            "Gasto_CasChile_original",
            "Gasto_real",
            "Centralizacion_CasChile_debe",
            "Centralizacion_CasChile_haber",
            "Centralizacion_CasChile_diff",
            "Haber_descuentos",
            "Haber_liquido",
            "Haber_aportes_patronales",
            "Asiento_haberes_ref",
            "Asiento_descuentos_ref",
            "Diff_propuesta",
        ],
        resumen,
    )
    append_sheet(
        wb,
        "Debitos",
        ["Proceso", "Proyecto", "Tipo_cargo", "Codigo", "Descripcion", "Monto"],
        debit_detail,
    )
    append_sheet(
        wb,
        "CreditosDescuentos",
        ["Proceso", "Codigo", "Descripcion", "Monto"],
        discount_detail,
    )
    append_sheet(
        wb,
        "CreditosPatronales",
        ["Proceso", "Codigo", "Descripcion", "Monto"],
        employer_detail,
    )
    append_sheet(
        wb,
        "CreditosLiquido",
        ["Proceso", "Codigo", "Descripcion", "Monto"],
        liquid_detail,
    )
    append_sheet(
        wb,
        "AsientoPropuesto",
        [
            "Proceso",
            "Lado",
            "Origen",
            "Proyecto",
            "Tipo_cargo",
            "Codigo",
            "Descripcion",
            "Monto",
            "Cuenta_contable",
            "Cuenta_presupuestaria_sugerida",
            "Glosa_sugerida",
        ],
        asiento,
    )
    append_sheet(
        wb,
        "PorFuenteFinanciamiento",
        ["Proceso", "Fuente_financiamiento", "Seccion", "Cuenta", "Concepto", "Debe"],
        funding_rows,
    )
    if fuente_corregida_rows:
        append_sheet(
            wb,
            "FuenteCorregida",
            [
                "Proceso", "Centro_Costo", "Tipo_Cargo", "Codigo", "Descripcion",
                "Fuente", "Monto_CAS", "Monto_Corregido", "Diferencia",
            ],
            fuente_corregida_rows,
        )
    append_sheet(
        wb,
        "PendientesCuenta",
        [
            "Lado",
            "Origen",
            "Tipo_cargo",
            "Codigo",
            "Descripcion",
            "Cuenta_contable",
            "Cuenta_presupuestaria_sugerida",
            "Observacion",
        ],
        pendientes,
    )
    formato = wb.create_sheet(title="CentralizacionFinal", index=1)
    formato["A1"] = ""
    formato["B1"] = ""
    formato["C1"] = title_text
    formato["A3"] = ""
    formato["B3"] = "Cta. Presupuestaria"
    formato["C3"] = "Concepto"
    formato["D3"] = "Debe"
    formato["E3"] = "Haber"
    for cell in formato[3]:
        cell.font = Font(bold=True)
    row_idx = 4
    for row in final_format_rows:
        for col_idx, value in enumerate(row, start=1):
            formato.cell(row=row_idx, column=col_idx, value=value)
        row_idx += 1
    autosize(formato)
    formato.freeze_panes = "A4"

    instrucciones = wb.create_sheet(title="Instrucciones", index=0)
    instrucciones["A1"] = "Uso sugerido"
    instrucciones["A1"].font = Font(bold=True)
    instrucciones["A2"] = "1. Revisa el archivo mapeo_cuentas.csv y completa la columna cuenta_contable."
    instrucciones["A3"] = "2. Las filas del debe salen del detalle del InformeGastoFinanciamiento, excluyendo la fila TOTALES."
    instrucciones["A4"] = "3. Las filas del haber salen de Descuentos, Liquido a pago y Aportes patronales 320xx."
    instrucciones["A5"] = "4. La hoja Resumen muestra la diferencia real versus CasChile por proceso."
    instrucciones["A6"] = "5. Si vuelves a ejecutar el script, conserva y reutiliza el mapeo ya completado."
    autosize(instrucciones)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def run_pipeline(
    process_files: List[str],
    gasto_files: List[str],
    central_files: List[str] | None,
    cuentas_pdf: str,
    salida: str,
    mapeo: str,
    mapeo_excel: str,
    title_text: str,
    asiento_files: List[str] | None = None,
    diccionario_path: str = "",
    process_name_map: Dict[str, str] | None = None,
) -> Tuple[Path, Path, Path, int, int, int]:
    """
    Retorna (salida, mapeo_csv, mapeo_excel,
             nuevos_en_mapeo, pendientes_sin_cuenta, nuevos_en_diccionario).
    process_name_map: {ruta_archivo: nombre_proceso} — permite forzar el proceso de cada archivo.
    """
    pnm = process_name_map or {}
    process_data = {
        item["process"]: item
        for item in (aggregate_process(p, pnm.get(p, "")) for p in process_files)
    }
    gasto_data = {
        item["process"]: item
        for item in (aggregate_gasto(p, pnm.get(p, "")) for p in gasto_files)
    }
    central_data = {
        item["process"]: item
        for item in (aggregate_central(p, pnm.get(p, "")) for p in (central_files or []))
    }
    asiento_data = {
        item["process"]: item
        for item in (aggregate_asiento_reference(p, pnm.get(p, "")) for p in (asiento_files or []))
    }
    budget_map = parse_existing_budget_map(central_files or [])
    pdf_budget_map = extract_pdf_budget_map(cuentas_pdf)

    debit_detail = [row for item in gasto_data.values() for row in item["detail"]]
    discount_rows = [row for item in process_data.values() for row in item["discount_credit_rows"]]
    employer_rows = [row for item in process_data.values() for row in item["employer_credit_rows"]]

    mapping_excel_path = Path(mapeo_excel)
    mapping_path = Path(mapeo)

    master_mapping = load_mapping_workbook(mapping_excel_path) or load_mapping_csv(mapping_path)
    candidates = collect_mapping_candidates(
        debit_detail, discount_rows, employer_rows, budget_map, pdf_budget_map
    )
    mapping_rows, new_count = merge_into_master(master_mapping, candidates)

    if new_count > 0 or not mapping_excel_path.exists():
        write_mapping_csv(mapping_path, mapping_rows)
        write_mapping_workbook(mapping_excel_path, mapping_rows)

    pending_count = sum(1 for r in mapping_rows if not r["cuenta_contable"])

    dic_new_count = 0
    fuente_corregida_rows: List[List[object]] | None = None
    if diccionario_path:
        dic_primary, dic_fallback = load_diccionario(diccionario_path)
        # Detectar haberes del InformeGasto que no están en el diccionario y agregarlos
        codes_in_gasto: List[Dict[str, str]] = []
        seen_codes: set[str] = set()
        for gd in gasto_data.values():
            for row in gd.get("detail_por_centro", []):
                _, _cc, _proj, _tc, codigo, descripcion, _monto = row
                if codigo not in seen_codes and codigo not in dic_fallback:
                    codes_in_gasto.append({"codigo": codigo, "descripcion": descripcion})
                    seen_codes.add(codigo)
        if codes_in_gasto:
            dic_new_count = update_diccionario(diccionario_path, codes_in_gasto)
            if dic_new_count > 0:
                dic_primary, dic_fallback = load_diccionario(diccionario_path)
        if dic_primary or dic_fallback:
            jornada_map = build_jornada_map(process_files)
            fuente_corregida_rows = recalculate_fuentes(
                gasto_data, jornada_map, dic_primary, dic_fallback
            )

    build_workbook(
        Path(salida),
        process_data,
        gasto_data,
        central_data,
        mapping_rows,
        asiento_data=asiento_data,
        title_text=title_text,
        fuente_corregida_rows=fuente_corregida_rows,
    )

    return Path(salida).resolve(), mapping_path.resolve(), mapping_excel_path.resolve(), new_count, pending_count, dic_new_count


def main():
    parser = argparse.ArgumentParser(description="Genera una centralizacion cuadrada de remuneraciones.")
    parser.add_argument("--salida", default="salidas/centralizacion_educacion_202602.xlsx")
    parser.add_argument("--mapeo", default="salidas/mapeo_cuentas.csv")
    parser.add_argument("--mapeo-excel", default="salidas/mapeo_cuentas_202602.xlsx")
    parser.add_argument("--maestros", nargs="*")
    parser.add_argument("--gastos", nargs="*")
    parser.add_argument("--centrales", nargs="*")
    parser.add_argument("--asientos", nargs="*")
    parser.add_argument("--cuentas-pdf", default="")
    parser.add_argument("--titulo", default="CENTRALIZACION REMUNERACIONES")
    args = parser.parse_args()

    if args.maestros and args.gastos:
        process_files = args.maestros
        gasto_files = args.gastos
        central_files = args.centrales or []
        cuentas_pdf = args.cuentas_pdf
    else:
        paths = discover_paths()
        process_files = paths.process_files
        gasto_files = paths.gasto_files
        central_files = paths.central_files
        cuentas_pdf = paths.cuentas_pdf

    salida_path, mapping_csv, mapping_xlsx = run_pipeline(
        process_files=process_files,
        gasto_files=gasto_files,
        central_files=central_files,
        cuentas_pdf=cuentas_pdf,
        salida=args.salida,
        mapeo=args.mapeo,
        mapeo_excel=args.mapeo_excel,
        title_text=args.titulo,
        asiento_files=args.asientos or [],
    )

    print(f"Workbook generado: {salida_path}")
    print(f"CSV de mapeo generado: {mapping_csv}")
    print(f"Excel de mapeo generado: {mapping_xlsx}")


if __name__ == "__main__":
    main()
