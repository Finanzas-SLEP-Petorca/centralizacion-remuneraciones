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


def aggregate_process(path: str):
    process = process_name_from_file(path)
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


def aggregate_gasto(path: str):
    process = process_name_from_file(path)
    sheet = open_sheet(path, sheet_index=0)
    headers, rows = rows_from_sheet(sheet)
    code_idx = find_header_index(headers, "Código", "Codigo")
    description_idx = find_header_index(headers, "Descripción", "Descripcion")
    project_idx = find_header_index(headers, "Proyecto")
    tipo_cargo_idx = find_header_index(headers, "Tipo de Cargo")
    monto_idx = find_header_index(headers, "Monto Debe")

    grouped: Dict[Tuple[str, str, str, str], float] = defaultdict(float)
    raw_total = 0.0
    for row in rows:
        raw_code = row[code_idx]
        code = str(int(raw_code)) if isinstance(raw_code, float) else normalize_text(raw_code)
        description = normalize_text(row[description_idx])
        project = normalize_text(row[project_idx])
        tipo_cargo = normalize_text(row[tipo_cargo_idx]).upper()
        if not code or not description or description.upper() == "TOTALES:":
            continue
        amount = float(row[monto_idx] or 0)
        raw_total += amount
        grouped[(project, tipo_cargo, code, description)] += amount

    detail = [
        (process, project, tipo_cargo, code, description, round(amount, 2))
        for (project, tipo_cargo, code, description), amount in sorted(grouped.items())
    ]
    return {
        "process": process,
        "detail": detail,
        "total_gasto_original": round(raw_total, 2),
        "total_gasto_real": round(raw_total, 2),
    }


def aggregate_asiento_reference(path: str):
    process = process_name_from_file(path)
    book = safe_open_workbook(path)
    totals = {"Haberes": 0.0, "Descuentos": 0.0}
    for sheet_name in ("Haberes", "Descuentos"):
        if sheet_name not in book.sheet_names():
            continue
        sh = book.sheet_by_name(sheet_name)
        if sh.nrows > 0 and sh.ncols > 1 and isinstance(sh.cell_value(0, 1), (int, float)):
            totals[sheet_name] = float(sh.cell_value(0, 1))
    return {"process": process, "asiento_haberes": totals["Haberes"], "asiento_descuentos": totals["Descuentos"]}


def aggregate_central(path: str):
    process = process_name_from_file(path)
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


def build_mapping_rows(
    debit_detail: List[Tuple[str, str, str, str, str, float]],
    discount_rows: List[Tuple[str, str, str, float]],
    employer_rows: List[Tuple[str, str, str, float]],
    existing_budget_map: Dict[Tuple[str, str], str],
    pdf_budget_map: Dict[str, Dict[str, str]],
    current_mapping: Dict[Tuple[str, str, str, str], Dict[str, str]],
) -> List[Dict[str, str]]:
    rows: Dict[Tuple[str, str, str, str], Dict[str, str]] = {}

    def upsert(lado: str, origen: str, tipo_cargo: str, codigo: str, descripcion: str, sugerida: str = ""):
        key = (lado, origen, tipo_cargo, codigo)
        current = current_mapping.get(key, {})
        rows[key] = {
            "lado": lado,
            "origen": origen,
            "tipo_cargo": tipo_cargo,
            "codigo": codigo,
            "descripcion": descripcion,
            "cuenta_contable": current.get("cuenta_contable", ""),
            "cuenta_presupuestaria_sugerida": current.get("cuenta_presupuestaria_sugerida") or sugerida,
            "glosa_sugerida": current.get("glosa_sugerida", descripcion),
            "observacion": current.get("observacion", ""),
        }

    for _, project, tipo_cargo, codigo, descripcion, _ in debit_detail:
        suggested = pdf_budget_map.get(codigo, {}).get(tipo_cargo, "")
        if not suggested:
            suggested = existing_budget_map.get((infer_section(tipo_cargo), descripcion.upper()), "")
        upsert("DEBE", "GASTO_FINANCIAMIENTO", tipo_cargo, codigo, descripcion, suggested)

    for _, codigo, descripcion, _ in discount_rows:
        upsert("HABER", "DESCUENTO", "", codigo, descripcion)

    for _, codigo, descripcion, _ in employer_rows:
        upsert("HABER", "APORTE_PATRONAL", "", codigo, descripcion)

    upsert("HABER", "LIQUIDO", "", LIQUIDO_CODE, "LIQUIDO A PAGO")
    return [rows[key] for key in sorted(rows)]


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
    guia["A1"] = "1. Completa o corrige la columna cuenta_contable en la hoja CompletarAqui."
    guia["A2"] = "2. Si aparece un haber o descuento nuevo, vuelve a ejecutar el script: se agregará como COMPLETAR."
    guia["A3"] = "3. Si prefieres trabajar en CSV, este Excel se regenera desde mapeo_cuentas.csv."
    guia["A4"] = "4. La cuenta_presupuestaria_sugerida fue precargada desde Cuentas contables.pdf cuando fue posible."
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

        employer_grouped: Dict[Tuple[str, str], float] = defaultdict(float)
        for process_name, codigo, descripcion, monto in p["employer_credit_rows"]:
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

        liquid_amount = round(sum(row[3] for row in p["liquid_rows"]), 2)
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
):
    process_data = {item["process"]: item for item in (aggregate_process(path) for path in process_files)}
    gasto_data = {item["process"]: item for item in (aggregate_gasto(path) for path in gasto_files)}
    central_data = {
        item["process"]: item for item in (aggregate_central(path) for path in (central_files or []))
    }
    asiento_data = {
        item["process"]: item for item in (aggregate_asiento_reference(path) for path in (asiento_files or []))
    }
    budget_map = parse_existing_budget_map(central_files or [])
    pdf_budget_map = extract_pdf_budget_map(cuentas_pdf)

    debit_detail = [row for item in gasto_data.values() for row in item["detail"]]
    discount_rows = [row for item in process_data.values() for row in item["discount_credit_rows"]]
    employer_rows = [row for item in process_data.values() for row in item["employer_credit_rows"]]

    mapping_path = Path(mapeo)
    mapping_excel_path = Path(mapeo_excel)
    current_mapping = load_mapping_workbook(mapping_excel_path) or load_mapping_csv(mapping_path)
    mapping_rows = build_mapping_rows(
        debit_detail,
        discount_rows,
        employer_rows,
        budget_map,
        pdf_budget_map,
        current_mapping,
    )
    write_mapping_csv(mapping_path, mapping_rows)
    write_mapping_workbook(mapping_excel_path, mapping_rows)
    build_workbook(
        Path(salida),
        process_data,
        gasto_data,
        central_data,
        mapping_rows,
        asiento_data=asiento_data,
        title_text=title_text,
    )

    return Path(salida).resolve(), mapping_path.resolve(), mapping_excel_path.resolve()


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
